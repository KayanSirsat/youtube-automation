#!/usr/bin/env python3
"""
video_breakdown.py

Builds the per-video production breakdown table for two videos per channel:
  1. the channel's most viewed long-form video in the latest yt_intel export
  2. the channel's most viewed video with "private jet" or "jet" in the title

Reads exports/csv/Videos.csv (run yt_intel.py first), so the selection itself
costs no quota. When a channel has no jet video in the sampled uploads, it
searches the channel's full catalogue (100 quota units per channel; skip with
--no-search or by not setting YT_API_KEY).

Filled automatically: URL, channel, title, length, views, age, views per day,
and from transcripts: word count, words per minute, first sponsor mention.
Filled by hand: seconds per visual category, cut count. Hook length, hook type and
sponsor start are guessed from the transcript; hand-checked values in
hook_labels.json replace the guesses.
The Classified Sec, Pct and Cuts per Min columns are Excel formulas, so they
calculate as soon as you type in the seconds.

Usage:
  python video_breakdown.py
  python video_breakdown.py --no-search
  python video_breakdown.py --no-transcripts
"""

import argparse
import json
import os
import re
import sys
import time
from datetime import date, datetime, timezone

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import yt_intel

HEADER = ["Video URL", "Channel", "Title", "Length Sec", "Hook Length Sec", "Hook Type", "Sec Data Graphics",
          "Sec Real Footage or Photos", "Sec Stock or Generated Imagery", "Sec Maps", "Sec Other", "Classified Sec",
          "Pct Data Graphics", "Pct Real", "Pct Stock or Generated", "Pct Maps", "Cut Count", "Cuts per Min",
          "Word Count", "Words per Min", "Sponsor Read At Sec", "Views", "Days Since Publish", "Views per Day"]

PRIVATE_JET_RX = re.compile(r"\bprivate[\s-]+jets?\b", re.I)
# "jet lag" is not about jets (Wendover's "Jet Lag: The Game")
JET_RX = re.compile(r"\bjets?\b(?!\s*lag)", re.I)


# sponsor reads rarely say "sponsor": they name the brand and give a vanity URL or code
SPONSOR_SIGNAL_RX = re.compile(
    r"\b(sponsor(ed|s|ship)?|brought to you by|promo code|use (the )?code|discount code|free trial|sign up (at|for|with)|"
    r"at no extra cost|first \d+ (people|users|customers)|link (is )?in the description|"
    r"nebula|curiosity ?stream|squarespace|brilliant(\.org)?|skillshare|nord ?vpn|surfshark|express ?vpn|audible|"
    r"ground news|incogni|hello ?fresh|aura|masterclass|blinkist|raycon|honey|delete ?me|shopify|wix|"
    r"betterhelp|private internet access|morning brew|factor|athletic greens|ag1|rocket money|opera gx)\b"
    r"|\b[a-z0-9-]+\s?(\.|dot\s)\s?(com|org|io|co)\s?(/|slash\s)\s?\w+",
    re.I,
)


def sponsor_start(segs, window=60, min_hits=2):
    """Start second of the first sponsor read: the first signal that has another signal within `window` sec."""
    hits = [s[1] for s in segs if SPONSOR_SIGNAL_RX.search(s[0].replace("\xa0", " "))]
    for i, t in enumerate(hits):
        if sum(1 for u in hits[i:] if u - t <= window) >= min_hits:
            return round(t)
    return None


def jet_rank(title):
    """2 = private jet in title, 1 = jet in title, 0 = no match."""
    t = title or ""
    if PRIVATE_JET_RX.search(t):
        return 2
    if JET_RX.search(t):
        return 1
    return 0


def pick_jet(df, exclude_id):
    """Most viewed video in df with the best keyword match, excluding exclude_id."""
    d = df[df["Video ID"] != exclude_id].copy()
    d["_rank"] = d["Title"].apply(jet_rank)
    d = d[d["_rank"] > 0]
    if d.empty:
        return None
    # long-form before Shorts, then private jet before jet, then views
    d["_long"] = (d["Format"] == "Long-form").astype(int)
    return d.sort_values(["_long", "_rank", "Views"], ascending=False).iloc[0]


def search_channel_for_jet(yt, channel_id, exclude_id, now, cfg):
    data = yt.get("search", part="snippet", channelId=channel_id, q="private jet|jet", type="video",
                  maxResults=50, order="viewCount")
    ids = [it["id"]["videoId"] for it in data.get("items", []) if it.get("id", {}).get("videoId")]
    if not ids:
        return None
    rows = pd.DataFrame(yt_intel.fetch_videos(yt, ids, now, cfg))
    if rows.empty:
        return None
    return pick_jet(rows, exclude_id)


BLOCKED_STATUSES = ("IpBlocked", "RequestBlocked", "TooManyRequests")


def fetch_subs_ytdlp(video_id, langs):
    """Captions through yt-dlp's player API. A separate code path from youtube-transcript-api,
    so when one is rate limited the other sometimes still answers. No video is downloaded."""
    try:
        from yt_dlp import YoutubeDL
    except ImportError:
        return None, "yt-dlp not installed"
    import urllib.request
    opts = {"skip_download": True, "writesubtitles": True, "writeautomaticsub": True,
            "quiet": True, "no_warnings": True, "subtitleslangs": [f"{l}.*" for l in langs]}
    try:
        with YoutubeDL(opts) as ydl:
            info = ydl.extract_info(f"https://www.youtube.com/watch?v={video_id}", download=False)
            tracks = {**(info.get("subtitles") or {}), **(info.get("automatic_captions") or {})}
            pick = next((tracks[k] for k in langs if k in tracks),
                        next((v for k, v in tracks.items() if k.split("-")[0] in langs), None))
            if not pick:
                return None, "no captions"
            url = next((f["url"] for f in pick if f.get("ext") == "json3"), pick[0]["url"])
            with urllib.request.urlopen(url, timeout=30) as r:
                data = json.loads(r.read().decode("utf-8"))
    except Exception as e:
        return None, type(e).__name__
    segs = []
    for ev in data.get("events", []):
        text = "".join(s.get("utf8", "") for s in ev.get("segs", [])).strip()
        if text:
            segs.append([text, (ev.get("tStartMs", 0)) / 1000, (ev.get("dDurationMs", 0)) / 1000])
    return (segs, "ok (yt-dlp)") if segs else (None, "empty captions")


def get_transcript(video_id, langs, cache_dir):
    """Transcript segments, cached on disk so reruns never refetch (YouTube rate limits repeat callers)."""
    path = os.path.join(cache_dir, f"{video_id}.json")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as fh:
            return json.load(fh), "ok (cached)"
    segs, status = yt_intel.fetch_transcript(video_id, langs)
    if not segs:
        segs, status = fetch_subs_ytdlp(video_id, langs)
    if segs:
        os.makedirs(cache_dir, exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(segs, fh)
    return segs, status


# closed list: percentages are only comparable across channels if the labels are fixed
HOOK_TYPES = [
    ("Contradiction", r"\b(but|however|yet|except|actually|in fact|turns out|despite|although)\b.{0,60}\b(not|isn'?t|wasn'?t|never|wrong|opposite|no longer)\b"
                      r"|\b(everything you (know|think)|common sense|you('| a)?d think|most people (think|believe|assume))\b"),
    ("Big number", r"(\$\s?\d|\b\d[\d,\.]*\s?(million|billion|trillion)\b|\b\d{1,3}(,\d{3}){2,}\b|\b\d+\s?(percent|%)\b)"),
    ("Question", r"^[^.?!]{0,160}\?|\b(what if|why (is|do|does|are)|how (is|do|does|did|can)|who (is|are|owns))\b"),
    ("Stakes", r"\b(collapse|crisis|disaster|catastroph|died|death|killed|war|destroy|bankrupt|lost everything|"
               r"never recover|too late|at risk|threat|danger|failing|failed)\b"),
    ("Personal story", r"^\s*(i|we|my|when i|last (year|week|month) i)\b"),
    ("Cold open", r"^\s*(it (all )?(happened|started|began)|on (january|february|march|april|may|june|july|august|"
                  r"september|october|november|december)|in (19|20)\d\d|at \d{1,2}(:\d\d)?\s?(am|pm)|picture this|imagine)\b"),
]


def classify_hook(text):
    """First matching label from the closed list, or Thesis when the hook just states its subject."""
    t = re.sub(r"\s+", " ", (text or "")).strip()
    if not t:
        return None
    for name, rx in HOOK_TYPES:
        if re.search(rx, t, re.I):
            return name
    return "Thesis"


def hook_length(segs, max_sec=90, min_sec=5, gap=1.2, chapters=None):
    """Approximate end of the cold open: the first chapter break, else the first real pause in the
    narration after min_sec.

    Without the video there is no title card to detect, so this is a proxy and worth a spot check.
    """
    if not segs:
        return None
    # a chapter break inside the opening is where the creator says the intro ends
    for _, start in chapters or []:
        if min_sec <= start <= max_sec:
            return round(start)
    for prev, nxt in zip(segs, segs[1:]):
        end = prev[1] + prev[2]
        if end >= min_sec and nxt[1] - end >= gap:
            return round(end)
        if nxt[1] > max_sec:
            break
    # imported transcripts run each segment up to the next, so they contain no
    # pauses at all; fall back to where the first sentence finishes
    for s in segs:
        if s[1] > max_sec:
            break
        end = s[1] + s[2]
        if end >= min_sec and re.search(r"""[.!?]["')\]]*\s*$""", s[0]):
            return round(end)
    return None


def transcript_stats(video_id, langs, cache_dir):
    segs, status = get_transcript(video_id, langs, cache_dir)
    if not segs:
        return {"status": status}
    status = "ok" if status.startswith("ok") else status
    # [Music], [clears throat] and the like are caption tags, not narration
    words = sum(len(re.sub(r"\[[^\]]*\]", " ", s[0]).split()) for s in segs)
    sponsor = sponsor_start(segs)
    hook = " ".join(s[0] for s in segs if s[1] < 15).replace("\n", " ")
    hook = re.sub(r"\s+", " ", hook).strip()
    chap_path = os.path.join(cache_dir, f"{video_id}.chapters.json")
    chapters = json.load(open(chap_path, encoding="utf-8")) if os.path.exists(chap_path) else None
    hlen = hook_length(segs, chapters=chapters)
    # classify on the spoken hook itself when we found its end, else on the first 15 seconds
    htext = " ".join(s[0] for s in segs if s[1] < hlen).replace("\n", " ") if hlen else hook
    return {"status": status, "words": words, "sponsor": sponsor, "hook": hook,
            "hook_len": hlen, "hook_type": classify_hook(htext)}


def write_workbook(rows, helper, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Breakdown"
    ws.append(HEADER)
    col = {h: get_column_letter(i + 1) for i, h in enumerate(HEADER)}

    for i, r in enumerate(rows, start=2):
        L, G, H, I, J, K = (col[c] + str(i) for c in ("Length Sec", "Sec Data Graphics", "Sec Real Footage or Photos",
                                                       "Sec Stock or Generated Imagery", "Sec Maps", "Sec Other"))
        cl, cc = col["Classified Sec"] + str(i), col["Cut Count"] + str(i)
        r["Classified Sec"] = f"=IF(COUNT({G}:{K})=0,\"\",SUM({G}:{K}))"
        # percentages are of the seconds you actually classified, so a partial or sampled coding still works
        r["Pct Data Graphics"] = f"=IFERROR({G}/{cl},\"\")"
        r["Pct Real"] = f"=IFERROR({H}/{cl},\"\")"
        r["Pct Stock or Generated"] = f"=IFERROR({I}/{cl},\"\")"
        r["Pct Maps"] = f"=IFERROR({J}/{cl},\"\")"
        r["Cuts per Min"] = f"=IF({cc}=\"\",\"\",IFERROR({cc}/({L}/60),\"\"))"
        ws.append([r.get(h) for h in HEADER])

    hdr_fill = PatternFill("solid", start_color="1F3A5F")
    manual_fill = PatternFill("solid", start_color="FFF4CC")
    manual = {"Sec Data Graphics", "Sec Real Footage or Photos",
              "Sec Stock or Generated Imagery", "Sec Maps", "Sec Other", "Cut Count"}
    auto_check = {"Hook Length Sec", "Hook Type"}  # derived from the transcript, worth a spot check
    check_fill = PatternFill("solid", start_color="E2EFDA")
    for cell in ws[1]:
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = hdr_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for h, letter in col.items():
        for row in ws.iter_rows(min_row=2, min_col=HEADER.index(h) + 1, max_col=HEADER.index(h) + 1):
            c = row[0]
            if h in manual:
                c.fill = manual_fill
            elif h in auto_check:
                c.fill = check_fill
            if h.startswith("Pct"):
                c.number_format = "0.0%"
            elif h in ("Views", "Length Sec", "Word Count", "Views per Day", "Days Since Publish"):
                c.number_format = "#,##0"
            elif h in ("Words per Min", "Cuts per Min"):
                c.number_format = "0.0"
        ws.column_dimensions[letter].width = 60 if h == "Title" else 45 if h == "Video URL" else 14
    ws.row_dimensions[1].height = 45
    ws.freeze_panes = "D2"

    hs = wb.create_sheet("Selection_and_Hooks")
    hcols = ["Video URL", "Channel", "Selected As", "Source", "Transcript", "Hook Text (first 15 sec)"]
    hs.append(hcols)
    for h in helper:
        hs.append([h.get(c) for c in hcols])
    for cell in hs[1]:
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = hdr_fill
    for letter, w in zip("ABCDEF", (45, 22, 30, 28, 14, 100)):
        hs.column_dimensions[letter].width = w
    for row in hs.iter_rows(min_row=2, min_col=6, max_col=6):
        row[0].alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description="Two-video-per-channel production breakdown table")
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--no-search", action="store_true", help="never spend quota searching for jet videos")
    ap.add_argument("--no-transcripts", action="store_true")
    args = ap.parse_args()

    with open(args.config, encoding="utf-8") as fh:
        cfg = json.load(fh)
    src = os.path.join(cfg["output_dir"], "csv", "Videos.csv")
    if not os.path.exists(src):
        sys.exit(f"{src} not found. Run yt_intel.py first.")
    videos = pd.read_csv(src)
    top_path = os.path.join(cfg["output_dir"], "csv", "Top_Videos.csv")
    top_urls = set(pd.read_csv(top_path)["URL"]) if os.path.exists(top_path) else set()

    key = os.environ.get("YT_API_KEY") or cfg.get("api_key")
    yt = yt_intel.YT(key, cfg["quota_budget_per_run"]) if key and not args.no_search else None
    now = datetime.now(timezone.utc)
    today = date.today()
    langs = cfg.get("transcripts", {}).get("languages", ["en"])

    picks = []  # (video row, selected as, source)
    for ch, g in videos.groupby("Channel", sort=True):
        longs = g[g["Format"] == "Long-form"]
        pool = longs if not longs.empty else g
        top = pool.sort_values("Views", ascending=False).iloc[0]
        picks.append((top, "Most viewed", "Top_Videos" if top["URL"] in top_urls else "Videos sample"))

        in_top = g[g["URL"].isin(top_urls)]
        jet, source = pick_jet(in_top, top["Video ID"]), "Top_Videos"
        if jet is None:
            jet, source = pick_jet(g, top["Video ID"]), "Videos sample"
        if jet is None and yt is not None:
            try:
                jet, source = search_channel_for_jet(yt, g["Channel ID"].iloc[0], top["Video ID"], now, cfg), "Channel search"
            except (yt_intel.QuotaExceeded, RuntimeError) as e:
                print(f"{ch}: search failed ({e})")
        if jet is None:
            print(f"{ch}: no video with 'private jet' or 'jet' in the title"
                  + ("" if yt else " in the sample (set YT_API_KEY to search the full channel)"))
            continue
        label = "Keyword: private jet" if jet_rank(jet["Title"]) == 2 else "Keyword: jet"
        picks.append((jet, label, source))

    cache_dir = os.path.join(cfg["output_dir"], "transcripts")
    labels = {}
    if os.path.exists("hook_labels.json"):
        with open("hook_labels.json", encoding="utf-8") as fh:
            labels = {k: v for k, v in json.load(fh).items() if not k.startswith("_")}
    blocked = False
    rows, helper = [], []
    for v, label, source in picks:
        published = date.fromisoformat(str(v["Published"])[:10])
        days = max((today - published).days, 1)
        length = int(v["Duration Sec"]) if pd.notna(v["Duration Sec"]) else None
        views = int(v["Views"])
        url = v["URL"]
        row = {"Video URL": url, "Channel": v["Channel"], "Title": v["Title"], "Length Sec": length,
               "Views": views, "Days Since Publish": days, "Views per Day": round(views / days)}
        h = {"Video URL": url, "Channel": v["Channel"], "Selected As": label, "Source": source}
        if not args.no_transcripts:
            cached = os.path.exists(os.path.join(cache_dir, f"{v['Video ID']}.json"))
            if blocked and not cached:
                ts = {"status": "skipped, IP blocked"}
            else:
                ts = transcript_stats(v["Video ID"], langs, cache_dir)
                if ts["status"] in BLOCKED_STATUSES:
                    blocked = True
                    print("YouTube is blocking transcript requests from this IP; skipping the rest. "
                          "Rerun in an hour or two: transcripts already fetched are cached and not refetched.")
                elif not cached:
                    time.sleep(cfg.get("transcripts", {}).get("delay_seconds", 2))
            h["Transcript"] = ts["status"]
            if ts["status"] == "ok":
                row["Word Count"] = ts["words"]
                row["Words per Min"] = round(ts["words"] / (length / 60), 1) if length else None
                row["Sponsor Read At Sec"] = ts["sponsor"]
                row["Hook Length Sec"] = ts["hook_len"]
                row["Hook Type"] = ts["hook_type"]
                for key, col in (("hook_len", "Hook Length Sec"), ("hook_type", "Hook Type"),
                                 ("sponsor", "Sponsor Read At Sec")):
                    if key in labels.get(v["Video ID"], {}):
                        row[col] = labels[v["Video ID"]][key]
                h["Hook Text (first 15 sec)"] = ts["hook"]
        print(f"{v['Channel']:<22} {label:<22} {views:>12,}  {v['Title']}")
        rows.append(row)
        helper.append(h)

    out = os.path.join(cfg["output_dir"], "video_breakdown.xlsx")
    try:
        write_workbook(rows, helper, out)
    except PermissionError:
        sys.exit(f"Could not write {out}; close it in Excel and rerun.")
    pd.DataFrame(rows, columns=HEADER).to_csv(os.path.join(cfg["output_dir"], "csv", "Video_Breakdown.csv"),
                                              index=False, encoding="utf-8-sig")
    print(f"\nWrote {out} ({len(rows)} videos)" + (f", quota used: {yt.used}" if yt else ""))
    # this rebuilds the table from scratch, so any visual columns already applied are gone until reapplied
    if os.path.exists("frame_labels.json"):
        print("frame_labels.json exists: run  python apply_frame_labels.py  to put the visual columns back.")


if __name__ == "__main__":
    main()
