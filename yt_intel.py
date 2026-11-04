#!/usr/bin/env python3
"""
yt_intel.py

Competitor and niche intelligence for the asset-economics YouTube channel.

What it does in one run:
  1. Channel stats for every competitor in config.json (subs, views, videos).
  2. Recent uploads per channel with views, length, format, title pattern and
     outlier multiple versus the channel's long-form median.
  3. Daily snapshots in SQLite, so the second run onward gives real
     "views per day" for channels and for every sampled video (evergreen test).
  4. Top comments on each channel's best videos, tagged by theme, with
     requested topics extracted (what the audience wants next).
  5. Optional transcript hook metrics (first 15 seconds, words per minute,
     sponsor timing, number density).
  6. Saturation scan: search queries per asset class, with views, publisher
     and age for every matching video, plus a publisher summary that shows
     template farms.
  7. Exports a tidy Excel workbook and CSVs for Power BI.

Usage:
  set the API key first
    macOS/Linux:  export YT_API_KEY="your-key"
    Windows:      set YT_API_KEY=your-key
  python yt_intel.py                   full run
  python yt_intel.py --estimate        print estimated quota cost, no API calls
  python yt_intel.py --snapshot-only   cheap daily run: channels and videos only
  python yt_intel.py --skip-comments --skip-transcripts --skip-search
"""

import argparse
import json
import math
import os
import random
import re
import sqlite3
import sys
import time
import traceback
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from statistics import median

import pandas as pd
import requests
from openpyxl.styles import Alignment, Font, PatternFill

API = "https://www.googleapis.com/youtube/v3"
COST = {"channels": 1, "playlistItems": 1, "videos": 1, "commentThreads": 1, "search": 100}


# --------------------------------------------------------------------------- API client
class QuotaExceeded(Exception):
    pass


class YT:
    """Thin YouTube Data API v3 client with a per-run quota budget."""

    ATTEMPTS = 6
    RETRY_STATUS = (408, 425, 429, 500, 502, 503, 504)

    def __init__(self, key, budget):
        self.key = key
        self.budget = budget
        self.used = 0
        self.session = requests.Session()
        # a fresh pool per host, no urllib3-level retries: retries are handled below
        # so that quota and permission errors are never retried blindly
        adapter = requests.adapters.HTTPAdapter(pool_connections=4, pool_maxsize=8, max_retries=0)
        self.session.mount("https://", adapter)
        self.session.headers.update({"Accept-Encoding": "gzip"})

    def get(self, endpoint, **params):
        cost = COST[endpoint]
        if self.used + cost > self.budget:
            raise QuotaExceeded(f"Run budget of {self.budget} units reached")
        params["key"] = self.key
        last = ""
        for attempt in range(self.ATTEMPTS):
            if attempt:
                time.sleep(min(2 ** attempt + random.uniform(0, 1), 30))
            try:
                # (connect timeout, read timeout)
                r = self.session.get(f"{API}/{endpoint}", params=params, timeout=(10, 60))
            except requests.exceptions.RequestException as e:
                # dropped connection, DNS blip, read timeout, SSL reset: retry on a new connection
                last = f"{type(e).__name__}: {e}"
                self.session.close()
                continue
            if r.status_code == 200:
                try:
                    payload = r.json()
                except ValueError as e:
                    last = f"bad JSON: {e}"
                    continue
                self.used += cost
                return payload
            reason = ""
            try:
                reason = r.json()["error"]["errors"][0].get("reason", "")
            except Exception:
                pass
            if r.status_code == 403 and reason in ("quotaExceeded", "dailyLimitExceeded", "rateLimitExceeded"):
                raise QuotaExceeded(f"Google API quota error: {reason}")
            if r.status_code == 403 and reason in ("commentsDisabled", "videoNotFound", "forbidden", "insufficientPermissions"):
                self.used += cost
                return {"items": [], "commentsDisabled": reason == "commentsDisabled", "skipped": reason}
            if r.status_code == 404:
                self.used += cost
                return {"items": []}
            if r.status_code in self.RETRY_STATUS:
                last = f"HTTP {r.status_code}"
                continue
            raise RuntimeError(f"{endpoint} HTTP {r.status_code} ({reason or 'no reason'}): {r.text[:300]}")
        raise RuntimeError(f"{endpoint} failed after {self.ATTEMPTS} attempts, last error: {last}")


# --------------------------------------------------------------------------- helpers
def parse_duration(iso):
    """ISO 8601 duration (PT1H2M3S, P1DT2H) to seconds."""
    m = re.match(r"^P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$", iso or "")
    if not m:
        return None
    d, h, mi, s = (int(x) if x else 0 for x in m.groups())
    return d * 86400 + h * 3600 + mi * 60 + s


def parse_dt(s):
    return datetime.fromisoformat(s.replace("Z", "+00:00"))


def to_int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def safe_div(a, b):
    if a is None or b in (None, 0):
        return None
    return a / b


TITLE_PATTERNS = [
    ("ELI5", r"explained like you"),
    ("Economics of", r"\beconomics of\b"),
    ("Real cost", r"\b(real cost|true cost|cost of owning|cost to own)\b"),
    ("Tier list", r"\btier list\b"),
    ("Every level", r"\bevery level\b"),
    ("Superlative", r"\b(most|biggest|largest|richest|worst|best|fastest|deadliest|only)\b"),
    ("How it works", r"^(how|why)\b"),
    ("Hypothetical", r"^(could|what if|what happens|can you|would)\b"),
    ("Question", r"\?\s*$"),
    ("Number hook", r"(\$\s?\d|\d[\d,\.]*\s?(million|billion|trillion|k|m|bn)\b|\b\d{2,}\b)"),
    ("Rise or fall", r"\b(collapse|fall|downfall|rise and fall|bankrupt|lost everything|destroyed|died)\b"),
    ("Scandal", r"\b(scam|fraud|scandal|lie|lied|exposed)\b"),
]


def title_patterns(title):
    t = (title or "").lower().strip()
    tags = [name for name, rx in TITLE_PATTERNS if re.search(rx, t)]
    return "; ".join(tags) if tags else "Other"


THEME_PATTERNS = {
    "Question": r"\?",
    "AI or slop complaint": r"\b(ai|a\.i\.|chatgpt|gpt|ai voice|ai generated|slop|robotic|robot voice|text to speech|tts)\b",
    "Correction or accuracy": r"\b(wrong|incorrect|mistake|not true|inaccurate|misleading|fact check|you forgot|you missed)\b",
    "Topic request": r"\b(do a video|make a video|next video|please (do|make|cover)|you should (do|cover|make)|can you (do|cover|make)|part 2)\b",
    "Praise": r"\b(great video|love (this|these|your)|amazing|underrated|best channel|well done|so good|fantastic|masterpiece)\b",
    "Voice or narration": r"\b(voice|narrator|narration|accent)\b",
    "Visuals or editing": r"\b(animation|animations|graphics|visuals|editing|edited|maps?)\b",
    "Money or numbers": r"(\$\d|\bmillion|\bbillion|\btrillion|\bcost|\bprice|\bexpensive|\bcheap|\bafford)",
    "Personal experience": r"\b(i work(ed)? (on|at|in|for)|as a (pilot|captain|crew|broker|owner|stewardess|deckhand)|my (dad|father|uncle|boss|family) (owns|owned|worked))\b",
    "Sponsor reaction": r"\b(sponsor|sponsored|ad read|skip the ad)\b",
}

TOPIC_REQUEST_RX = re.compile(
    r"(?:do a video (?:on|about)|make a video (?:on|about)|next (?:video|one) (?:on|about|should be)|"
    r"(?:please|you should|can you) (?:do|cover|make) (?:a video (?:on|about) )?)\s*(.{3,70}?)(?:[.!?\n]|$)",
    re.IGNORECASE,
)

STOPWORDS = set(
    """a an the and or but if then so to of in on for with at by from as is are was were be been being it its this that
    these those i you he she they we me my your our their them his her not no yes do does did done have has had just very
    really can could would should will about into over than too also what which who whom how why when where all any some
    more most much many there here out up down off only same such don im ive youre dont doesnt didnt cant wont isnt arent
    thats theres whats gonna get got like one even still make made video videos channel people thing things lot""".split()
)

SPONSOR_RX = re.compile(r"\b(sponsor|sponsored|brought to you by|use code|promo code|link in the description)\b", re.I)
NUMBER_RX = re.compile(r"\$?\d[\d,\.]*|\b(hundred|thousand|million|billion|trillion)\b", re.I)


# --------------------------------------------------------------------------- stages
def fetch_channels(yt, handles, log):
    rows = []
    for h in handles:
        h = h.strip()
        params = {"part": "snippet,statistics,contentDetails"}
        if h.startswith("UC") and len(h) == 24:
            params["id"] = h
        else:
            params["forHandle"] = h if h.startswith("@") else "@" + h
        data = yt.get("channels", **params)
        items = data.get("items", [])
        if not items:
            log.append(f"Channel not found: {h}")
            continue
        it = items[0]
        sn, st, cd = it["snippet"], it.get("statistics", {}), it.get("contentDetails", {})
        rows.append({
            "Channel": sn.get("title"),
            "Channel ID": it["id"],
            "Handle": h,
            "Country": sn.get("country"),
            "Created": (sn.get("publishedAt") or "")[:10],
            "Subscribers": None if st.get("hiddenSubscriberCount") else to_int(st.get("subscriberCount")),
            "Total Views": to_int(st.get("viewCount")),
            "Videos": to_int(st.get("videoCount")),
            "_uploads": cd.get("relatedPlaylists", {}).get("uploads"),
        })
    return rows


def fetch_videos(yt, ids, now, cfg):
    rows = []
    for i in range(0, len(ids), 50):
        batch = ids[i:i + 50]
        data = yt.get("videos", part="snippet,statistics,contentDetails", id=",".join(batch), maxResults=50)
        for it in data.get("items", []):
            sn, st, cd = it["snippet"], it.get("statistics", {}), it.get("contentDetails", {})
            if sn.get("liveBroadcastContent") == "upcoming":
                continue
            pub = parse_dt(sn["publishedAt"])
            age = max((now - pub).total_seconds() / 86400, 1 / 24)
            dur = parse_duration(cd.get("duration"))
            views = to_int(st.get("viewCount")) or 0
            likes = to_int(st.get("likeCount"))
            comments = to_int(st.get("commentCount"))
            is_short = dur is not None and dur <= cfg["shorts_max_seconds"]
            rows.append({
                "Channel": sn.get("channelTitle"),
                "Channel ID": sn.get("channelId"),
                "Video ID": it["id"],
                "Title": sn.get("title"),
                "Published": pub.date().isoformat(),
                "Age Days": round(age, 1),
                "Duration Sec": dur,
                "Length Min": round(dur / 60, 2) if dur is not None else None,
                "Format": "Short (heuristic)" if is_short else "Long-form",
                "Views": views,
                "Likes": likes,
                "Comments": comments,
                "Views per Day": round(views / age, 1),
                "Like Rate": safe_div(likes, views),
                "Comment Rate": safe_div(comments, views),
                "Title Pattern": title_patterns(sn.get("title")),
                "URL": f"https://www.youtube.com/watch?v={it['id']}",
            })
    return rows


def fetch_channel_uploads(yt, ch, cfg, now):
    if not ch.get("_uploads"):
        return []
    ids, token = [], None
    while len(ids) < cfg["max_videos_per_channel"]:
        data = yt.get("playlistItems", part="contentDetails", playlistId=ch["_uploads"], maxResults=50, pageToken=token)
        ids += [it["contentDetails"]["videoId"] for it in data.get("items", [])]
        token = data.get("nextPageToken")
        if not token:
            break
    rows = fetch_videos(yt, ids[:cfg["max_videos_per_channel"]], now, cfg)
    for r in rows:
        r["Channel"], r["Channel ID"] = ch["Channel"], ch["Channel ID"]
    return rows


def enrich_channels_and_videos(ch_rows, vid_rows, now):
    by_ch = {}
    for v in vid_rows:
        by_ch.setdefault(v["Channel ID"], []).append(v)
    cutoff = (now - timedelta(days=90)).date().isoformat()
    for ch in ch_rows:
        vids = by_ch.get(ch["Channel ID"], [])
        longs = [v for v in vids if v["Format"] == "Long-form"]
        med = median([v["Views"] for v in longs]) if longs else None
        for v in vids:
            v["Multiple of Channel Median"] = round(v["Views"] / med, 2) if med and v["Format"] == "Long-form" else None
        ch["Lifetime Avg Views per Video"] = safe_div(ch["Total Views"], ch["Videos"])
        ch["Views per Subscriber"] = safe_div(ch["Total Views"], ch["Subscribers"])
        s = safe_div(ch["Subscribers"], ch["Total Views"])
        ch["Subs per 1K Views"] = s * 1000 if s is not None else None
        ch["Sampled Videos"] = len(vids)
        ch["Sampled Long-form"] = len(longs)
        ch["Median Long-form Views in Sample"] = med
        lens = [v["Length Min"] for v in longs if v["Length Min"] is not None]
        ch["Median Length Min"] = round(median(lens), 1) if lens else None
        recent = [v for v in vids if v["Published"] >= cutoff]
        ch["Uploads per Week 90d"] = round(len(recent) / (90 / 7), 2)
        ch["Shorts Share of Sample"] = safe_div(len(vids) - len(longs), len(vids))


# --------------------------------------------------------------------------- snapshots
def db_connect(path):
    con = sqlite3.connect(path)
    con.execute("""CREATE TABLE IF NOT EXISTS channel_snapshots(
        run_date TEXT, channel_id TEXT, channel TEXT, subscribers INTEGER, views INTEGER, videos INTEGER,
        PRIMARY KEY(run_date, channel_id))""")
    con.execute("""CREATE TABLE IF NOT EXISTS video_snapshots(
        run_date TEXT, video_id TEXT, channel_id TEXT, views INTEGER, likes INTEGER, comments INTEGER,
        PRIMARY KEY(run_date, video_id))""")
    con.commit()
    return con


def previous_views(con, table, key_col, run_date):
    q = f"""SELECT t.{key_col}, t.run_date, t.views FROM {table} t
            JOIN (SELECT {key_col}, MAX(run_date) md FROM {table} WHERE run_date < ? GROUP BY {key_col}) m
            ON t.{key_col} = m.{key_col} AND t.run_date = m.md"""
    return {k: (d, v) for k, d, v in con.execute(q, (run_date,))}


def apply_snapshots(con, run_date, ch_rows, vid_rows):
    prev_ch = previous_views(con, "channel_snapshots", "channel_id", run_date)
    prev_v = previous_views(con, "video_snapshots", "video_id", run_date)
    today = date.fromisoformat(run_date)
    for ch in ch_rows:
        p = prev_ch.get(ch["Channel ID"])
        ch["Recent Views per Day"], ch["Snapshot Window Days"] = None, None
        if p and ch["Total Views"] is not None and p[1] is not None:
            days = (today - date.fromisoformat(p[0])).days
            if days > 0:
                ch["Recent Views per Day"] = round((ch["Total Views"] - p[1]) / days, 1)
                ch["Snapshot Window Days"] = days
    for v in vid_rows:
        p = prev_v.get(v["Video ID"])
        v["Recent Views per Day"], v["Recent vs Lifetime Daily Rate"] = None, None
        if p and p[1] is not None:
            days = (today - date.fromisoformat(p[0])).days
            if days > 0:
                rvpd = (v["Views"] - p[1]) / days
                v["Recent Views per Day"] = round(rvpd, 1)
                v["Recent vs Lifetime Daily Rate"] = round(safe_div(rvpd, v["Views per Day"]) or 0, 3)
    con.executemany("INSERT OR REPLACE INTO channel_snapshots VALUES (?,?,?,?,?,?)",
                    [(run_date, c["Channel ID"], c["Channel"], c["Subscribers"], c["Total Views"], c["Videos"]) for c in ch_rows])
    con.executemany("INSERT OR REPLACE INTO video_snapshots VALUES (?,?,?,?,?,?)",
                    [(run_date, v["Video ID"], v["Channel ID"], v["Views"], v["Likes"], v["Comments"]) for v in vid_rows])
    con.commit()


# --------------------------------------------------------------------------- comments
def fetch_comments(yt, video_id, max_n):
    out, token = [], None
    while len(out) < max_n:
        data = yt.get("commentThreads", part="snippet", videoId=video_id, maxResults=100,
                      order="relevance", textFormat="plainText", pageToken=token)
        if data.get("commentsDisabled"):
            break
        for it in data.get("items", []):
            top = it.get("snippet", {}).get("topLevelComment", {}).get("snippet")
            if not top:
                continue
            out.append({
                "Comment ID": it["id"],
                "Text": top.get("textDisplay", ""),
                "Comment Likes": to_int(top.get("likeCount")) or 0,
                "Replies": to_int(it["snippet"].get("totalReplyCount")) or 0,
                "Comment Date": (top.get("publishedAt") or "")[:10],
            })
        token = data.get("nextPageToken")
        if not token:
            break
    return out[:max_n]


def get_sentiment_fn():
    try:
        from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
        sia = SentimentIntensityAnalyzer()
        return lambda t: round(sia.polarity_scores(t)["compound"], 3)
    except ImportError:
        return None


def tag_comment(text):
    t = text.lower()
    themes = [name for name, rx in THEME_PATTERNS.items() if re.search(rx, t)]
    m = TOPIC_REQUEST_RX.search(text)
    topic = m.group(1).strip(" ,:;-") if m else None
    return themes, topic


def bigrams(texts, top_n):
    c = Counter()
    for t in texts:
        words = [w for w in re.findall(r"[a-z$][a-z0-9$']+", t.lower()) if w not in STOPWORDS and len(w) > 2]
        c.update(" ".join(p) for p in zip(words, words[1:]))
    return c.most_common(top_n)


# --------------------------------------------------------------------------- transcripts
def fetch_transcript(video_id, langs):
    try:
        from youtube_transcript_api import YouTubeTranscriptApi
    except ImportError:
        return None, "youtube-transcript-api not installed"
    try:
        api = YouTubeTranscriptApi()
        if hasattr(api, "fetch"):
            ft = api.fetch(video_id, languages=langs)
            segs = [(s.text, s.start, s.duration) for s in ft]
        else:
            raw = YouTubeTranscriptApi.get_transcript(video_id, languages=langs)
            segs = [(d["text"], d["start"], d["duration"]) for d in raw]
        return segs, "ok"
    except Exception as e:
        return None, type(e).__name__


def hook_metrics(video, segs, hook_seconds):
    words_all = sum(len(s[0].split()) for s in segs)
    hook = " ".join(s[0] for s in segs if s[1] < hook_seconds).replace("\n", " ")
    first60 = " ".join(s[0] for s in segs if s[1] < 60)
    sponsor_at = next((round(s[1]) for s in segs if SPONSOR_RX.search(s[0])), None)
    dur = video.get("Duration Sec") or (segs[-1][1] + segs[-1][2] if segs else None)
    return {
        "Channel": video["Channel"],
        "Video ID": video["Video ID"],
        "Title": video["Title"],
        "Views": video["Views"],
        "Hook Text": hook[:500],
        "Hook Words": len(hook.split()),
        "Hook Has Number": "Yes" if NUMBER_RX.search(hook) else "No",
        "Numbers in First 60 Sec": len(NUMBER_RX.findall(first60)),
        "Total Words": words_all,
        "Words per Min": round(words_all / (dur / 60), 1) if dur else None,
        "First Sponsor Mention Sec": sponsor_at,
        "Transcript Status": "ok",
    }


# --------------------------------------------------------------------------- saturation scan
def saturation_scan(yt, cfg, now, log):
    sc = cfg["saturation_scan"]
    after = (now - timedelta(days=sc["published_within_days"])).strftime("%Y-%m-%dT%H:%M:%SZ")
    found = {}
    for asset in sc["assets"]:
        for tpl in sc["query_templates"]:
            q = tpl.format(asset=asset["query"])
            try:
                data = yt.get("search", part="snippet", q=q, type="video", maxResults=sc["max_results_per_query"],
                              order="relevance", publishedAfter=after, relevanceLanguage=sc.get("relevance_language", "en"))
            except RuntimeError as e:
                log.append(f"Search query failed, skipped: {q!r} ({type(e).__name__}: {e})")
                continue
            for it in data.get("items", []):
                vid = it.get("id", {}).get("videoId")
                if vid and vid not in found:
                    found[vid] = (asset["name"], q, [t.lower() for t in asset["terms"]])
    rows = fetch_videos(yt, list(found), now, cfg)
    intent = re.compile(sc["intent_regex"], re.I)
    for r in rows:
        asset, q, terms = found[r["Video ID"]]
        title = (r["Title"] or "").lower()
        r["Asset Class"] = asset
        r["Query"] = q
        r["Title Match"] = "Yes" if any(t in title for t in terms) and intent.search(title) else "No"
        r["Under 150 Days"] = "Yes" if r["Age Days"] < 150 else "No"
    log.append(f"Saturation scan: {len(rows)} videos from {len(sc['assets']) * len(sc['query_templates'])} queries")
    return rows


def saturation_summaries(scan_rows):
    df = pd.DataFrame(scan_rows)
    if df.empty:
        return df, df
    m = df[df["Title Match"] == "Yes"]
    if m.empty:
        return pd.DataFrame(), pd.DataFrame()
    summ = m.groupby("Asset Class").agg(
        **{"Videos Found": ("Video ID", "count"),
           "Distinct Channels": ("Channel ID", "nunique"),
           "Under 150 Days": ("Under 150 Days", lambda s: (s == "Yes").sum()),
           "Median Views": ("Views", "median"),
           "Mean Views": ("Views", "mean"),
           "Videos Over 100K Views": ("Views", lambda s: (s >= 100000).sum()),
           "Max Views": ("Views", "max")}).reset_index()
    summ["Share Under 150 Days"] = summ["Under 150 Days"] / summ["Videos Found"]
    summ["Videos per Channel"] = summ["Videos Found"] / summ["Distinct Channels"]
    pubs = m.groupby(["Channel", "Channel ID"]).agg(
        **{"Videos in Scan": ("Video ID", "count"),
           "Asset Classes Covered": ("Asset Class", "nunique"),
           "Total Views in Scan": ("Views", "sum"),
           "Median Views": ("Views", "median"),
           "Newest Upload": ("Published", "max")}).reset_index()
    pubs["Likely Template Publisher"] = pubs["Asset Classes Covered"].apply(lambda n: "Yes" if n >= 3 else "No")
    pubs = pubs.sort_values(["Videos in Scan", "Total Views in Scan"], ascending=False)
    return summ, pubs


# --------------------------------------------------------------------------- quota estimate
def estimate(cfg, args):
    n = len(cfg["channels"])
    pages = math.ceil(cfg["max_videos_per_channel"] / 50)
    base = n + n * pages * 2
    com = 0 if args.skip_comments else n * cfg["comments"]["top_videos_per_channel"] * math.ceil(cfg["comments"]["max_comments_per_video"] / 100)
    sc = cfg["saturation_scan"]
    q = len(sc["assets"]) * len(sc["query_templates"])
    srch = 0 if args.skip_search else q * 100 + math.ceil(q * sc["max_results_per_query"] / 50)
    total = base + com + srch
    return (f"Estimated quota units: channels and uploads {base}, comments {com}, saturation scan {srch}, "
            f"total {total} (run budget {cfg['quota_budget_per_run']}, default Google daily quota is usually 10,000)")


# --------------------------------------------------------------------------- export
VIDEO_COLS = ["Run Date", "Channel", "Channel ID", "Video ID", "Title", "Published", "Age Days", "Duration Sec", "Length Min",
              "Format", "Views", "Likes", "Comments", "Views per Day", "Like Rate", "Comment Rate", "Title Pattern",
              "Multiple of Channel Median", "Recent Views per Day", "Recent vs Lifetime Daily Rate", "URL"]
CHANNEL_COLS = ["Run Date", "Channel", "Channel ID", "Handle", "Country", "Created", "Subscribers", "Total Views", "Videos",
                "Lifetime Avg Views per Video", "Views per Subscriber", "Subs per 1K Views", "Sampled Videos", "Sampled Long-form",
                "Median Long-form Views in Sample", "Median Length Min", "Uploads per Week 90d", "Shorts Share of Sample",
                "Recent Views per Day", "Snapshot Window Days"]


def frame(rows, cols=None):
    df = pd.DataFrame(rows)
    if cols is not None:
        for c in cols:
            if c not in df.columns:
                df[c] = None
        df = df[cols]
    return df


def write_outputs(sheets, out_dir, run_date):
    os.makedirs(os.path.join(out_dir, "csv"), exist_ok=True)
    dated = os.path.join(out_dir, f"yt_intel_{run_date}.xlsx")
    latest = os.path.join(out_dir, "latest.xlsx")
    hdr_fill = PatternFill("solid", start_color="1F3A5F")
    for path in (dated, latest):
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as xw:
                for name, df in sheets.items():
                    if df is None or (df.empty and len(df.columns) == 0):
                        df = pd.DataFrame({"Note": ["No data this run"]})
                    df.to_excel(xw, sheet_name=name[:31], index=False)
                    ws = xw.book[name[:31]]
                    ws.freeze_panes = "A2"
                    for cell in ws[1]:
                        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
                        cell.fill = hdr_fill
                        cell.alignment = Alignment(wrap_text=True, vertical="center")
                    for col in ws.columns:
                        sample = [len(str(c.value)) for c in list(col)[:300] if c.value is not None]
                        ws.column_dimensions[col[0].column_letter].width = max(10, min(60, (max(sample) if sample else 10) + 2))
            print(f"Wrote {path}")
        except PermissionError:
            print(f"Could not write {path}; close it in Excel and rerun the export.")
    for name, df in sheets.items():
        if df is not None and not df.empty:
            df.to_csv(os.path.join(out_dir, "csv", f"{name}.csv"), index=False, encoding="utf-8-sig")


# --------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="YouTube competitor and niche intelligence")
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--estimate", action="store_true", help="print quota estimate and exit")
    ap.add_argument("--snapshot-only", action="store_true", help="channels and videos only (cheap daily run)")
    ap.add_argument("--skip-comments", action="store_true")
    ap.add_argument("--skip-transcripts", action="store_true")
    ap.add_argument("--skip-search", action="store_true")
    args = ap.parse_args()
    if args.snapshot_only:
        args.skip_comments = args.skip_transcripts = args.skip_search = True

    with open(args.config, encoding="utf-8") as fh:
        cfg = json.load(fh)
    if args.estimate:
        print(estimate(cfg, args))
        return
    key = os.environ.get("YT_API_KEY") or cfg.get("api_key")
    if not key:
        sys.exit("Set the YT_API_KEY environment variable (or api_key in config.json).")

    yt = YT(key, cfg["quota_budget_per_run"])
    now = datetime.now(timezone.utc)
    run_date = now.date().isoformat()
    con = db_connect(cfg["db_path"])
    log = [f"Run started {now.isoformat(timespec='seconds')}"]
    ch_rows, vid_rows, com_rows, hook_rows, scan_rows = [], [], [], [], []

    try:
        ch_rows = fetch_channels(yt, cfg["channels"], log)
        for ch in ch_rows:
            try:
                vid_rows += fetch_channel_uploads(yt, ch, cfg, now)
            except (RuntimeError, KeyError) as e:
                log.append(f"Uploads failed for {ch['Channel']}: {type(e).__name__}: {e}")
                print(log[-1])
            print(f"{ch['Channel']}: {sum(1 for v in vid_rows if v['Channel ID'] == ch['Channel ID'])} videos sampled")
        enrich_channels_and_videos(ch_rows, vid_rows, now)
        apply_snapshots(con, run_date, ch_rows, vid_rows)

        top_by_ch = {}
        for ch in ch_rows:
            longs = sorted([v for v in vid_rows if v["Channel ID"] == ch["Channel ID"] and v["Format"] == "Long-form"],
                           key=lambda v: v["Views"], reverse=True)
            top_by_ch[ch["Channel ID"]] = longs

        if not args.skip_comments:
            sent = get_sentiment_fn()
            cc = cfg["comments"]
            for ch in ch_rows:
                for v in top_by_ch[ch["Channel ID"]][:cc["top_videos_per_channel"]]:
                    try:
                        fetched = fetch_comments(yt, v["Video ID"], cc["max_comments_per_video"])
                    except (RuntimeError, KeyError) as e:
                        log.append(f"Comments failed for {v['Video ID']}: {type(e).__name__}: {e}")
                        continue
                    for c in fetched:
                        themes, topic = tag_comment(c["Text"])
                        c.update({"Channel": v["Channel"], "Video ID": v["Video ID"], "Video Title": v["Title"],
                                  "Themes": "; ".join(themes) if themes else "None",
                                  "Is Question": "Yes" if "Question" in themes else "No",
                                  "Requested Topic": topic,
                                  "Sentiment": sent(c["Text"]) if sent else None})
                        com_rows.append(c)
            log.append(f"Comments collected: {len(com_rows)}" + ("" if sent else " (install vaderSentiment for sentiment)"))

        tc = cfg["transcripts"]
        if not args.skip_transcripts and tc.get("enabled"):
            for ch in ch_rows:
                for v in top_by_ch[ch["Channel ID"]][:tc["top_videos_per_channel"]]:
                    segs, status = fetch_transcript(v["Video ID"], tc["languages"])
                    if segs:
                        hook_rows.append(hook_metrics(v, segs, tc["hook_seconds"]))
                    else:
                        hook_rows.append({"Channel": v["Channel"], "Video ID": v["Video ID"], "Title": v["Title"],
                                          "Views": v["Views"], "Transcript Status": status})
                    time.sleep(tc.get("delay_seconds", 2))
        if not args.skip_search:
            scan_rows = saturation_scan(yt, cfg, now, log)
    except QuotaExceeded as e:
        log.append(f"STOPPED EARLY: {e}. Exporting what was collected.")
        print(log[-1])
    except KeyboardInterrupt:
        log.append("STOPPED EARLY: interrupted by user. Exporting what was collected.")
        print(log[-1])
    except Exception as e:
        log.append(f"STOPPED EARLY: {type(e).__name__}: {e}. Exporting what was collected.")
        print(log[-1])
        traceback.print_exc()
    finally:
        log.append(f"Quota units used this run: {yt.used}")
        con.close()

    for r in ch_rows:
        r["Run Date"] = run_date
    for r in vid_rows:
        r["Run Date"] = run_date

    channels = frame(ch_rows, CHANNEL_COLS)
    videos = frame(vid_rows, VIDEO_COLS)

    top = videos[videos["Format"] == "Long-form"].copy() if not videos.empty else videos
    if not top.empty:
        top = top.sort_values(["Channel", "Multiple of Channel Median"], ascending=[True, False])
        top["Rank"] = top.groupby("Channel").cumcount() + 1
        top = top[top["Rank"] <= cfg["top_videos_sheet_per_channel"]]
        top["Outlier"] = top["Multiple of Channel Median"].apply(lambda m: "Yes" if m is not None and m >= cfg["outlier_multiple"] else "No")
        top = top[["Channel", "Rank", "Title", "Views", "Age Days", "Multiple of Channel Median", "Outlier", "Title Pattern",
                   "Length Min", "Views per Day", "Recent Views per Day", "URL"]]

    comments = frame(com_rows, ["Channel", "Video ID", "Video Title", "Comment ID", "Text", "Comment Likes", "Replies",
                                "Comment Date", "Themes", "Is Question", "Requested Topic", "Sentiment"])
    theme_rows, ngram_rows = [], []
    if not comments.empty:
        for chname, g in comments.groupby("Channel"):
            total = len(g)
            for theme in THEME_PATTERNS:
                mask = g["Themes"].str.contains(re.escape(theme), regex=True)
                theme_rows.append({"Channel": chname, "Theme": theme, "Comments": int(mask.sum()),
                                   "Share of Comments": mask.sum() / total if total else None,
                                   "Likes on Theme Comments": int(g.loc[mask, "Comment Likes"].sum())})
            for i, (bg, n) in enumerate(bigrams(g["Text"].tolist(), cfg["comments"]["top_bigrams"]), start=1):
                ngram_rows.append({"Channel": chname, "Rank": i, "Bigram": bg, "Count": n})
    requested = comments[comments["Requested Topic"].notna()][["Channel", "Requested Topic", "Comment Likes", "Video Title", "Text"]] \
        .sort_values("Comment Likes", ascending=False) if not comments.empty else pd.DataFrame()

    scan = frame(scan_rows, ["Asset Class", "Query", "Title Match", "Channel", "Channel ID", "Video ID", "Title", "Published",
                             "Age Days", "Under 150 Days", "Length Min", "Views", "Views per Day", "Title Pattern", "URL"])
    sat_summary, sat_pubs = saturation_summaries(scan_rows)

    sheets = {
        "Run_Log": pd.DataFrame({"Log": log}),
        "Channels": channels,
        "Videos": videos,
        "Top_Videos": top,
        "Comments": comments,
        "Comment_Themes": pd.DataFrame(theme_rows),
        "Requested_Topics": requested,
        "Comment_Bigrams": pd.DataFrame(ngram_rows),
        "Transcript_Hooks": pd.DataFrame(hook_rows),
        "Saturation_Scan": scan,
        "Saturation_Summary": sat_summary,
        "Saturation_Publishers": sat_pubs,
    }
    write_outputs(sheets, cfg["output_dir"], run_date)
    print("\n".join(log))


if __name__ == "__main__":
    main()
