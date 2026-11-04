#!/usr/bin/env python3
"""
import_transcripts.py

Fills the transcript columns without calling YouTube, by importing transcripts
you copy out of YouTube's own "Show transcript" panel.

YouTube rate limits caption downloads per IP, which can block the automated
path for hours. The transcript panel on the watch page is unaffected, so this
is the reliable fallback.

How to use:
  1. python import_transcripts.py --list
     prints the videos still missing a transcript, with their URLs.
  2. For each one: open the video, "..." under the title -> Show transcript,
     then select the panel and copy it.
  3. Save each as transcripts_in/<video id>.txt
     (the id is the v= part of the URL, e.g. ccjn4tXifFA.txt)
  4. python import_transcripts.py
  5. python video_breakdown.py --no-search   (reads the cache, no caption calls)
     python apply_frame_labels.py            (restores the visual columns)

Accepted layouts, mixed freely:
  0:00 some words        timestamp and text on one line
  0:00                   timestamp on its own line, text on the next
  some words
  1:02:03 some words     hours are fine

  0:022 secondssome words timestamp fused with its screen-reader label, as the
                         panel sometimes copies; the label is stripped
  Chapter title          a line that does not follow a bare timestamp and
                         is followed by a timestamp is a chapter heading: it is kept out
                         of the word count and saved as <id>.chapters.json

Other lines without a timestamp are appended to the previous entry, so wrapped
text survives. Segment durations come from the next timestamp.
"""

import argparse
import json
import os
import re
import sys

import pandas as pd

TS = re.compile(r"^\s*(?:(\d+):)?(\d{1,2}):(\d{2})\s*(.*)$")
# the screen-reader label YouTube sometimes glues on: "2 seconds", "5 minutes, 58 seconds", "1 hour, 1 minute"
LABEL = re.compile(r"^(?:(\d+) hours?,? ?)?(?:(\d+) minutes?,? ?)?(?:(\d+) seconds?)?")
DEFAULT_TAIL = 5.0


def strip_label(rest, start):
    """Drop a leading duration label, but only when it spells out this exact timestamp."""
    m = LABEL.match(rest)
    if not m or not m.group(0):
        return rest
    h, mnt, sec = (int(g or 0) for g in m.groups())
    return rest[m.end():] if h * 3600 + mnt * 60 + sec == start else rest


def parse(text):
    """YouTube transcript panel text -> ([[text, start_sec, duration], ...], [[title, start_sec], ...])"""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    entries, chapters = [], []
    for i, line in enumerate(lines):
        m = TS.match(line)
        if m:
            h, mnt, sec, rest = m.groups()
            start = int(h or 0) * 3600 + int(mnt) * 60 + int(sec)
            entries.append([strip_label(rest.strip(), start).strip(), float(start)])
            continue
        # text belongs to a bare timestamp line above it; after a line that already has text it is a heading
        prev = TS.match(lines[i - 1]) if i > 0 else None
        after_ts = prev is not None and not prev.group(4).strip()
        before_ts = i + 1 < len(lines) and TS.match(lines[i + 1])
        if not after_ts and (before_ts or not entries):
            nxt = TS.match(lines[i + 1]) if before_ts else None
            start = int(nxt.group(1) or 0) * 3600 + int(nxt.group(2)) * 60 + int(nxt.group(3)) if nxt else 0
            chapters.append([line, float(start)])
        elif entries:
            entries[-1][0] = (entries[-1][0] + " " + line).strip()
    entries = [e for e in entries if e[0]]
    segs = []
    for i, (txt, start) in enumerate(entries):
        end = entries[i + 1][1] if i + 1 < len(entries) else start + DEFAULT_TAIL
        segs.append([txt, start, round(max(end - start, 0.1), 2)])
    return segs, chapters


def main():
    ap = argparse.ArgumentParser(description="Import pasted YouTube transcripts into the cache")
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--dir", default="transcripts_in", help="folder of <video id>.txt files")
    ap.add_argument("--list", action="store_true", help="show which videos still need a transcript")
    args = ap.parse_args()

    with open(args.config, encoding="utf-8") as fh:
        cfg = json.load(fh)
    cache_dir = os.path.join(cfg["output_dir"], "transcripts")
    csv = os.path.join(cfg["output_dir"], "csv", "Video_Breakdown.csv")
    if not os.path.exists(csv):
        sys.exit(f"{csv} not found. Run video_breakdown.py first.")
    df = pd.read_csv(csv)
    df["Video ID"] = df["Video URL"].str.extract(r"v=([\w-]+)")

    if args.list:
        missing = [r for _, r in df.iterrows()
                   if not os.path.exists(os.path.join(cache_dir, f"{r['Video ID']}.json"))]
        if not missing:
            print("Every video already has a cached transcript.")
            return
        print(f"{len(missing)} video(s) need a transcript. Save each as {args.dir}/<id>.txt:\n")
        for r in missing:
            print(f"  {r['Video ID']}.txt   {r['Channel']}")
            print(f"      {r['Video URL']}")
        return

    if not os.path.isdir(args.dir):
        os.makedirs(args.dir, exist_ok=True)
        sys.exit(f"Created {args.dir}/. Put <video id>.txt files there, then rerun. "
                 f"Use --list to see which ids are needed.")

    os.makedirs(cache_dir, exist_ok=True)
    known = set(df["Video ID"])
    imported = 0
    for name in sorted(os.listdir(args.dir)):
        if not name.endswith(".txt"):
            continue
        vid = name[:-4]
        with open(os.path.join(args.dir, name), encoding="utf-8") as fh:
            segs, chapters = parse(fh.read())
        if not segs:
            print(f"{name}: no timestamped lines found, skipped")
            continue
        if vid not in known:
            print(f"{name}: {vid} is not in the breakdown, importing anyway")
        with open(os.path.join(cache_dir, f"{vid}.json"), "w", encoding="utf-8") as fh:
            json.dump(segs, fh)
        chap_path = os.path.join(cache_dir, f"{vid}.chapters.json")
        if chapters:
            with open(chap_path, "w", encoding="utf-8") as fh:
                json.dump(chapters, fh)
        elif os.path.exists(chap_path):
            os.remove(chap_path)
        words = sum(len(s[0].split()) for s in segs)
        span = segs[-1][1] + segs[-1][2]
        print(f"{vid}: {len(segs)} segments, {words} words, {len(chapters)} chapters, "
              f"spans {int(span // 60)}:{int(span % 60):02d}")
        imported += 1

    still = [r["Video ID"] for _, r in df.iterrows()
             if not os.path.exists(os.path.join(cache_dir, f"{r['Video ID']}.json"))]
    print(f"\nImported {imported} transcript(s) into {cache_dir}")
    if still:
        print(f"Still missing {len(still)}: {', '.join(still)}")
    else:
        print("All videos covered. Next: python video_breakdown.py --no-search"
              "  then  python apply_frame_labels.py")


if __name__ == "__main__":
    main()
