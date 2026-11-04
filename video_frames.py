#!/usr/bin/env python3
"""
video_frames.py

Prepares the visual columns of the breakdown table that no API can answer:
cut count, and sampled frames to classify into the five Sec categories.

For each video in exports/csv/Video_Breakdown.csv it
  1. downloads a video-only copy at 360p or lower (small, no audio, no ffmpeg needed),
  2. counts hard cuts with PySceneDetect,
  3. samples a frame every N seconds and tiles them into labelled 5x5 contact
     sheets, so a reviewer classifies 25 timestamps per image instead of one.

Each sample stands for N seconds of runtime, so counting the samples in a
category and multiplying by N gives the Sec columns.

The interval adapts to the runtime, aiming for ~75 samples per video, so a
10 minute video and a 79 minute one cost about the same to review.

Downloaded video files are never executed or imported; only OpenCV reads them.

Usage:
  python video_frames.py                     all videos in the breakdown
  python video_frames.py --only ccjn4tXifFA  one video
  python video_frames.py --keep-video        do not delete the mp4 afterwards
"""

import argparse
import json
import math
import os
import re
import shutil
import sys

import cv2
import pandas as pd

TARGET_SAMPLES = 75
MIN_INTERVAL = 3.0
MAX_INTERVAL = 40.0
GRID_COLS, GRID_ROWS = 5, 5
CELL_W, CELL_H = 320, 180


def sample_interval(duration):
    return round(min(max(duration / TARGET_SAMPLES, MIN_INTERVAL), MAX_INTERVAL), 1)


def download(video_id, video_dir):
    """Smallest usable video-only stream. Video-only keeps it to one file, so ffmpeg is not needed."""
    from yt_dlp import YoutubeDL
    os.makedirs(video_dir, exist_ok=True)
    existing = [f for f in os.listdir(video_dir) if f.startswith(video_id + ".")]
    if existing:
        return os.path.join(video_dir, existing[0])
    opts = {
        # avc1 (H.264) only: OpenCV here cannot decode AV1, which YouTube now serves by default
        "format": "bv*[height<=360][vcodec^=avc1]/bv*[height<=480][vcodec^=avc1]/18/"
                  "b[height<=480][vcodec^=avc1]/worst[ext=mp4]",
        "outtmpl": os.path.join(video_dir, "%(id)s.%(ext)s"),
        "quiet": True, "no_warnings": True, "retries": 5, "fragment_retries": 5,
    }
    with YoutubeDL(opts) as ydl:
        info = ydl.extract_info(f"https://www.youtube.com/watch?v={video_id}", download=True)
    return ydl.prepare_filename(info)


def count_cuts(path):
    from scenedetect import ContentDetector, SceneManager, open_video
    video = open_video(path)
    sm = SceneManager()
    sm.add_detector(ContentDetector(threshold=27.0, min_scene_len=8))
    sm.detect_scenes(video, show_progress=False)
    scenes = sm.get_scene_list()
    # N scenes means N-1 boundaries between them
    return max(len(scenes) - 1, 0), [round(s[0].get_seconds(), 1) for s in scenes]


def label(img, text):
    cv2.rectangle(img, (0, CELL_H - 20), (74, CELL_H), (0, 0, 0), -1)
    cv2.putText(img, text, (4, CELL_H - 6), cv2.FONT_HERSHEY_SIMPLEX, 0.45, (255, 255, 255), 1, cv2.LINE_AA)
    return img


def contact_sheets(path, video_id, interval, out_dir):
    """Tile sampled frames into labelled grids. Returns (grid files, sample timestamps)."""
    cap = cv2.VideoCapture(path)
    fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    frames = cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0
    duration = frames / fps if fps else 0
    os.makedirs(out_dir, exist_ok=True)

    stamps, cells = [], []
    t = interval / 2  # sample mid-interval, not on the boundary
    while t < duration:
        cap.set(cv2.CAP_PROP_POS_MSEC, t * 1000)
        ok, frame = cap.read()
        if not ok:
            break
        cell = cv2.resize(frame, (CELL_W, CELL_H), interpolation=cv2.INTER_AREA)
        cells.append(label(cell, f"{int(t) // 60}:{int(t) % 60:02d}"))
        stamps.append(round(t, 1))
        t += interval
    cap.release()

    per_sheet = GRID_COLS * GRID_ROWS
    files = []
    for i in range(0, len(cells), per_sheet):
        chunk = cells[i:i + per_sheet]
        rows = []
        for r in range(GRID_ROWS):
            row = chunk[r * GRID_COLS:(r + 1) * GRID_COLS]
            if not row:
                break
            while len(row) < GRID_COLS:
                row.append((cells[0] * 0))  # black filler, same shape
            rows.append(cv2.hconcat(row))
        sheet = cv2.vconcat(rows)
        f = os.path.join(out_dir, f"{video_id}_sheet{i // per_sheet + 1:02d}.jpg")
        cv2.imwrite(f, sheet, [cv2.IMWRITE_JPEG_QUALITY, 72])
        files.append(f)
    return files, stamps, round(duration), round(fps, 2)


def main():
    ap = argparse.ArgumentParser(description="Download, count cuts, build contact sheets")
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--only", help="a single video id")
    ap.add_argument("--work-dir", default=os.environ.get("YT_WORK_DIR", "work"))
    ap.add_argument("--keep-video", action="store_true")
    args = ap.parse_args()

    with open(args.config, encoding="utf-8") as fh:
        cfg = json.load(fh)
    src = os.path.join(cfg["output_dir"], "csv", "Video_Breakdown.csv")
    if not os.path.exists(src):
        sys.exit(f"{src} not found. Run video_breakdown.py first.")
    df = pd.read_csv(src)
    df["Video ID"] = df["Video URL"].str.extract(r"v=([\w-]+)")
    if args.only:
        df = df[df["Video ID"] == args.only]
        if df.empty:
            sys.exit(f"{args.only} is not in the breakdown")

    video_dir = os.path.join(args.work_dir, "videos")
    sheet_dir = os.path.join(args.work_dir, "sheets")
    manifest_path = os.path.join(args.work_dir, "manifest.json")
    manifest = {}
    if os.path.exists(manifest_path):
        with open(manifest_path, encoding="utf-8") as fh:
            manifest = json.load(fh)

    for _, v in df.iterrows():
        vid = v["Video ID"]
        if vid in manifest:
            print(f"{vid}  already done, skipping")
            continue
        print(f"{vid}  {v['Channel']}: downloading")
        try:
            path = download(vid, video_dir)
        except Exception as e:
            print(f"  download failed: {type(e).__name__}: {e}")
            continue
        size_mb = os.path.getsize(path) / 1e6
        interval = sample_interval(v["Length Sec"])
        print(f"  {size_mb:.0f} MB, counting cuts")
        cuts, _ = count_cuts(path)
        print(f"  {cuts} cuts, sampling every {interval}s")
        sheets, stamps, duration, fps = contact_sheets(path, vid, interval, sheet_dir)
        if not stamps:
            print("  no frames decoded, skipping (codec not supported by this OpenCV build)")
            continue
        manifest[vid] = {"channel": v["Channel"], "title": v["Title"], "length_sec": int(v["Length Sec"]),
                         "decoded_sec": duration, "fps": fps, "cuts": cuts, "interval": interval,
                         "samples": len(stamps), "stamps": stamps, "sheets": sheets}
        with open(manifest_path, "w", encoding="utf-8") as fh:
            json.dump(manifest, fh, indent=1)
        print(f"  {len(stamps)} samples across {len(sheets)} sheets")
        if not args.keep_video:
            os.remove(path)

    done = sum(len(m["samples"] if isinstance(m["samples"], list) else [0] * m["samples"]) for m in manifest.values())
    print(f"\n{len(manifest)} videos, {done} samples, manifest at {manifest_path}")


if __name__ == "__main__":
    main()
