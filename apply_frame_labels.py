#!/usr/bin/env python3
"""
apply_frame_labels.py

Turns per-frame visual labels into the Sec and Cut columns of the breakdown table.

video_frames.py samples one frame every N seconds and tiles them into contact
sheets. A reviewer reads those sheets and records one letter per sample:

  D  Data graphics    charts, diagrams, animated figures, on-screen text and
                      documents used as evidence, title and quote cards
  R  Real footage     camera footage or photographs of real people, places, events
  S  Stock/generated  generic stock clips, AI imagery, 3D renders and recreations
  M  Maps             any map, still or animated, including 3D terrain
  O  Other            logos, idents, endcards, plain colour, host-only talking head

Each sample stands for N seconds, so seconds in a category = count * interval.
That makes the five Sec columns add up to the sampled runtime rather than the
exact runtime, which is why Classified Sec is its own column.

Labels live in frame_labels.json as one string per video, one character per
sample, in timestamp order. Relabelling is just editing that file and rerunning,
so a change of rule costs nothing.

Usage:
  python apply_frame_labels.py --work-dir work
"""

import argparse
import json
import os
import sys

import pandas as pd
from openpyxl import load_workbook

CATEGORY = {
    "D": "Sec Data Graphics",
    "R": "Sec Real Footage or Photos",
    "S": "Sec Stock or Generated Imagery",
    "M": "Sec Maps",
    "O": "Sec Other",
}


def main():
    ap = argparse.ArgumentParser(description="Apply frame labels to the breakdown table")
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--work-dir", default=os.environ.get("YT_WORK_DIR", "work"))
    ap.add_argument("--labels", default="frame_labels.json")
    args = ap.parse_args()

    with open(args.config, encoding="utf-8") as fh:
        cfg = json.load(fh)
    manifest_path = os.path.join(args.work_dir, "manifest.json")
    for p in (manifest_path, args.labels):
        if not os.path.exists(p):
            sys.exit(f"{p} not found")
    with open(manifest_path, encoding="utf-8") as fh:
        manifest = json.load(fh)
    with open(args.labels, encoding="utf-8") as fh:
        labels = json.load(fh)

    xlsx = os.path.join(cfg["output_dir"], "video_breakdown.xlsx")
    csv = os.path.join(cfg["output_dir"], "csv", "Video_Breakdown.csv")
    df = pd.read_csv(csv)
    df["Video ID"] = df["Video URL"].str.extract(r"v=([\w-]+)")

    for vid, m in manifest.items():
        row = df.index[df["Video ID"] == vid]
        if not len(row):
            print(f"{vid} not in the breakdown, skipped")
            continue
        i = row[0]
        df.loc[i, "Cut Count"] = m["cuts"]
        s = "".join(labels.get(vid, "").split()).upper()
        if not s:
            print(f"{vid} has no labels yet, cut count only")
            continue
        n = len(m["stamps"])
        if len(s) != n:
            print(f"{vid}: {len(s)} labels for {n} samples, check frame_labels.json")
            continue
        bad = set(s) - set(CATEGORY)
        if bad:
            print(f"{vid}: unknown label(s) {sorted(bad)}")
            continue
        interval = m["interval"]
        for letter, col in CATEGORY.items():
            df.loc[i, col] = round(s.count(letter) * interval)

    df["Classified Sec"] = df[list(CATEGORY.values())].sum(axis=1, min_count=1)
    for letter, col in CATEGORY.items():
        pct = {"D": "Pct Data Graphics", "R": "Pct Real", "S": "Pct Stock or Generated", "M": "Pct Maps"}.get(letter)
        if pct:
            df[pct] = (df[col] / df["Classified Sec"]).round(4)
    df["Cuts per Min"] = (df["Cut Count"] / (df["Length Sec"] / 60)).round(1)

    out_cols = [c for c in df.columns if c != "Video ID"]
    df[out_cols].to_csv(csv, index=False, encoding="utf-8-sig")

    # write the same values into the workbook, replacing the placeholder formulas
    if os.path.exists(xlsx):
        wb = load_workbook(xlsx)
        ws = wb["Breakdown"]
        header = [c.value for c in ws[1]]
        for r in range(2, ws.max_row + 1):
            url = ws.cell(row=r, column=header.index("Video URL") + 1).value
            match = df[df["Video URL"] == url]
            if match.empty:
                continue
            rec = match.iloc[0]
            for col in list(CATEGORY.values()) + ["Cut Count", "Cuts per Min", "Classified Sec",
                                                  "Pct Data Graphics", "Pct Real", "Pct Stock or Generated", "Pct Maps"]:
                v = rec[col]
                if pd.notna(v):
                    keep_decimals = "Pct" in col or col == "Cuts per Min"
                    ws.cell(row=r, column=header.index(col) + 1).value = float(v) if keep_decimals else int(v)
        try:
            wb.save(xlsx)
        except PermissionError:
            sys.exit(f"Could not write {xlsx}; close it in Excel and rerun.")
        print(f"Updated {xlsx}")
    print(f"Updated {csv}")

    done = df["Classified Sec"].notna().sum()
    print(f"{done} of {len(df)} videos have visual categories, {df['Cut Count'].notna().sum()} have cut counts")


if __name__ == "__main__":
    main()
