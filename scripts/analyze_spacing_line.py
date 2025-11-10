#!/usr/bin/env python3
import argparse
import json
from pathlib import Path

parser = argparse.ArgumentParser(description="Inspect baseline vs candidate spacing deltas for a single line")
parser.add_argument("--spacing", required=True, help="Path to spacing.json (output of extract-spacing.py)")
group = parser.add_mutually_exclusive_group(required=True)
group.add_argument("--line-index", type=int, help="1-based index of the line (sorted by baseline Y)")
group.add_argument("--y", type=float, help="Exact baseline Y (points) to match")
args = parser.parse_args()

data = json.loads(Path(args.spacing).read_text())
base_words = data["base"]["pages"][0]["words"]
cand_words = data["candidate"]["pages"][0]["words"]

def group_by_line(words):
    lines = {}
    for word in words:
        y = word["y"]
        lines.setdefault(y, []).append(word)
    return {y: sorted(words, key=lambda w: w["x"]) for y, words in lines.items()}

base_lines = group_by_line(base_words)
cand_lines = group_by_line(cand_words)
sorted_y = sorted(base_lines.keys())

if args.line_index:
    if args.line_index < 1 or args.line_index > len(sorted_y):
        raise SystemExit(f"line-index must be between 1 and {len(sorted_y)}")
    target_y = sorted_y[args.line_index - 1]
else:
    target_y = min(sorted_y, key=lambda y: abs(y - args.y))
    if abs(target_y - args.y) > 0.5:
        raise SystemExit(f"No baseline line near y={args.y}")

def fetch_line(lines, target_y):
    if target_y in lines:
        return lines[target_y]
    nearest_y = min(lines.keys(), key=lambda y: abs(y - target_y))
    if abs(nearest_y - target_y) > 1.5:
        raise SystemExit(f"No candidate line found near y={target_y}")
    return lines[nearest_y]

base_line = fetch_line(base_lines, target_y)
cand_line = fetch_line(cand_lines, target_y)
if len(base_line) != len(cand_line):
    raise SystemExit(f"Line at y={target_y} has {len(base_line)} baseline words but {len(cand_line)} candidate words.")

def compute_gaps(words):
    gaps = []
    widths = []
    for left, right in zip(words, words[1:]):
        left_end = left["x"] + left["width"]
        gaps.append(right["x"] - left_end)
        widths.append(left["width"])
    return gaps, widths

base_gaps, base_widths = compute_gaps(base_line)
cand_gaps, _ = compute_gaps(cand_line)

print(f"Line y={target_y} (word count {len(base_line)})")
print("-" * 80)

print("Word deltas (candidate x - baseline x):")
for bw, cw in zip(base_line, cand_line):
    dx = cw["x"] - bw["x"]
    print(f"{bw['text']:<12} dx={dx:+.3f}")

if base_gaps:
    total_delta = sum(cg - bg for bg, cg in zip(base_gaps, cand_gaps))
    print("\nGap deltas (candidate - baseline):")
    for i, (bg, cg, width) in enumerate(zip(base_gaps, cand_gaps, base_widths), 1):
        delta = cg - bg
        share = (delta / total_delta) if total_delta else 0.0
        print(f"gap{i:02d}: base={bg:>8.3f} cand={cg:>8.3f} delta={delta:+.3f} share={share:+.3f} prevWordWidth={width:.3f}")

total_extra = sum(cand_gaps) - sum(base_gaps)
print(f"\nTotal extra gap width: {total_extra:+.3f} pt")
