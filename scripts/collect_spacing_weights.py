#!/usr/bin/env python3
import argparse
import html
import json
from pathlib import Path

parser = argparse.ArgumentParser(description="Collect gap spacing deltas per alignment")
parser.add_argument("--spacing", required=True, help="spacing.json from extract-spacing.py")
parser.add_argument("--alignment-map", required=True, help="JSON produced by alignment-map script")
parser.add_argument("--sample", required=True, help="sample name key inside alignment-map JSON")
parser.add_argument("--alignments", default="both,distribute", help="comma-separated list of paragraph alignments to include")
args = parser.parse_args()

spacing = json.loads(Path(args.spacing).read_text())
align_map = json.loads(Path(args.alignment_map).read_text())[args.sample]
target_alignments = {token.strip() for token in args.alignments.split(",") if token.strip()}

def tokenize(text):
    return [tok for tok in html.unescape(text).split() if tok]

paragraphs = []
for entry in align_map:
    tokens = tokenize(entry["text"])
    if tokens:
        paragraphs.append({"alignment": entry["alignment"], "tokens": tokens})

def group_lines(words):
    lines = {}
    for word in words:
        y = round(word["y"], 3)
        lines.setdefault(y, []).append(word)
    sorted_lines = []
    for y in sorted(lines.keys()):
        sorted_lines.append({"y": y, "words": sorted(lines[y], key=lambda w: w["x"])})
    return sorted_lines

base_lines = group_lines(spacing["base"]["pages"][0]["words"])
cand_lines = group_lines(spacing["candidate"]["pages"][0]["words"])
if len(base_lines) != len(cand_lines):
    raise SystemExit("Line count differs between baseline and candidate.")

line_alignments = []
para_idx = 0
token_idx = 0

for line in base_lines:
    tokens = [w["text"] for w in line["words"]]
    while para_idx < len(paragraphs) and token_idx >= len(paragraphs[para_idx]["tokens"]):
        para_idx += 1
        token_idx = 0
    if para_idx >= len(paragraphs):
        line_alignments.append("unknown")
        continue
    para_tokens = paragraphs[para_idx]["tokens"]
    slice_tokens = para_tokens[token_idx : token_idx + len(tokens)]
    if slice_tokens == tokens:
        token_idx += len(tokens)
        line_alignments.append(paragraphs[para_idx]["alignment"])
    else:
        # if mismatch, try restarting at paragraph boundary
        if tokens == para_tokens[: len(tokens)]:
            para_idx += 1
            token_idx = len(tokens)
            line_alignments.append(paragraphs[para_idx - 1]["alignment"])
        else:
            line_alignments.append("unknown")

entries = []
for line_index, (base_line, cand_line, align) in enumerate(zip(base_lines, cand_lines, line_alignments), 1):
    if align not in target_alignments:
        continue
    base_words = base_line["words"]
    cand_words = cand_line["words"]
    if len(base_words) != len(cand_words):
        continue
    for i in range(len(base_words) - 1):
        left_base = base_words[i]
        right_base = base_words[i + 1]
        left_cand = cand_words[i]
        right_cand = cand_words[i + 1]
        base_gap = right_base["x"] - (left_base["x"] + left_base["width"])
        cand_gap = right_cand["x"] - (left_cand["x"] + left_cand["width"])
        delta = cand_gap - base_gap
        entries.append(
            {
                "line_index": line_index,
                "alignment": align,
                "word": left_base["text"],
                "width": left_base["width"],
                "base_gap": base_gap,
                "cand_gap": cand_gap,
                "delta": delta,
            }
        )

if not entries:
    print("No entries collected.")
    raise SystemExit(0)

total_delta = sum(entry["delta"] for entry in entries)
abs_delta = sum(abs(entry["delta"]) for entry in entries)

print(f"Entries collected: {len(entries)}")
print(f"Sum delta: {total_delta:+.3f} pt, sum |delta|: {abs_delta:.3f} pt")

entries.sort(key=lambda e: abs(e["delta"]), reverse=True)
for entry in entries[: min(10, len(entries))]:
    print(
        f"line {entry['line_index']:02d} align={entry['alignment']:<10} word={entry['word']:<12} "
        f"width={entry['width']:.2f} delta={entry['delta']:+.3f}"
    )

from collections import defaultdict
groups = defaultdict(list)
for entry in entries:
    groups[entry["alignment"]].append(entry)

print("\nPer-alignment summary:")
for alignment, group in groups.items():
    total = sum(entry["delta"] for entry in group)
    abs_total = sum(abs(entry["delta"]) for entry in group)
    width_sum = sum(entry["width"] for entry in group)
    avg_ratio = total / width_sum if width_sum else 0.0
    print(
        f"{alignment:<10} count={len(group):3d} sum_delta={total:+.3f} "
        f"sum|delta|={abs_total:.3f} delta/width={avg_ratio:+.4f}"
    )
