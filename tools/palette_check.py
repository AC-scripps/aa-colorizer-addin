#!/usr/bin/env python3
"""Palette analysis for the amino acid colorizer.

Reads AA_CATEGORIES straight out of docs/aa-colors.js so it can never drift
from what ships, then reports:

  * contrast against white for every category  (legibility)
  * the distance from every color to every other  (confusability)
  * optionally, candidate colors for a new category

Usage:
    python3 tools/palette_check.py
    python3 tools/palette_check.py --suggest 120 175   # hue range to search

Why distance matters: the palette's tightest existing pair is blue vs cyan at
80. Treat that as the floor. A color that clears the numbers can still look
wrong in use -- check tools/color-swatch.html before committing.
"""

import argparse
import colorsys
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
AA_COLORS_JS = os.path.join(HERE, "..", "docs", "aa-colors.js")


def load_palette():
    """Parse AA_CATEGORIES out of aa-colors.js. Ignores commented-out lines."""
    with open(AA_COLORS_JS, encoding="utf-8") as fh:
        source = fh.read()

    block = re.search(r"var AA_CATEGORIES\s*=\s*\[(.*?)\];", source, re.S)
    if not block:
        sys.exit("could not find AA_CATEGORIES in " + AA_COLORS_JS)

    out = []
    for line in block.group(1).splitlines():
        if line.strip().startswith("//"):
            continue
        m = re.search(r'name:\s*"([^"]+)".*?residues:\s*"([^"]+)".*?color:\s*"([0-9A-Fa-f]{6})"', line)
        if m:
            out.append((m.group(1), m.group(2), m.group(3).upper()))
    return out


def rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def luminance(h):
    chan = [c / 255 for c in rgb(h)]
    chan = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in chan]
    return 0.2126 * chan[0] + 0.7152 * chan[1] + 0.0722 * chan[2]


def contrast(a, b="FFFFFF"):
    l1, l2 = luminance(a), luminance(b)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def distance(a, b):
    return sum((x - y) ** 2 for x, y in zip(rgb(a), rgb(b))) ** 0.5


def report(palette):
    print(f"{'residues':<8}{'hex':<9}{'contrast':>9}{'nearest':>10}{'dist':>7}   category")
    print("-" * 74)
    for name, residues, hexv in palette:
        others = [(distance(hexv, o[2]), o[1]) for o in palette if o[2] != hexv]
        d, near = min(others) if others else (float("inf"), "-")
        flag = "  <-- tight" if d < 80 else ""
        print(f"{residues:<8}{hexv:<9}{contrast(hexv):>9.2f}{near:>10}{d:>7.0f}   {name}{flag}")

    pairs = [
        (distance(a[2], b[2]), a[1], b[1])
        for i, a in enumerate(palette) for b in palette[i + 1:]
    ]
    if pairs:
        d, x, y = min(pairs)
        print(f"\ntightest pair: {x} vs {y} = {d:.0f}")
        print("(treat this as the floor for any new color)")

    low = [p for p in palette if contrast(p[2]) < 2.2]
    if low:
        print("\nlow contrast on white (may be hard to read):")
        for name, residues, hexv in low:
            print(f"  {residues:<8}{hexv}  {contrast(hexv):.2f}")


def suggest(palette, hue_lo, hue_hi, min_contrast=2.3, max_contrast=3.1, floor=80):
    existing = [p[2] for p in palette]
    found = []
    for hdeg in range(hue_lo, hue_hi + 1, 3):
        for sat in range(45, 101, 5):
            for lig in range(35, 65, 2):
                r, g, b = colorsys.hls_to_rgb(hdeg / 360, lig / 100, sat / 100)
                hexv = "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))
                ct = contrast(hexv)
                if not (min_contrast <= ct <= max_contrast):
                    continue
                dmin = min(distance(hexv, o) for o in existing)
                if dmin < floor:
                    continue
                found.append((dmin, ct, hexv, hdeg, sat, lig))

    found.sort(reverse=True)
    print(f"\ncandidates in hue {hue_lo}-{hue_hi}, contrast {min_contrast}-{max_contrast}, "
          f"min distance >= {floor}:")
    print(f"{'hex':<9}{'minDist':>9}{'contrast':>10}{'hue':>6}{'sat':>5}{'lig':>5}")
    print("-" * 44)
    seen = set()
    for dmin, ct, hexv, hdeg, sat, lig in found:
        if hexv in seen:
            continue
        seen.add(hexv)
        print(f"{hexv:<9}{dmin:>9.0f}{ct:>10.2f}{hdeg:>6}{sat:>5}{lig:>5}")
        if len(seen) >= 12:
            break
    if not found:
        print("  none — widen the hue range or relax the contrast window")


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--suggest", nargs=2, type=int, metavar=("HUE_LO", "HUE_HI"),
                    help="search this hue range for a new category color")
    args = ap.parse_args()

    palette = load_palette()
    print(f"palette from {os.path.relpath(AA_COLORS_JS, os.path.join(HERE, '..'))} "
          f"({len(palette)} categories)\n")
    report(palette)
    if args.suggest:
        suggest(palette, args.suggest[0], args.suggest[1])


if __name__ == "__main__":
    main()
