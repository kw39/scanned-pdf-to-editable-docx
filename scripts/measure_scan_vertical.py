#!/usr/bin/env python3
"""Measure vertical spacing on a scanned page image: printed text lines and ruled answer lines.

Give the scale as --pt-per-px, or as --column-pt with --column-px: the width of the same full
justified text line in the rendered PDF (points) and in the scan (pixels). Restrict the region with
--box x0 y0 x1 y1 (pixels) to one column of printed text; exclude figures, and expect handwriting
on answer lines to add noise. Results are measurements for the manifest, not a pass/fail verdict:
check them against the page image. Requires Pillow and numpy.
"""
import argparse
import json
import math
import sys

import numpy as np
from PIL import Image


def bands(flags, min_height):
    """Runs of consecutive True rows, as (start, end) pairs at least min_height tall."""
    result, start = [], None
    for y, flag in enumerate(list(flags) + [False]):
        if flag and start is None:
            start = y
        elif not flag and start is not None:
            if y - start >= min_height:
                result.append((start, y - 1))
            start = None
    return result


def measure(image, box, threshold, pt_per_px):
    if not math.isfinite(pt_per_px) or pt_per_px <= 0:
        raise ValueError("scale must be finite and positive")
    if not 0 <= threshold <= 255:
        raise ValueError("threshold must be between 0 and 255")
    gray = np.asarray(image.convert("L"), dtype=np.int16)
    x0, y0, x1, y1 = box or (0, 0, gray.shape[1], gray.shape[0])
    if not (0 <= x0 < x1 <= gray.shape[1] and 0 <= y0 < y1 <= gray.shape[0]):
        raise ValueError("box must have positive dimensions and stay within the image")
    region = gray[y0:y1, x0:x1] < threshold
    width = region.shape[1]
    coverage = region.mean(axis=1)
    # A ruled line: one thin band where most of the width is dark.
    rules = [(a, b) for a, b in bands(coverage > 0.45, 1) if b - a <= 8]
    rule_rows = np.zeros(len(coverage), bool)
    for a, b in rules:
        rule_rows[max(0, a - 2):b + 3] = True
    # A text line: a band of rows with some ink, ignoring ruled-line rows.
    text = bands((coverage > max(0.01, 8 / width)) & ~rule_rows, 6)
    to_pt = lambda px: round(px * pt_per_px, 2)
    centres = [(a + b) / 2 for a, b in text]
    rule_y = [(a + b) / 2 for a, b in rules]
    return {
        "pt_per_px": pt_per_px,
        "box_px": [x0, y0, x1, y1],
        "text_lines": [{"y_pt": to_pt(c), "height_pt": to_pt(b - a + 1)} for c, (a, b) in zip(centres, text)],
        "text_line_steps_pt": [to_pt(b - a) for a, b in zip(centres, centres[1:])],
        "ruled_lines": len(rules),
        "ruled_line_y_pt": [to_pt(y) for y in rule_y],
        "ruled_line_steps_pt": [to_pt(b - a) for a, b in zip(rule_y, rule_y[1:])],
        "ruled_line_median_pitch_pt": to_pt(float(np.median(np.diff(rule_y)))) if len(rule_y) > 1 else None,
        "note": "Text y values are ink-band centres, not recovered typographic baselines; confirm baseline measurements on the image. y values are relative to the top of the box. A step about one line pitch is a line inside a "
                "paragraph; a larger step includes a paragraph gap (gap = step - pitch). Confirm on the image.",
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("image")
    parser.add_argument("--box", nargs=4, type=int, metavar=("X0", "Y0", "X1", "Y1"))
    scale = parser.add_mutually_exclusive_group(required=True)
    scale.add_argument("--pt-per-px", type=float)
    scale.add_argument("--column-pt", type=float, help="rendered width of a full justified line, in points")
    parser.add_argument("--column-px", type=float, help="scan width of the same line, in pixels")
    parser.add_argument("--threshold", type=int, default=150, help="grey level below which a pixel is ink")
    args = parser.parse_args()
    try:
        if args.column_pt is not None:
            if args.column_px is None or not math.isfinite(args.column_px) or args.column_px <= 0 or not math.isfinite(args.column_pt) or args.column_pt <= 0:
                raise ValueError("--column-pt and --column-px must both be finite and positive")
            pt_per_px = args.column_pt / args.column_px
        else:
            pt_per_px = args.pt_per_px
        if not pt_per_px or pt_per_px <= 0:
            raise ValueError("scale must be positive")
        with Image.open(args.image) as image:
            print(json.dumps(measure(image, args.box, args.threshold, pt_per_px), indent=2))
        return 0
    except (OSError, ValueError) as error:
        print("Measurement failed: " + json.dumps(str(error)), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
