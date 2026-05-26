#!/usr/bin/env python3
"""Compare a source image with a rendered VBA screenshot."""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path


def load_image(path: Path):
    try:
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("Pillow is required to load images. Install with: python -m pip install pillow") from exc

    try:
        return Image.open(path).convert("RGB")
    except OSError as exc:
        raise RuntimeError(f"Cannot open image {path}: {exc}") from exc


def resize_to_match(image, size):
    try:
        from PIL import Image
        resample = getattr(Image, "Resampling", Image).BICUBIC
        return image.resize(size, resample)
    except Exception:
        return image.resize(size)


def channel_stats(img_a, img_b):
    width, height = img_a.size
    count = width * height
    total_abs = [0, 0, 0]
    total_pixel_delta = 0.0
    strong_diff = 0
    quadrants = {
        "top_left": 0.0,
        "top_right": 0.0,
        "bottom_left": 0.0,
        "bottom_right": 0.0,
    }
    quadrant_counts = {key: 0 for key in quadrants}

    pix_a = img_a.load()
    pix_b = img_b.load()
    for y in range(height):
        for x in range(width):
            a = pix_a[x, y]
            b = pix_b[x, y]
            diffs = [abs(a[i] - b[i]) for i in range(3)]
            for i, diff in enumerate(diffs):
                total_abs[i] += diff
            pixel_delta = sum(diffs) / 3
            total_pixel_delta += pixel_delta
            if pixel_delta > 32:
                strong_diff += 1
            if y < height / 2 and x < width / 2:
                key = "top_left"
            elif y < height / 2:
                key = "top_right"
            elif x < width / 2:
                key = "bottom_left"
            else:
                key = "bottom_right"
            quadrants[key] += pixel_delta
            quadrant_counts[key] += 1

    mean_channel_abs = [round(value / count, 3) for value in total_abs]
    mean_pixel_abs = round(total_pixel_delta / count, 3)
    strong_diff_ratio = round(strong_diff / count, 4)
    quadrant_scores = {
        key: round(value / max(quadrant_counts[key], 1), 3)
        for key, value in quadrants.items()
    }
    return mean_channel_abs, mean_pixel_abs, strong_diff_ratio, quadrant_scores


def try_ssim(img_a, img_b):
    try:
        import numpy as np
        from skimage.metrics import structural_similarity
    except ImportError:
        return None, "skimage is not installed; SSIM skipped."

    arr_a = np.asarray(img_a)
    arr_b = np.asarray(img_b)
    try:
        score = structural_similarity(arr_a, arr_b, channel_axis=2)
    except TypeError:
        score = structural_similarity(arr_a, arr_b, multichannel=True)
    except Exception as exc:
        return None, f"SSIM failed: {exc}"
    return round(float(score), 4), None


def suggestions(size_delta, mean_pixel_abs, strong_diff_ratio, quadrant_scores):
    notes = []
    if size_delta["width"] != 0 or size_delta["height"] != 0:
        notes.append("Match the Office canvas or export size before tuning individual shapes.")
    if mean_pixel_abs > 35:
        notes.append("Large average color difference: review background, palette, gradients, and transparency first.")
    elif mean_pixel_abs > 18:
        notes.append("Moderate color difference: tune key fills, text colors, and shadow transparency.")
    if strong_diff_ratio > 0.2:
        notes.append("Many pixels differ strongly: check object positions, missing regions, and layer order.")
    worst_region = max(quadrant_scores, key=quadrant_scores.get)
    if quadrant_scores[worst_region] > 20:
        notes.append(f"Focus the next visual pass on the {worst_region.replace('_', ' ')} region.")
    if not notes:
        notes.append("Images are broadly close; refine text metrics, small icons, and stroke weights.")
    return notes


def compare(source_path: Path, rendered_path: Path, sample_width: int):
    source = load_image(source_path)
    rendered = load_image(rendered_path)
    original_sizes = {"source": source.size, "rendered": rendered.size}
    size_delta = {
        "width": rendered.size[0] - source.size[0],
        "height": rendered.size[1] - source.size[1],
    }

    if source.size != rendered.size:
        rendered_for_compare = resize_to_match(rendered, source.size)
    else:
        rendered_for_compare = rendered

    if sample_width > 0 and source.size[0] > sample_width:
        scale = sample_width / source.size[0]
        sample_size = (sample_width, max(1, int(round(source.size[1] * scale))))
        source_for_compare = resize_to_match(source, sample_size)
        rendered_for_compare = resize_to_match(rendered_for_compare, sample_size)
    else:
        source_for_compare = source

    mean_channel_abs, mean_pixel_abs, strong_diff_ratio, quadrant_scores = channel_stats(
        source_for_compare,
        rendered_for_compare,
    )
    ssim, ssim_note = try_ssim(source_for_compare, rendered_for_compare)

    result = {
        "passed": True,
        "sizes": {
            "source": {"width": original_sizes["source"][0], "height": original_sizes["source"][1]},
            "rendered": {"width": original_sizes["rendered"][0], "height": original_sizes["rendered"][1]},
            "delta_rendered_minus_source": size_delta,
        },
        "mean_color_difference_rgb": {
            "red": mean_channel_abs[0],
            "green": mean_channel_abs[1],
            "blue": mean_channel_abs[2],
        },
        "mean_pixel_difference": mean_pixel_abs,
        "strong_pixel_difference_ratio": strong_diff_ratio,
        "region_difference_scores": quadrant_scores,
        "ssim": ssim,
        "notes": [] if ssim_note is None else [ssim_note],
        "suggested_modifications": suggestions(size_delta, mean_pixel_abs, strong_diff_ratio, quadrant_scores),
    }
    if math.isnan(mean_pixel_abs):
        result["passed"] = False
        result["notes"].append("Comparison produced NaN; check image files.")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Compare an original image with a screenshot exported from VBA-rendered Office Shapes.",
    )
    parser.add_argument("source_image", help="Path to the original reference image.")
    parser.add_argument("rendered_image", help="Path to the VBA-rendered screenshot/export.")
    parser.add_argument(
        "--sample-width",
        type=int,
        default=900,
        help="Resize comparison copies to this width for speed; use 0 for full size.",
    )
    parser.add_argument("--pretty", action="store_true", help="Pretty-print JSON output.")
    args = parser.parse_args()

    source_path = Path(args.source_image)
    rendered_path = Path(args.rendered_image)
    for path in (source_path, rendered_path):
        if not path.exists():
            print(json.dumps({"passed": False, "errors": [f"File not found: {path}"]}))
            return 2
        if not path.is_file():
            print(json.dumps({"passed": False, "errors": [f"Not a file: {path}"]}))
            return 2

    try:
        result = compare(source_path, rendered_path, args.sample_width)
    except RuntimeError as exc:
        print(json.dumps({"passed": False, "errors": [str(exc)]}, indent=2 if args.pretty else None))
        return 2

    print(json.dumps(result, indent=2 if args.pretty else None))
    return 0 if result.get("passed") else 1


if __name__ == "__main__":
    sys.exit(main())
