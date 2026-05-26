#!/usr/bin/env python3
"""Calculate image-to-Office point coordinate scaling."""

from __future__ import annotations

import argparse
import json
import sys


def positive_number(value: str) -> float:
    try:
        number = float(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"Expected a number, got {value!r}") from exc
    if number <= 0:
        raise argparse.ArgumentTypeError("Value must be greater than zero.")
    return number


def calculate(image_width: float, image_height: float, target_width_pt: float, target_height_pt: float) -> dict:
    scale_x = target_width_pt / image_width
    scale_y = target_height_pt / image_height
    uniform_scale = min(scale_x, scale_y)
    fitted_width = image_width * uniform_scale
    fitted_height = image_height * uniform_scale
    offset_x = (target_width_pt - fitted_width) / 2
    offset_y = (target_height_pt - fitted_height) / 2

    example_px = {
        "x": image_width * 0.1,
        "y": image_height * 0.1,
        "width": image_width * 0.25,
        "height": image_height * 0.2,
    }
    example_stretched = {
        "left_pt": example_px["x"] * scale_x,
        "top_pt": example_px["y"] * scale_y,
        "width_pt": example_px["width"] * scale_x,
        "height_pt": example_px["height"] * scale_y,
    }
    example_uniform = {
        "left_pt": offset_x + example_px["x"] * uniform_scale,
        "top_pt": offset_y + example_px["y"] * uniform_scale,
        "width_pt": example_px["width"] * uniform_scale,
        "height_pt": example_px["height"] * uniform_scale,
    }

    return {
        "input": {
            "image_width": image_width,
            "image_height": image_height,
            "target_width_pt": target_width_pt,
            "target_height_pt": target_height_pt,
        },
        "scale_x": scale_x,
        "scale_y": scale_y,
        "suggested_uniform_scale": uniform_scale,
        "uniform_fit": {
            "width_pt": fitted_width,
            "height_pt": fitted_height,
            "offset_x_pt": offset_x,
            "offset_y_pt": offset_y,
        },
        "coordinate_formulas": {
            "stretched": "x_pt = x_px * scale_x; y_pt = y_px * scale_y",
            "uniform": "x_pt = offset_x_pt + x_px * suggested_uniform_scale; y_pt = offset_y_pt + y_px * suggested_uniform_scale",
        },
        "example_source_box_px": example_px,
        "example_stretched_box_pt": {key: round(value, 3) for key, value in example_stretched.items()},
        "example_uniform_box_pt": {key: round(value, 3) for key, value in example_uniform.items()},
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Compute scaling factors from image pixels to Office point coordinates.",
    )
    parser.add_argument("image_width", type=positive_number, help="Original image width in pixels.")
    parser.add_argument("image_height", type=positive_number, help="Original image height in pixels.")
    parser.add_argument("target_width_pt", type=positive_number, help="Target Office canvas width in points.")
    parser.add_argument("target_height_pt", type=positive_number, help="Target Office canvas height in points.")
    parser.add_argument("--pretty", action="store_true", help="Pretty-print JSON output.")
    args = parser.parse_args()

    result = calculate(
        args.image_width,
        args.image_height,
        args.target_width_pt,
        args.target_height_pt,
    )
    print(json.dumps(result, indent=2 if args.pretty else None))
    return 0


if __name__ == "__main__":
    sys.exit(main())
