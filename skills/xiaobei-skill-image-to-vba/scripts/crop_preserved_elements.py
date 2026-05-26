#!/usr/bin/env python3
"""Crop preserved raster elements for hybrid image-to-VBA reconstruction."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


def load_image(path: Path):
    try:
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("Pillow is required. Install with: python -m pip install pillow") from exc

    try:
        return Image.open(path).convert("RGBA")
    except OSError as exc:
        raise RuntimeError(f"Cannot open image {path}: {exc}") from exc


def safe_name(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", value.strip())
    cleaned = cleaned.strip("_")
    return cleaned or "preserved_element"


def parse_region(value: str) -> dict:
    parts = value.split(":")
    if len(parts) != 5:
        raise argparse.ArgumentTypeError(
            "Region must be name:x:y:w:h, for example logo:120:80:240:160"
        )
    name, x, y, width, height = parts
    try:
        box = {
            "name": name,
            "x": float(x),
            "y": float(y),
            "width": float(width),
            "height": float(height),
        }
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"Invalid numeric region values: {value}") from exc
    if box["width"] <= 0 or box["height"] <= 0:
        raise argparse.ArgumentTypeError("Region width and height must be greater than zero.")
    return box


def load_regions(args: argparse.Namespace) -> list[dict]:
    regions = list(args.region or [])
    if args.regions_json:
        path = Path(args.regions_json)
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except OSError as exc:
            raise RuntimeError(f"Cannot read regions JSON {path}: {exc}") from exc
        except json.JSONDecodeError as exc:
            raise RuntimeError(f"Invalid JSON in {path}: {exc}") from exc
        if not isinstance(payload, list):
            raise RuntimeError("Regions JSON must contain a list of region objects.")
        for item in payload:
            if not isinstance(item, dict):
                raise RuntimeError("Each region must be an object.")
            required = {"name", "x", "y", "width", "height"}
            missing = required - set(item)
            if missing:
                raise RuntimeError(f"Region is missing keys: {', '.join(sorted(missing))}")
            regions.append(
                {
                    "name": str(item["name"]),
                    "x": float(item["x"]),
                    "y": float(item["y"]),
                    "width": float(item["width"]),
                    "height": float(item["height"]),
                    "reason": item.get("reason", ""),
                }
            )
    return regions


def clamp_box(region: dict, image_size: tuple[int, int], padding: float) -> tuple[int, int, int, int]:
    img_w, img_h = image_size
    left = max(0, int(round(region["x"] - padding)))
    top = max(0, int(round(region["y"] - padding)))
    right = min(img_w, int(round(region["x"] + region["width"] + padding)))
    bottom = min(img_h, int(round(region["y"] + region["height"] + padding)))
    if right <= left or bottom <= top:
        raise RuntimeError(f"Region {region['name']!r} is outside the image bounds.")
    return left, top, right, bottom


def crop_regions(source_path: Path, output_dir: Path, regions: list[dict], padding: float) -> dict:
    image = load_image(source_path)
    output_dir.mkdir(parents=True, exist_ok=True)
    assets = []

    for index, region in enumerate(regions, start=1):
        box = clamp_box(region, image.size, padding)
        crop = image.crop(box)
        filename = f"{index:02d}_{safe_name(region['name'])}.png"
        asset_path = output_dir / filename
        crop.save(asset_path)
        left, top, right, bottom = box
        assets.append(
            {
                "name": region["name"],
                "reason": region.get("reason", ""),
                "asset_path": str(asset_path.resolve()),
                "source_box_px": {
                    "x": left,
                    "y": top,
                    "width": right - left,
                    "height": bottom - top,
                },
            }
        )

    return {
        "source_image": str(source_path.resolve()),
        "source_size_px": {"width": image.size[0], "height": image.size[1]},
        "output_dir": str(output_dir.resolve()),
        "assets": assets,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Crop complex graphic elements that should be preserved as local image "
            "assets while the rest of a figure is reconstructed as editable VBA Shapes."
        )
    )
    parser.add_argument("source_image", help="Path to the source image.")
    parser.add_argument("output_dir", help="Directory where cropped assets should be written.")
    parser.add_argument(
        "--region",
        action="append",
        type=parse_region,
        help="Preserved region as name:x:y:w:h in source pixels. May be repeated.",
    )
    parser.add_argument(
        "--regions-json",
        help="Optional JSON file with region objects containing name, x, y, width, height, and optional reason.",
    )
    parser.add_argument("--padding", type=float, default=0, help="Padding in source pixels around every crop.")
    parser.add_argument("--manifest", help="Optional path to write a JSON manifest.")
    parser.add_argument("--pretty", action="store_true", help="Pretty-print JSON output.")
    args = parser.parse_args()

    source_path = Path(args.source_image)
    if not source_path.is_file():
        print(json.dumps({"passed": False, "errors": [f"File not found: {source_path}"]}))
        return 2

    try:
        regions = load_regions(args)
        if not regions:
            raise RuntimeError("At least one --region or --regions-json entry is required.")
        result = crop_regions(source_path, Path(args.output_dir), regions, args.padding)
        result["passed"] = True
        if args.manifest:
            manifest_path = Path(args.manifest)
            manifest_path.parent.mkdir(parents=True, exist_ok=True)
            manifest_path.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
            result["manifest_path"] = str(manifest_path.resolve())
    except RuntimeError as exc:
        print(json.dumps({"passed": False, "errors": [str(exc)]}, indent=2 if args.pretty else None))
        return 2

    print(json.dumps(result, indent=2 if args.pretty else None, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
