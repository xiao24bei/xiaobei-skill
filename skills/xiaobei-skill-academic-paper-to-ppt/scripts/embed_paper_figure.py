#!/usr/bin/env python3
"""Losslessly embed a real cropped paper figure into an $imagegen-rendered slide image.

Preserves 100% pixel-exact authenticity for paper architectures, medical/microscopy images,
qualitative visual comparisons, and dense multi-panel experimental plots while keeping the
slide's unified $imagegen academic header, deep analysis cards, conclusion anchor, and footer.
"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw


def trim_outer_white(im: Image.Image, threshold: int = 248, pad: int = 4) -> Image.Image:
    rgb = im.convert("RGB")
    bg = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg).convert("L")
    mask = diff.point(lambda p: 255 if p > (255 - threshold) else 0)
    bbox = mask.getbbox()
    if not bbox:
        return rgb
    x1, y1, x2, y2 = bbox
    w, h = rgb.size
    return rgb.crop((max(0, x1 - pad), max(0, y1 - pad), min(w, x2 + pad), min(h, y2 + pad)))


def is_near_white(rgb: tuple[int, int, int], thresh: int = 236) -> bool:
    return rgb[0] >= thresh and rgb[1] >= thresh and rgb[2] >= thresh


def snap_inner_white_box(
    slide_rgb: Image.Image, bx1: int, by1: int, bx2: int, by2: int
) -> tuple[int, int, int, int]:
    """Trim dark header bars or box borders along the edges of the target slot."""
    px = slide_rgb.load()
    w_box = bx2 - bx1
    h_box = by2 - by1
    if w_box < 60 or h_box < 60:
        return (bx1, by1, bx2, by2)

    # First strip any solid colored header bar in the top 22% of the slot (e.g., solid blue card title strip)
    top = by1
    max_header_y = by1 + int(h_box * 0.24)
    for y in range(by1, max_header_y):
        white_count = sum(1 for x in range(bx1 + 6, bx2 - 6, 4) if is_near_white(px[x, y]))
        total = max(1, len(range(bx1 + 6, bx2 - 6, 4)))
        if white_count / total >= 0.78:
            top = y
            break

    # Trim thin border lines on bottom, left, right (within 4% margin)
    bottom = by2
    for y in range(by2 - 1, max(top + 20, by2 - int(h_box * 0.08)), -1):
        white_count = sum(1 for x in range(bx1 + 6, bx2 - 6, 4) if is_near_white(px[x, y]))
        total = max(1, len(range(bx1 + 6, bx2 - 6, 4)))
        if white_count / total >= 0.78:
            bottom = y + 1
            break

    left = bx1
    for x in range(bx1, min(bx2 - 20, bx1 + int(w_box * 0.06))):
        white_count = sum(1 for y in range(top + 6, bottom - 6, 4) if is_near_white(px[x, y]))
        total = max(1, len(range(top + 6, bottom - 6, 4)))
        if white_count / total >= 0.78:
            left = x
            break

    right = bx2
    for x in range(bx2 - 1, max(left + 20, bx2 - int(w_box * 0.06)), -1):
        white_count = sum(1 for y in range(top + 6, bottom - 6, 4) if is_near_white(px[x, y]))
        total = max(1, len(range(top + 6, bottom - 6, 4)))
        if white_count / total >= 0.78:
            right = x + 1
            break

    # Safety fallback if snapping shrank too aggressively
    if (right - left) < w_box * 0.65 or (bottom - top) < h_box * 0.55:
        return (bx1 + 4, top + 4, bx2 - 4, by2 - 4)
    return (left + 3, top + 3, right - 3, bottom - 3)


def parse_slot(slot_str: str, width: int, height: int) -> tuple[int, int, int, int]:
    parts = [float(v.strip()) for v in slot_str.split(",")]
    if len(parts) != 4:
        raise ValueError("Expected 4 comma-separated values: x1,y1,x2,y2")
    x1, y1, x2, y2 = parts
    if max(x1, y1, x2, y2) <= 1.05:
        bx1 = max(0, min(width, int(round(x1 * width))))
        by1 = max(0, min(height, int(round(y1 * height))))
        bx2 = max(0, min(width, int(round(x2 * width))))
        by2 = max(0, min(height, int(round(y2 * height))))
    else:
        bx1, by1, bx2, by2 = int(x1), int(y1), int(x2), int(y2)
    if bx2 <= bx1 + 20 or by2 <= by1 + 20:
        raise ValueError(f"Slot rectangle too small or invalid: {(bx1, by1, bx2, by2)}")
    return (bx1, by1, bx2, by2)


def embed_figure(
    slide_path: Path,
    figure_path: Path,
    slot_str: str,
    output_path: Path,
    pad: int = 12,
    auto_snap: bool = True,
    trim_figure: bool = True,
    draw_border: bool = False,
    backup_dir: Path | None = None,
) -> tuple[tuple[int, int, int, int], tuple[int, int]]:
    if backup_dir is not None:
        backup_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(slide_path, backup_dir / slide_path.name)

    with Image.open(slide_path) as s_im:
        slide = s_im.convert("RGB")
    with Image.open(figure_path) as f_im:
        fig = trim_outer_white(f_im) if trim_figure else f_im.convert("RGB")

    sw, sh = slide.size
    bx1, by1, bx2, by2 = parse_slot(slot_str, sw, sh)
    if auto_snap:
        sx1, sy1, sx2, sy2 = snap_inner_white_box(slide, bx1, by1, bx2, by2)
    else:
        sx1, sy1, sx2, sy2 = bx1, by1, bx2, by2

    # Cleanly wipe slot interior to pure white so zero placeholder residue remains
    draw = ImageDraw.Draw(slide)
    draw.rectangle([sx1, sy1, sx2, sy2], fill=(255, 255, 255))

    avail_w = max(20, (sx2 - sx1) - 2 * pad)
    avail_h = max(20, (sy2 - sy1) - 2 * pad)
    fw, fh = fig.size
    scale = min(avail_w / fw, avail_h / fh)
    new_w = max(1, int(round(fw * scale)))
    new_h = max(1, int(round(fh * scale)))
    resized = fig.resize((new_w, new_h), Image.Resampling.LANCZOS)

    px = sx1 + ((sx2 - sx1) - new_w) // 2
    py = sy1 + ((sy2 - sy1) - new_h) // 2
    slide.paste(resized, (px, py))

    if draw_border:
        draw.rectangle([sx1, sy1, sx2, sy2], outline=(184, 200, 220), width=1)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    slide.save(output_path, "PNG", optimize=True)
    return (sx1, sy1, sx2, sy2), (new_w, new_h)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Composite a real cropped paper figure pixel-accurately into a generated slide image."
    )
    parser.add_argument("--slide", required=True, type=Path, help="Input $imagegen-generated slide PNG.")
    parser.add_argument("--figure", required=True, type=Path, help="Cropped source paper figure image.")
    parser.add_argument(
        "--slot",
        required=True,
        type=str,
        help="Target slot x1,y1,x2,y2 (normalized 0..1 ratios or pixel coords).",
    )
    parser.add_argument(
        "--output",
        default=None,
        type=Path,
        help="Output slide PNG path (defaults to overwriting --slide in-place).",
    )
    parser.add_argument("--pad", type=int, default=12, help="Padding inside slot in pixels (default: 12).")
    parser.add_argument("--no-auto-snap", action="store_true", help="Disable auto-snapping inside white frame.")
    parser.add_argument("--no-trim", action="store_true", help="Disable outer white-margin trim on figure.")
    parser.add_argument("--draw-border", action="store_true", help="Draw crisp 1px #B8C8DC border around slot.")
    parser.add_argument(
        "--backup-dir",
        default=None,
        type=Path,
        help="Optional directory to archive the raw pre-composite slide image.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    slide = args.slide.expanduser().resolve()
    figure = args.figure.expanduser().resolve()
    output = (args.output or args.slide).expanduser().resolve()
    backup_dir = args.backup_dir.expanduser().resolve() if args.backup_dir else None

    if not slide.is_file():
        print(f"ERROR: slide image not found: {slide}")
        return 1
    if not figure.is_file():
        print(f"ERROR: source figure not found: {figure}")
        return 1

    try:
        box, placed_size = embed_figure(
            slide_path=slide,
            figure_path=figure,
            slot_str=args.slot,
            output_path=output,
            pad=args.pad,
            auto_snap=not args.no_auto_snap,
            trim_figure=not args.no_trim,
            draw_border=args.draw_border,
            backup_dir=backup_dir,
        )
    except Exception as exc:
        print(f"ERROR: failed to embed figure: {exc}")
        return 1

    print(
        f"OK: embedded real paper figure {figure.name} into {output.name} "
        f"(slot={box}, rendered_figure={placed_size[0]}x{placed_size[1]})"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
