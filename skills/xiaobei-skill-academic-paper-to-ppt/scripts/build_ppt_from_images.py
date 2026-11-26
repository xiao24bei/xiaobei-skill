#!/usr/bin/env python3
"""Build a 16:9 strict image-only PowerPoint deck from slide images.

Features:
- Inserts one full-bleed 16:9 bitmap per slide (zero editable visible text on slide canvas).
- Optionally reads `### 演讲备注` from `ppt_outline.md` and populates PPTX Speaker Notes
  (stored in ppt/notesSlides/, keeping the slide canvas 100% strict image-only).
- Optionally builds a numbered 3-column review contact sheet PNG for fast visual QA.
- Checks aspect ratio against 16:9 and warns if an image risks edge cropping.
"""

from __future__ import annotations

import argparse
import math
import re
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches


SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}
DEFAULT_OUTPUT = "defense_presentation.pptx"
SLIDE_HEADING_RE = re.compile(r"^##\s+第\s*(\d+)\s*页[：:].*$", re.MULTILINE)
SPEAKER_NOTES_RE = re.compile(
    r"^###\s*演讲备注\s*\n(.*?)(?=^###\s|\Z)", re.MULTILINE | re.DOTALL
)


def natural_key(path: Path) -> list[object]:
    parts = re.split(r"(\d+)", path.name.lower())
    return [int(part) if part.isdigit() else part for part in parts]


def collect_images(image_dir: Path) -> list[Path]:
    if not image_dir.exists():
        raise FileNotFoundError(f"Image directory does not exist: {image_dir}")
    if not image_dir.is_dir():
        raise NotADirectoryError(f"Image path is not a directory: {image_dir}")

    images = [
        path
        for path in image_dir.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    ]
    if not images:
        supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        raise ValueError(f"No slide images found in {image_dir}. Supported formats: {supported}")

    return sorted(images, key=natural_key)


def parse_speaker_notes(outline_path: Path | None) -> dict[int, str]:
    if outline_path is None or not outline_path.is_file():
        return {}

    text = outline_path.read_text(encoding="utf-8")
    matches = list(SLIDE_HEADING_RE.finditer(text))
    notes: dict[int, str] = {}
    for idx, match in enumerate(matches):
        slide_num = int(match.group(1))
        start = match.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
        section = text[start:end]
        note_match = SPEAKER_NOTES_RE.search(section)
        if note_match:
            cleaned = re.sub(r"\n{3,}", "\n\n", note_match.group(1).strip())
            if cleaned:
                notes[slide_num] = cleaned
    return notes


def normalized_image_path(path: Path, temp_dir: Path) -> Path:
    if path.suffix.lower() != ".webp":
        return path

    output = temp_dir / f"{path.stem}.png"
    with Image.open(path) as image:
        image.convert("RGBA").save(output, "PNG")
    return output


def image_size(path: Path) -> tuple[int, int]:
    with Image.open(path) as image:
        width, height = image.size
    if width <= 0 or height <= 0:
        raise ValueError(f"Invalid image dimensions for {path}: {width}x{height}")
    return width, height


def add_full_slide_image(
    slide, image_path: Path, slide_width: int, slide_height: int, slide_index: int
) -> None:
    width_px, height_px = image_size(image_path)
    slide_ratio = slide_width / slide_height
    image_ratio = width_px / height_px
    drift = abs(image_ratio - slide_ratio) / slide_ratio

    if drift <= 0.015:
        # Exact or near-exact 16:9: stretch cleanly to exact slide bounds without clipping edges.
        left, top, pic_width, pic_height = 0, 0, slide_width, slide_height
    else:
        print(
            f"WARNING: slide {slide_index:02d} ({image_path.name}) aspect ratio "
            f"{width_px}x{height_px} ({image_ratio:.3f}) differs from 16:9 ({slide_ratio:.3f}) "
            f"by {drift * 100:.1f}%; edge cropping may occur."
        )
        if image_ratio >= slide_ratio:
            pic_height = slide_height
            pic_width = int(slide_height * image_ratio)
            left = int((slide_width - pic_width) / 2)
            top = 0
        else:
            pic_width = slide_width
            pic_height = int(slide_width / image_ratio)
            left = 0
            top = int((slide_height - pic_height) / 2)

    slide.shapes.add_picture(
        str(image_path),
        left,
        top,
        width=pic_width,
        height=pic_height,
    )


def build_contact_sheet(images: list[Path], out_path: Path, cols: int = 3) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    tile_w, tile_h = 640, 360
    pad = 16
    label_h = 26
    rows = math.ceil(len(images) / cols)
    canvas_w = cols * tile_w + (cols + 1) * pad
    canvas_h = rows * (tile_h + label_h) + (rows + 1) * pad

    sheet = Image.new("RGB", (canvas_w, canvas_h), (236, 240, 246))
    draw = ImageDraw.Draw(sheet)

    for idx, img_path in enumerate(images, start=1):
        row, col = divmod(idx - 1, cols)
        x = pad + col * (tile_w + pad)
        y = pad + row * (tile_h + label_h + pad)

        with Image.open(img_path) as im:
            orig_w, orig_h = im.size
            thumb = im.convert("RGB").resize((tile_w, tile_h), Image.Resampling.LANCZOS)

        sheet.paste(thumb, (x, y))
        draw.rectangle([x, y, x + tile_w - 1, y + tile_h - 1], outline=(180, 192, 208), width=1)
        label = f"Slide {idx:02d}/{len(images):02d}  ({orig_w}x{orig_h})  {img_path.name}"
        draw.text((x + 4, y + tile_h + 6), label, fill=(15, 44, 89))

    sheet.save(out_path, "PNG")


def build_ppt_from_images(
    image_dir: Path,
    output: Path,
    outline: Path | None = None,
    contact_sheet: Path | None = None,
) -> tuple[int, int]:
    images = collect_images(image_dir)
    speaker_notes = parse_speaker_notes(outline)

    presentation = Presentation()
    presentation.slide_width = Inches(13.333333)
    presentation.slide_height = Inches(7.5)
    blank_layout = presentation.slide_layouts[6]

    output.parent.mkdir(parents=True, exist_ok=True)
    notes_attached = 0

    with tempfile.TemporaryDirectory(prefix="ppt_slide_images_") as temp_root:
        temp_dir = Path(temp_root)
        for idx, image in enumerate(images, start=1):
            slide = presentation.slides.add_slide(blank_layout)
            normalized = normalized_image_path(image, temp_dir)
            add_full_slide_image(
                slide,
                normalized,
                presentation.slide_width,
                presentation.slide_height,
                idx,
            )
            note_text = speaker_notes.get(idx)
            if note_text:
                notes_slide = slide.notes_slide
                notes_slide.notes_text_frame.text = note_text
                notes_attached += 1

    presentation.save(output)

    if contact_sheet is not None:
        build_contact_sheet(images, contact_sheet)

    return len(images), notes_attached


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create a 16:9 strict image-only PowerPoint deck from full-slide images."
    )
    parser.add_argument(
        "--image-dir",
        required=True,
        type=Path,
        help="Directory containing slide images in png, jpg, jpeg, or webp format.",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT,
        type=Path,
        help=f"Output PPTX path. Defaults to {DEFAULT_OUTPUT}.",
    )
    parser.add_argument(
        "--outline",
        default=None,
        type=Path,
        help="Optional ppt_outline.md path; attaches '### 演讲备注' to PPTX speaker notes.",
    )
    parser.add_argument(
        "--contact-sheet",
        default=None,
        type=Path,
        help="Optional output PNG path for a numbered 3-column review contact sheet.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        slide_count, notes_count = build_ppt_from_images(
            args.image_dir,
            args.output,
            outline=args.outline,
            contact_sheet=args.contact_sheet,
        )
    except Exception as exc:
        print(f"ERROR: {exc}")
        return 1

    print(f"Created PowerPoint deck: {args.output} ({slide_count} slide(s))")
    if args.outline is not None:
        print(f"Attached speaker notes to {notes_count}/{slide_count} slide(s)")
    if args.contact_sheet is not None:
        print(f"Created visual QA contact sheet: {args.contact_sheet}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
