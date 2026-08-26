#!/usr/bin/env python3
"""Validate required outputs for the academic defense PPT workflow."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


SLIDE_HEADING_RE = re.compile(r"^##\s+第\s*(\d+)\s*页[：:]", re.MULTILINE)
ALLOWED_METHODS = {"built_in_image_gen", "imagegen_cli_fallback"}
FORBIDDEN_FINAL_PROMPT_TERMS = (
    "text-free",
    "no readable words",
    "background template",
    "blank areas",
    "editable text",
    "editable chinese text",
    "text to be added later",
)
XML_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def check_file(path: Path, label: str) -> bool:
    if path.is_file():
        print(f"OK: {label} exists: {path}")
        return True
    print(f"MISSING: {label}: {path}")
    return False


def count_outline_slides(path: Path) -> int | None:
    if not path.is_file():
        print(f"MISSING: ppt_outline.md: {path}")
        return None

    text = path.read_text(encoding="utf-8")
    numbers = [int(match.group(1)) for match in SLIDE_HEADING_RE.finditer(text)]
    if not numbers:
        print(f"ERROR: no slide headings found in {path}")
        return None

    expected = list(range(1, len(numbers) + 1))
    if numbers != expected:
        print(f"ERROR: slide headings are not consecutive from 1: {numbers[:12]}")
        return None

    print(f"OK: ppt_outline.md contains {len(numbers)} slide(s)")
    return len(numbers)


def resolve_manifest_path(
    project_root: Path, output_dir: Path, value: object
) -> Path | None:
    if not isinstance(value, str) or not value:
        return None
    raw_path = Path(value).expanduser()
    candidates = (
        [raw_path.resolve()]
        if raw_path.is_absolute()
        else [(project_root / raw_path).resolve(), (output_dir / raw_path).resolve()]
    )
    for path in candidates:
        try:
            path.relative_to(project_root)
        except ValueError:
            continue
        if path.is_file():
            return path
    return None


def manifest_path_ok(project_root: Path, output_dir: Path, value: object) -> bool:
    path = resolve_manifest_path(project_root, output_dir, value)
    return bool(path and path.is_file())


def load_manifest(output_dir: Path) -> dict[str, object] | None:
    manifest_file = output_dir / "imagegen_manifest.json"
    if not manifest_file.is_file():
        print(f"MISSING: imagegen_manifest.json: {manifest_file}")
        return None

    try:
        manifest = json.loads(manifest_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        print(f"ERROR: invalid imagegen_manifest.json: {exc}")
        return None

    if not isinstance(manifest, dict):
        print("ERROR: imagegen_manifest.json must contain a JSON object")
        return None
    return manifest


def check_imagegen_manifest(
    project_root: Path,
    output_dir: Path,
    manifest: dict[str, object] | None,
    expected_count: int | None,
) -> bool:
    if manifest is None:
        return False

    ok = True
    method = manifest.get("generation_method")
    if method not in ALLOWED_METHODS:
        print(
            "ERROR: imagegen_manifest.json generation_method must be one of "
            f"{sorted(ALLOWED_METHODS)!r}, found {method!r}"
        )
        ok = False
    else:
        print(f"OK: generation_method is {method}")

    if method == "imagegen_cli_fallback":
        for key in ("model", "quality", "size"):
            if not isinstance(manifest.get(key), str) or not manifest.get(key):
                print(f"ERROR: CLI fallback manifest must include non-empty {key!r}")
                ok = False

    preview = manifest.get("visual_preview")
    if not manifest_path_ok(project_root, output_dir, preview):
        print(f"ERROR: visual_preview in manifest is missing or invalid: {preview!r}")
        ok = False
    else:
        print(f"OK: visual preview exists: {preview}")

    slides = manifest.get("slides")
    if not isinstance(slides, list) or not slides:
        print("ERROR: imagegen_manifest.json must contain a non-empty slides list")
        return False

    if expected_count is not None and len(slides) != expected_count:
        print(
            f"ERROR: manifest slide count mismatch: {len(slides)} manifest item(s), "
            f"{expected_count} outline slide(s)"
        )
        ok = False

    for expected_number, item in enumerate(slides, start=1):
        if not isinstance(item, dict):
            print(f"ERROR: manifest slide entry {expected_number} is not an object")
            ok = False
            continue

        slide_number = item.get("slide")
        file_value = item.get("file")
        item_method = item.get("method")

        if slide_number != expected_number:
            print(
                f"ERROR: manifest slide order mismatch at entry {expected_number}: "
                f"found slide {slide_number!r}"
            )
            ok = False
        if item_method not in ALLOWED_METHODS:
            print(
                f"ERROR: manifest slide {slide_number!r} method must be one of "
                f"{sorted(ALLOWED_METHODS)!r}, found {item_method!r}"
            )
            ok = False
        elif method in ALLOWED_METHODS and item_method != method:
            print(
                f"ERROR: manifest slide {slide_number!r} method {item_method!r} "
                f"does not match generation_method {method!r}"
            )
            ok = False
        if not manifest_path_ok(project_root, output_dir, file_value):
            print(f"ERROR: manifest slide file is missing or invalid: {file_value!r}")
            ok = False
        elif isinstance(file_value, str):
            expected_stem = f"slide_{expected_number:02d}"
            if Path(file_value).stem != expected_stem:
                print(
                    f"ERROR: manifest slide {slide_number!r} file must be named "
                    f"{expected_stem}.*, found {file_value!r}"
                )
                ok = False
        if method == "imagegen_cli_fallback":
            prompt = item.get("prompt")
            if not isinstance(prompt, str) or not prompt.strip():
                print(
                    f"ERROR: CLI fallback manifest slide {slide_number!r} "
                    "must include the prompt used for that page"
                )
                ok = False
            elif any(term in prompt.lower() for term in FORBIDDEN_FINAL_PROMPT_TERMS):
                print(
                    f"ERROR: manifest slide {slide_number!r} prompt describes a "
                    "background/template instead of a complete final slide image"
                )
                ok = False

    if ok:
        print(f"OK: imagegen manifest validates {len(slides)} slide image(s)")
    return ok


def slide_xml_sort_key(name: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", name)
    return int(match.group(1)) if match else 0


def check_image_only_pptx(path: Path, expected_count: int | None) -> bool:
    if not path.is_file():
        print(f"MISSING: defense_presentation.pptx: {path}")
        return False

    ok = True
    try:
        with zipfile.ZipFile(path) as archive:
            slide_names = sorted(
                [
                    name
                    for name in archive.namelist()
                    if re.match(r"ppt/slides/slide\d+\.xml$", name)
                ],
                key=slide_xml_sort_key,
            )
            if expected_count is not None and len(slide_names) != expected_count:
                print(
                    f"ERROR: PPTX slide count mismatch: {len(slide_names)} slide XML file(s), "
                    f"{expected_count} outline slide(s)"
                )
                ok = False

            for index, slide_name in enumerate(slide_names, start=1):
                root = ET.fromstring(archive.read(slide_name))
                visible_text = [
                    node.text.strip()
                    for node in root.findall(".//a:t", XML_NS)
                    if node.text and node.text.strip()
                ]
                pictures = root.findall(".//p:pic", XML_NS)

                if visible_text:
                    preview = " / ".join(visible_text[:3])
                    print(
                        f"ERROR: PPTX slide {index} contains editable text; "
                        f"final decks must be image-only. Text preview: {preview!r}"
                    )
                    ok = False
                if len(pictures) != 1:
                    print(
                        f"ERROR: PPTX slide {index} must contain exactly one full-slide image; "
                        f"found {len(pictures)} picture object(s)"
                    )
                    ok = False
    except (zipfile.BadZipFile, ET.ParseError) as exc:
        print(f"ERROR: could not inspect PPTX image-only structure: {exc}")
        return False

    if ok:
        print("OK: defense_presentation.pptx is strict image-only: one image per slide, no editable text")
    return ok


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate outputs from the academic defense PPT workflow."
    )
    parser.add_argument(
        "--base-dir",
        default=".",
        type=Path,
        help=(
            "Project directory or output directory containing ppt_outline.md, "
            "defense_presentation.pptx, and imagegen_manifest.json."
        ),
    )
    return parser.parse_args()


def locate_output_dir(base_dir: Path) -> tuple[Path, Path]:
    """Return (project_root, output_dir) for root-level or outputs/ layouts."""

    required = ("ppt_outline.md", "defense_presentation.pptx", "imagegen_manifest.json")
    nested_output_dir = base_dir / "outputs"
    direct_score = sum((base_dir / name).is_file() for name in required)
    nested_score = sum((nested_output_dir / name).is_file() for name in required)

    if nested_score > direct_score:
        return base_dir, nested_output_dir
    if base_dir.name == "outputs":
        return base_dir.parent, base_dir
    return base_dir, base_dir


def main() -> int:
    args = parse_args()
    base_dir = args.base_dir.resolve()
    project_root, output_dir = locate_output_dir(base_dir)
    outline_count = count_outline_slides(output_dir / "ppt_outline.md")
    manifest = load_manifest(output_dir)

    checks = [
        outline_count is not None,
        check_file(output_dir / "defense_presentation.pptx", "defense_presentation.pptx"),
        check_imagegen_manifest(project_root, output_dir, manifest, outline_count),
        check_image_only_pptx(output_dir / "defense_presentation.pptx", outline_count),
    ]

    if all(checks):
        print("VALIDATION PASSED")
        return 0

    print("VALIDATION FAILED")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
