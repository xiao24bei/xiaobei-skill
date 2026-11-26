#!/usr/bin/env python3
"""Validate outputs, Stage-1 outline academic rigor, equations, provenance, and PPTX structure (v5.0).

Key v5.0 Advisor-Rigor Gates:
- `--outline-only` mode for mandatory Stage-1 gating BEFORE spending imagegen budget in Stages 2-3.
- Pre-outline Evidence Ledger checks (Equation Mapping Ledger + Baseline/Protocol Ledger).
- Per-slide substantive `上屏文案` density check (rejects <115-char 3-bullet flashcard shells!).
- Cited-Equation Presence Gate (if a slide cites `式(6)-(8)` in rationale/evidence, `上屏文案` MUST
  actually contain genuine mathematical equations rather than omitting them or only writing `T = 1000`).
- Textbook Metric Formula Padding Ban (blocks `PSNR = 10 log10(MAX^2/MSE)` / `FID = ||mu_r - mu_g||^2`
  filler formulas on experimental/ablation slides and prompts).
- Verbatim Banner/Card Echo Detection & Generic AI Cliché Warnings.
"""

from __future__ import annotations

import argparse
import json
import re
import struct
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


SLIDE_HEADING_RE = re.compile(r"^##\s+第\s*(\d+)\s*页[：:].*$", re.MULTILINE)
ON_SCREEN_RE = re.compile(r"^###\s*上屏文案\s*\n(.*?)(?=^###\s|\Z)", re.MULTILINE | re.DOTALL)
LAYOUT_RE = re.compile(r"^###\s*版式原型\s*\n([^\n]+)", re.MULTILINE)
TITLE_RE = re.compile(r"^###\s*页面标题\s*\n([^\n]+)", re.MULTILINE)
SPEAKER_NOTES_RE = re.compile(r"^###\s*演讲备注\s*\n(.*?)(?=^###\s|\Z)", re.MULTILINE | re.DOTALL)
SECTION_NUM_RE = re.compile(r"^\s*(?:\d+\.\d+|\d+\s*[|｜、.]|第[一二三四五六七八九十]+章)")
LAYOUT_ID_RE = re.compile(r"\b(L\d{2}[a-c]?)\b", re.IGNORECASE)

CITED_EQ_RE = re.compile(
    r"(?:式|公式|Eq(?:uation)?s?\.?)\s*[\(（]?\s*\d+",
    re.IGNORECASE,
)
SCALAR_ONLY_EQ_RE = re.compile(
    r"^[A-Za-zα-ωΑ-Ω_0-9\s]{1,12}=\s*[-+]?\d+(?:\.\d+)?(?:e[-+]?\d+)?\s*$",
    re.IGNORECASE,
)
MATH_OPERATOR_HINTS = (
    "+",
    "−",
    "-",
    "⊙",
    "⊗",
    "∑",
    "∫",
    "√",
    "\\sqrt",
    "\\frac",
    "\\mathcal",
    "\\mathbb",
    "\\nabla",
    "∇",
    "‖",
    "||",
    "\\|",
    "Softmax",
    "softmax",
    "Attention",
    "log",
    "exp",
    "_",
    "^",
    "∼",
    "~",
    "∈",
)

TEXTBOOK_METRIC_FORMULA_PATTERNS = (
    (
        "PSNR textbook definition",
        re.compile(r"10\s*\\?log(?:_?\{?10\}?)?\s*[\(\[（].{0,40}(?:MAX|255).{0,40}MSE", re.IGNORECASE),
    ),
    (
        "FID textbook definition",
        re.compile(r"(?:μ_r|mu_r)\s*[-−]\s*(?:μ_g|mu_g).{0,80}(?:Tr|trace)\s*[\(\[（]", re.IGNORECASE),
    ),
    (
        "SSIM textbook definition",
        re.compile(r"\(\s*2\s*\\?mu_x\s*\\?mu_y\s*\+\s*C_?1\s*\)", re.IGNORECASE),
    ),
)

AI_CLICHE_PATTERNS = (
    "提升特征一致性",
    "增强全局建模能力",
    "进一步提升重建精度",
    "增强结构约束",
    "提升模型表征能力",
    "有效提升模型性能",
)

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
META_LEAK_PATTERNS = (
    re.compile(r"原文未明确说明"),
    re.compile(r"原文位置未明确说明"),
    re.compile(r"原文第\s*\d+\s*页"),
    re.compile(r"基于原文归纳"),
)
ANCHOR_INSTITUTION_NAMES = (
    "南京审计大学",
    "北京交通大学",
    "NIH/NIDCD",
)
NEGATION_CLAUSE_RE = re.compile(
    r"(?:严禁|禁止|绝不|绝对不|不得|不要|不含|避免|勿|never|do not|must not|without)[^。\n；;！!]{0,180}",
    re.IGNORECASE,
)
AESTHETIC_DRIFT_TERMS = (
    "photorealistic background",
    "cinematic photo",
    "3d commercial render",
    "glossy neon",
    "stock photo background",
    "mint green pill",
    "薄荷绿胶囊",
)
STRUCTURAL_LAYOUTS = {"L01", "L02A", "L02B", "L02C", "L10B"}

XML_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def strip_negation_clauses(text: str) -> str:
    return NEGATION_CLAUSE_RE.sub(" ", text)


def clean_onscreen_for_length(raw: str) -> str:
    """Strip markdown scaffolding and color-tag wrappers to measure real substantive content."""
    lines = []
    for line in raw.splitlines():
        s = line.strip()
        if not s or s.startswith(">"):
            continue
        # Remove leading bullet markers and boilerplate field prefixes
        s = re.sub(r"^[-*•]\s*", "", s)
        s = re.sub(
            r"^(?:顶栏核心结论横幅|顶栏结论横幅|结构化卡片\s*\d+|模块/剖析卡\s*\d+|卡\s*\d+|左卡|中卡|右卡|底部总结条(?:与角标)?|左下角标)\s*(?:[\(（][^\)）]{0,40}[\)）])?\s*[：:]\s*",
            "",
            s,
        )
        # Unwrap 【红：...】 / 【蓝：...】 keeping the inner content
        s = re.sub(r"【(?:红|蓝|加粗|重点)[：:]\s*([^】]+)】", r"\1", s)
        lines.append(s)
    joined = "".join(lines)
    return re.sub(r"\s+", "", joined)


def has_genuine_math_formula(text: str) -> bool:
    """Return True if `text` contains a substantive mathematical equation (not just `T = 1000`)."""
    # Check explicit $...$ math blocks containing = or ~
    for m in re.finditer(r"\$?\$([^$]+)\$?\$", text):
        expr = m.group(1).strip()
        if any(rel in expr for rel in ("=", "≈", "∼", "~", "∈", "→")) and not SCALAR_ONLY_EQ_RE.match(expr):
            if any(op in expr for op in MATH_OPERATOR_HINTS):
                return True

    # Also check plain-text mathematical lines (e.g. Unicode formulas without $)
    for line in text.splitlines():
        s = line.strip()
        if "=" in s or "∼" in s:
            # Check around '='
            for seg in re.split(r"[；;，,。]", s):
                seg_clean = seg.strip()
                if "=" in seg_clean and not SCALAR_ONLY_EQ_RE.match(seg_clean):
                    if sum(1 for op in MATH_OPERATOR_HINTS if op in seg_clean) >= 1 and re.search(
                        r"[A-Za-zα-ωΑ-Ωℒ∇ϵεΨΦ]", seg_clean
                    ):
                        return True
    return False


def find_verbatim_echoes(title: str, raw_onscreen: str) -> list[str]:
    """Detect verbatim duplicate clauses (>=12 chars) repeated across banner and body cards."""
    clauses: list[str] = []
    for src in [title, *raw_onscreen.splitlines()]:
        s = src.strip()
        if not s or s.startswith(">"):
            continue
        s = re.sub(r"^[-*•]\s*", "", s)
        s = re.sub(r"^[^：:\n]{2,18}[：:]\s*", "", s)
        s = re.sub(r"【(?:红|蓝)[：:]\s*([^】]+)】", r"\1", s)
        norm = re.sub(r"[\s，,。；;：:！!？?·\"'“”‘’\(\)（）]", "", s)
        if len(norm) >= 12:
            clauses.append(norm)

    seen: set[str] = set()
    dups: list[str] = []
    for c in clauses:
        if c in seen and c not in dups:
            dups.append(c)
        seen.add(c)
    return dups


def read_image_dimensions(path: Path) -> tuple[int, int] | None:
    try:
        from PIL import Image  # type: ignore

        with Image.open(path) as im:
            return im.size
    except Exception:
        pass

    data = path.read_bytes()
    if len(data) >= 24 and data[:8] == b"\x89PNG\r\n\x1a\n":
        w, h = struct.unpack(">II", data[16:24])
        return (w, h)
    if len(data) >= 4 and data[:2] == b"\xff\xd8":
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xC0, 0xC1, 0xC2):
                h, w = struct.unpack(">HH", data[i + 5 : i + 9])
                return (w, h)
            if marker in (0xD8, 0xD9, 0x01) or (0xD0 <= marker <= 0xD7):
                i += 2
            else:
                seg_len = struct.unpack(">H", data[i + 2 : i + 4])[0]
                i += 2 + seg_len
    return None


def check_file(path: Path, label: str) -> bool:
    if path.is_file():
        print(f"OK: {label} exists: {path}")
        return True
    print(f"MISSING: {label}: {path}")
    return False


def inspect_outline(
    path: Path, base_dir: Path, warnings: list[str]
) -> tuple[int | None, bool, str]:
    if not path.is_file():
        print(f"MISSING: ppt_outline.md: {path}")
        return None, False, ""

    text = path.read_text(encoding="utf-8")
    matches = list(SLIDE_HEADING_RE.finditer(text))
    if not matches:
        print(f"ERROR: no slide headings found in {path}")
        return None, False, text

    numbers = [int(m.group(1)) for m in matches]
    expected = list(range(1, len(numbers) + 1))
    if numbers != expected:
        print(f"ERROR: slide headings are not consecutive from 1: {numbers[:12]}")
        return None, False, text

    ok = True
    is_short_mode = any(k in text[:800] for k in ("单图", "极短", "精简版", "简短材料", "短笔记"))
    preamble = text[: matches[0].start()]

    # Check Stage-1 Evidence Ledgers in preamble
    inv_path = base_dir / "source_assets" / "inventory.json"
    source_eq_count = 0
    if inv_path.is_file():
        try:
            inv_data = json.loads(inv_path.read_text(encoding="utf-8"))
            if isinstance(inv_data, dict):
                source_eq_count = len(inv_data.get("numbered_equations") or [])
        except Exception:
            source_eq_count = 0

    if not is_short_mode:
        # Exclude the LOCKED STYLE BIBLE block when checking whether Stage-1 ledgers exist
        ledger_region = re.split(r"##\s*(?:2\.\s*)?全套视觉", preamble)[0]
        has_eq_ledger = any(
            k in ledger_region for k in ("公式与符号", "公式台账", "公式映射", "式(1)", "式（1）", "Equation")
        )
        has_baseline_ledger = any(
            k in ledger_region for k in ("基线", "技术路线", "冻结", "训练/推理")
        )
        if source_eq_count >= 2 and not has_eq_ledger:
            print(
                f"ERROR: source paper has {source_eq_count} numbered equations, but ppt_outline.md "
                "preamble lacks a Core Equation Mapping Ledger ('1.2 核心编号公式与符号上屏映射台账')."
            )
            ok = False
        elif not has_eq_ledger:
            warnings.append(
                "ppt_outline.md preamble does not include an explicit Equation Mapping Ledger; "
                "ensure all core paper formulas are mapped onto method slides."
            )
        if not has_baseline_ledger:
            warnings.append(
                "ppt_outline.md preamble lacks a Baseline Positioning & Train/Infer Protocol Ledger "
                "('1.3 既有代表基线卡点与本文训练/推理协议台账')."
            )

    layouts: list[tuple[int, str]] = []
    content_titles_total = 0
    content_titles_numbered = 0
    question_title_slides: list[int] = []
    shallow_slides: list[tuple[int, int]] = []
    content_lengths: list[int] = []
    has_positioning_slide = False

    for idx, match in enumerate(matches):
        slide_num = int(match.group(1))
        start = match.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
        section = text[start:end]

        layout_id = ""
        layout_m = LAYOUT_RE.search(section)
        if layout_m:
            id_m = LAYOUT_ID_RE.search(layout_m.group(1))
            if id_m:
                layout_id = id_m.group(1).upper()
                layouts.append((slide_num, layout_id))

        title_m = TITLE_RE.search(section)
        raw_title = title_m.group(1).strip() if title_m else ""
        if title_m and layout_id not in STRUCTURAL_LAYOUTS and 2 < slide_num < len(numbers):
            content_titles_total += 1
            if SECTION_NUM_RE.match(raw_title):
                content_titles_numbered += 1
            if ("?" in raw_title or "？" in raw_title) and slide_num > 3:
                question_title_slides.append(slide_num)

        if slide_num in (3, 4, 5) and (
            layout_id == "L03B"
            or any(k in section for k in ("既有", "相关工作", "代表方法", "技术路线", "定位", "表1"))
        ):
            has_positioning_slide = True

        on_screen = ON_SCREEN_RE.search(section)
        if not on_screen:
            print(f"ERROR: ppt_outline.md slide {slide_num} is missing '### 上屏文案'")
            ok = False
            continue

        raw_onscreen = on_screen.group(1)
        visible_positive = strip_negation_clauses(raw_onscreen)

        # 1. Meta-leak check
        for pattern in META_LEAK_PATTERNS:
            if pattern.search(visible_positive):
                print(
                    f"ERROR: ppt_outline.md slide {slide_num} '上屏文案' leaks internal "
                    f"meta-text matching {pattern.pattern!r}; move audit notes to '原文依据'."
                )
                ok = False

        # 2. Textbook metric formula padding ban
        for label, tb_pat in TEXTBOOK_METRIC_FORMULA_PATTERNS:
            if tb_pat.search(visible_positive) and "新指标" not in section:
                print(
                    f"ERROR: ppt_outline.md slide {slide_num} '上屏文案' pads with {label}; "
                    "replace textbook metric definitions with the paper's own equations or multi-baseline analysis."
                )
                ok = False

        # 3. Verbatim echo check (banner vs card repetition on content slides)
        check_title = "" if layout_id in STRUCTURAL_LAYOUTS or slide_num == 1 else raw_title
        dups = find_verbatim_echoes(check_title, visible_positive)
        if dups:
            print(
                f"ERROR: ppt_outline.md slide {slide_num} repeats identical text across title/banner/cards "
                f"({dups[0][:24]!r}…); each layer must add distinct academic information."
            )
            ok = False

        # 4. Generic AI cliché check
        for cliche in AI_CLICHE_PATTERNS:
            if cliche in visible_positive:
                warnings.append(
                    f"Slide {slide_num:02d} '上屏文案' contains generic AI cliché {cliche!r}; "
                    "replace with concrete mechanism or metric-specific causal analysis."
                )

        # 5. Slide 2 TOC completeness check
        substantive_len = len(clean_onscreen_for_length(visible_positive))
        if slide_num == 2 and layout_id == "L02A" and not is_short_mode:
            if substantive_len < 85:
                print(
                    f"ERROR: ppt_outline.md slide 2 (TOC) '上屏文案' is too bare ({substantive_len} chars); "
                    "include each chapter's concrete scientific sub-question/method anchor and page range (P03–P04)."
                )
                ok = False

        # 6. Core content slide density gate (blocks 40-80 char flashcard shells!)
        is_core_content = (
            layout_id not in STRUCTURAL_LAYOUTS and 2 < slide_num < len(numbers)
        )
        if is_core_content and not is_short_mode:
            content_lengths.append(substantive_len)
            min_chars = 95 if slide_num == len(numbers) - 1 else 115
            if substantive_len < min_chars:
                shallow_slides.append((slide_num, substantive_len))

        # 7. Cited-equation presence gate (blocks citing 式(6)-(8) while rendering 0 formulas!)
        non_onscreen = section[: on_screen.start()] + section[on_screen.end() :]
        if is_core_content and CITED_EQ_RE.search(non_onscreen):
            if not has_genuine_math_formula(visible_positive):
                print(
                    f"ERROR: ppt_outline.md slide {slide_num} cites numbered paper equation(s) in "
                    "rationale/evidence, yet '上屏文案' contains NO genuine mathematical equation "
                    "(only prose or a bare scalar!). Write the actual paper equation(s) on-screen."
                )
                ok = False

        # 8. Speaker notes presence check
        notes_m = SPEAKER_NOTES_RE.search(section)
        if not notes_m or len(notes_m.group(1).strip()) < 45:
            warnings.append(
                f"Slide {slide_num:02d} speaker notes ('演讲备注') are missing or too brief (<45 chars); "
                "recommended 130-220 chars of natural oral explanation."
            )

    if shallow_slides:
        formatted = ", ".join(f"P{s:02d}({n} chars)" for s, n in shallow_slides)
        print(
            f"ERROR: {len(shallow_slides)} core content slide(s) in ppt_outline.md are too shallow/skeletal "
            f"(minimum 115 substantive chars, target 140-260 chars + formulas/data): {formatted}"
        )
        ok = False

    if content_lengths and not is_short_mode:
        avg_len = sum(content_lengths) / len(content_lengths)
        if avg_len < 140:
            print(
                f"ERROR: average core content slide '上屏文案' length is only {avg_len:.0f} chars "
                "(minimum average 140 chars, target 160-240 chars). Expand equations, tensor flows, "
                "multi-baseline metrics, and WHY-attribution!"
            )
            ok = False
        else:
            print(
                f"OK: core content density passed (avg {avg_len:.0f} substantive chars across "
                f"{len(content_lengths)} content slides)"
            )

    if not is_short_mode and len(numbers) < 15 and any(
        k in text for k in ("组会", "精读", "开题", "毕业", "答辩", "基金")
    ):
        if len(numbers) < 12:
            print(
                f"ERROR: full academic deck has only {len(numbers)} slides (minimum 15 recommended, "
                "hard floor 12)."
            )
            ok = False
        else:
            warnings.append(
                f"Deck has {len(numbers)} slides; v5.0 recommends 16-18 slides (min 15) for full papers "
                "so related-work positioning (1.2), complete formulas, and multi-baseline trade-offs fit cleanly."
            )

    if not is_short_mode and len(numbers) >= 14 and not has_positioning_slide:
        warnings.append(
            "No dedicated related-work / SOTA baseline positioning slide detected in slides 3-5 "
            "(recommended '1.2 既有技术路线局限与本文定位' using layout L03b)."
        )

    if content_titles_total >= 4 and content_titles_numbered < content_titles_total * 0.6:
        warnings.append(
            f"Only {content_titles_numbered}/{content_titles_total} content slide titles use "
            "academic section numbering (e.g., '1.1', '2.1')."
        )

    if question_title_slides:
        warnings.append(
            f"Slides {question_title_slides} use question marks in method/result/conclusion titles; "
            "switch to authoritative assertion titles."
        )

    if layouts:
        content_sequence = [(s, lid) for s, lid in layouts if lid not in STRUCTURAL_LAYOUTS]
        for (s1, l1), (s2, l2) in zip(layouts, layouts[1:]):
            if l1 == l2 and l1 not in {"L02B"}:
                warnings.append(f"Adjacent slides {s1} and {s2} share identical layout {l1}.")
        distinct_content = {lid for _, lid in content_sequence}
        if len(content_sequence) >= 10 and len(distinct_content) < 6:
            warnings.append(
                f"Low layout diversity: only {len(distinct_content)} distinct content layout(s) "
                f"({sorted(distinct_content)}); recommended >= 6."
            )

    print(f"OK: ppt_outline.md parsed ({len(numbers)} slide(s))")
    return len(numbers), ok, text


def resolve_manifest_path(base_dir: Path, value: object) -> Path | None:
    if not isinstance(value, str) or not value:
        return None
    path = (base_dir / value).resolve()
    try:
        path.relative_to(base_dir)
    except ValueError:
        return None
    return path


def load_manifest(base_dir: Path) -> dict[str, object] | None:
    manifest_file = base_dir / "imagegen_manifest.json"
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


def prompt_has_page_number(prompt: str, page_num: int, total_slides: int) -> bool:
    tokens = (
        f"{page_num:02d}/{total_slides:02d}",
        f"{page_num:02d}/{total_slides}",
        f"{page_num}/{total_slides}",
        f"页码“{page_num:02d}”",
        f'页码"{page_num:02d}"',
        f"页码 {page_num:02d}",
        f"白字“{page_num:02d}”",
        f"居中“{page_num:02d}”",
    )
    return any(tok in prompt for tok in tokens)


def check_imagegen_manifest(
    base_dir: Path,
    manifest: dict[str, object] | None,
    expected_count: int | None,
    outline_text: str,
    warnings: list[str],
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

    preview_method = manifest.get("visual_preview_method")
    if preview_method is not None and preview_method not in ALLOWED_METHODS:
        print(f"ERROR: invalid visual_preview_method: {preview_method!r}")
        ok = False

    preview = manifest.get("visual_preview")
    preview_path = resolve_manifest_path(base_dir, preview)
    if not (preview_path and preview_path.is_file()):
        print(f"ERROR: visual_preview in manifest is missing or invalid: {preview!r}")
        ok = False
    else:
        dims = read_image_dimensions(preview_path)
        if dims is None:
            print(f"ERROR: visual_preview is not a readable image: {preview!r}")
            ok = False
        else:
            print(f"OK: visual preview exists: {preview} ({dims[0]}x{dims[1]})")

    slides = manifest.get("slides")
    if not isinstance(slides, list) or not slides:
        print("ERROR: imagegen_manifest.json must contain a non-empty slides list")
        return False

    total_slides = len(slides)
    if expected_count is not None and total_slides != expected_count:
        print(
            f"ERROR: manifest slide count mismatch: {total_slides} manifest item(s), "
            f"{expected_count} outline slide(s)"
        )
        ok = False

    target_ratio = 16.0 / 9.0
    outline_positive = strip_negation_clauses(outline_text)
    embedded_fig_slides = 0

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
        elif expected_number == 1 and method in ALLOWED_METHODS and item_method != method:
            print(
                f"ERROR: manifest slide {slide_number!r} method {item_method!r} "
                f"does not match generation_method {method!r}"
            )
            ok = False

        embedded = item.get("embedded_source_figures")
        if embedded is not None:
            if not isinstance(embedded, list):
                print(f"ERROR: manifest slide {slide_number!r} embedded_source_figures must be a list")
                ok = False
            elif embedded:
                embedded_fig_slides += 1
                for fig_rel in embedded:
                    fig_path = resolve_manifest_path(base_dir, fig_rel)
                    if not (fig_path and fig_path.is_file()):
                        print(
                            f"ERROR: manifest slide {slide_number!r} embedded source figure "
                            f"missing or invalid: {fig_rel!r}"
                        )
                        ok = False

        slide_path = resolve_manifest_path(base_dir, file_value)
        if not (slide_path and slide_path.is_file()):
            print(f"ERROR: manifest slide file is missing or invalid: {file_value!r}")
            ok = False
        else:
            if isinstance(file_value, str):
                expected_stem = f"slide_{expected_number:02d}"
                if Path(file_value).stem != expected_stem:
                    print(
                        f"ERROR: manifest slide {slide_number!r} file must be named "
                        f"{expected_stem}.*, found {file_value!r}"
                    )
                    ok = False
            dims = read_image_dimensions(slide_path)
            if dims is None:
                print(f"ERROR: manifest slide {slide_number!r} is not a valid image: {file_value!r}")
                ok = False
            else:
                w, h = dims
                if w < 960 or h < 540:
                    print(
                        f"ERROR: manifest slide {slide_number!r} resolution {w}x{h} is too low "
                        "(minimum 960x540)"
                    )
                    ok = False
                ratio = w / h if h else 0.0
                drift = abs(ratio - target_ratio) / target_ratio
                if drift > 0.08:
                    print(
                        f"ERROR: manifest slide {slide_number!r} aspect ratio {w}x{h} "
                        f"deviates {drift * 100:.1f}% from 16:9"
                    )
                    ok = False
                elif drift > 0.02:
                    warnings.append(
                        f"Slide {expected_number:02d} aspect ratio {w}x{h} deviates "
                        f"{drift * 100:.1f}% from 16:9."
                    )

        if item_method == "imagegen_cli_fallback":
            for key in ("model", "quality", "size"):
                value = item.get(key, manifest.get(key))
                if not isinstance(value, str) or not value.strip():
                    print(f"ERROR: CLI slide {slide_number!r} requires non-empty {key!r}")
                    ok = False

        prompt = item.get("prompt")
        if not isinstance(prompt, str) or not prompt.strip():
            print(f"ERROR: manifest slide {slide_number!r} must include its actual prompt")
            ok = False
        else:
            prompt_lower = prompt.lower()
            if any(term in prompt_lower for term in FORBIDDEN_FINAL_PROMPT_TERMS):
                print(
                    f"ERROR: manifest slide {slide_number!r} prompt describes a "
                    "background/template instead of a complete final slide image"
                )
                ok = False

            positive_prompt = strip_negation_clauses(prompt)
            for pattern in META_LEAK_PATTERNS:
                if pattern.search(positive_prompt):
                    print(
                        f"ERROR: manifest slide {slide_number!r} prompt appears to place internal "
                        f"audit text ({pattern.pattern}) onto the slide canvas"
                    )
                    ok = False

            for label, tb_pat in TEXTBOOK_METRIC_FORMULA_PATTERNS:
                if tb_pat.search(positive_prompt) and "新指标" not in outline_positive:
                    print(
                        f"ERROR: manifest slide {slide_number!r} prompt pads slide with {label} "
                        "instead of paper-specific formulas/analysis"
                    )
                    ok = False

            for school in ANCHOR_INSTITUTION_NAMES:
                if school in positive_prompt and school not in outline_positive:
                    print(
                        f"ERROR: manifest slide {slide_number!r} prompt appears to copy sample "
                        f"anchor institution {school!r} not present in ppt_outline.md"
                    )
                    ok = False

            positive_lower = positive_prompt.lower()
            for drift_term in AESTHETIC_DRIFT_TERMS:
                if drift_term in positive_lower:
                    warnings.append(
                        f"Slide {expected_number:02d} prompt contains aesthetic drift risk term "
                        f"{drift_term!r}."
                    )

            for cliche in AI_CLICHE_PATTERNS:
                if cliche in positive_prompt:
                    warnings.append(
                        f"Slide {expected_number:02d} prompt contains generic AI cliché {cliche!r}."
                    )

            if 1 < expected_number < total_slides:
                if not prompt_has_page_number(prompt, expected_number, total_slides):
                    warnings.append(
                        f"Slide {expected_number:02d} prompt does not explicitly specify page "
                        f"number ({expected_number:02d}/{total_slides:02d} or “{expected_number:02d}”)."
                    )

    figs_dir = base_dir / "source_assets" / "figures"
    if figs_dir.is_dir():
        available_figs = [
            p
            for p in figs_dir.iterdir()
            if p.is_file() and p.suffix.lower() in {".png", ".jpg", ".jpeg"}
        ]
        if available_figs and embedded_fig_slides == 0:
            warnings.append(
                f"Found {len(available_figs)} extracted paper figure(s) in source_assets/figures/ "
                "but 0 slides recorded embedded_source_figures; composite real paper figures via "
                "scripts/embed_paper_figure.py."
            )

    if ok:
        print(
            f"OK: imagegen manifest validates {len(slides)} slide image(s) "
            f"({embedded_fig_slides} slide(s) embed lossless real paper figure(s))"
        )
    return ok


def slide_xml_sort_key(name: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", name)
    return int(match.group(1)) if match else 0


def check_image_only_pptx(path: Path, expected_count: int | None) -> bool:
    if not path.is_file():
        print(f"MISSING: defense_presentation.pptx: {path}")
        return False

    ok = True
    notes_count = 0
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            slide_names = sorted(
                [name for name in names if re.match(r"ppt/slides/slide\d+\.xml$", name)],
                key=slide_xml_sort_key,
            )
            notes_names = [
                name for name in names if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", name)
            ]
            notes_count = len(notes_names)

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
        print(
            "OK: defense_presentation.pptx is strict image-only: one image per slide, "
            f"no editable visible text (speaker notes attached on {notes_count} slide(s))"
        )
    return ok


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate Stage-1 outline depth and Stage-3 PPT outputs (v5.0)."
    )
    parser.add_argument(
        "--base-dir",
        default=".",
        type=Path,
        help="Directory containing ppt_outline.md (and Stage-3 outputs when not --outline-only).",
    )
    parser.add_argument(
        "--outline-only",
        action="store_true",
        help="Validate Stage-1 ppt_outline.md depth, equations, and anti-echo gates only.",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Treat presentation quality warnings as errors.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    base_dir = args.base_dir.resolve()
    warnings: list[str] = []

    outline_count, outline_ok, outline_text = inspect_outline(
        base_dir / "ppt_outline.md", base_dir, warnings
    )

    if args.outline_only:
        checks = [outline_count is not None and outline_ok]
    else:
        manifest = load_manifest(base_dir)
        checks = [
            outline_count is not None and outline_ok,
            check_file(base_dir / "defense_presentation.pptx", "defense_presentation.pptx"),
            check_imagegen_manifest(base_dir, manifest, outline_count, outline_text, warnings),
            check_image_only_pptx(base_dir / "defense_presentation.pptx", outline_count),
        ]

    if warnings:
        print(f"\nQUALITY WARNINGS ({len(warnings)}):")
        for item in warnings:
            print(f"  WARN: {item}")

    if all(checks) and (not args.strict or not warnings):
        label = "STAGE-1 OUTLINE VALIDATION PASSED" if args.outline_only else "VALIDATION PASSED"
        print(f"\n{label}")
        return 0

    label = "STAGE-1 OUTLINE VALIDATION FAILED" if args.outline_only else "VALIDATION FAILED"
    print(f"\n{label}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
