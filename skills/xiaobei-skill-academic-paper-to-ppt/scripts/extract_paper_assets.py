#!/usr/bin/env python3
"""Extract full text, OMML/PDF formulas, structured tables, high-res pages, and figures from papers (v5.0).

Key v5.0 upgrades over v4.x:
- Full Word (.docx) OMML math (<m:oMath> / <m:oMathPara>) conversion into clean LaTeX-style notation
  (instead of silently dropping all Word equations).
- Automatic recognition of Chinese/English journal 2-column equation tables [ formula | (eq_no) ]
  and extraction of a Numbered Equations Catalog into inventory.json + figure_guide.md.
- Structured GitHub Markdown table preservation for .docx data tables (Tables 1..N) instead of
  flattening cells into unstructured lines.
- Automatic stitching of multi-panel .docx table figures (e.g., (a)..(h) comparison grids) with
  CJK-capable subcaptions into ready-to-embed composite PNGs under figures/.

Usage:
  1) Extract full paper assets (text + equations + tables + pages + figures + figure_guide.md):
     python3 extract_paper_assets.py --input paper.pdf --out-dir ./outputs/source_assets
  2) Crop any specific vector/raster figure cleanly from a rendered PDF page (auto-trims margins):
     python3 extract_paper_assets.py crop \
       --page ./outputs/source_assets/pages/page_03.png \
       --box 0.06,0.08,0.94,0.48 \
       --out ./outputs/source_assets/figures/fig_01.png
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


FIG_TABLE_RE = re.compile(
    r"(?:^|\n)\s*((?:Fig(?:ure)?\.?\s*\d+[a-zA-Z]?|图\s*\d+[a-zA-Z]?|Table\s*\d+|表\s*\d+)\s*[:：.\-—\s]\s*[^\n]{1,180})",
    re.IGNORECASE,
)
EQ_NUM_CELL_RE = re.compile(r"^\s*[\(（]\s*(\d+[a-zA-Z]?)\s*[\)）]\s*$")
TEXT_EQ_LINE_RE = re.compile(
    r"(?:^|\n)\s*([^\n]{4,220}?)\s*[\(（](\d{1,2}[a-zA-Z]?)[\)）]\s*(?=\n|$)"
)

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
V_NS = "{urn:schemas-microsoft-com:vml}"

ACCENT_MAP = {
    "^": r"\hat",
    "̂": r"\hat",
    "~": r"\tilde",
    "˜": r"\tilde",
    "˙": r"\dot",
    "¨": r"\ddot",
    "¯": r"\bar",
    "̅": r"\overline",
    "→": r"\vec",
    "⃗": r"\vec",
}


def detect_captions(text: str) -> list[str]:
    found: list[str] = []
    seen: set[str] = set()
    for match in FIG_TABLE_RE.finditer(text):
        line = re.sub(r"\s+", " ", match.group(1)).strip()
        if len(line) < 4 or len(line) > 180:
            continue
        # Skip body prose sentences like "图6展示了不同方法..." or "表2给出了..."
        if re.match(
            r"^(?:图|表|Fig(?:ure)?\.?|Table)\s*\d+\s*(?:展示|给出|显示|表明|说明|对比了|列出了|汇总了|中|的|可|如)",
            line,
            re.IGNORECASE,
        ):
            continue
        key = line[:42].lower()
        if key not in seen:
            seen.add(key)
            found.append(line)
    return found


def detect_text_equations(text: str) -> list[dict[str, str]]:
    """Best-effort detection of numbered equations in PDF/plain-text extractions."""
    results: list[dict[str, str]] = []
    seen_nums: set[str] = set()
    for match in TEXT_EQ_LINE_RE.finditer(text):
        expr = re.sub(r"\s+", " ", match.group(1)).strip()
        num = match.group(2).strip()
        if num in seen_nums:
            continue
        # Must look mathematical rather than a numbered list item
        if any(sym in expr for sym in ("=", "≈", "≤", "≥", "∈", "∑", "∫", "⊙", "+", "−", "-")):
            if len(expr) <= 180 and not expr.startswith(("（", "(", "表", "图")):
                seen_nums.add(num)
                results.append({"number": num, "expression": expr})
    return results


def trim_white_margins(image, threshold: int = 246, pad: int = 12):
    """Trim uniform near-white outer borders while keeping neat breathing padding."""
    from PIL import Image, ImageChops

    rgb = image.convert("RGB")
    bg = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg).convert("L")
    mask = diff.point(lambda p: 255 if p > (255 - threshold) else 0)
    bbox = mask.getbbox()
    if not bbox:
        return rgb
    x1, y1, x2, y2 = bbox
    w, h = rgb.size
    x1 = max(0, x1 - pad)
    y1 = max(0, y1 - pad)
    x2 = min(w, x2 + pad)
    y2 = min(h, y2 + pad)
    return rgb.crop((x1, y1, x2, y2))


def load_cjk_font(size: int = 22):
    """Load a CJK-capable TrueType/OpenType font if available on macOS/Linux/Windows."""
    from PIL import ImageFont

    candidates = [
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
        "/Library/Fonts/Arial Unicode.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "C:/Windows/Fonts/msyh.ttc",
    ]
    for path in candidates:
        if Path(path).is_file():
            try:
                return ImageFont.truetype(path, size=size)
            except Exception:
                continue
    return ImageFont.load_default()


def run_crop(args: argparse.Namespace) -> int:
    from PIL import Image

    page_path = args.page.expanduser().resolve()
    if not page_path.is_file():
        print(f"ERROR: rendered page image not found: {page_path}")
        return 1

    try:
        parts = [float(v.strip()) for v in args.box.split(",")]
        if len(parts) != 4:
            raise ValueError("Expected 4 comma-separated values: x1,y1,x2,y2")
        x1, y1, x2, y2 = parts
    except ValueError as exc:
        print(f"ERROR: invalid --box {args.box!r}: {exc}")
        return 1

    with Image.open(page_path) as im:
        w, h = im.size
        if max(x1, y1, x2, y2) <= 1.05:
            px1 = max(0, min(w, int(round(x1 * w))))
            py1 = max(0, min(h, int(round(y1 * h))))
            px2 = max(0, min(w, int(round(x2 * w))))
            py2 = max(0, min(h, int(round(y2 * h))))
        else:
            px1, py1, px2, py2 = int(x1), int(y1), int(x2), int(y2)

        if px2 <= px1 or py2 <= py1:
            print(f"ERROR: empty crop rectangle: {(px1, py1, px2, py2)}")
            return 1

        cropped = im.convert("RGB").crop((px1, py1, px2, py2))
        if not args.no_trim:
            cropped = trim_white_margins(cropped, pad=args.pad)

        out_path = args.out.expanduser().resolve()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        cropped.save(out_path, "PNG", optimize=True)
        cw, ch = cropped.size
        print(
            f"OK: cropped paper figure saved to {out_path} "
            f"({cw}x{ch}, aspect={cw / max(1, ch):.2f})"
        )
    return 0


# ---------------------------------------------------------------------------
# Word (.docx) OMML math, structured table, and multi-panel figure extraction
# ---------------------------------------------------------------------------


def _omml_children(elem: ET.Element | None) -> str:
    if elem is None:
        return ""
    return "".join(omml_to_latex(child) for child in elem)


def omml_to_latex(elem: ET.Element) -> str:
    """Convert Word Office Math Markup Language (OMML) XML into readable LaTeX-style math."""
    tag = elem.tag.split("}")[-1]

    if tag == "t":
        return elem.text or ""
    if tag in ("rPr", "ctrlPr", "fPr", "radPr", "naryPr", "dPr", "sSubPr", "sSupPr", "sSubSupPr", "mPr", "eqArrPr", "accPr", "barPr", "limLowPr", "limUppPr", "funcPr"):
        return ""
    if tag == "f":
        num = _omml_children(elem.find(f"{M_NS}num")).strip()
        den = _omml_children(elem.find(f"{M_NS}den")).strip()
        return rf"\frac{{{num}}}{{{den}}}"
    if tag == "sSub":
        base = _omml_children(elem.find(f"{M_NS}e")).strip()
        sub = _omml_children(elem.find(f"{M_NS}sub")).strip()
        return f"{base}_{{{sub}}}"
    if tag == "sSup":
        base = _omml_children(elem.find(f"{M_NS}e")).strip()
        sup = _omml_children(elem.find(f"{M_NS}sup")).strip()
        return f"{base}^{{{sup}}}"
    if tag == "sSubSup":
        base = _omml_children(elem.find(f"{M_NS}e")).strip()
        sub = _omml_children(elem.find(f"{M_NS}sub")).strip()
        sup = _omml_children(elem.find(f"{M_NS}sup")).strip()
        return f"{base}_{{{sub}}}^{{{sup}}}"
    if tag == "rad":
        deg = _omml_children(elem.find(f"{M_NS}deg")).strip()
        radicand = _omml_children(elem.find(f"{M_NS}e")).strip()
        return rf"\sqrt[{deg}]{{{radicand}}}" if deg else rf"\sqrt{{{radicand}}}"
    if tag == "nary":
        chr_el = elem.find(f".//{M_NS}chr")
        sym = chr_el.attrib.get(f"{M_NS}val", "∑") if chr_el is not None else "∑"
        sub = _omml_children(elem.find(f"{M_NS}sub")).strip()
        sup = _omml_children(elem.find(f"{M_NS}sup")).strip()
        body = _omml_children(elem.find(f"{M_NS}e")).strip()
        limits = ""
        if sub:
            limits += f"_{{{sub}}}"
        if sup:
            limits += f"^{{{sup}}}"
        return f"{sym}{limits} {body}".strip()
    if tag == "d":
        beg_el = elem.find(f".//{M_NS}begChr")
        end_el = elem.find(f".//{M_NS}endChr")
        beg = beg_el.attrib.get(f"{M_NS}val", "(") if beg_el is not None else "("
        end = end_el.attrib.get(f"{M_NS}val", ")") if end_el is not None else ")"
        items = [_omml_children(x).strip() for x in elem.findall(f"{M_NS}e")]
        return f"{beg}{', '.join(items)}{end}"
    if tag == "acc":
        chr_el = elem.find(f".//{M_NS}chr")
        acc_char = chr_el.attrib.get(f"{M_NS}val", "^") if chr_el is not None else "^"
        cmd = ACCENT_MAP.get(acc_char, r"\hat")
        body = _omml_children(elem.find(f"{M_NS}e")).strip()
        return f"{cmd}{{{body}}}"
    if tag == "bar":
        body = _omml_children(elem.find(f"{M_NS}e")).strip()
        return rf"\overline{{{body}}}"
    if tag == "limLow":
        base = _omml_children(elem.find(f"{M_NS}e")).strip()
        lim = _omml_children(elem.find(f"{M_NS}lim")).strip()
        return f"{base}_{{{lim}}}"
    if tag == "limUpp":
        base = _omml_children(elem.find(f"{M_NS}e")).strip()
        lim = _omml_children(elem.find(f"{M_NS}lim")).strip()
        return f"{base}^{{{lim}}}"
    if tag == "func":
        fname = _omml_children(elem.find(f"{M_NS}fName")).strip()
        body = _omml_children(elem.find(f"{M_NS}e")).strip()
        return f"{fname}({body})"
    if tag == "eqArr":
        rows = [_omml_children(x).strip() for x in elem.findall(f"{M_NS}e")]
        return " ; ".join(r for r in rows if r)
    if tag == "m":
        rows = []
        for mr in elem.findall(f"{M_NS}mr"):
            cols = [_omml_children(me).strip() for me in mr.findall(f"{M_NS}e")]
            rows.append(" & ".join(cols))
        return r"\begin{bmatrix} " + r" \\ ".join(rows) + r" \end{bmatrix}"

    return _omml_children(elem)


def _clean_math_str(expr: str) -> str:
    cleaned = expr.replace("\u2005", " ").replace("\u2009", " ").replace("\xa0", " ")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def _parse_docx_paragraph(
    p_elem: ET.Element, rels: dict[str, str]
) -> tuple[str, list[str], list[str]]:
    """Return (formatted_text, referenced_media_filenames, omml_expressions_in_p)."""
    parts: list[str] = []
    media_refs: list[str] = []
    math_exprs: list[str] = []

    for child in p_elem:
        ctag = child.tag
        if ctag == f"{W_NS}r":
            for rc in child:
                if rc.tag == f"{W_NS}t" and rc.text:
                    parts.append(rc.text)
                elif rc.tag == f"{W_NS}tab":
                    parts.append(" ")
                elif rc.tag == f"{W_NS}br":
                    parts.append("\n")
            for blip in child.findall(f".//{A_NS}blip"):
                rid = blip.attrib.get(f"{R_NS}embed")
                if rid and rid in rels:
                    mname = Path(rels[rid]).name
                    media_refs.append(mname)
                    parts.append(f" ![figure](figures/{mname}) ")
            for imagedata in child.findall(f".//{V_NS}imagedata"):
                rid = imagedata.attrib.get(f"{R_NS}id")
                if rid and rid in rels:
                    mname = Path(rels[rid]).name
                    media_refs.append(mname)
                    parts.append(f" ![figure](figures/{mname}) ")
        elif ctag in (f"{M_NS}oMath", f"{M_NS}oMathPara"):
            mstr = _clean_math_str(omml_to_latex(child))
            if mstr:
                math_exprs.append(mstr)
                parts.append(f" ${mstr}$ ")
        elif ctag == f"{W_NS}hyperlink":
            for t in child.findall(f".//{W_NS}t"):
                if t.text:
                    parts.append(t.text)

    text = "".join(parts)
    text = re.sub(r"[ \t]+", " ", text).strip()
    return text, media_refs, math_exprs


def _stitch_docx_figure_grid(
    grid_cells: list[list[tuple[list[str], str]]],
    figs_dir: Path,
    out_dir: Path,
    composite_idx: int,
) -> dict[str, object] | None:
    """Stitch a multi-panel Word figure table ((a)..(h) subfigures + labels) into one crisp PNG."""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return None

    rows = len(grid_cells)
    cols = max((len(r) for r in grid_cells), default=0)
    if rows == 0 or cols == 0:
        return None

    # Load all first-image panels per cell
    loaded_grid: list[list[tuple[Image.Image | None, str]]] = []
    sample_aspects: list[float] = []
    for row in grid_cells:
        loaded_row: list[tuple[Image.Image | None, str]] = []
        for media_list, label in row:
            im_obj = None
            for mname in media_list:
                candidate = figs_dir / mname
                if candidate.is_file():
                    try:
                        with Image.open(candidate) as raw_im:
                            im_obj = raw_im.convert("RGB")
                        sample_aspects.append(im_obj.width / max(1, im_obj.height))
                        break
                    except Exception:
                        im_obj = None
            loaded_row.append((im_obj, label))
        loaded_grid.append(loaded_row)

    if not sample_aspects:
        return None

    avg_aspect = sum(sample_aspects) / len(sample_aspects)
    max_canvas_w = 2400
    pad = 22
    cell_w = max(260, min(620, (max_canvas_w - (cols + 1) * pad) // cols))
    cell_h = max(200, int(round(cell_w / max(0.45, min(2.4, avg_aspect)))))
    has_labels = any(lbl.strip() for row in loaded_grid for _, lbl in row)
    label_h = 46 if has_labels else 8

    canvas_w = cols * cell_w + (cols + 1) * pad
    canvas_h = rows * (cell_h + label_h) + (rows + 1) * pad
    canvas = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
    draw = ImageDraw.Draw(canvas)
    font = load_cjk_font(size=max(18, min(26, cell_w // 15)))

    for r_idx, row in enumerate(loaded_grid):
        for c_idx, (im_obj, label) in enumerate(row):
            x0 = pad + c_idx * (cell_w + pad)
            y0 = pad + r_idx * (cell_h + label_h + pad)
            if im_obj is not None:
                scale = min(cell_w / im_obj.width, cell_h / im_obj.height)
                nw = max(1, int(round(im_obj.width * scale)))
                nh = max(1, int(round(im_obj.height * scale)))
                resized = im_obj.resize((nw, nh), Image.Resampling.LANCZOS)
                px = x0 + (cell_w - nw) // 2
                py = y0 + (cell_h - nh) // 2
                canvas.paste(resized, (px, py))
            if label.strip():
                clean_lbl = re.sub(r"\s+", " ", label).strip()
                try:
                    bbox = draw.textbbox((0, 0), clean_lbl, font=font)
                    tw = bbox[2] - bbox[0]
                except Exception:
                    tw = len(clean_lbl) * 14
                tx = x0 + max(0, (cell_w - tw) // 2)
                ty = y0 + cell_h + 8
                draw.text((tx, ty), clean_lbl, fill=(26, 32, 44), font=font)

    out_file = figs_dir / f"composite_panel_{composite_idx:02d}.png"
    canvas.save(out_file, "PNG", optimize=True)
    return {
        "file": str(out_file.relative_to(out_dir)),
        "page": 0,
        "width": canvas_w,
        "height": canvas_h,
        "kind": "stitched_docx_panel_grid",
        "grid": f"{rows}x{cols}",
    }


def extract_docx(docx_path: Path, out_dir: Path) -> dict[str, object]:
    figs_dir = out_dir / "figures"
    figs_dir.mkdir(parents=True, exist_ok=True)

    blocks: list[str] = []
    extracted_figures: list[dict[str, object]] = []
    numbered_equations: list[dict[str, str]] = []
    data_tables_summary: list[dict[str, object]] = []
    omml_total = 0
    composite_idx = 0

    with zipfile.ZipFile(docx_path) as zf:
        # 1. Export all embedded media files with real pixel dimensions
        media_dims: dict[str, tuple[int, int]] = {}
        for name in sorted(zf.namelist()):
            if name.startswith("word/media/") and not name.endswith("/"):
                dest = figs_dir / Path(name).name
                dest.write_bytes(zf.read(name))
                iw, ih = 0, 0
                try:
                    from PIL import Image

                    with Image.open(dest) as im:
                        iw, ih = im.size
                except Exception:
                    iw, ih = 0, 0
                media_dims[dest.name] = (iw, ih)
                extracted_figures.append(
                    {
                        "file": str(dest.relative_to(out_dir)),
                        "page": 0,
                        "width": iw,
                        "height": ih,
                        "kind": "embedded_media",
                    }
                )

        # 2. Parse document relationships (rId -> media/*)
        rels: dict[str, str] = {}
        if "word/_rels/document.xml.rels" in zf.namelist():
            rroot = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
            for rel in rroot:
                rid = rel.attrib.get("Id")
                target = rel.attrib.get("Target")
                if rid and target:
                    rels[rid] = target

        # 3. Walk document body in exact reading order (paragraphs, equation tables, figure grids, data tables)
        if "word/document.xml" in zf.namelist():
            root = ET.fromstring(zf.read("word/document.xml"))
            body = root.find(f"{W_NS}body")
            elements = list(body) if body is not None else []

            for elem in elements:
                if elem.tag == f"{W_NS}p":
                    txt, _, m_list = _parse_docx_paragraph(elem, rels)
                    omml_total += len(m_list)
                    if txt:
                        blocks.append(txt)

                elif elem.tag == f"{W_NS}tbl":
                    raw_rows: list[list[tuple[str, list[str], list[str]]]] = []
                    total_tbl_images = 0
                    for tr in elem.findall(f"{W_NS}tr"):
                        row_cells: list[tuple[str, list[str], list[str]]] = []
                        for tc in tr.findall(f"{W_NS}tc"):
                            p_texts: list[str] = []
                            c_media: list[str] = []
                            c_math: list[str] = []
                            for p in tc.findall(f"{W_NS}p"):
                                ptxt, pmed, pmth = _parse_docx_paragraph(p, rels)
                                omml_total += len(pmth)
                                if ptxt:
                                    p_texts.append(ptxt)
                                c_media.extend(pmed)
                                c_math.extend(pmth)
                            total_tbl_images += len(c_media)
                            joined_cell = " ".join(p_texts).strip()
                            row_cells.append((joined_cell, c_media, c_math))
                        if row_cells:
                            raw_rows.append(row_cells)

                    if not raw_rows:
                        continue

                    # Case A: Equation layout table (e.g. [ formula | (1) ])
                    is_eq_table = all(
                        len(r) >= 2
                        and EQ_NUM_CELL_RE.match(r[-1][0]) is not None
                        and not any(c[1] for c in r)
                        for r in raw_rows
                    )
                    if is_eq_table:
                        eq_lines: list[str] = []
                        for r in raw_rows:
                            eq_no_m = EQ_NUM_CELL_RE.match(r[-1][0])
                            eq_no = eq_no_m.group(1) if eq_no_m else ""
                            lhs_math = " ".join(m for c in r[:-1] for m in c[2]).strip()
                            if not lhs_math:
                                lhs_math = " ".join(c[0] for c in r[:-1]).strip().strip("$")
                            lhs_math = _clean_math_str(lhs_math)
                            numbered_equations.append({"number": eq_no, "expression": lhs_math})
                            eq_lines.append(f"$$ {lhs_math} \\quad ({eq_no}) $$")
                        blocks.append("\n".join(eq_lines))
                        continue

                    # Case B: Multi-panel figure comparison table (e.g. Fig. 6 / Fig. 7 / Fig. 8)
                    if total_tbl_images >= 2:
                        composite_idx += 1
                        grid_spec: list[list[tuple[list[str], str]]] = []
                        for r in raw_rows:
                            g_row: list[tuple[list[str], str]] = []
                            for cell_txt, cell_media, _ in r:
                                clean_subcap = re.sub(r"!\[figure\]\([^)]+\)", "", cell_txt).strip()
                                g_row.append((cell_media, clean_subcap))
                            grid_spec.append(g_row)
                        comp_info = _stitch_docx_figure_grid(
                            grid_spec, figs_dir, out_dir, composite_idx
                        )
                        if comp_info is not None:
                            extracted_figures.append(comp_info)
                            subcaps = [
                                lbl
                                for row in grid_spec
                                for _, lbl in row
                                if lbl
                            ]
                            blocks.append(
                                f"![composite_figure]({comp_info['file']})\n"
                                f"> [自动拼合多子图面板 `{comp_info['file']}` ({comp_info['width']}x{comp_info['height']}, "
                                f"网格 {comp_info['grid']})，子图标签：{' | '.join(subcaps)}]"
                            )
                            continue

                    # Case C: Genuine academic data table -> GitHub Markdown table
                    col_count = max(len(r) for r in raw_rows)
                    md_rows: list[str] = []
                    for r_i, r in enumerate(raw_rows):
                        padded = [r[c_i][0].replace("|", r"\|") if c_i < len(r) else "" for c_i in range(col_count)]
                        md_rows.append("| " + " | ".join(padded) + " |")
                        if r_i == 0:
                            md_rows.append("| " + " | ".join(["---"] * col_count) + " |")
                    blocks.append("\n".join(md_rows))
                    header_preview = [c[0] for c in raw_rows[0][:6]]
                    data_tables_summary.append(
                        {
                            "index": len(data_tables_summary) + 1,
                            "rows": len(raw_rows),
                            "cols": col_count,
                            "headers": header_preview,
                        }
                    )

    text = "\n\n".join(blocks)
    (out_dir / "extracted_text.md").write_text(text, encoding="utf-8")

    # Map standalone & stitched figures to their immediately following captions in extracted_text.md
    figure_caption_links: list[dict[str, str]] = []
    for i, blk in enumerate(blocks):
        img_m = re.search(r"!\[(?:figure|composite_figure)\]\((figures/[^)]+)\)", blk)
        if img_m:
            fig_rel = img_m.group(1)
            caption_guess = ""
            for look_ahead in range(i + 1, min(len(blocks), i + 4)):
                cand = blocks[look_ahead].strip()
                if re.match(r"^(?:图\s*\d+|Fig(?:ure)?\.?\s*\d+)", cand, re.IGNORECASE):
                    caption_guess = cand
                    break
            figure_caption_links.append({"file": fig_rel, "caption": caption_guess})

    return {
        "source_file": str(docx_path),
        "source_type": "docx",
        "extractor": "docx-omml-table-panel-v5",
        "block_count": len(blocks),
        "omml_math_count": omml_total,
        "numbered_equations": numbered_equations,
        "data_tables": data_tables_summary,
        "captions": detect_captions(text),
        "figure_caption_links": figure_caption_links,
        "extracted_figures": extracted_figures,
    }


# ---------------------------------------------------------------------------
# PDF and Plain-Text Extractors
# ---------------------------------------------------------------------------


def extract_pdf_with_fitz(
    pdf_path: Path, out_dir: Path, dpi: int, render_all: bool, max_pages: int
) -> dict[str, object] | None:
    try:
        import fitz  # type: ignore
    except ImportError:
        return None

    pages_dir = out_dir / "pages"
    figs_dir = out_dir / "figures"
    pages_dir.mkdir(parents=True, exist_ok=True)
    figs_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    page_records: list[dict[str, object]] = []
    full_text_parts: list[str] = []
    extracted_figures: list[dict[str, object]] = []
    seen_xrefs: set[int] = set()

    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)

    for idx in range(min(total_pages, max_pages)):
        page = doc[idx]
        page_num = idx + 1
        text = page.get_text("text") or ""
        captions = detect_captions(text)
        images = page.get_images(full=True)
        drawings = page.get_drawings()
        has_visuals = bool(images) or len(drawings) >= 12 or bool(captions)

        rendered_rel: str | None = None
        if render_all or has_visuals or page_num == 1:
            png_path = pages_dir / f"page_{page_num:02d}.png"
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            pix.save(str(png_path))
            rendered_rel = str(png_path.relative_to(out_dir))

        page_fig_rels: list[str] = []
        for img_idx, img_info in enumerate(images, start=1):
            xref = img_info[0]
            if xref in seen_xrefs:
                continue
            seen_xrefs.add(xref)
            try:
                base = doc.extract_image(xref)
                iw, ih = base.get("width", 0), base.get("height", 0)
                if iw < 140 or ih < 100:
                    continue
                ext = base.get("ext", "png")
                fig_file = figs_dir / f"raw_p{page_num:02d}_img{img_idx:02d}.{ext}"
                fig_file.write_bytes(base["image"])
                rel = str(fig_file.relative_to(out_dir))
                page_fig_rels.append(rel)
                extracted_figures.append(
                    {"file": rel, "page": page_num, "width": iw, "height": ih}
                )
            except Exception:
                continue

        page_records.append(
            {
                "page": page_num,
                "char_count": len(text.strip()),
                "embedded_image_count": len(images),
                "vector_drawing_count": len(drawings),
                "captions": captions,
                "rendered_page": rendered_rel,
                "extracted_raw_images": page_fig_rels,
            }
        )
        full_text_parts.append(f"## Page {page_num}\n\n{text.strip()}\n")

    doc.close()
    joined_text = "\n".join(full_text_parts)
    (out_dir / "extracted_text.md").write_text(joined_text, encoding="utf-8")
    return {
        "source_file": str(pdf_path),
        "source_type": "pdf",
        "extractor": "pymupdf",
        "render_dpi": dpi,
        "total_pages": total_pages,
        "processed_pages": len(page_records),
        "numbered_equations": detect_text_equations(joined_text),
        "extracted_figures": extracted_figures,
        "pages": page_records,
    }


def extract_pdf_with_cli(
    pdf_path: Path, out_dir: Path, dpi: int, render_all: bool, max_pages: int
) -> dict[str, object]:
    pages_dir = out_dir / "pages"
    figs_dir = out_dir / "figures"
    pages_dir.mkdir(parents=True, exist_ok=True)
    figs_dir.mkdir(parents=True, exist_ok=True)

    page_texts: list[str] = []
    text_method = "none"
    try:
        from pypdf import PdfReader  # type: ignore

        reader = PdfReader(str(pdf_path))
        page_texts = [(p.extract_text() or "") for p in reader.pages[:max_pages]]
        text_method = "pypdf"
    except Exception:
        if shutil.which("pdftotext"):
            proc = subprocess.run(
                ["pdftotext", str(pdf_path), "-"],
                capture_output=True,
                text=True,
                check=False,
            )
            if proc.returncode == 0:
                page_texts = proc.stdout.split("\f")
                if page_texts and not page_texts[-1].strip():
                    page_texts.pop()
                page_texts = page_texts[:max_pages]
                text_method = "pdftotext"

    rendered_files: dict[int, str] = {}
    render_method = "none"
    if shutil.which("pdftoppm"):
        prefix = pages_dir / "page"
        cmd = [
            "pdftoppm",
            "-png",
            "-r",
            str(dpi),
            "-f",
            "1",
            "-l",
            str(max_pages),
            str(pdf_path),
            str(prefix),
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
        if proc.returncode == 0:
            render_method = "pdftoppm"
            for png in sorted(pages_dir.glob("page-*.png")):
                m = re.search(r"page-(\d+)\.png$", png.name)
                if not m:
                    continue
                pnum = int(m.group(1))
                normalized = pages_dir / f"page_{pnum:02d}.png"
                if png != normalized:
                    png.rename(normalized)
                rendered_files[pnum] = str(normalized.relative_to(out_dir))

    extracted_figures: list[dict[str, object]] = []
    if shutil.which("pdfimages"):
        prefix = figs_dir / "raw_img"
        proc = subprocess.run(
            ["pdfimages", "-png", "-p", str(pdf_path), str(prefix)],
            capture_output=True,
            text=True,
            check=False,
        )
        if proc.returncode == 0:
            for png in sorted(figs_dir.glob("raw_img-*.png")):
                try:
                    from PIL import Image

                    with Image.open(png) as im:
                        iw, ih = im.size
                    if iw < 140 or ih < 100:
                        png.unlink(missing_ok=True)
                        continue
                except Exception:
                    iw, ih = 0, 0
                pm = re.search(r"raw_img-(\d+)-", png.name)
                pnum = int(pm.group(1)) if pm else 0
                extracted_figures.append(
                    {
                        "file": str(png.relative_to(out_dir)),
                        "page": pnum,
                        "width": iw,
                        "height": ih,
                    }
                )

    total = max(len(page_texts), len(rendered_files))
    page_records: list[dict[str, object]] = []
    full_text_parts: list[str] = []

    for pnum in range(1, total + 1):
        text = page_texts[pnum - 1] if pnum <= len(page_texts) else ""
        captions = detect_captions(text)
        rendered_rel = rendered_files.get(pnum)
        if not render_all and rendered_rel and pnum > 1 and not captions and len(rendered_files) > 16:
            if pnum < total:
                (out_dir / rendered_rel).unlink(missing_ok=True)
                rendered_rel = None

        page_records.append(
            {
                "page": pnum,
                "char_count": len(text.strip()),
                "captions": captions,
                "rendered_page": rendered_rel,
            }
        )
        if text.strip():
            full_text_parts.append(f"## Page {pnum}\n\n{text.strip()}\n")

    joined_text = "\n".join(full_text_parts)
    if full_text_parts:
        (out_dir / "extracted_text.md").write_text(joined_text, encoding="utf-8")

    return {
        "source_file": str(pdf_path),
        "source_type": "pdf",
        "extractor": f"text={text_method},render={render_method}",
        "render_dpi": dpi,
        "total_pages": total,
        "processed_pages": len(page_records),
        "numbered_equations": detect_text_equations(joined_text),
        "extracted_figures": extracted_figures,
        "pages": page_records,
    }


def extract_text_file(path: Path, out_dir: Path) -> dict[str, object]:
    text = path.read_text(encoding="utf-8", errors="replace")
    (out_dir / "extracted_text.md").write_text(text, encoding="utf-8")
    return {
        "source_file": str(path),
        "source_type": path.suffix.lower().lstrip(".") or "text",
        "extractor": "plain-text",
        "char_count": len(text),
        "numbered_equations": detect_text_equations(text),
        "captions": detect_captions(text),
        "extracted_figures": [],
    }


def write_figure_guide(out_dir: Path, inventory: dict[str, object]) -> None:
    lines = [
        "# Source Paper Asset, Equation & Table Catalog (v5.0)",
        "",
        f"- Source: `{inventory.get('source_file')}`",
        f"- Type: `{inventory.get('source_type')}` ({inventory.get('extractor')})",
    ]
    if inventory.get("omml_math_count") is not None:
        lines.append(f"- Extracted Word OMML math expressions: `{inventory.get('omml_math_count')}`")

    lines += ["", "## 1. Detected Figure & Table Captions"]
    pages = inventory.get("pages")
    if isinstance(pages, list):
        any_cap = False
        for p in pages:
            if not isinstance(p, dict):
                continue
            caps = p.get("captions") or []
            if caps:
                any_cap = True
                rpage = p.get("rendered_page") or "not rendered"
                lines.append(f"- **Page {p.get('page')}** (`{rpage}`):")
                for c in caps:
                    lines.append(f"  - {c}")
        if not any_cap:
            lines.append("- (No captions detected via regex; inspect rendered pages visually.)")
    else:
        caps = inventory.get("captions") or []
        if caps:
            for c in caps:
                lines.append(f"- {c}")
        else:
            lines.append("- (No captions detected via regex; inspect `extracted_text.md`.)")

    links = inventory.get("figure_caption_links") or []
    if links:
        lines += ["", "## 2. Direct Figure-to-Caption Mapping (Ready for `embed_paper_figure.py`)"]
        for item in links:
            if isinstance(item, dict):
                cap = item.get("caption") or "(inspect context in extracted_text.md)"
                lines.append(f"- `{item.get('file')}` → **{cap}**")

    eqs = inventory.get("numbered_equations") or []
    lines += [
        "",
        f"## 3. Extracted Numbered Equations ({len(eqs)}) — MUST Map Key Equations Onto Method Slides!",
    ]
    if eqs:
        for eq in eqs:
            if isinstance(eq, dict):
                lines.append(f"- **式({eq.get('number')})**: `{eq.get('expression')}`")
    else:
        lines.append(
            "- (Check `extracted_text.md` and rendered PDF pages for mathematical formulas; "
            "every core method/loss formula MUST appear on-screen, never replaced by textbook PSNR/FID definitions!)"
        )

    dtbls = inventory.get("data_tables") or []
    if dtbls:
        lines += ["", f"## 4. Extracted Structured Data Tables ({len(dtbls)})"]
        for tbl in dtbls:
            if isinstance(tbl, dict):
                hdrs = " | ".join(str(h) for h in (tbl.get("headers") or []))
                lines.append(
                    f"- Table #{tbl.get('index')}: {tbl.get('rows')} rows × {tbl.get('cols')} cols — Headers: `{hdrs}`"
                )

    figs = inventory.get("extracted_figures") or []
    lines += ["", f"## 5. Extracted Figure Files ({len(figs)})"]
    for f in figs:
        if isinstance(f, dict):
            kind = f" [{f.get('kind')}]" if f.get("kind") else ""
            lines.append(
                f"- `{f.get('file')}` (page {f.get('page', '?')}, "
                f"{f.get('width', '?')}x{f.get('height', '?')}){kind}"
            )

    lines += [
        "",
        "## 6. Cropping Complex Vector/Multi-Panel Figures (PDF Pages)",
        "Inspect rendered pages in `pages/page_XX.png` and crop any complete figure cleanly:",
        "```bash",
        "python3 scripts/extract_paper_assets.py crop --page pages/page_03.png --box 0.06,0.08,0.94,0.48 --out figures/fig_01.png",
        "```",
    ]
    (out_dir / "figure_guide.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def run_extract(args: argparse.Namespace) -> int:
    src = args.input.expanduser().resolve()
    if not src.is_file():
        print(f"ERROR: source document not found: {src}")
        return 1

    out_dir = args.out_dir.expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    suffix = src.suffix.lower()

    if suffix == ".pdf":
        inventory = extract_pdf_with_fitz(src, out_dir, args.dpi, args.render_all, args.max_pages)
        if inventory is None:
            inventory = extract_pdf_with_cli(src, out_dir, args.dpi, args.render_all, args.max_pages)
    elif suffix == ".docx":
        inventory = extract_docx(src, out_dir)
    else:
        inventory = extract_text_file(src, out_dir)

    inv_path = out_dir / "inventory.json"
    inv_path.write_text(json.dumps(inventory, ensure_ascii=False, indent=2), encoding="utf-8")
    write_figure_guide(out_dir, inventory)

    print(f"OK: extracted source assets to {out_dir}")
    print(f"OK: inventory written to {inv_path}")
    print(f"OK: asset/equation/table catalog written to {out_dir / 'figure_guide.md'}")
    if inventory.get("source_type") == "pdf":
        pages = inventory.get("pages", [])
        rendered = [p for p in pages if isinstance(p, dict) and p.get("rendered_page")]
        caption_count = sum(len(p.get("captions", [])) for p in pages if isinstance(p, dict))
        fig_count = len(inventory.get("extracted_figures") or [])
        eq_count = len(inventory.get("numbered_equations") or [])
        print(
            f"SUMMARY: {inventory.get('processed_pages')} page(s) parsed via "
            f"{inventory.get('extractor')}, {len(rendered)} page PNG(s) rendered at {args.dpi} DPI, "
            f"{fig_count} embedded figure(s), {caption_count} caption(s), {eq_count} numbered equation(s)"
        )
    elif inventory.get("source_type") == "docx":
        print(
            f"SUMMARY: DOCX parsed ({inventory.get('block_count')} blocks, "
            f"{inventory.get('omml_math_count')} OMML math expressions, "
            f"{len(inventory.get('numbered_equations') or [])} numbered equations, "
            f"{len(inventory.get('data_tables') or [])} structured data tables, "
            f"{len(inventory.get('extracted_figures') or [])} figure files incl. stitched multi-panel grids)"
        )
    return 0


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] == "crop":
        parser = argparse.ArgumentParser(
            prog="extract_paper_assets.py crop",
            description="Crop a figure from a rendered PDF page with optional white-margin auto-trimming.",
        )
        parser.add_argument("--page", required=True, type=Path, help="Rendered page PNG path.")
        parser.add_argument(
            "--box",
            required=True,
            type=str,
            help="Crop box x1,y1,x2,y2 (normalized 0..1 ratios or pixel coords).",
        )
        parser.add_argument("--out", required=True, type=Path, help="Output cropped figure PNG path.")
        parser.add_argument("--pad", type=int, default=12, help="Outer white padding after auto-trim (default: 12).")
        parser.add_argument("--no-trim", action="store_true", help="Disable automatic white-margin trimming.")
        args = parser.parse_args(sys.argv[2:])
        return run_crop(args)

    parser = argparse.ArgumentParser(
        description="Extract text, equations, structured tables, page renders, and figures from a paper."
    )
    parser.add_argument("--input", required=True, type=Path, help="Path to PDF, DOCX, MD, or TXT file.")
    parser.add_argument(
        "--out-dir",
        default=Path("./outputs/source_assets"),
        type=Path,
        help="Output directory for extracted text, renders, figures, and inventory.json.",
    )
    parser.add_argument("--dpi", type=int, default=220, help="Render DPI for PDF pages (default: 220).")
    parser.add_argument(
        "--render-all",
        action="store_true",
        help="Render all PDF pages (up to --max-pages) rather than only figure-bearing pages.",
    )
    parser.add_argument("--max-pages", type=int, default=30, help="Maximum PDF pages to process (default: 30).")
    args = parser.parse_args()
    return run_extract(args)


if __name__ == "__main__":
    raise SystemExit(main())
