"""Fast native PowerPoint COM writers, CJK-safe no-squiggle typography, spatial-zone choreography, and audit."""
import math
import re

try:
    import pythoncom
except ImportError:
    pythoncom = None

PRESETS = {
    "rectangle": 1,
    "rect": 1,
    "box": 1,
    "parallelogram": 2,
    "trapezoid": 3,
    "diamond": 4,
    "rounded_rectangle": 5,
    "round_rect": 5,
    "rounded_rect": 5,
    "round_rectangle": 5,
    "pill": 5,
    "octagon": 6,
    "triangle": 7,
    "isosceles_triangle": 7,
    "right_triangle": 8,
    "oval": 9,
    "ellipse": 9,
    "circle": 9,
    "hexagon": 10,
    "cross": 11,
    "plus": 11,
    "regular_pentagon": 12,
    "cylinder": 13,
    "can": 13,
    "cube": 14,
    "donut": 18,
    "right_arrow": 33,
    "left_arrow": 34,
    "up_arrow": 35,
    "down_arrow": 36,
    "left_right_arrow": 37,
    "up_down_arrow": 38,
    "bent_arrow": 41,
    "u_turn_arrow": 42,
    "curved_right_arrow": 45,
    "curved_left_arrow": 46,
    "pentagon": 51,
    "pentagon_arrow": 51,
    "chevron": 52,
    "star5": 92,
    "callout_rect": 105,
    "callout_rounded": 106,
    "cloud": 179,
}

ARROWS = {"none": 1, "triangle": 2, "open": 3, "arrow": 3, "stealth": 4, "diamond": 5, "oval": 6}
ARROW_SIZES = {"narrow": 1, "short": 1, "sm": 1, "medium": 2, "med": 2, "wide": 3, "long": 3, "lg": 3}
DASHES = {
    "solid": 1,
    "square_dot": 2,
    "dot": 3,
    "round_dot": 3,
    "dash": 4,
    "dashed": 4,
    "dash_dot": 5,
    "long_dash": 7,
    "long_dash_dot": 8,
    "long_dash_dot_dot": 9,
}
ALIGNMENTS = {"left": 1, "center": 2, "centre": 2, "right": 3, "justify": 4}
V_ALIGNMENTS = {"top": 1, "middle": 3, "center": 3, "bottom": 4}

TYPE_ALIASES = {
    "shape": "add_shape",
    "add_shape": "add_shape",
    "rect": "add_shape",
    "box": "add_shape",
    "panel": "add_shape",
    "text": "add_textbox",
    "textbox": "add_textbox",
    "add_text": "add_textbox",
    "add_textbox": "add_textbox",
    "label": "add_textbox",
    "image": "add_image",
    "picture": "add_image",
    "add_picture": "add_image",
    "add_image": "add_image",
    "line": "add_line",
    "arrow": "add_line",
    "connector": "add_line",
    "add_connector": "add_line",
    "add_line": "add_line",
    "freeform": "freeform",
    "curve": "freeform",
    "path": "freeform",
    "duplicate": "duplicate_shape",
    "duplicate_shape": "duplicate_shape",
    "group": "group_shapes",
    "group_shapes": "group_shapes",
}

STAGE_NAMES = {
    0: "Stage 1/4 · Slide Skeleton",
    1: "Stage 2/4 · Title & Navigation",
    2: "Stage 3/4 · Parallel Modules",
    3: "Stage 3/4 · Annotations & Copy",
    4: "Stage 4/4 · Polish & Finalize",
}

STAGE_ALIASES = {
    "structure": 0,
    "skeleton": 0,
    "background": 0,
    "headers": 1,
    "header": 1,
    "titles": 1,
    "hierarchy": 1,
    "visuals": 2,
    "visual": 2,
    "diagrams": 2,
    "charts": 2,
    "modules": 2,
    "annotations": 3,
    "copy": 3,
    "labels": 3,
    "connectors": 3,
    "polish": 4,
    "footer": 4,
    "finalize": 4,
}

_TICK_TEXT_RE = re.compile(r"^[\s0-9.\-+eE×x^∗*⁻¹²³⁰⁴⁵⁶⁷⁸⁹]+$")


def rgb(hex_color):
    h = str(hex_color).strip().lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return int.from_bytes(bytes.fromhex(h), "little")


def has_cjk(text):
    return any("\u4e00" <= ch <= "\u9fff" or "\u3000" <= ch <= "\u303f" or "\uff00" <= ch <= "\uffef" for ch in text)


def extract_plain_text(action):
    if action.get("runs"):
        raw = "".join(str(r.get("text", "")) for r in action["runs"])
    else:
        raw = str(action.get("text", ""))
    return raw.replace("\r\n", "\n").replace("\r", "\n")


def char_em_width(ch, bold=False):
    """Calibrated em-width ratio matching Microsoft YaHei (CJK) and Arial/Times (Latin)."""
    b_mul = 1.05 if bold else 1.0
    if "\u4e00" <= ch <= "\u9fff" or "\u3000" <= ch <= "\u303f" or "\uff00" <= ch <= "\uffef":
        return 1.04 * b_mul
    if ch in "ilIj1.,:;|!`' ":
        return 0.32 * b_mul
    if ch in "mwMW@":
        return 0.84 * b_mul
    if ch.isupper():
        return 0.68 * b_mul
    return 0.54 * b_mul


def estimate_text_metrics(action):
    """Estimate max single-line width and total wrapped height (in slide points) for a text action."""
    runs = action.get("runs")
    base_fs = float(action.get("font_size", 14.0))
    base_bold = bool(action.get("bold", False))
    ml = float(action.get("margin_left", 0.0))
    mr = float(action.get("margin_right", 0.0))
    mt = float(action.get("margin_top", 0.0))
    mb = float(action.get("margin_bottom", 0.0))
    line_sp = float(action.get("line_spacing", 1.0))

    stream = []
    if runs:
        for r in runs:
            rtext = str(r.get("text", "")).replace("\r\n", "\n").replace("\r", "\n")
            rfs = float(r.get("font_size", base_fs))
            rbold = bool(r.get("bold", base_bold))
            scale = 0.72 if (r.get("superscript") or r.get("subscript")) else 1.0
            for ch in rtext:
                stream.append((ch, rfs * scale, rbold))
    else:
        raw = str(action.get("text", "")).replace("\r\n", "\n").replace("\r", "\n")
        for ch in raw:
            stream.append((ch, base_fs, base_bold))

    para_widths = []
    cur_w = 0.0
    max_fs = base_fs
    for ch, fs, b in stream:
        if ch == "\n":
            para_widths.append(cur_w)
            cur_w = 0.0
        else:
            cur_w += char_em_width(ch, b) * fs
            if fs > max_fs:
                max_fs = fs
    para_widths.append(cur_w)
    max_line_w = max(para_widths) if para_widths else 0.0

    w = max(1.0, float(action.get("width", 100.0)) - ml - mr)
    wrap = bool(action.get("word_wrap", True))
    if wrap:
        total_lines = sum(max(1, math.ceil(pw / max(4.0, w))) for pw in para_widths)
    else:
        total_lines = max(1, len(para_widths))
    est_h = total_lines * max_fs * 1.28 * line_sp + mt + mb
    return {
        "max_line_w": round(max_line_w + ml + mr, 2),
        "est_height": round(est_h, 2),
        "para_count": len(para_widths),
        "est_lines": total_lines,
        "char_count": sum(1 for ch, _, _ in stream if ch != "\n"),
    }


def normalize_action(raw, index=1):
    """Accept both canonical swarm-plan actions and powerpoint_draw_sequence style dicts."""
    a = dict(raw)
    raw_type = str(a.get("type") or a.get("op") or a.get("action") or "add_shape").strip().lower()
    kind = TYPE_ALIASES.get(raw_type, raw_type)
    a["type"] = kind
    a["name"] = str(a.get("name") or a.get("id") or f"obj_{index:03d}")

    for dst, aliases in (
        ("left", ("left", "x")),
        ("top", ("top", "y")),
        ("width", ("width", "w")),
        ("height", ("height", "h")),
    ):
        if dst not in a:
            for k in aliases:
                if k in a and a[k] is not None:
                    a[dst] = float(a[k])
                    break

    if kind == "add_line":
        for dst, aliases in (
            ("begin_x", ("begin_x", "x1", "start_x")),
            ("begin_y", ("begin_y", "y1", "start_y")),
            ("end_x", ("end_x", "x2", "finish_x")),
            ("end_y", ("end_y", "y2", "finish_y")),
        ):
            if dst not in a:
                for k in aliases:
                    if k in a and a[k] is not None:
                        a[dst] = float(a[k])
                        break
        if "end_arrow" not in a and a.get("arrow_end"):
            a["end_arrow"] = a["arrow_end"]
        if "start_arrow" not in a and a.get("arrow_start"):
            a["start_arrow"] = a["arrow_start"]
        if raw_type == "arrow" and "end_arrow" not in a:
            a["end_arrow"] = "triangle"

    if kind == "add_image" and "image_path" not in a:
        for k in ("asset", "path", "src", "file"):
            if a.get(k):
                a["image_path"] = a[k]
                break

    if kind == "add_shape" and "shape" not in a and "shape_type_id" not in a:
        a["shape"] = a.get("preset") or a.get("shape_type") or "rectangle"
    if kind == "add_shape" and isinstance(a.get("shape"), str):
        a["shape"] = a["shape"].strip().lower().replace("-", "_").replace(" ", "_")

    if "fill_color" not in a and a.get("fill") not in (None, "", "none", "transparent"):
        a["fill_color"] = a["fill"]
    if "line_color" not in a:
        for k in ("stroke_color", "stroke", "border_color"):
            if a.get(k) not in (None, "", "none", "transparent"):
                a["line_color"] = a[k]
                break
    if "line_width" not in a:
        for k in ("stroke_width", "border_width", "weight"):
            if a.get(k) is not None:
                a["line_width"] = float(a[k])
                break
    if kind == "add_line" and not a.get("line_width"):
        a["line_width"] = 1.5
    if "font_color" not in a and a.get("color") and kind in ("add_textbox", "add_shape"):
        a["font_color"] = a["color"]
    if "z_index" not in a and a.get("layer") is not None:
        a["z_index"] = a["layer"]

    if "wave" not in a and a.get("stage"):
        st = str(a["stage"]).strip().lower()
        if st in STAGE_ALIASES:
            a["wave"] = STAGE_ALIASES[st]

    return a


def set_adjustment(shape, value, index=1):
    if pythoncom is None:
        return
    adj = shape.Adjustments._oleobj_
    adj.Invoke(adj.GetIDsOfNames("Item"), 0, pythoncom.DISPATCH_PROPERTYPUT, 0, index, float(value))


def set_adjustments(shape, action, width, height):
    if action.get("adjustments"):
        for idx, val in enumerate(action["adjustments"], 1):
            try:
                set_adjustment(shape, val, index=idx)
            except Exception:
                pass
    elif action.get("radius") is not None:
        ratio = max(0.0, min(0.5, float(action["radius"]) / max(0.1, min(width, height))))
        try:
            set_adjustment(shape, ratio, index=1)
        except Exception:
            pass


def fill(shape, action):
    f = shape.Fill
    stops = action.get("gradient_stops")
    grad = action.get("gradient")
    if stops and len(stops) >= 2:
        f.Visible = -1
        f.TwoColorGradient(action.get("gradient_style", 1), 1)
        if action.get("gradient_angle") is not None:
            try:
                f.GradientAngle = float(action["gradient_angle"]) % 360.0
            except Exception:
                pass
        first, last = stops[0], stops[-1]
        s1 = f.GradientStops.Item(1)
        s1.Color.RGB = rgb(first["color"])
        s1.Position = float(first.get("pos", 0.0))
        if first.get("transparency") is not None:
            s1.Transparency = float(first["transparency"]) / 100.0
        sn = f.GradientStops.Item(f.GradientStops.Count)
        sn.Color.RGB = rgb(last["color"])
        sn.Position = float(last.get("pos", 1.0))
        if last.get("transparency") is not None:
            sn.Transparency = float(last["transparency"]) / 100.0
        for mid in stops[1:-1]:
            try:
                f.GradientStops.Insert(
                    rgb(mid["color"]),
                    float(mid["pos"]),
                    float(mid.get("transparency", 0.0)) / 100.0,
                )
            except Exception:
                pass
        return

    if grad and len(grad) >= 2:
        f.Visible = -1
        f.TwoColorGradient(action.get("gradient_style", 1), 1)
        if action.get("gradient_angle") is not None:
            try:
                f.GradientAngle = float(action["gradient_angle"]) % 360.0
            except Exception:
                pass
        f.GradientStops.Item(1).Color.RGB = rgb(grad[0])
        f.GradientStops.Item(f.GradientStops.Count).Color.RGB = rgb(grad[-1])
        if len(grad) > 2:
            step = 1.0 / (len(grad) - 1)
            for i, col in enumerate(grad[1:-1], 1):
                try:
                    f.GradientStops.Insert(rgb(col), i * step)
                except Exception:
                    pass
        if action.get("fill_transparency"):
            f.Transparency = float(action["fill_transparency"]) / 100.0
        return

    color = action.get("fill_color")
    if color is None or str(color).lower() in ("none", "transparent"):
        f.Visible = 0
        return
    f.Visible = -1
    f.Solid()
    f.ForeColor.RGB = rgb(color)
    f.Transparency = float(action.get("fill_transparency", 0)) / 100.0


def apply_shadow(shape, action):
    cfg = action.get("shadow")
    if not cfg:
        try:
            shape.Shadow.Visible = 0
        except Exception:
            pass
        return
    if cfg is True:
        cfg = {}
    try:
        sh = shape.Shadow
        sh.Visible = -1
        sh.Style = 2
        sh.Blur = float(cfg.get("blur", 6.0))
        sh.OffsetX = float(cfg.get("offset_x", 0.0))
        sh.OffsetY = float(cfg.get("offset_y", 2.0))
        sh.ForeColor.RGB = rgb(cfg.get("color", "000000"))
        sh.Transparency = float(cfg.get("transparency", 82.0)) / 100.0
    except Exception:
        pass


def stroke(shape, action):
    line = shape.Line
    width = action.get("line_width", 0)
    if not width or float(width) <= 0:
        line.Visible = 0
        return
    line.Visible = -1
    line.Weight = float(width)
    line.ForeColor.RGB = rgb(action.get("line_color", "333333"))
    line.Transparency = float(action.get("line_transparency", 0)) / 100.0
    if action.get("line_dash"):
        line.DashStyle = DASHES.get(str(action["line_dash"]).lower(), 1)
    if action.get("start_arrow"):
        line.BeginArrowheadStyle = ARROWS.get(str(action["start_arrow"]).lower(), 1)
        if action.get("arrow_width"):
            line.BeginArrowheadWidth = ARROW_SIZES.get(str(action["arrow_width"]).lower(), 2)
        if action.get("arrow_length"):
            line.BeginArrowheadLength = ARROW_SIZES.get(str(action["arrow_length"]).lower(), 2)
    if action.get("end_arrow"):
        line.EndArrowheadStyle = ARROWS.get(str(action["end_arrow"]).lower(), 1)
        if action.get("arrow_width"):
            line.EndArrowheadWidth = ARROW_SIZES.get(str(action["arrow_width"]).lower(), 2)
        if action.get("arrow_length"):
            line.EndArrowheadLength = ARROW_SIZES.get(str(action["arrow_length"]).lower(), 2)


def _resolve_font_spec(spec, defaults, text_sample=""):
    latin_name = spec.get("font_name", defaults.get("font_name", "Arial"))
    ea_name = spec.get("font_name_east_asian", defaults.get("font_name_east_asian"))
    if not ea_name and has_cjk(text_sample):
        ea_name = "Microsoft YaHei"
    size = float(spec.get("font_size", defaults.get("font_size", 14)))
    bold = bool(spec.get("bold", defaults.get("bold", False)))
    italic = bool(spec.get("italic", defaults.get("italic", False)))
    color = str(spec.get("font_color", spec.get("color", defaults.get("font_color", "161616"))))
    underline = bool(spec.get("underline", defaults.get("underline", False)))
    if spec.get("superscript"):
        baseline = 0.3
    elif spec.get("subscript"):
        baseline = -0.25
    else:
        baseline = 0.0
    spacing = float(spec.get("char_spacing", defaults.get("char_spacing", 0.0)))
    # 2052 = Simplified Chinese (preserves CJK shaping), 1024 = msoLanguageIDNoProofing (suppresses red squigglies under math variables like Kb, Unb, γτ!)
    lang_id = 2052 if has_cjk(text_sample) else 1024
    return {
        "latin": latin_name,
        "ea": ea_name,
        "size": size,
        "bold": bold,
        "italic": italic,
        "color": color,
        "underline": underline,
        "baseline": baseline,
        "spacing": spacing,
        "lang_id": lang_id,
    }


def _apply_font_props(font_obj, spec, defaults, text_sample="", base_resolved=None, tf2_font=None, tr_obj=None, tf2_tr=None):
    """Apply font properties cleanly, skipping unchanged attributes on runs to minimize COM IPC."""
    res = _resolve_font_spec(spec, defaults, text_sample)
    if base_resolved is None:
        font_obj.Name = res["latin"]
        if res["ea"]:
            try:
                font_obj.NameFarEast = res["ea"]
            except Exception:
                pass
        font_obj.Size = res["size"]
        font_obj.Bold = -int(res["bold"])
        font_obj.Italic = -int(res["italic"])
        font_obj.Color.RGB = rgb(res["color"])
        if res["underline"]:
            try:
                font_obj.Underline = -1
            except Exception:
                pass
        if res["baseline"] != 0.0:
            try:
                font_obj.BaselineOffset = res["baseline"]
            except Exception:
                pass
        if tr_obj is not None:
            try:
                tr_obj.LanguageID = res["lang_id"]
            except Exception:
                pass
        if tf2_tr is not None:
            try:
                tf2_tr.LanguageID = res["lang_id"]
            except Exception:
                pass
        if tf2_font is not None:
            try:
                tf2_font.NameAscii = res["latin"]
                if res["ea"]:
                    tf2_font.NameFarEast = res["ea"]
                if res["spacing"] != 0.0 or has_cjk(text_sample):
                    tf2_font.Spacing = res["spacing"]
            except Exception:
                pass
        return res

    if res["latin"] != base_resolved["latin"]:
        font_obj.Name = res["latin"]
    if res["ea"] and res["ea"] != base_resolved["ea"]:
        try:
            font_obj.NameFarEast = res["ea"]
        except Exception:
            pass
    if abs(res["size"] - base_resolved["size"]) > 1e-4:
        font_obj.Size = res["size"]
    if res["bold"] != base_resolved["bold"]:
        font_obj.Bold = -int(res["bold"])
    if res["italic"] != base_resolved["italic"]:
        font_obj.Italic = -int(res["italic"])
    if res["color"].strip().lstrip("#").lower() != base_resolved["color"].strip().lstrip("#").lower():
        font_obj.Color.RGB = rgb(res["color"])
    if res["underline"] != base_resolved["underline"]:
        try:
            font_obj.Underline = -int(res["underline"])
        except Exception:
            pass
    if abs(res["baseline"] - base_resolved["baseline"]) > 1e-4:
        try:
            font_obj.BaselineOffset = res["baseline"]
        except Exception:
            pass
    if tr_obj is not None and res["lang_id"] != base_resolved["lang_id"]:
        try:
            tr_obj.LanguageID = res["lang_id"]
        except Exception:
            pass
    if tf2_tr is not None and res["lang_id"] != base_resolved["lang_id"]:
        try:
            tf2_tr.LanguageID = res["lang_id"]
        except Exception:
            pass
    if tf2_font is not None:
        try:
            if res["latin"] != base_resolved["latin"]:
                tf2_font.NameAscii = res["latin"]
            if res["ea"] and res["ea"] != base_resolved["ea"]:
                tf2_font.NameFarEast = res["ea"]
            if abs(res["spacing"] - base_resolved["spacing"]) > 1e-4:
                tf2_font.Spacing = res["spacing"]
        except Exception:
            pass
    return res


def write_text_fast(shape, action):
    """Write complete styled text with CJK-safe TextFrame + TextFrame2 metrics and zero red spellcheck squigglies."""
    frame = shape.TextFrame
    frame.AutoSize = 0
    wrap = bool(action.get("word_wrap", True))
    frame.WordWrap = -int(wrap)

    tf2 = None
    try:
        tf2 = shape.TextFrame2
        tf2.AutoSize = 0
        tf2.WordWrap = -int(wrap)
    except Exception:
        tf2 = None

    frame.VerticalAnchor = V_ALIGNMENTS.get(str(action.get("vertical_alignment", "middle")).lower(), 3)
    for side in ("Left", "Top", "Right", "Bottom"):
        setattr(frame, "Margin" + side, float(action.get("margin_" + side.lower(), 0)))

    runs = action.get("runs")
    if runs:
        full_text = "".join(str(r.get("text", "")) for r in runs).replace("\r\n", "\r").replace("\n", "\r")
    else:
        full_text = str(action.get("text", "")).replace("\r\n", "\r").replace("\n", "\r")

    tr = frame.TextRange
    tr.Text = full_text
    tf2_tr = None
    if tf2 is not None:
        try:
            tf2_tr = tf2.TextRange
        except Exception:
            tf2_tr = None

    base_res = _apply_font_props(
        tr.Font,
        action,
        action,
        full_text,
        base_resolved=None,
        tf2_font=(tf2_tr.Font if tf2_tr is not None else None),
        tr_obj=tr,
        tf2_tr=tf2_tr,
    )

    p = tr.ParagraphFormat
    align_key = str(action.get("alignment", "left")).lower()
    if align_key == "justify" and has_cjk(full_text) and (not wrap or "\r" not in full_text):
        align_key = "left"
    p.Alignment = ALIGNMENTS.get(align_key, 1)
    sp_before = float(action.get("space_before", 0))
    sp_after = float(action.get("space_after", 0))
    line_sp = float(action.get("line_spacing", 1))
    if sp_before != 0.0:
        p.SpaceBefore = sp_before
    if sp_after != 0.0:
        p.SpaceAfter = sp_after
    if abs(line_sp - 1.0) > 1e-4:
        p.LineRuleWithin = -1
        p.SpaceWithin = line_sp

    if runs:
        cursor = 1
        prev_baseline = base_res["baseline"]
        for rspec in runs:
            rtext = str(rspec.get("text", "")).replace("\r\n", "\r").replace("\n", "\r")
            if not rtext:
                continue
            rlen = len(rtext)
            sub = tr.Characters(cursor, rlen)
            cand_res = _resolve_font_spec(rspec, action, rtext)
            need_tf2 = (
                tf2_tr is not None
                and (
                    cand_res["latin"] != base_res["latin"]
                    or (cand_res["ea"] and cand_res["ea"] != base_res["ea"])
                    or abs(cand_res["spacing"] - base_res["spacing"]) > 1e-4
                    or cand_res["lang_id"] != base_res["lang_id"]
                )
            )
            tf2_sub = None
            tf2_sub_font = None
            if need_tf2:
                try:
                    tf2_sub = tf2_tr.Characters(cursor, rlen)
                    tf2_sub_font = tf2_sub.Font
                except Exception:
                    tf2_sub = tf2_sub_font = None
            res = _apply_font_props(
                sub.Font,
                rspec,
                action,
                rtext,
                base_resolved=base_res,
                tf2_font=tf2_sub_font,
                tr_obj=sub,
                tf2_tr=tf2_sub,
            )
            if prev_baseline != 0.0 and res["baseline"] == 0.0:
                try:
                    sub.Font.BaselineOffset = 0.0
                except Exception:
                    pass
            prev_baseline = res["baseline"]
            cursor += rlen


def apply_background(slide, bg_spec):
    if not bg_spec:
        return
    slide.FollowMasterBackground = 0
    bg_fill = slide.Background.Fill
    grad = bg_spec.get("gradient")
    if grad and len(grad) >= 2:
        bg_fill.Visible = -1
        bg_fill.TwoColorGradient(bg_spec.get("gradient_style", 1), 1)
        if bg_spec.get("gradient_angle") is not None:
            try:
                bg_fill.GradientAngle = float(bg_spec["gradient_angle"]) % 360.0
            except Exception:
                pass
        bg_fill.GradientStops.Item(1).Color.RGB = rgb(grad[0])
        bg_fill.GradientStops.Item(bg_fill.GradientStops.Count).Color.RGB = rgb(grad[-1])
        return
    color = bg_spec.get("fill_color") or bg_spec.get("color")
    if color:
        bg_fill.Visible = -1
        bg_fill.Solid()
        bg_fill.ForeColor.RGB = rgb(color)


def path_commands(action):
    if action["type"] == "add_line":
        return [("M", action["begin_x"], action["begin_y"]), ("L", action["end_x"], action["end_y"])]
    return action["path"]


def path_segments(commands):
    first = tuple(commands[0][1:])
    current = first
    for cmd, *coords in commands[1:]:
        if cmd == "Z":
            if math.dist(current, first) > 1e-7:
                yield ("L", current, first)
            current = first
        else:
            nodes = [tuple(coords[i : i + 2]) for i in range(0, len(coords), 2)]
            yield (cmd, current, *nodes)
            current = nodes[-1]


def draw_path(shapes, segments):
    segs = list(segments)
    start = segs[0][1]
    if len(segs) == 1 and segs[0][0] == "L":
        return shapes.AddLine(*start, *segs[0][-1])
    builder = shapes.BuildFreeform(1, *start)
    for cmd, _, *nodes in segs:
        if cmd == "C":
            builder.AddNodes(1, 1, *[c for pt in nodes for c in pt])
        else:
            builder.AddNodes(0, 0, *[c for pt in nodes for c in pt])
    return builder.ConvertToShape()


def curve_samples(segment):
    cmd, p0, *nodes = segment
    if cmd == "L":
        return [p0, nodes[0]]
    p1, p2, p3 = nodes
    perimeter = sum(math.dist(a, b) for a, b in zip((p0, p1, p2), (p1, p2, p3)))
    count = max(8, min(72, math.ceil(perimeter / 3.0)))

    def value(step):
        t = step / count
        s = 1 - t
        return tuple(
            s**3 * a + 3 * s * s * t * b + 3 * s * t * t * c + t**3 * d
            for a, b, c, d in zip(p0, p1, p2, p3)
        )

    return [p0] + [value(i) for i in range(1, count + 1)]


def bounds(action):
    if action["type"] in ("add_line", "freeform"):
        pts = [pt for seg in path_segments(path_commands(action)) for pt in seg[1:]]
        xs, ys = zip(*pts)
        return min(xs), min(ys), max(xs), max(ys)
    x, y = float(action["left"]), float(action["top"])
    return x, y, x + float(action["width"]), y + float(action["height"])


def intersects(a, b):
    return min(a[2], b[2]) + 0.25 >= max(a[0], b[0]) and min(a[3], b[3]) + 0.25 >= max(a[1], b[1])


def _overlap_ratio(a, b):
    ix = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    inter = ix * iy
    if inter <= 0:
        return 0.0
    area_a = max(1e-6, (a[2] - a[0]) * (a[3] - a[1]))
    area_b = max(1e-6, (b[2] - b[0]) * (b[3] - b[1]))
    return inter / min(area_a, area_b)


def is_micro_tick_label(action):
    """Detect tiny numeric axis tick labels (0.1, 0.5, 1.0, 10^1, 10^-2) suitable for smooth axis sweep."""
    if action["type"] != "add_textbox":
        return False
    if "micro" in action:
        return bool(action["micro"])
    txt = extract_plain_text(action).strip()
    if not txt or len(txt) > 6 or "\n" in txt:
        return False
    fs = float(action.get("font_size", 14.0))
    if fs > 11.5:
        return False
    return bool(_TICK_TEXT_RE.match(txt))


def is_micro_action(action):
    """Detect small repetitive visual primitives (chart ticks, data dots, hatch lines, spheres)."""
    if "micro" in action:
        return bool(action["micro"])
    kind = action["type"]
    if kind == "add_textbox":
        return is_micro_tick_label(action)
    if kind in ("add_image", "group_shapes", "duplicate_shape"):
        return False
    if extract_plain_text(action).strip():
        return False
    x0, y0, x1, y1 = bounds(action)
    w, h = x1 - x0, y1 - y0
    if kind == "add_shape":
        return max(w, h) <= 18.0 and (w * h) <= 260.0
    if kind == "add_line":
        if action.get("from") or action.get("to"):
            return False
        length = math.hypot(action["end_x"] - action["begin_x"], action["end_y"] - action["begin_y"])
        has_arrow = any(
            action.get(k) and str(action.get(k)).lower() != "none"
            for k in ("start_arrow", "end_arrow")
        )
        return length <= 24.0 and not has_arrow
    if kind == "freeform":
        return max(w, h) <= 18.0
    return False


def _default_z_layer(action, idx, slide_area):
    kind = action["type"]
    if kind == "add_shape":
        x0, y0, x1, y1 = bounds(action)
        area = max(0.0, (x1 - x0) * (y1 - y0))
        nm = action.get("name", "").lower()
        has_text = bool(extract_plain_text(action).strip())
        is_bg_name = any(tok in nm for tok in ("_bg", "panel", "backdrop", "container", "zone", "section", "frame", "background", "sidebar", "card", "banner"))
        if not has_text and (area >= slide_area * 0.035 or (is_bg_name and area >= slide_area * 0.02)):
            return 100 + idx
        return 200 + idx
    if kind == "add_image":
        return 300 + idx
    if kind in ("add_line", "freeform"):
        return 400 + idx
    if kind == "add_textbox":
        return 500 + idx
    return 600 + idx


def default_wave(action, slide_w=960.0, slide_h=540.0):
    """4-phase architectural slide build order:
    0: Skeleton (sidebar, header rule, subtitle banner, card/panel containers, major dividers)
    1: Top Header & Sidebar Title (main title, subtitle banner text, sidebar title)
    2: Coherent Zoned Modules (cards/panels built smoothly left->right in parallel)
    3: Remaining Global Annotations / Connectors (if outside module zones)
    4: Footer, Page Pill & Polish
    """
    if "wave" in action and action["wave"] is not None:
        return int(action["wave"])

    slide_area = max(1.0, float(slide_w) * float(slide_h))
    kind = action["type"]
    nm = action.get("name", "").lower()
    if kind == "group_shapes":
        return 4

    x0, y0, x1, y1 = bounds(action)
    w, h = x1 - x0, y1 - y0
    area = max(0.0, w * h)
    txt = extract_plain_text(action).strip()
    fs = float(action.get("font_size", 14.0))

    is_footer_name = any(tok in nm for tok in ("footer", "page_num", "page_badge", "citation", "footnote", "slide_num"))
    if is_footer_name or (y0 >= slide_h * 0.91 and (txt or area < slide_area * 0.025)):
        return 4

    if kind == "add_shape":
        is_bg_name = any(
            tok in nm
            for tok in ("_bg", "panel", "backdrop", "container", "zone", "section", "frame", "background", "sidebar", "card", "banner")
        )
        if not txt and (area >= slide_area * 0.035 or (is_bg_name and area >= slide_area * 0.02)):
            return 0
        return 2

    if kind == "add_image":
        return 2

    if kind == "add_line":
        length = math.hypot(action["end_x"] - action["begin_x"], action["end_y"] - action["begin_y"])
        has_arrow = any(
            action.get(k) and str(action.get(k)).lower() != "none"
            for k in ("start_arrow", "end_arrow")
        )
        is_divider = any(tok in nm for tok in ("divider", "rule", "sep", "border", "header_line", "footer_line"))
        if is_divider or (not has_arrow and not action.get("from") and length >= min(slide_w, slide_h) * 0.16):
            if y0 < slide_h * 0.18 or y0 > slide_h * 0.89 or x0 < slide_w * 0.18 or is_divider:
                return 0
        return 2

    if kind == "freeform":
        if any(tok in nm for tok in ("sidebar", "bg", "decor", "arc", "ring", "accent")) and area >= slide_area * 0.015:
            return 0
        return 2

    if kind == "add_textbox":
        is_main_header = any(tok in nm for tok in ("main_title", "slide_title", "banner_text", "subtitle_text", "sidebar_text", "sidebar_title"))
        is_top_banner = y0 < slide_h * 0.22 and fs >= 15.0
        is_sidebar_title = x1 <= slide_w * 0.19 and fs >= 18.0
        if is_main_header or is_top_banner or is_sidebar_title:
            return 1
        return 3

    return 4


def _detect_panel_zones(actions, raw_boxes, slide_w, slide_h):
    """Detect Wave 0 card/panel containers (4 horizontal rows, 3 vertical columns, or 2x2 grid)."""
    slide_area = max(1.0, slide_w * slide_h)
    candidates = []
    for i, a in enumerate(actions):
        if a["type"] != "add_shape" or int(a.get("wave", 2)) != 0 or raw_boxes[i] is None:
            continue
        x0, y0, x1, y1 = raw_boxes[i]
        w, h = x1 - x0, y1 - y0
        area = w * h
        # Exclude full-height narrow sidebars and thin top banners
        if w < slide_w * 0.22 and h > slide_h * 0.55:
            continue
        if y1 < slide_h * 0.24 and h < slide_h * 0.14:
            continue
        if slide_area * 0.04 <= area <= slide_area * 0.65:
            candidates.append((i, raw_boxes[i]))

    # Keep non-overlapping primary panels (drop nested sub-panels inside a larger panel)
    zones = []
    for i, box in sorted(candidates, key=lambda p: (p[1][1] // max(1.0, slide_h * 0.12), p[1][0])):
        if any(_overlap_ratio(box, zbox) > 0.50 for _, zbox in zones):
            continue
        zones.append((i, box))
    return zones


def auto_assign_waves_and_lanes(actions, lanes, slide_w, slide_h=540.0):
    """Calm, spatially-zoned Codex choreography:
    - Wave 0: Slide skeleton (top-to-bottom cascade)
    - Wave 1: Top header title, subtitle & top-to-bottom sidebar title
    - Wave 2: Each card/panel zone owned cleanly by ONE dedicated cursor flowing smoothly left->right!
    - Wave 4: Footer & page pill signature
    """
    slide_area = max(1.0, float(slide_w) * float(slide_h))
    explicit_wave = [bool("wave" in a and a["wave"] is not None) for a in actions]
    explicit_lane = [bool(a.get("lane")) for a in actions]

    for idx, a in enumerate(actions):
        a["wave"] = default_wave(a, slide_w, slide_h)
        if "z_index" not in a or a["z_index"] is None:
            a["z_index"] = _default_z_layer(a, idx, slide_area)

    raw_boxes = [bounds(a) if a["type"] != "group_shapes" else None for a in actions]

    # 1. Promote large textless background shapes sitting behind higher-z items to wave 0
    for i, a in enumerate(actions):
        if a["type"] == "add_shape" and raw_boxes[i] is not None and not extract_plain_text(a).strip():
            x0, y0, x1, y1 = raw_boxes[i]
            if (x1 - x0) * (y1 - y0) >= slide_area * 0.035:
                has_above = any(
                    raw_boxes[j] is not None
                    and float(actions[j]["z_index"]) > float(a["z_index"])
                    and _overlap_ratio(raw_boxes[i], raw_boxes[j]) > 0.10
                    for j in range(len(actions))
                    if j != i
                )
                if has_above:
                    a["wave"] = 0

    # 2. Enforce strict top-to-bottom reading chain on vertically stacked sidebar textboxes (prevents "纲" before "汇"!)
    sidebar_texts = []
    for i, a in enumerate(actions):
        if a["type"] == "add_textbox" and raw_boxes[i] is not None:
            x0, y0, x1, y1 = raw_boxes[i]
            if x1 <= slide_w * 0.19 and y0 < slide_h * 0.88:
                a["wave"] = 1
                sidebar_texts.append((y0, i, a))
    sidebar_texts.sort(key=lambda t: t[0])
    for prev_t, cur_t in zip(sidebar_texts, sidebar_texts[1:]):
        prev_a, cur_a = prev_t[2], cur_t[2]
        cur_a.setdefault("after", [])
        if prev_a["name"] not in cur_a["after"]:
            cur_a["after"].append(prev_a["name"])
        cur_a["_chain_parent"] = prev_a["name"]
        # Ensure z_index order agrees with top-to-bottom dependency chain when input was reversed
        if float(cur_a["z_index"]) <= float(prev_a["z_index"]):
            cur_a["z_index"] = float(prev_a["z_index"]) + 0.1

    # 3. Automatic Spatial Panel-Zone Ownership (eliminates erratic cross-slide jumping!)
    zones = _detect_panel_zones(actions, raw_boxes, slide_w, slide_h)
    zoned_mode = len(zones) >= 2

    if zoned_mode:
        for i, a in enumerate(actions):
            if raw_boxes[i] is None or a["type"] == "group_shapes":
                continue
            w_cur = int(a["wave"])
            if w_cur in (0, 4):
                continue
            bx0, by0, bx1, by1 = raw_boxes[i]
            cx, cy = (bx0 + bx1) * 0.5, (by0 + by1) * 0.5
            # Keep top slide header & sidebar text in Wave 1
            if w_cur == 1 and (cy < slide_h * 0.22 or bx1 <= slide_w * 0.19):
                continue
            matched_zone = None
            for z_idx, (_, zb) in enumerate(zones):
                if zb[0] - 8.0 <= cx <= zb[2] + 8.0 and zb[1] - 6.0 <= cy <= zb[3] + 6.0:
                    matched_zone = z_idx
                    break
            if matched_zone is not None:
                a["_zone"] = matched_zone
                # Unify all elements inside a card/panel into Wave 2 so ONE cursor sweeps that card L->R seamlessly!
                if not explicit_wave[i]:
                    a["wave"] = 2
                elif w_cur in (1, 3):
                    # Even if author put badge in wave 1 and copy in wave 3, allow smooth card-level flow
                    # unless author explicitly set custom multi-stage waves for non-card reasons
                    a["wave"] = 2

    # 4. Badge-Text Coupling: pair any short textbox strictly inside a small badge/pill shape or freeform ribbon
    for j, tb in enumerate(actions):
        if tb["type"] != "add_textbox" or raw_boxes[j] is None:
            continue
        txt = extract_plain_text(tb).strip()
        if not txt or len(txt) > 14:
            continue
        tb_box = raw_boxes[j]
        tb_area = max(1e-6, (tb_box[2] - tb_box[0]) * (tb_box[3] - tb_box[1]))
        best_parent = None
        best_parent_area = float("inf")
        for i, sh in enumerate(actions):
            if i == j or sh["type"] not in ("add_shape", "freeform") or raw_boxes[i] is None:
                continue
            if extract_plain_text(sh).strip():
                continue
            sb = raw_boxes[i]
            sw_box, sh_box = sb[2] - sb[0], sb[3] - sb[1]
            if sw_box > slide_w * 0.16 or sh_box > slide_h * 0.11:
                continue
            ix = max(0.0, min(tb_box[2], sb[2]) - max(tb_box[0], sb[0]))
            iy = max(0.0, min(tb_box[3], sb[3]) - max(tb_box[1], sb[1]))
            if (ix * iy) / tb_area >= 0.72 and float(sh["z_index"]) <= float(tb["z_index"]):
                area_s = sw_box * sh_box
                if area_s < best_parent_area:
                    best_parent_area = area_s
                    best_parent = i
        if best_parent is not None:
            parent = actions[best_parent]
            target_w = min(parent["wave"], tb["wave"]) if not explicit_wave[best_parent] else parent["wave"]
            if target_w == 0:
                target_w = 1
            parent["wave"] = target_w
            tb["wave"] = target_w
            tb.setdefault("after", [])
            if parent["name"] not in tb["after"]:
                tb["after"].append(parent["name"])
            tb["_paired_badge"] = parent["name"]
            if "_zone" in parent:
                tb["_zone"] = parent["_zone"]

    # 5. Enforce strict Z-order wave monotonicity for overlapping pairs
    changed = True
    passes = 0
    while changed and passes < 8:
        changed = False
        passes += 1
        for i in range(len(actions)):
            bi = raw_boxes[i]
            if bi is None:
                continue
            zi = float(actions[i]["z_index"])
            for j in range(len(actions)):
                if i == j or raw_boxes[j] is None:
                    continue
                zj = float(actions[j]["z_index"])
                if zi < zj and _overlap_ratio(bi, raw_boxes[j]) > 0.10:
                    if actions[i]["wave"] > actions[j]["wave"]:
                        actions[i]["wave"] = actions[j]["wave"]
                        changed = True

    # 6. Compute smooth column-bucketed left-to-right reading order (`_seq_idx`) and full-swarm balanced lanes
    zone_boxes = {z_idx: zb for z_idx, (_, zb) in enumerate(zones)}

    def _smooth_order_key(idx):
        a = actions[idx]
        w_idx = int(a["wave"])
        b = raw_boxes[idx]
        if b is None:
            return (w_idx, 99, 99, 1e9, 1e9, float(a.get("z_index", 0)))
        x0, y0, x1, y1 = b
        cx, cy = (x0 + x1) * 0.5, (y0 + y1) * 0.5
        z_id = a.get("_zone", -1)
        if z_id >= 0 and z_id in zone_boxes:
            zb = zone_boxes[z_id]
            zw = max(1.0, zb[2] - zb[0])
            zh = max(1.0, zb[3] - zb[1])
            rel_x = (cx - zb[0]) / zw
            if rel_x < 0.34:
                col_b = 0
            elif rel_x < 0.68:
                col_b = 1
            else:
                col_b = 2
            row_b = int(((y0 - zb[1]) / zh) * 4.0) if col_b != 1 else 0
            kind_pri = {"add_shape": 0, "add_image": 1, "add_line": 2, "freeform": 2, "add_textbox": 3}.get(a["type"], 4)
            return (w_idx, z_id, col_b, row_b if col_b != 1 else kind_pri, round(y0 / 12.0), x0)
        return (w_idx, -1, 0, round(y0 / 24.0), x0, float(a.get("z_index", 0)))

    ordered_indices = sorted(range(len(actions)), key=_smooth_order_key)
    for seq_i, idx in enumerate(ordered_indices):
        actions[idx]["_seq_idx"] = seq_i

    if lanes <= 1:
        for a in actions:
            a["lane"] = 1
        return

    name_to_action = {a["name"]: a for a in actions}
    wave0_loads = {ln: 0 for ln in range(1, lanes + 1)}
    num_zones = len(zones)

    # First pass for Wave 0: spread primary root structural panels (sidebar, banner, cards)
    # 1-per-lane across ALL 1..lanes so EVERY cursor launches simultaneously at t=0.0s!
    w0_primary_indices = []
    for idx in ordered_indices:
        a = actions[idx]
        if int(a["wave"]) == 0 and not explicit_lane[idx] and a["type"] == "add_shape" and raw_boxes[idx] is not None:
            b = raw_boxes[idx]
            if (b[2] - b[0]) * (b[3] - b[1]) >= slide_area * 0.03:
                w0_primary_indices.append(idx)
    for k_p, idx in enumerate(w0_primary_indices):
        ln_p = (k_p % lanes) + 1
        actions[idx]["lane"] = ln_p
        wave0_loads[ln_p] += 1

    for idx in ordered_indices:
        a = actions[idx]
        if a.get("lane"):
            if explicit_lane[idx]:
                ln_e = max(1, min(lanes, int(a["lane"])))
                a["lane"] = ln_e
                if int(a["wave"]) == 0:
                    wave0_loads[ln_e] += 1
            continue
        parent_nm = a.get("_chain_parent") or a.get("_paired_badge")
        if parent_nm and parent_nm in name_to_action and name_to_action[parent_nm].get("lane"):
            a["lane"] = int(name_to_action[parent_nm]["lane"])
            continue

        w_idx = int(a["wave"])
        b = raw_boxes[idx]
        z_id = a.get("_zone", -1)

        if w_idx == 0:
            if b is not None:
                cy = (b[1] + b[3]) * 0.5
                pref = int(max(0, min(lanes - 1, (cy / max(1.0, slide_h)) * lanes))) + 1
                chosen = min(range(1, lanes + 1), key=lambda ln: (wave0_loads[ln], abs(ln - pref)))
                a["lane"] = chosen
                wave0_loads[chosen] += 1
            else:
                chosen = min(range(1, lanes + 1), key=lambda ln: wave0_loads[ln])
                a["lane"] = chosen
                wave0_loads[chosen] += 1
        elif w_idx == 1:
            # Spread header & sidebar titles cleanly so they run in parallel without blocking Wave 2
            if b is not None and b[2] <= slide_w * 0.19:
                a["lane"] = 1
            else:
                a["lane"] = min(2, lanes)
        elif z_id >= 0:
            if lanes > num_zones and b is not None and z_id in zone_boxes:
                # Extra cursors (e.g. 6 cursors on 4 cards) cleanly own the Right annotation column (rel_x >= 0.68)
                # while primary cursors 1..num_zones glide smoothly through Left+Center of each card!
                zb = zone_boxes[z_id]
                rel_x = (((b[0] + b[2]) * 0.5) - zb[0]) / max(1.0, zb[2] - zb[0])
                if rel_x >= 0.68:
                    extra_slots = lanes - num_zones
                    a["lane"] = num_zones + (z_id % extra_slots) + 1
                else:
                    a["lane"] = z_id + 1
            else:
                a["lane"] = (z_id % lanes) + 1
        else:
            if b is not None:
                cy = (b[1] + b[3]) * 0.5
                a["lane"] = int(max(0, min(lanes - 1, (cy / max(1.0, slide_h)) * lanes))) + 1
            else:
                a["lane"] = 1


def build_micro_bursts(actions, deps, boxes, slide_w, slide_h):
    """Group spatially clustered micro-elements (ticks, dots, spheres, hatches, tick numbers) into smooth single-cursor sweeps."""
    used = set()
    clusters = []

    # 1. Explicit `batch_group` clusters
    explicit_groups = {}
    for a in actions:
        bg = a.get("batch_group")
        if bg:
            explicit_groups.setdefault((int(a.get("wave", 2)), a.get("_zone", -1), str(bg)), []).append(a)
    for (_, _, bg), members in explicit_groups.items():
        if len(members) >= 2:
            clusters.append((f"burst_{bg}", members))
            for m in members:
                used.add(m["name"])

    # 2. Automatic clustering of remaining micro-actions (separately for geometry primitives vs tiny numeric tick labels)
    for bucket_name, pred, rad_mul, min_sz in (
        ("geom", lambda a: a["type"] != "add_textbox" and is_micro_action(a), 0.14, 4),
        ("ticks", lambda a: is_micro_tick_label(a), 0.22, 4),
    ):
        by_group = {}
        for a in actions:
            if a["name"] in used:
                continue
            if pred(a):
                by_group.setdefault((int(a.get("wave", 2)), a.get("_zone", -1), int(a.get("lane", 1))), []).append(a)

        cluster_rad = min(slide_w, slide_h) * rad_mul
        auto_idx = 0
        for (w_idx, z_id, _), items in sorted(by_group.items()):
            if len(items) < min_sz:
                continue
            visited = set()
            for seed in items:
                sn = seed["name"]
                if sn in visited:
                    continue
                comp = [seed]
                visited.add(sn)
                q_i = 0
                while q_i < len(comp):
                    cur = comp[q_i]
                    q_i += 1
                    cb = boxes[cur["name"]]
                    cc = ((cb[0] + cb[2]) * 0.5, (cb[1] + cb[3]) * 0.5)
                    for cand in items:
                        cn = cand["name"]
                        if cn in visited:
                            continue
                        ob = boxes[cn]
                        oc = ((ob[0] + ob[2]) * 0.5, (ob[1] + ob[3]) * 0.5)
                        if math.dist(cc, oc) <= cluster_rad:
                            visited.add(cn)
                            comp.append(cand)
                if len(comp) >= min_sz:
                    member_names = {m["name"] for m in comp}
                    ext_deps = set()
                    for m in comp:
                        ext_deps.update(deps[m["name"]] - member_names)
                    if not any(deps.get(ed, set()) & member_names for ed in ext_deps):
                        auto_idx += 1
                        clusters.append((f"micro_{bucket_name}_{w_idx}_{z_id}_{auto_idx}", comp))
                        used.update(member_names)

    # Assemble ordered units preserving topological painter order inside each cluster
    member_to_cluster = {}
    for cname, members in clusters:
        member_map = {m["name"]: m for m in members}
        member_names = set(member_map.keys())
        in_deg = {mn: len(deps.get(mn, set()) & member_names) for mn in member_names}
        children = {mn: [] for mn in member_names}
        for mn in member_names:
            for d in deps.get(mn, set()) & member_names:
                children[d].append(mn)

        def sweep_tie_key(mn):
            b = boxes[mn]
            return (round(b[1] / 14.0), b[0], float(member_map[mn].get("z_index", 0)))

        ready = sorted([mn for mn, deg in in_deg.items() if deg == 0], key=sweep_tie_key)
        members_sorted = []
        while ready:
            cur_n = ready.pop(0)
            members_sorted.append(member_map[cur_n])
            for ch in children[cur_n]:
                in_deg[ch] -= 1
                if in_deg[ch] == 0:
                    ready.append(ch)
            ready.sort(key=sweep_tie_key)
        if len(members_sorted) < len(members):
            seen_m = {m["name"] for m in members_sorted}
            members_sorted.extend(m for m in members if m["name"] not in seen_m)

        union_box = (
            min(boxes[m["name"]][0] for m in members_sorted),
            min(boxes[m["name"]][1] for m in members_sorted),
            max(boxes[m["name"]][2] for m in members_sorted),
            max(boxes[m["name"]][3] for m in members_sorted),
        )
        union_deps = set()
        for m in members_sorted:
            union_deps.update(deps[m["name"]] - member_names)
        burst_unit = {
            "type": "micro_burst",
            "name": cname,
            "members": members_sorted,
            "member_names": [m["name"] for m in members_sorted],
            "wave": int(members_sorted[0].get("wave", 2)),
            "lane": int(members_sorted[0].get("lane", 1)),
            "_zone": members_sorted[0].get("_zone", -1),
            "_seq_idx": min(int(m.get("_seq_idx", 0)) for m in members_sorted),
            "z_index": min(float(m.get("z_index", 0)) for m in members_sorted),
            "left": union_box[0],
            "top": union_box[1],
            "width": max(1.0, union_box[2] - union_box[0]),
            "height": max(1.0, union_box[3] - union_box[1]),
            "_burst_deps": union_deps,
            "_burst_box": union_box,
        }
        for m in members_sorted:
            member_to_cluster[m["name"]] = burst_unit

    units = []
    emitted_clusters = set()
    for a in actions:
        cl = member_to_cluster.get(a["name"])
        if cl is not None:
            if cl["name"] not in emitted_clusters:
                emitted_clusters.add(cl["name"])
                units.append(cl)
        else:
            units.append(a)
    units.sort(key=lambda u: (int(u.get("wave", 2)), int(u.get("_seq_idx", 0))))
    return units


def apply_z_order(shapes, ordered_names, boxes=None):
    """Overlap-only zero-flash Z-order verifier."""
    if len(ordered_names) < 2:
        return
    present = {}
    for i in range(1, shapes.Count + 1):
        sh = shapes.Item(i)
        present[sh.Name] = (i, sh)
    target = [name for name in ordered_names if name in present]
    if len(target) < 2:
        return

    rects = {}
    for name in target:
        if boxes and name in boxes:
            rects[name] = boxes[name]
        else:
            sh = present[name][1]
            rects[name] = (float(sh.Left), float(sh.Top), float(sh.Left + sh.Width), float(sh.Top + sh.Height))

    inverted_fg = set()
    n = len(target)
    for i in range(n):
        ni = target[i]
        pi = present[ni][0]
        ri = rects[ni]
        for j in range(i + 1, n):
            nj = target[j]
            if present[nj][0] < pi and intersects(ri, rects[nj]):
                inverted_fg.add(nj)

    if not inverted_fg:
        return

    to_lift = set(inverted_fg)
    for j in range(n):
        nj = target[j]
        if nj in to_lift:
            rj = rects[nj]
            for k in range(j + 1, n):
                nk = target[k]
                if intersects(rj, rects[nk]):
                    to_lift.add(nk)

    for name in target:
        if name in to_lift:
            try:
                present[name][1].ZOrder(0)
            except Exception:
                pass


def resolve_connector_endpoints(actions, existing_boxes=None):
    """If an add_line specifies `from`/`to` object names instead of coordinates, derive clean boundary endpoints."""
    boxes = dict(existing_boxes or {})
    for a in actions:
        if a["type"] == "add_line" and "begin_x" not in a and a.get("from") in boxes and a.get("to") in boxes:
            b1 = boxes[a["from"]]
            b2 = boxes[a["to"]]
            c1 = ((b1[0] + b1[2]) * 0.5, (b1[1] + b1[3]) * 0.5)
            c2 = ((b2[0] + b2[2]) * 0.5, (b2[1] + b2[3]) * 0.5)
            dx, dy = c2[0] - c1[0], c2[1] - c1[1]
            if abs(dx) >= abs(dy):
                sx = b1[2] if dx >= 0 else b1[0]
                ex = b2[0] if dx >= 0 else b2[2]
                sy, ey = c1[1], c2[1]
            else:
                sy = b1[3] if dy >= 0 else b1[1]
                ey = b2[1] if dy >= 0 else b2[3]
                sx, ex = c1[0], c2[0]
            a["begin_x"], a["begin_y"], a["end_x"], a["end_y"] = round(sx, 2), round(sy, 2), round(ex, 2), round(ey, 2)
            a.setdefault("after", [])
            for dep in (a["from"], a["to"]):
                if dep not in a["after"]:
                    a["after"].append(dep)
        if a["type"] not in ("group_shapes", "duplicate_shape") and (
            a["type"] != "add_line" or all(k in a for k in ("begin_x", "begin_y", "end_x", "end_y"))
        ):
            boxes[a["name"]] = bounds(a)


def validate(actions, existing, lanes):
    """Validate actions and build fast wave-parallel dependency graph."""
    known = set(existing)
    folded = {n.casefold() for n in known}
    rectangles = {}
    waves = {}
    dependencies = {}
    for a in actions:
        n = a["name"]
        kind = a["type"]
        if not isinstance(n, str) or not n.strip() or n.casefold() in folded:
            raise ValueError(f"Missing or duplicate object name: {n}")
        if not 1 <= int(a.get("lane", 1)) <= lanes:
            raise ValueError(f"Invalid lane {a.get('lane')} on {n} (must be 1..{lanes})")
        wave = int(a.get("wave", 1))
        deps = set(a.get("after", []))
        if kind == "duplicate_shape":
            deps.add(a["source"])
        if kind == "group_shapes":
            group = a["members"]
            if len(group) < 2 or len(group) != len(set(group)) or any(m not in rectangles for m in group):
                raise ValueError(f"Group requires >=2 distinct prior members: {n}")
            deps.update(group)
            boxes = [rectangles[m] for m in group]
            rect = (
                min(b[0] for b in boxes),
                min(b[1] for b in boxes),
                max(b[2] for b in boxes),
                max(b[3] for b in boxes),
            )
        elif kind in ("add_shape", "add_textbox", "add_image", "duplicate_shape", "freeform", "add_line"):
            if kind == "freeform":
                path = a["path"]
                allowed = {"M": 3, "L": 3, "C": 7, "Z": 1}
                if (
                    len(path) < 2
                    or path[0][0] != "M"
                    or any(cmd[0] == "M" for cmd in path[1:])
                    or any(
                        len(cmd) != allowed.get(cmd[0]) or not all(math.isfinite(v) for v in cmd[1:])
                        for cmd in path
                    )
                ):
                    raise ValueError(f"Expected finite M/L/C/Z path on {n}")
            if kind == "add_shape" and "shape_type_id" not in a and a.get("shape") not in PRESETS:
                raise ValueError(f"Unknown shape preset {a.get('shape')!r} on {n}; pass a valid name or shape_type_id")
            rect = bounds(a)
            if not all(math.isfinite(v) for v in rect) or rect[2] < rect[0] or rect[3] < rect[1]:
                raise ValueError(f"Invalid geometry on {n}")
            if kind not in ("add_line", "freeform") and (rect[2] == rect[0] or rect[3] == rect[1]):
                raise ValueError(f"Width and height must be positive on {n}")
        else:
            raise ValueError(f"Unsupported operation type: {kind}")

        if not deps <= known:
            raise ValueError(f"{n} depends on missing or later objects: {deps - known}")
        if lanes > 1:
            for k, b in rectangles.items():
                kw = waves.get(k, 1)
                if kw < wave and intersects(b, rect):
                    deps.add(k)
                elif kw == wave and _overlap_ratio(b, rect) > 0.15:
                    deps.add(k)
        dependencies[n] = deps
        rectangles[n] = rect
        waves[n] = wave
        folded.add(n.casefold())
        known.add(n)
        if kind == "group_shapes":
            known.difference_update(a["members"])
    return dependencies, rectangles


def lint_plan(plan, actions, boxes, slide_w, slide_h, units=None):
    """Preflight quality linter catching CJK width overflow, split-line text, late titles, and pacing issues."""
    warnings = []
    textboxes = []
    body_waves = []
    title_waves = []

    for a in actions:
        kind = a["type"]
        nm = a["name"]
        b = boxes.get(nm)
        if b is not None:
            if b[0] < -2.0 or b[1] < -2.0 or b[2] > slide_w + 2.0 or b[3] > slide_h + 2.0:
                warnings.append({
                    "code": "out_of_canvas",
                    "name": nm,
                    "detail": f"Bounds {[round(v, 1) for v in b]} exceed canvas {slide_w}x{slide_h}.",
                })
        txt = extract_plain_text(a).strip()
        if kind == "add_shape" and txt:
            preset_id = a.get("shape_type_id") or PRESETS.get(a.get("shape", "rectangle"))
            if preset_id not in (1, 5, 9):
                warnings.append({
                    "code": "non_rect_embedded_text_risk",
                    "name": nm,
                    "detail": (
                        f"Non-rectangular shape (preset={a.get('shape', preset_id)}) embeds text '{txt[:12]}'; "
                        "PowerPoint often shifts/clips embedded text on slanted or custom AutoShapes. "
                        "Prefer a freeform/shape polygon (no text) + a coupled centered add_textbox on top."
                    ),
                })
        if kind == "add_textbox" and txt:
            textboxes.append((a, b))
            fs = float(a.get("font_size", 14.0))
            if (b[1] < slide_h * 0.22 and fs >= 16.0) or fs >= 20.0:
                title_waves.append((nm, int(a.get("wave", 2))))
            elif len(txt) >= 10 and fs <= 14.5:
                body_waves.append(int(a.get("wave", 2)))

            metrics = estimate_text_metrics(a)
            w = b[2] - b[0]
            h = b[3] - b[1]
            wrap = bool(a.get("word_wrap", True))
            if has_cjk(txt) and str(a.get("alignment", "left")).lower() == "justify" and (not wrap or "\n" not in txt):
                warnings.append({
                    "code": "justify_cjk_single_line",
                    "name": nm,
                    "detail": "Single-line CJK textbox uses alignment='justify'; engine will downgrade to 'left' to prevent glyph crushing.",
                })
            if not wrap and metrics["max_line_w"] > w * 1.04:
                warnings.append({
                    "code": "cjk_width_risk",
                    "name": nm,
                    "detail": f"Single-line text '{txt[:18]}' est_width={metrics['max_line_w']}pt > box width={w:.1f}pt (+10% CJK headroom recommended).",
                })
            elif wrap and metrics["est_height"] > h * 1.18:
                warnings.append({
                    "code": "text_height_risk",
                    "name": nm,
                    "detail": f"Wrapped text '{txt[:18]}' est_height={metrics['est_height']}pt > box height={h:.1f}pt (risk of unintended wrap/clip).",
                })

    for i in range(len(textboxes)):
        a1, b1 = textboxes[i]
        for j in range(i + 1, len(textboxes)):
            a2, b2 = textboxes[j]
            if abs(b1[1] - b2[1]) <= 3.0 and abs(b1[3] - b2[3]) <= 4.0:
                gap = max(b2[0] - b1[2], b1[0] - b2[2])
                if -4.0 <= gap <= 4.0:
                    t1 = extract_plain_text(a1).strip()
                    t2 = extract_plain_text(a2).strip()
                    if has_cjk(t1 + t2) and len(t1) >= 2 and len(t2) >= 2:
                        warnings.append({
                            "code": "split_inline_text_risk",
                            "name": f"{a1['name']} + {a2['name']}",
                            "detail": "Adjacent same-baseline CJK textboxes detected; prefer one add_textbox with `runs` to prevent overlap.",
                        })

    if title_waves and body_waves:
        min_body_w = min(body_waves)
        for nm, tw in title_waves:
            if tw > min_body_w:
                warnings.append({
                    "code": "late_title_order",
                    "name": nm,
                    "detail": f"Hero title '{nm}' is in wave {tw} after body copy (wave {min_body_w}); move titles to wave 1.",
                })

    burst_count = sum(1 for u in (units or []) if u.get("type") == "micro_burst")
    burst_members = sum(len(u["members"]) for u in (units or []) if u.get("type") == "micro_burst")
    zones_used = sorted({a.get("_zone") for a in actions if a.get("_zone", -1) >= 0})
    return {
        "warning_count": len(warnings),
        "warnings": warnings,
        "detected_zones": len(zones_used),
        "micro_bursts": burst_count,
        "micro_burst_members": burst_members,
    }


def audit_slide_shapes(shapes, actions, slide_w, slide_h, auto_fix_text_overflow=True):
    """Run real PowerPoint COM geometry & text-fit integrity audit right before saving."""
    action_by_name = {a["name"]: a for a in actions}
    present = {}
    for i in range(1, shapes.Count + 1):
        sh = shapes.Item(i)
        present[sh.Name] = sh

    hard_findings = []
    warnings = []
    auto_fixes = []
    textbox_rects = []
    checked = 0

    for a in actions:
        nm = a["name"]
        if nm not in present or a["type"] == "group_shapes":
            continue
        sh = present[nm]
        checked += 1
        left = float(sh.Left)
        top = float(sh.Top)
        width = float(sh.Width)
        height = float(sh.Height)
        right = left + width
        bottom = top + height

        if left < -1.5 or top < -1.5 or right > slide_w + 1.5 or bottom > slide_h + 1.5:
            hard_findings.append({
                "code": "out_of_canvas",
                "name": nm,
                "rect": [round(left, 1), round(top, 1), round(width, 1), round(height, 1)],
            })

        txt = extract_plain_text(a).strip()
        if not txt or a["type"] not in ("add_textbox", "add_shape"):
            continue

        try:
            tr = sh.TextFrame.TextRange
            bw = float(tr.BoundWidth)
            bh = float(tr.BoundHeight)
            bl = float(getattr(tr, "BoundLeft", left))
            bt = float(getattr(tr, "BoundTop", top))
        except Exception:
            continue

        if a["type"] == "add_shape" and (
            bl < left - 2.0 or bt < top - 2.0 or (bl + bw) > right + 2.0 or (bt + bh) > bottom + 2.0
        ):
            hard_findings.append({
                "code": "shape_text_out_of_bounds",
                "name": nm,
                "shape_rect": [round(left, 1), round(top, 1), round(right, 1), round(bottom, 1)],
                "text_rect": [round(bl, 1), round(bt, 1), round(bl + bw, 1), round(bt + bh, 1)],
                "text": txt[:24],
            })

        # COM reports page-aligned text bounds. Compare a quarter-turned
        # textbox in its local axes, retaining the original page-space bounds above.
        quarter_turn = abs(float(sh.Rotation) % 180.0 - 90.0) < 0.001
        if quarter_turn:
            bw, bh = bh, bw

        if a["type"] == "add_textbox":
            textbox_rects.append((nm, (left, top, right, bottom), txt))

        wrap = bool(a.get("word_wrap", True))
        if a["type"] == "add_textbox" and bw > width + 1.5 and (not wrap or "\n" not in txt):
            needed_w = bw + 4.0
            max_right = slide_w - 4.0
            for other_name, other_sh in present.items():
                if other_name == nm:
                    continue
                other_a = action_by_name.get(other_name)
                if other_a and int(other_a.get("wave", 2)) == 0:
                    o_left = float(other_sh.Left)
                    o_top = float(other_sh.Top)
                    o_right = o_left + float(other_sh.Width)
                    o_bottom = o_top + float(other_sh.Height)
                    if o_left <= left + 4.0 and o_right >= right - 4.0 and o_top <= top + 4.0 and o_bottom >= bottom - 4.0:
                        max_right = min(max_right, o_right - 4.0)
                    continue
                ol = float(other_sh.Left)
                ot = float(other_sh.Top)
                ob = ot + float(other_sh.Height)
                if ol >= right - 1.0 and min(bottom, ob) > max(top, ot) + 2.0:
                    max_right = min(max_right, ol - 3.0)

            safe_w = max(width, max_right - left)
            if auto_fix_text_overflow and not quarter_turn and needed_w <= safe_w + 0.5:
                sh.Width = needed_w
                auto_fixes.append({
                    "code": "auto_widened_textbox",
                    "name": nm,
                    "old_width": round(width, 1),
                    "new_width": round(needed_w, 1),
                    "bound_width": round(bw, 1),
                })
                width = needed_w
                right = left + width
                if textbox_rects and textbox_rects[-1][0] == nm:
                    textbox_rects[-1] = (nm, (left, top, right, bottom), txt)
            else:
                hard_findings.append({
                    "code": "text_horizontal_overflow",
                    "name": nm,
                    "width": round(width, 1),
                    "bound_width": round(bw, 1),
                    "available_width": round(safe_w, 1),
                    "text": txt[:24],
                })

        if bh > height + 3.0 and bh > float(a.get("font_size", 14.0)) * 1.45:
            hard_findings.append({
                "code": "text_vertical_overflow",
                "name": nm,
                "height": round(height, 1),
                "bound_height": round(bh, 1),
                "text": txt[:24],
            })

    for i in range(len(textbox_rects)):
        n1, r1, t1 = textbox_rects[i]
        for j in range(i + 1, len(textbox_rects)):
            n2, r2, t2 = textbox_rects[j]
            if _overlap_ratio(r1, r2) > 0.20:
                warnings.append({
                    "code": "textbox_overlap",
                    "names": [n1, n2],
                    "overlap_ratio": round(_overlap_ratio(r1, r2), 2),
                    "texts": [t1[:16], t2[:16]],
                })

    return {
        "checked_shapes": checked,
        "hard_finding_count": len(hard_findings),
        "warning_count": len(warnings),
        "auto_fix_count": len(auto_fixes),
        "hard_findings": hard_findings,
        "warnings": warnings,
        "auto_fixes": auto_fixes,
    }
