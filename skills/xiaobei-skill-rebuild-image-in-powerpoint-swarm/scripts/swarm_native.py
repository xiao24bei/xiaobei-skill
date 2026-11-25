"""Fast native PowerPoint COM writers, schema normalizer, and wave scheduler."""
import math

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


def rgb(hex_color):
    h = str(hex_color).strip().lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return int.from_bytes(bytes.fromhex(h), "little")


def has_cjk(text):
    return any("\u4e00" <= ch <= "\u9fff" or "\u3000" <= ch <= "\u303f" or "\uff00" <= ch <= "\uffef" for ch in text)


def normalize_action(raw, index=1):
    """Accept both canonical swarm-plan actions and powerpoint_draw_sequence style dicts."""
    a = dict(raw)
    raw_type = str(a.get("type") or a.get("op") or a.get("action") or "add_shape").strip().lower()
    kind = TYPE_ALIASES.get(raw_type, raw_type)
    a["type"] = kind
    a["name"] = str(a.get("name") or a.get("id") or f"obj_{index:03d}")

    # Normalize box coordinates (support x/y/w/h aliases)
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

    # Normalize line coordinates (support x1/y1/x2/y2 aliases)
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

    # Normalize image path aliases
    if kind == "add_image" and "image_path" not in a:
        for k in ("asset", "path", "src", "file"):
            if a.get(k):
                a["image_path"] = a[k]
                break

    # Normalize shape preset aliases
    if kind == "add_shape" and "shape" not in a and "shape_type_id" not in a:
        a["shape"] = a.get("preset") or a.get("shape_type") or "rectangle"
    if kind == "add_shape" and isinstance(a.get("shape"), str):
        a["shape"] = a["shape"].strip().lower().replace("-", "_").replace(" ", "_")

    # Normalize style aliases
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


def _apply_font_props(font_obj, spec, defaults, text_sample=""):
    latin_name = spec.get("font_name", defaults.get("font_name", "Arial"))
    ea_name = spec.get("font_name_east_asian", defaults.get("font_name_east_asian"))
    if not ea_name and has_cjk(text_sample):
        ea_name = "Microsoft YaHei"
    font_obj.Name = latin_name
    if ea_name:
        try:
            font_obj.NameFarEast = ea_name
        except Exception:
            pass
    font_obj.Size = float(spec.get("font_size", defaults.get("font_size", 14)))
    font_obj.Bold = -int(spec.get("bold", defaults.get("bold", False)))
    font_obj.Italic = -int(spec.get("italic", defaults.get("italic", False)))
    color = spec.get("font_color", spec.get("color", defaults.get("font_color", "161616")))
    font_obj.Color.RGB = rgb(color)
    if spec.get("underline", defaults.get("underline", False)):
        try:
            font_obj.Underline = -1
        except Exception:
            pass
    if spec.get("superscript"):
        try:
            font_obj.BaselineOffset = 0.3
        except Exception:
            pass
    elif spec.get("subscript"):
        try:
            font_obj.BaselineOffset = -0.25
        except Exception:
            pass


def write_text_fast(shape, action):
    """Write complete styled text (plain or rich runs) in minimal atomic COM calls."""
    frame = shape.TextFrame
    frame.AutoSize = 0
    try:
        shape.TextFrame2.AutoSize = 0
    except Exception:
        pass
    frame.WordWrap = -int(action.get("word_wrap", True))
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
    _apply_font_props(tr.Font, action, action, full_text)
    p = tr.ParagraphFormat
    p.Alignment = ALIGNMENTS.get(str(action.get("alignment", "left")).lower(), 1)
    p.SpaceBefore = float(action.get("space_before", 0))
    p.SpaceAfter = float(action.get("space_after", 0))
    p.LineRuleWithin = -1
    p.SpaceWithin = float(action.get("line_spacing", 1))

    if runs:
        cursor = 1
        for rspec in runs:
            rtext = str(rspec.get("text", "")).replace("\r\n", "\r").replace("\n", "\r")
            if not rtext:
                continue
            sub = tr.Characters(cursor, len(rtext))
            _apply_font_props(sub.Font, rspec, action, rtext)
            cursor += len(rtext)


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


def default_wave(action, slide_area=518400.0):
    if "wave" in action and action["wave"] is not None:
        return int(action["wave"])
    kind = action["type"]
    if kind == "add_shape":
        # Large background panels / zone containers go in wave 0 so they finish cleanly
        # before any assets or nested cards begin (preventing mid-show occlusion or Z-flashes).
        x0, y0, x1, y1 = bounds(action)
        area = max(0.0, (x1 - x0) * (y1 - y0))
        nm = action.get("name", "").lower()
        is_bg_name = any(tok in nm for tok in ("_bg", "panel", "backdrop", "container", "zone", "section", "frame", "background"))
        has_text = bool(action.get("text") or action.get("runs"))
        if (area >= slide_area * 0.04 and not has_text) or (is_bg_name and area >= slide_area * 0.025 and not has_text):
            return 0
        return 1
    if kind == "add_image":
        return 2
    if kind in ("add_line", "freeform"):
        return 3
    if kind == "add_textbox":
        return 4
    return 5


def auto_assign_waves_and_lanes(actions, lanes, slide_w, slide_h=540.0):
    """Ensure every action has a Z-order-consistent wave, z_index, and spatially balanced lane."""
    slide_area = max(1.0, float(slide_w) * float(slide_h))
    for idx, a in enumerate(actions):
        a["wave"] = default_wave(a, slide_area)
        if "z_index" not in a or a["z_index"] is None:
            a["z_index"] = a["wave"] * 1000 + idx * 10

    # Promote large textless background shapes that sit behind other objects to wave 0.
    raw_boxes = [bounds(a) if a["type"] != "group_shapes" else None for a in actions]
    for i, a in enumerate(actions):
        if a["type"] == "add_shape" and raw_boxes[i] is not None and not (a.get("text") or a.get("runs")):
            x0, y0, x1, y1 = raw_boxes[i]
            if (x1 - x0) * (y1 - y0) >= slide_area * 0.04:
                has_above = any(
                    raw_boxes[j] is not None
                    and float(actions[j]["z_index"]) > float(a["z_index"])
                    and _overlap_ratio(raw_boxes[i], raw_boxes[j]) > 0.10
                    for j in range(len(actions))
                    if j != i
                )
                if has_above:
                    a["wave"] = 0

    # Enforce strict Z-order wave monotonicity for overlapping pairs:
    # if A has lower z_index than overlapping B, A must never be in a later wave than B.
    # This guarantees PowerPoint creates every overlapping pair in exact painter's order natively!
    changed = True
    passes = 0
    while changed and passes < 6:
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

    if lanes <= 1:
        for a in actions:
            a["lane"] = 1
        return

    by_wave = {}
    for a in actions:
        by_wave.setdefault(a["wave"], []).append(a)

    for w, group in sorted(by_wave.items()):
        unassigned = [a for a in group if not a.get("lane")]
        if not unassigned:
            continue

        def spatial_key(a):
            x0, y0, x1, y1 = bounds(a)
            return ((x0 + x1) * 0.5) / max(1.0, slide_w) + ((y0 + y1) * 0.5) / max(1.0, slide_w) * 0.18

        unassigned.sort(key=spatial_key)
        active = min(lanes, len(unassigned))
        for i, a in enumerate(unassigned):
            a["lane"] = min(active, int(i * active / len(unassigned)) + 1)


def _overlap_ratio(a, b):
    ix = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    inter = ix * iy
    if inter <= 0:
        return 0.0
    area_a = max(1e-6, (a[2] - a[0]) * (a[3] - a[1]))
    area_b = max(1e-6, (b[2] - b[0]) * (b[3] - b[1]))
    return inter / min(area_a, area_b)


def apply_z_order(shapes, ordered_names, boxes=None):
    """Overlap-only zero-flash Z-order verifier.

    Non-overlapping shapes do NOT affect each other visually, so we NEVER shuffle them
    (avoiding multi-second Z-order thrashing or accidental hiding behind white parent panels).
    If every overlapping pair is already in exact painter's order (the normal case thanks to
    wave monotonicity), this returns immediately with ZERO COM mutations.
    """
    if len(ordered_names) < 2:
        return
    present = {}
    for i in range(1, shapes.Count + 1):
        sh = shapes.Item(i)
        present[sh.Name] = (i, sh)
    target = [name for name in ordered_names if name in present]
    if len(target) < 2:
        return

    # Build bounding boxes for overlap checks
    rects = {}
    for name in target:
        if boxes and name in boxes:
            rects[name] = boxes[name]
        else:
            sh = present[name][1]
            rects[name] = (float(sh.Left), float(sh.Top), float(sh.Left + sh.Width), float(sh.Top + sh.Height))

    # Find ONLY foreground shapes that are genuinely hidden behind an overlapping lower-z shape
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

    # Also include any higher-rank shapes overlapping the moved shapes so they stay on top of them
    to_lift = set(inverted_fg)
    for j in range(n):
        nj = target[j]
        if nj in to_lift:
            rj = rects[nj]
            for k in range(j + 1, n):
                nk = target[k]
                if intersects(rj, rects[nk]):
                    to_lift.add(nk)

    # Lift ONLY the minimal affected foreground chain in ascending target rank order
    for name in target:
        if name in to_lift:
            try:
                present[name][1].ZOrder(0)  # msoBringToFront
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
            # 1. Always wait for overlapping shapes from EARLIER waves (ensures backgrounds finish
            #    before assets, and assets finish before arrows/labels on top of them).
            # 2. Within the SAME wave, serialize ONLY truly nested/overlapping shapes (>15% area
            #    overlap) in z_index order so shapes land natively in exact painter's order with
            #    ZERO post-hoc Z-order flashing, while non-overlapping shapes stay 100% parallel!
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
