"""Silky 60fps Lockstep Codex-in-PowerPoint multi-agent live drawing engine.

Architecture:
  - Main Thread: Lightweight COM-free 60fps Tkinter transparent overlay (with Win32 1ms
    multimedia timer precision `timeBeginPeriod(1)` and background-matched chroma key to
    eliminate dark ClearType anti-aliasing halos), cropped tightly to slide + HUD margin
    while hardware DXGI `ddagrab` captures the full PowerPoint window at 60fps.
  - Hidden Pre-Build + Priority Reveal Lockstep:
    * The moment a cursor claims an object (before travel/drag/type), the background STA COM
      worker pre-builds the native PowerPoint object with `Visible = 0` in parallel with the
      cursor's smooth 60fps motion.
    * When the cursor completes its stroke/word/chunk, a Priority-0 `reveal` flips `Visible = -1`
      in ~2–8ms (with an 8ms paint breather so PowerPoint flushes `WM_PAINT` within 1 frame),
      and the Tk preview clears IMMEDIATELY (`hold_s = 0.0`) — guaranteeing:
      1. ZERO double-image ghosting ("重影") — Tk preview and PowerPoint native object are
         NEVER double-rendered on top of each other.
      2. ZERO cursor-to-object lag — no cursor ever leaves an object (or micro-burst chunk)
         before PowerPoint has visibly painted it directly under the cursor tip.
      3. ZERO dead-stop freezes — elastic follow-through gently eases the final 12% of a
         stroke if COM is momentarily saturated.
  - Built-in Post-Draw Integrity Audit: Inspects real PowerPoint COM `BoundLeft`/`BoundTop`/
    `BoundWidth`/`BoundHeight` before saving & exporting `preview.png`.
"""
import argparse
from bisect import bisect_left
from collections import deque
from contextlib import contextmanager
import ctypes
import hashlib
import json
import math
from pathlib import Path
import queue
import random
import sys
import threading
import time
import traceback

try:
    import tkinter as tk
    import tkinter.font as tkfont
except ImportError:
    tk = None
    tkfont = None

try:
    import pythoncom
    import win32api
    import win32con as wc
    import win32gui as gui
    import win32process
    from win32com.client import GetActiveObject
    HAS_WIN32 = True
except ImportError:
    pythoncom = win32api = wc = gui = win32process = GetActiveObject = None
    HAS_WIN32 = False

from swarm_native import (
    PRESETS,
    STAGE_NAMES,
    apply_background,
    apply_shadow,
    apply_z_order,
    audit_slide_shapes,
    auto_assign_waves_and_lanes,
    bounds,
    build_micro_bursts,
    char_em_width,
    curve_samples,
    draw_path,
    extract_plain_text,
    fill,
    has_cjk,
    is_micro_action,
    lint_plan,
    normalize_action,
    path_commands,
    path_segments,
    resolve_connector_endpoints,
    set_adjustments,
    stroke,
    validate,
    write_text_fast,
)
from swarm_record import Recorder

KEY_LIGHT = "#FDFDFE"
KEY_DARK = "#010203"
KEY = KEY_LIGHT


def _pick_chroma_key(bg_spec):
    """Choose near-white chroma key on light slides (eliminates dark ClearType halos) and near-black on dark slides."""
    col = None
    if isinstance(bg_spec, dict):
        col = bg_spec.get("fill_color") or bg_spec.get("color")
        if not col and bg_spec.get("gradient"):
            col = bg_spec["gradient"][0]
    if col:
        h = str(col).strip().lstrip("#")
        if len(h) == 3:
            h = "".join(c * 2 for c in h)
        if len(h) == 6:
            try:
                r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
                if (0.299 * r + 0.587 * g + 0.114 * b) < 128.0:
                    return KEY_DARK
            except ValueError:
                pass
    return KEY_LIGHT

DEFAULT_PROFILES = [
    {"name": "Codex-1", "color": "#2563EB", "tint": "#93C5FD"},
    {"name": "Codex-2", "color": "#7C3AED", "tint": "#C4B5FD"},
    {"name": "Codex-3", "color": "#059669", "tint": "#6EE7B7"},
    {"name": "Codex-4", "color": "#0891B2", "tint": "#67E8F9"},
    {"name": "Codex-5", "color": "#EA580C", "tint": "#FDBA74"},
    {"name": "Codex-6", "color": "#DB2777", "tint": "#F9A8D4"},
    {"name": "Codex-7", "color": "#4F46E5", "tint": "#A5B4FC"},
    {"name": "Codex-8", "color": "#D97706", "tint": "#FCD34D"},
]

VERB_BY_TYPE = {
    "add_shape": "Frame",
    "add_image": "Place",
    "add_line": "Route",
    "freeform": "Trace",
    "add_textbox": "Type",
    "micro_burst": "Sweep",
    "duplicate_shape": "Copy",
    "group_shapes": "Group",
}


def _lighten_hex(hex_color, factor=0.55):
    h = str(hex_color).strip().lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"#{r:02x}{g:02x}{b:02x}"


def _safe_hex(hex_color, fallback="161616"):
    if not hex_color:
        return "#" + fallback
    h = str(hex_color).strip().lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    if len(h) != 6:
        return "#" + fallback
    hl = h.lower()
    if hl == "010203":
        return "#020304"
    if hl == "fdfdfe":
        return "#fcfcfd"
    return "#" + h


def blend(a, b, t):
    return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)


def ease_out_cubic(t):
    u = 1.0 - max(0.0, min(1.0, t))
    return 1.0 - u * u * u


def smooth(t):
    """C2-continuous quintic minimum-jerk curve for silky human-like cursor motion."""
    t = max(0.0, min(1.0, t))
    return t * t * t * (10 + t * (-15 + 6 * t))


def clamp(v, lo, hi):
    return max(lo, min(hi, v))


def compute_tight_overlay_rect(window_rect, slide_rect, record_scope="window"):
    """Keep WS_EX_LAYERED Tk overlay tightly cropped to slide + small cursor/HUD margin (~45% less DWM cost!)."""
    if record_scope == "slide":
        return slide_rect
    wx, wy, ww, wh = window_rect
    vx, vy, vw, vh = slide_rect
    pad_x = 28
    pad_top = 26
    pad_bot = 58
    x0 = max(wx, vx - pad_x)
    y0 = max(wy, vy - pad_top)
    x1 = min(wx + ww, vx + vw + pad_x)
    y1 = min(wy + wh, vy + vh + pad_bot)
    return (int(x0), int(y0), max(32, int(x1 - x0)), max(32, int(y1 - y0)))


def _style_tuple_for_run(spec, defaults, sample_ch):
    fs_pt = float(spec.get("font_size", defaults.get("font_size", 14.0)))
    sup = bool(spec.get("superscript", False))
    sub = bool(spec.get("subscript", False))
    eff_fs = fs_pt * (0.68 if (sup or sub) else 1.0)
    bold = bool(spec.get("bold", defaults.get("bold", False)))
    italic = bool(spec.get("italic", defaults.get("italic", False)))
    color = _safe_hex(spec.get("font_color") or spec.get("color") or defaults.get("font_color") or defaults.get("color"), "161616")
    if has_cjk(sample_ch):
        fam = spec.get("font_name_east_asian") or defaults.get("font_name_east_asian") or "Microsoft YaHei"
    else:
        fam = spec.get("font_name") or defaults.get("font_name") or "Arial"
    y_shift = -0.28 * fs_pt if sup else (0.20 * fs_pt if sub else 0.0)
    return (fam, round(eff_fs, 2), bold, italic, color, round(y_shift, 2))


def _expand_styled_chars(action):
    runs = action.get("runs")
    out = []
    if runs:
        for rspec in runs:
            rtext = str(rspec.get("text", "")).replace("\r\n", "\n").replace("\r", "\n")
            for ch in rtext:
                out.append((ch, _style_tuple_for_run(rspec, action, ch)))
    else:
        raw = str(action.get("text", "")).replace("\r\n", "\n").replace("\r", "\n")
        for ch in raw:
            out.append((ch, _style_tuple_for_run(action, action, ch)))
    return out


def compute_rich_text_layout(action, rect, scale_x=1.0, scale_y=1.0, measure_fn=None):
    """Pre-compute pixel-accurate wrapped lines, run segments, and per-char caret coordinates."""
    x0, y0, x1, y1 = rect
    ml = float(action.get("margin_left", 0.0))
    mr = float(action.get("margin_right", 0.0))
    mt = float(action.get("margin_top", 0.0))
    mb = float(action.get("margin_bottom", 0.0))
    ix0, iy0 = x0 + ml, y0 + mt
    ix1, iy1 = max(ix0 + 4.0, x1 - mr), max(iy0 + 4.0, y1 - mb)
    iw, ih = ix1 - ix0, iy1 - iy0

    base_fs = float(action.get("font_size", 14.0))
    line_sp = float(action.get("line_spacing", 1.0))
    align = str(action.get("alignment", "left")).lower()
    valign = str(action.get("vertical_alignment", "middle")).lower()
    wrap = bool(action.get("word_wrap", True))

    chars = _expand_styled_chars(action)

    def char_width_pt(ch, st):
        if ch == "\n":
            return 0.0
        fam, eff_fs, bold, italic, _, _ = st
        if measure_fn is not None and scale_x > 1e-6:
            return measure_fn(ch, fam, eff_fs, bold, italic, scale_y) / scale_x
        return char_em_width(ch, bold) * eff_fs

    lines = [[]]
    cur_w = 0.0
    for idx, (ch, st) in enumerate(chars):
        if ch == "\n":
            lines.append([])
            cur_w = 0.0
            continue
        cw = char_width_pt(ch, st)
        if wrap and lines[-1] and cur_w + cw > iw * 1.015:
            lines.append([(idx, ch, st, cw)])
            cur_w = cw
        else:
            lines[-1].append((idx, ch, st, cw))
            cur_w += cw

    line_heights = []
    line_widths = []
    for ln in lines:
        mx_fs = max((st[1] for _, _, st, _ in ln), default=base_fs)
        line_heights.append(max(base_fs, mx_fs) * 1.24 * line_sp)
        line_widths.append(sum(cw for _, _, _, cw in ln))

    total_h = sum(line_heights) if line_heights else base_fs * 1.24 * line_sp
    if valign == "top":
        top_y = iy0
    elif valign == "bottom":
        top_y = iy1 - total_h
    else:
        top_y = iy0 + (ih - total_h) * 0.5

    positioned_lines = []
    char_carets = []
    cur_y = top_y

    first_w = line_widths[0] if line_widths else 0.0
    first_h = line_heights[0] if line_heights else base_fs * 1.24
    if align in ("center", "centre"):
        first_lx = ix0 + (iw - min(iw, first_w)) * 0.5
    elif align == "right":
        first_lx = ix1 - min(iw, first_w)
    else:
        first_lx = ix0
    char_carets.append((clamp(first_lx, ix0, ix1), clamp(cur_y + first_h * 0.5, iy0, iy1), first_h))

    caret_by_idx = {}
    for ln, lw, lh in zip(lines, line_widths, line_heights):
        if align in ("center", "centre"):
            lx = ix0 + (iw - min(iw, lw)) * 0.5
        elif align == "right":
            lx = ix1 - min(iw, lw)
        else:
            lx = ix0
        cy = cur_y + lh * 0.5
        x_cursor = lx
        placed = []
        for idx, ch, st, cw in ln:
            placed.append((idx, x_cursor, cy, ch, st))
            x_cursor += cw
            caret_by_idx[idx] = (clamp(x_cursor, ix0, ix1), clamp(cy, iy0, iy1), lh)
        positioned_lines.append((lx, cy, lh, placed))
        cur_y += lh

    last_caret = char_carets[0]
    for idx in range(len(chars)):
        if idx in caret_by_idx:
            last_caret = caret_by_idx[idx]
        char_carets.append(last_caret)

    return {
        "rect": rect,
        "inner_rect": (ix0, iy0, ix1, iy1),
        "char_count": len(chars),
        "lines": positioned_lines,
        "carets": char_carets,
    }


def extract_visible_segments(layout, shown_count):
    segments = []
    for _, _, _, placed in layout["lines"]:
        cur_x = None
        cur_y = None
        cur_st = None
        buf = []
        for idx, cx, cy, ch, st in placed:
            if idx >= shown_count:
                break
            if cur_st == st and cur_y == cy:
                buf.append(ch)
            else:
                if buf:
                    segments.append((cur_x, cur_y, "".join(buf), cur_st))
                cur_x, cur_y, cur_st = cx, cy, st
                buf = [ch]
        if buf:
            segments.append((cur_x, cur_y, "".join(buf), cur_st))
    return segments


def find_smart_guides(rect, committed_boxes, tol=1.6):
    """Subtle structural alignment smart guides (used ONLY during Wave 0 skeleton layout!)."""
    x0, y0, x1, y1 = rect
    if max(x1 - x0, y1 - y0) < 40.0:
        return []
    cx, cy = (x0 + x1) * 0.5, (y0 + y1) * 0.5
    v_guides = []
    h_guides = []

    for ob in committed_boxes:
        ox0, oy0, ox1, oy1 = ob
        if max(ox1 - ox0, oy1 - oy0) < 40.0:
            continue
        ocx, ocy = (ox0 + ox1) * 0.5, (oy0 + oy1) * 0.5
        for tx in (x0, cx, x1):
            for ox in (ox0, ocx, ox1):
                if abs(tx - ox) <= tol:
                    v_guides.append((ox, min(y0, oy0), max(y1, oy1), abs(tx - ox)))
                    break
        for ty in (y0, cy, y1):
            for oy in (oy0, ocy, oy1):
                if abs(ty - oy) <= tol:
                    h_guides.append((oy, min(x0, ox0), max(x1, ox1), abs(ty - oy)))
                    break

    guides = []
    seen_x = []
    for gx, gy0, gy1, _ in sorted(v_guides, key=lambda g: g[3]):
        if all(abs(gx - px) > 3.0 for px in seen_x):
            seen_x.append(gx)
            guides.append(("v", gx, gy0, gy1))
        if len(seen_x) >= 2:
            break
    seen_y = []
    for gy, gx0, gx1, _ in sorted(h_guides, key=lambda g: g[3]):
        if all(abs(gy - py) > 3.0 for py in seen_y):
            seen_y.append(gy)
            guides.append(("h", gy, gx0, gx1))
        if len(seen_y) >= 1:
            break
    return guides


def estimate_plan_duration(units, lanes, sw):
    """Estimate realistic wall-clock duration at pace=1.0 (accounting for zero-freeze pinned handoffs)."""
    by_wave = {}
    for u in units:
        by_wave.setdefault(int(u.get("wave", 2)), []).append(u)

    total = 0.7
    for _, items in sorted(by_wave.items()):
        lane_totals = {}
        for u in items:
            kind = u["type"]
            if kind == "micro_burst":
                n = len(u["members"])
                cost = 0.16 + clamp(0.38 + n * 0.028, 0.38, 1.25)
            elif kind == "add_textbox":
                n_chars = len(extract_plain_text(u))
                if n_chars <= 8:
                    td = clamp(0.16 + n_chars * 0.025, 0.18, 0.34)
                elif n_chars <= 24:
                    td = clamp(0.24 + n_chars * 0.016, 0.26, 0.54)
                else:
                    td = clamp(0.36 + n_chars * 0.009, 0.42, 0.78)
                cost = 0.15 + td
            elif kind == "add_image":
                cost = 0.16 + 0.36
            elif kind == "add_shape":
                has_txt = bool(extract_plain_text(u).strip())
                cost = 0.15 + (0.06 if is_micro_action(u) else 0.24) + (0.28 if has_txt else 0.0)
            elif kind in ("add_line", "freeform"):
                cost = 0.14 + (0.05 if is_micro_action(u) else 0.20)
            else:
                cost = 0.18
            ln = int(u.get("lane", 1))
            lane_totals[ln] = lane_totals.get(ln, 0.0) + cost
        if lane_totals:
            total += max(lane_totals.values())
    return round(total, 2)


@contextmanager
def quiet_new_shapes(shapes):
    """Suppress ugly PowerPoint default blue fills/shadows during shape creation."""
    original = shapes.AddShape(1, 0, 0, 1, 1)
    original_line = quiet = quiet_line = None
    try:
        original.Visible = 0
        original_line = shapes.AddLine(0, 0, 0.5, 0)
        original_line.Visible = 0
        quiet = shapes.AddShape(1, 0, 0, 1, 1)
        quiet.Fill.Visible = quiet.Line.Visible = quiet.Shadow.Visible = 0
        quiet.SetShapesDefaultProperties()
        quiet.Delete()
        quiet = None
        quiet_line = shapes.AddLine(0, 0, 0.5, 0)
        quiet_line.Line.Visible = quiet_line.Shadow.Visible = 0
        quiet_line.SetShapesDefaultProperties()
        quiet_line.Delete()
        quiet_line = None
        yield
    finally:
        if quiet_line is not None:
            quiet_line.Delete()
        if quiet is not None:
            quiet.Delete()
        original.Visible = -1
        original.SetShapesDefaultProperties()
        original.Delete()
        if original_line is not None:
            original_line.Visible = -1
            original_line.SetShapesDefaultProperties()
            original_line.Delete()


def prepare_plan(plan, base, existing=()):
    raw_list = plan.get("actions") or plan.get("operations") or plan.get("elements") or []
    if not raw_list:
        raise ValueError("Swarm plan contains no actions/operations.")
    lanes = int(plan.get("cursors", 6))
    pace = float(plan.get("pace", 1.0))
    if not 1 <= lanes <= 24 or not 0.25 <= pace <= 5.0:
        raise ValueError("Use 1–24 cursors and pace between 0.25 and 5.0.")

    actions = [normalize_action(raw, i) for i, raw in enumerate(raw_list, 1)]
    for a in actions:
        if a["type"] == "add_image":
            img_p = (base / a["image_path"]).resolve(strict=True)
            if not img_p.is_file():
                raise FileNotFoundError(img_p)
            a["image_path"] = img_p

    canvas = plan.get("canvas") or [960, 540]
    sw, sh = float(canvas[0]), float(canvas[1])
    resolve_connector_endpoints(actions)
    auto_assign_waves_and_lanes(actions, lanes, sw, sh)
    actions.sort(key=lambda a: (int(a.get("wave", 2)), float(a.get("z_index", 0))))
    deps, boxes = validate(actions, set(existing), lanes)
    use_bursts = bool(plan.get("micro_bursts", True))
    units = build_micro_bursts(actions, deps, boxes, sw, sh) if use_bursts else list(actions)
    return actions, deps, boxes, units


def _find_ppt_hwnd(win, deck_path=None):
    caption = win.Caption.casefold()
    stem = deck_path.stem.casefold() if deck_path else ""
    name = deck_path.name.casefold() if deck_path else ""
    exact = []
    all_ppt = []

    def find(hwnd, _):
        if gui.GetClassName(hwnd) != "PPTFrameClass" or not gui.IsWindowVisible(hwnd):
            return
        all_ppt.append(hwnd)
        title = gui.GetWindowText(hwnd).casefold()
        if (
            title == caption
            or title.startswith(caption + " - ")
            or (stem and (title.startswith(stem + " - ") or title == stem))
            or (name and (title.startswith(name + " - ") or title == name))
        ):
            exact.append(hwnd)

    gui.EnumWindows(find, None)
    if exact:
        return exact[0]
    if len(all_ppt) == 1:
        return all_ppt[0]
    fg = gui.GetForegroundWindow()
    if fg and gui.GetClassName(fg) == "PPTFrameClass":
        return fg
    return all_ppt[0] if all_ppt else None


def _get_visible_window_rect(hwnd):
    """Get exact visible Windows window rect without Win10/11 invisible drop-shadow padding."""
    if not hwnd:
        return (0, 0, 0, 0)
    try:
        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long),
            ]

        rect = RECT()
        DWMWA_EXTENDED_FRAME_BOUNDS = 9
        hr = ctypes.windll.dwmapi.DwmGetWindowAttribute(
            ctypes.c_void_p(hwnd),
            ctypes.c_uint(DWMWA_EXTENDED_FRAME_BOUNDS),
            ctypes.byref(rect),
            ctypes.sizeof(rect),
        )
        if hr == 0 and rect.right > rect.left and rect.bottom > rect.top:
            return (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)
    except Exception:
        pass
    l, t, r, b = gui.GetWindowRect(hwnd)
    return (l, t, max(0, r - l), max(0, b - t))


def _bring_to_foreground(hwnd):
    if not hwnd or gui.GetForegroundWindow() == hwnd:
        return
    cur = win32api.GetCurrentThreadId()
    fg = gui.GetForegroundWindow()
    other = win32process.GetWindowThreadProcessId(fg)[0] if fg else cur
    user32 = ctypes.windll.user32
    attached = other != cur and user32.AttachThreadInput(cur, other, True)
    try:
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
    finally:
        if attached:
            user32.AttachThreadInput(cur, other, False)


def _prepare_native_action(shapes, a, skip_text=False, reveal=False):
    """Build one native PowerPoint object on the STA COM thread.

    When reveal=False (default for lockstep hidden pre-build), creates and styles the object
    while `Visible = 0` so PowerPoint does not show it early; Priority-0 `reveal` later flips
    `Visible = -1` in ~2-5ms the exact frame the cursor completes its stroke!
    """
    kind = a["type"]
    if kind == "add_shape":
        x, y, r, b = bounds(a)
        w, h = max(0.5, r - x), max(0.5, b - y)
        preset = a.get("shape_type_id") or PRESETS[a["shape"]]
        s = shapes.AddShape(preset, x, y, w, h)
        s.Visible = 0
        s.LockAspectRatio = 0
        if a.get("fill_color") not in (None, "", "none", "transparent") or a.get("gradient") or a.get("gradient_stops"):
            fill(s, a)
        if a.get("line_width", 0) and float(a.get("line_width", 0)) > 0:
            stroke(s, a)
        if a.get("shadow"):
            apply_shadow(s, a)
        if a.get("adjustments") or a.get("radius") is not None:
            set_adjustments(s, a, w, h)
        if a.get("rotation"):
            s.Rotation = float(a["rotation"])
        if not skip_text and (a.get("text") or a.get("runs")):
            write_text_fast(s, a)
        s.Name = a["name"]
        if reveal:
            s.Visible = -1
        return s

    if kind == "add_image":
        x, y, r, b = bounds(a)
        w, h = max(0.5, r - x), max(0.5, b - y)
        if reveal:
            s = shapes.AddPicture(str(a["image_path"]), 0, -1, x, y, w, h)
        else:
            # Create off-slide and immediately hide so zero 1-frame premature flash ever appears!
            s = shapes.AddPicture(str(a["image_path"]), 0, -1, -2400.0, -2400.0, w, h)
            s.Visible = 0
        s.LockAspectRatio = 0
        s.Left, s.Top, s.Width, s.Height = x, y, w, h
        s.LockAspectRatio = -int(a.get("lock_aspect_ratio", True))
        s.AlternativeText = a.get("alt_text") or a.get("visual_role") or a["name"]
        for field in ("raster_reason", "visual_role", "alpha_method"):
            if field in a:
                try:
                    s.Tags.Add(field, str(a[field]))
                except Exception:
                    pass
        s.Name = a["name"]
        return s

    if kind in ("add_line", "freeform"):
        segments = list(path_segments(path_commands(a)))
        s = draw_path(shapes, segments)
        s.Visible = 0
        if a.get("fill_color") not in (None, "", "none", "transparent") or a.get("gradient") or a.get("gradient_stops"):
            fill(s, a)
        if a.get("line_width", 0) and float(a.get("line_width", 0)) > 0:
            stroke(s, a)
        if a.get("shadow"):
            apply_shadow(s, a)
        s.Name = a["name"]
        if reveal:
            s.Visible = -1
        return s

    if kind == "add_textbox":
        x, y, r, b = bounds(a)
        w, h = max(2.0, r - x), max(2.0, b - y)
        s = shapes.AddTextbox(1, x, y, w, h)
        s.Visible = 0
        s.Fill.Visible = 0
        s.Line.Visible = 0
        s.Shadow.Visible = 0
        write_text_fast(s, a)
        if a.get("rotation"):
            s.Rotation = float(a["rotation"])
        s.Name = a["name"]
        if reveal:
            s.Visible = -1
        return s

    if kind == "duplicate_shape":
        x, y, r, b = bounds(a)
        w, h = max(0.5, r - x), max(0.5, b - y)
        src = shapes.Item(a["source"])
        s = src.Duplicate().Item(1)
        s.Left, s.Top, s.Width, s.Height = x, y, w, h
        s.Name = a["name"]
        return s

    if kind == "group_shapes":
        s = shapes.Range(tuple(a["members"])).Group()
        s.Name = a["name"]
        return s

    raise ValueError(f"Unsupported action type: {kind}")


def _commit_native_action(shapes, a, skip_text=False):
    """Backward-compatible immediate build + reveal helper."""
    s = _prepare_native_action(shapes, a, skip_text=skip_text, reveal=True)
    return s.Name


class ComWorker(threading.Thread):
    """Dedicated PowerPoint STA COM thread with Priority-0 Reveal & Hidden Pre-Build queues."""

    def __init__(self, deck_path, slide_index, background, expected_canvas):
        super().__init__(daemon=True, name="ppt-com-worker")
        self.deck_path = deck_path
        self.slide_index = slide_index
        self.background = background
        self.expected_canvas = expected_canvas
        self._cond = threading.Condition()
        self._hi_queue = deque()
        self._lo_queue = deque()
        self.done_queue = queue.Queue()
        self.ready_event = threading.Event()
        self.finished_event = threading.Event()
        self.init_info = None
        self.finish_info = None
        self.error = None

    def submit_hi(self, job):
        with self._cond:
            self._hi_queue.append(job)
            self._cond.notify()

    def submit_lo(self, job):
        with self._cond:
            self._lo_queue.append(job)
            self._cond.notify()

    def _pop_job(self, paint_breather=0.0):
        with self._cond:
            if paint_breather > 0.0 and not self._hi_queue and self._lo_queue:
                # Brief 8ms paint breather after a visible reveal so PowerPoint's UI thread
                # flushes WM_PAINT immediately before the next background pre-build starts.
                # Wakes up instantaneously if another Priority-0 reveal arrives!
                self._cond.wait(timeout=paint_breather)
            while not self._hi_queue and not self._lo_queue:
                self._cond.wait()
            if self._hi_queue:
                return self._hi_queue.popleft()
            return self._lo_queue.popleft()

    def run(self):
        pythoncom.CoInitialize()
        try:
            app = GetActiveObject("PowerPoint.Application")
            deck = app.ActivePresentation
            if Path(deck.FullName).resolve() != self.deck_path:
                found = None
                for i in range(1, app.Presentations.Count + 1):
                    pres = app.Presentations.Item(i)
                    if Path(pres.FullName).resolve() == self.deck_path:
                        found = pres
                        break
                if found is None:
                    raise RuntimeError(f"Open and activate the working presentation first: {self.deck_path}")
                deck = found
                deck.Windows.Item(1).Activate()

            win = app.ActiveWindow
            win.ViewType = 9  # ppViewNormal
            try:
                if float(win.SplitVertical) < 96.0:
                    win.SplitVertical = 100
            except Exception:
                pass
            slide = deck.Slides.Item(self.slide_index)
            win.View.GotoSlide(slide.SlideIndex)
            if self.background:
                apply_background(slide, self.background)
            win.View.ZoomToFit = True
            win.View.Zoom = win.View.Zoom
            win.Activate()

            ppt_hwnd = _find_ppt_hwnd(win, self.deck_path)
            _bring_to_foreground(ppt_hwnd)
            time.sleep(0.24)
            sw = float(deck.PageSetup.SlideWidth)
            sh = float(deck.PageSetup.SlideHeight)
            if self.expected_canvas and max(abs(a - b) for a, b in zip(self.expected_canvas, (sw, sh))) > 0.5:
                raise ValueError(f"Slide canvas ({sw}x{sh}) does not match plan canvas {self.expected_canvas}.")

            vx = win.PointsToScreenPixelsX(0)
            vy = win.PointsToScreenPixelsY(0)
            vw = round((win.PointsToScreenPixelsX(round(sw)) - vx) * sw / round(sw))
            vh = round((win.PointsToScreenPixelsY(round(sh)) - vy) * sh / round(sh))
            if min(vw, vh) < 40:
                raise RuntimeError("Target PowerPoint slide viewport is too small or minimized.")

            win_rect = _get_visible_window_rect(ppt_hwnd) if ppt_hwnd else (vx, vy, vw, vh)
            shapes = slide.Shapes
            existing = {shapes.Item(i).Name for i in range(1, shapes.Count + 1)}

            self.init_info = {
                "sw": sw,
                "sh": sh,
                "ppt_hwnd": ppt_hwnd,
                "slide_rect": (vx, vy, vw, vh),
                "window_rect": win_rect,
                "slide_offset": (vx - win_rect[0], vy - win_rect[1], vw, vh),
                "existing": existing,
            }
            self.ready_event.set()

            prepared_shapes = {}
            next_breather = 0.0
            with quiet_new_shapes(shapes):
                while True:
                    job = self._pop_job(paint_breather=next_breather)
                    next_breather = 0.0
                    kind = job[0]
                    if kind == "abort":
                        return
                    if kind == "prepare":
                        action, token = job[1], job[2]
                        s_obj = _prepare_native_action(shapes, action, skip_text=False, reveal=False)
                        prepared_shapes[action["name"]] = s_obj
                        self.done_queue.put(("prepared", [action["name"]], token))
                    elif kind == "prepare_shape_empty":
                        action, token = job[1], job[2]
                        s_obj = _prepare_native_action(shapes, action, skip_text=True, reveal=False)
                        prepared_shapes[action["name"]] = s_obj
                        self.done_queue.put(("prepared", [action["name"]], token))
                    elif kind == "prepare_batch":
                        batch_actions, token = job[1], job[2]
                        names = []
                        for act in batch_actions:
                            s_obj = _prepare_native_action(shapes, act, skip_text=False, reveal=False)
                            prepared_shapes[act["name"]] = s_obj
                            names.append(act["name"])
                        self.done_queue.put(("prepared", names, token))
                    elif kind == "reveal":
                        name, token = job[1], job[2]
                        s_obj = prepared_shapes.pop(name, None) or shapes.Item(name)
                        s_obj.Visible = -1
                        self.done_queue.put(("ok", [name], token))
                        next_breather = 0.008
                    elif kind == "reveal_empty":
                        name, token = job[1], job[2]
                        s_obj = prepared_shapes.get(name) or shapes.Item(name)
                        s_obj.Visible = -1
                        self.done_queue.put(("sub", [name], token))
                        next_breather = 0.008
                    elif kind == "reveal_batch":
                        names, token = job[1], job[2]
                        for nm in names:
                            s_obj = prepared_shapes.pop(nm, None) or shapes.Item(nm)
                            s_obj.Visible = -1
                        self.done_queue.put(("ok", list(names), token))
                        next_breather = 0.008
                    elif kind == "set_text":
                        action, token = job[1], job[2]
                        s_obj = prepared_shapes.pop(action["name"], None) or shapes.Item(action["name"])
                        write_text_fast(s_obj, action)
                        self.done_queue.put(("ok", [action["name"]], token))
                        next_breather = 0.008
                    elif kind in ("commit", "commit_sync"):
                        action, token = job[1], job[2]
                        committed_name = _commit_native_action(shapes, action, skip_text=False)
                        self.done_queue.put(("ok", [committed_name], token))
                        next_breather = 0.008
                    elif kind == "commit_batch":
                        batch_actions, token = job[1], job[2]
                        names = [_commit_native_action(shapes, act, skip_text=False) for act in batch_actions]
                        self.done_queue.put(("ok", names, token))
                        next_breather = 0.008
                    elif kind == "finish":
                        _, out_path, export_preview, z_ordered, boxes, all_actions, auto_fix = job
                        # Defensive guarantee: never save any pre-built shape hidden
                        for leftover in prepared_shapes.values():
                            try:
                                leftover.Visible = -1
                            except Exception:
                                pass
                        prepared_shapes.clear()
                        apply_z_order(shapes, z_ordered, boxes=boxes)
                        audit = audit_slide_shapes(
                            shapes,
                            all_actions,
                            sw,
                            sh,
                            auto_fix_text_overflow=auto_fix,
                        )
                        break
            # Persist after the default-style helper shapes are removed.
            if out_path == self.deck_path:
                deck.Save()
            else:
                out_path.parent.mkdir(parents=True, exist_ok=True)
                deck.SaveCopyAs(str(out_path), 24)
            exported = None
            if export_preview:
                export_preview.parent.mkdir(parents=True, exist_ok=True)
                slide.Export(str(export_preview), "PNG", int(round(sw * 2)), int(round(sh * 2)))
                exported = str(export_preview)
            self.finish_info = {
                "output": str(out_path),
                "preview_png": exported,
                "audit": audit,
            }
            self.finished_event.set()
        except Exception:
            self.error = traceback.format_exc()
            self.done_queue.put(("error", self.error, None))
            self.ready_event.set()
            self.finished_event.set()
        finally:
            pythoncom.CoUninitialize()


class AgentCursor:
    def __init__(self, show, index, count, profile=None):
        self.show = show
        self.index = index
        base_p = DEFAULT_PROFILES[index % len(DEFAULT_PROFILES)]
        prof = profile or {}
        self.name = prof.get("name") or (base_p["name"] if index < len(DEFAULT_PROFILES) else f"Codex-{index + 1}")
        self.color = prof.get("color") or base_p["color"]
        self.tint = prof.get("tint") or _lighten_hex(self.color, 0.58)
        # Start calmly along the left/top margin aligned with each agent's vertical band (never bottom edge!)
        band_y = show.sh * (0.18 + 0.66 * (index + 0.5) / max(1, count))
        self.home = (show.sw * 0.12, clamp(band_y, show.sh * 0.12, show.sh * 0.86))
        self.pos = self.home
        self.trail = deque(maxlen=6)
        self.tool = "pointer"
        self.status_verb = ""
        self.generator = None
        self.action = None
        self.target_center = self.home
        self.overlay = None
        self.rng = random.Random(index + 17)
        self.used = False
        self.active = False
        self.samples = 0
        self.cue_started = 0.0
        self.wall_start = 0.0
        self.idle_anchor = self.pos

    def move_to(self, pos, record_trail=True):
        self.pos = pos
        if record_trail:
            if not self.trail or math.dist(self.trail[-1], pos) >= 1.0:
                self.trail.append(pos)

    def animate(self, duration, update, easing=smooth):
        yield from self.animate_sync(duration, update, easing=easing, ready_token=None)

    def animate_sync(self, duration, update, easing=smooth, ready_token=None, hold_frac=0.88, max_wait_wall=25.0):
        """Animate smoothly at 60fps; if a hidden pre-build is still finishing near the end,
        gently ease through the final 12% of the stroke instead of freezing dead or leaving early!"""
        start = self.show.clock
        dur = max(0.016, duration)
        deadline = time.perf_counter() + max_wait_wall
        while True:
            raw_t = (self.show.clock - start) / dur
            if ready_token is not None and ready_token not in self.show.prepared_tokens:
                if raw_t >= hold_frac:
                    if time.perf_counter() > deadline:
                        raise TimeoutError(f"PowerPoint COM worker timed out preparing token {ready_token}.")
                    over = raw_t - hold_frac
                    t = hold_frac + (0.985 - hold_frac) * (1.0 - math.exp(-over * 2.2))
                    update(easing(t))
                    self.samples += 1
                    yield
                    continue
            t = min(1.0, raw_t)
            update(easing(t))
            self.samples += 1
            if t >= 1.0:
                break
            yield

    def wait_for_token_sync(self, token, max_wait_wall=25.0):
        """Ultra-short lockstep wait for Priority-0 reveal/commit (~2-8ms, 0-1 frames)."""
        deadline = time.perf_counter() + max_wait_wall
        while token not in self.show.completed_tokens:
            if time.perf_counter() > deadline:
                raise TimeoutError(f"PowerPoint COM worker timed out waiting for commit token {token}.")
            yield

    def travel(self, target, fast=False):
        start = self.pos
        dist = math.dist(start, target)
        if dist < 1.2:
            self.move_to(target, record_trail=False)
            return
        self.tool = "pointer"
        self.trail.clear()
        use_trail = dist > self.show.sw * 0.085 and not fast
        if use_trail:
            self.trail.append(start)
        bend = min(dist * 0.055, self.show.sw * 0.018) * self.rng.choice((-1.0, 1.0)) * self.rng.uniform(0.3, 0.85)
        perp = (-(target[1] - start[1]) / dist, (target[0] - start[0]) / dist)
        if fast:
            dur = clamp(0.05 + 0.032 * math.log2(1.0 + dist / 18.0), 0.05, 0.13)
        else:
            dur = clamp(0.09 + 0.052 * math.log2(1.0 + dist / 18.0), 0.10, 0.26)

        def step(t):
            x, y = blend(start, target, t)
            h = math.sin(math.pi * t) * bend
            self.move_to((x + perp[0] * h, y + perp[1] * h), record_trail=use_trail)

        # Quintic minimum-jerk curve = silky continuous acceleration & deceleration!
        yield from self.animate(dur, step, smooth)
        self.move_to(target, record_trail=False)
        self.trail.clear()

    def _live_type_sequence(self, rect, a, mode="reveal"):
        """Silky rich-run live typing in strict lockstep: zero double-text ghosting and zero post-cursor lag."""
        prep_tok = self.show.enqueue_prepare("prepare", a) if mode == "reveal" else None
        layout = self.show.build_text_layout(a, rect)
        n_chars = layout["char_count"]
        start_cx, start_cy, _ = layout["carets"][0]
        yield from self.travel((start_cx, start_cy))
        self.tool = "ibeam"
        self.status_verb = "Type"
        self.trail.clear()

        show_frame = n_chars > 8 and (rect[3] - rect[1]) > 16.0

        if n_chars <= 6:
            dur = clamp(0.16 + n_chars * 0.024, 0.17, 0.32)
        elif n_chars <= 22:
            dur = clamp(0.24 + n_chars * 0.016, 0.26, 0.54)
        else:
            dur = clamp(0.36 + n_chars * 0.009, 0.42, 0.78)

        def type_step(t):
            shown = max(1, min(n_chars, int(round(n_chars * t)))) if n_chars > 0 else 0
            cx, cy, lh = layout["carets"][shown]
            self.move_to((cx, cy), record_trail=False)
            self.overlay = {
                "kind": "live_text",
                "rect": rect,
                "layout": layout,
                "shown": shown,
                "show_frame": show_frame,
                "caret": (cx, cy, lh),
                "blink": int(self.show.clock * 7.0) % 2 == 0,
            }

        yield from self.animate_sync(dur, type_step, lambda t: t, ready_token=prep_tok)
        self.show.typed += n_chars

        # Priority-0 reveal or set_text right at final caret, then IMMEDIATELY clear live_text (ZERO ghosting!)
        if mode == "reveal":
            rev_tok = self.show.enqueue_reveal("reveal", a["name"])
        else:
            rev_tok = self.show.enqueue_reveal("set_text", a)
        yield from self.wait_for_token_sync(rev_tok)
        self.overlay = None
        self.show.committed_boxes.append(rect)
        self.tool = "pointer"
        self.idle_anchor = self.pos

    def do_shape(self, a):
        x, y, r, b = bounds(a)
        w, h = r - x, b - y
        has_text = bool(extract_plain_text(a).strip())
        micro = is_micro_action(a)
        is_wave0 = int(a.get("wave", 2)) == 0

        # Pre-build hidden in background while cursor travels & drags!
        prep_tok = self.show.enqueue_prepare("prepare_shape_empty" if has_text else "prepare", a)
        yield from self.travel((x, y), fast=micro)
        self.tool = "cross"
        self.status_verb = "Frame"
        self.trail.clear()
        dur = 0.055 if micro else clamp(0.16 + math.dist((x, y), (r, b)) / (self.show.sw * 2.0), 0.16, 0.34)
        # Smart guides ONLY for Wave 0 structural panels (clean, non-distracting!)
        guides = self.show.compute_smart_guides((x, y, r, b)) if (is_wave0 and not micro) else []

        def drag(t):
            cx, cy = blend((x, y), (r, b), t)
            self.move_to((cx, cy), record_trail=False)
            self.overlay = {
                "kind": "cad_box",
                "rect": (x, y, cx, cy),
                "shape": a.get("shape", "rectangle"),
                "dim": f"{round(cx - x)}×{round(cy - y)}",
                "show_dim": (is_wave0 and not micro and w >= 48.0 and h >= 24.0),
                "guides": guides if t > 0.72 else [],
            }

        yield from self.animate_sync(dur, drag, smooth, ready_token=prep_tok)

        if has_text:
            # Beat 1: reveal empty badge instantly under cursor before typing text inside it
            rev1 = self.show.enqueue_reveal("reveal_empty", a["name"])
            yield from self.wait_for_token_sync(rev1)
            self.overlay = None
            self.show.committed_boxes.append((x, y, r, b))
            # Beat 2: smoothly type embedded text inside the shape in lockstep
            yield from self._live_type_sequence((x, y, r, b), a, mode="set_text")
        else:
            rev = self.show.enqueue_reveal("reveal", a["name"])
            yield from self.wait_for_token_sync(rev)
            self.overlay = None
            self.show.committed_boxes.append((x, y, r, b))
            if is_wave0 and not micro:
                self.show.selection_locks.append((self.show.clock, (x, y, r, b), self.color))
            self.tool = "pointer"
            self.idle_anchor = self.pos

    def do_image(self, a):
        x, y, r, b = bounds(a)
        w, h = r - x, b - y
        # Pre-load picture off-slide & hidden during travel + scan!
        prep_tok = self.show.enqueue_prepare("prepare", a)
        yield from self.travel((x, y))
        self.tool = "pen"
        self.status_verb = "Place"
        self.trail.clear()
        dur = clamp(0.28 + (w + h) / (self.show.sw * 2.6), 0.28, 0.44)

        def scan(t):
            scan_y = y + h * t
            wave_x = x + w * (0.20 + 0.60 * (0.5 + 0.5 * math.sin(t * math.pi * 1.5)))
            self.move_to((wave_x, scan_y), record_trail=False)
            self.overlay = {"kind": "asset_scan", "rect": (x, y, r, b), "progress": t, "guides": []}

        yield from self.animate_sync(dur, scan, smooth, ready_token=prep_tok)
        rev = self.show.enqueue_reveal("reveal", a["name"])
        yield from self.wait_for_token_sync(rev)
        self.overlay = None
        self.show.committed_boxes.append((x, y, r, b))
        self.show.selection_locks.append((self.show.clock, (x, y, r, b), self.color))
        self.tool = "pointer"
        self.idle_anchor = self.pos

    def do_path(self, a):
        segments = list(path_segments(path_commands(a)))
        micro = is_micro_action(a)
        prep_tok = self.show.enqueue_prepare("prepare", a)
        yield from self.travel(segments[0][1], fast=micro)
        self.tool = "pen"
        self.status_verb = "Route"
        self.trail.clear()
        stroke_col = a.get("line_color") if a.get("line_width", 0) else None
        ink_col = stroke_col or (a.get("gradient") or [a.get("fill_color", self.color)])[-1]
        ink_w = a.get("line_width") or 1.6

        all_pts = []
        for idx, seg in enumerate(segments):
            s_pts = curve_samples(seg)
            all_pts.extend(s_pts if idx == 0 else s_pts[1:])
        dists = [0.0]
        for p, q in zip(all_pts, all_pts[1:]):
            dists.append(dists[-1] + math.dist(p, q))
        total_len = max(0.1, dists[-1])
        dur = 0.05 if micro else clamp(0.14 + total_len / (self.show.sw * 1.75), 0.14, 0.34)

        def trace(t):
            target = total_len * t
            i = max(1, min(len(all_pts) - 1, bisect_left(dists, target)))
            f = (target - dists[i - 1]) / max(1e-8, dists[i] - dists[i - 1])
            tip = blend(all_pts[i - 1], all_pts[i], f)
            self.move_to(tip, record_trail=False)
            self.overlay = {
                "kind": "vector_guide",
                "points": all_pts[:i] + [tip],
                "color": ink_col,
                "width": ink_w,
                "dash": a.get("line_dash"),
            }

        yield from self.animate_sync(dur, trace, smooth, ready_token=prep_tok)
        rev = self.show.enqueue_reveal("reveal", a["name"])
        yield from self.wait_for_token_sync(rev)
        self.overlay = None
        self.show.committed_boxes.append(bounds(a))
        self.tool = "pointer"
        self.idle_anchor = self.pos

    def do_micro_burst(self, burst):
        """Continuous lockstep pen sweep: 1-chunk hidden lookahead + instant per-chunk reveal directly under pen tip!"""
        members = burst["members"]
        centers = []
        for m in members:
            bx0, by0, bx1, by1 = self.show.boxes[m["name"]]
            centers.append(((bx0 + bx1) * 0.5, (by0 + by1) * 0.5))

        n = len(members)
        if n <= 6:
            chunk_size = 1
        elif n <= 18:
            chunk_size = 2
        elif n <= 32:
            chunk_size = 3
        else:
            chunk_size = 4
        chunks = [members[i : i + chunk_size] for i in range(0, n, chunk_size)]
        total_dur = clamp(0.36 + n * 0.026, 0.36, 1.35)
        step_dur = total_dur / max(1, len(chunks))
        ub = burst["_burst_box"]

        # Pre-build chunk 0 hidden while traveling to the cluster!
        prep_toks = {0: self.show.enqueue_prepare("prepare_batch", chunks[0])}
        yield from self.travel(centers[0])
        self.tool = "pen"
        self.status_verb = "Sweep"
        self.trail.clear()

        for ci, chunk in enumerate(chunks):
            # 1-chunk lookahead keeps COM saturated without ever outrunning the pen!
            if ci + 1 < len(chunks):
                prep_toks[ci + 1] = self.show.enqueue_prepare("prepare_batch", chunks[ci + 1])

            end_idx = min(n - 1, (ci + 1) * chunk_size - 1)
            end_pt = centers[end_idx]
            start_pt = self.pos

            def sweep_step(t, sp=start_pt, ep=end_pt, prog=(ci + 1) / max(1, len(chunks))):
                pt = blend(sp, ep, t)
                self.move_to(pt, record_trail=False)
                self.overlay = {"kind": "burst_sweep", "rect": ub, "tip": pt, "progress": prog}

            yield from self.animate_sync(step_dur, sweep_step, lambda t: t, ready_token=prep_toks[ci])
            rev_tok = self.show.enqueue_reveal("reveal_batch", [m["name"] for m in chunk])
            yield from self.wait_for_token_sync(rev_tok)

        self.overlay = None
        self.show.committed_boxes.append(ub)
        self.tool = "pointer"
        self.idle_anchor = self.pos

    def do_textbox(self, a):
        rect = bounds(a)
        yield from self._live_type_sequence(rect, a, mode="reveal")

    def do_duplicate(self, a):
        x, y, r, b = bounds(a)
        w, h = r - x, b - y
        sb = self.show.boxes.get(a["source"], (x, y, r, b))
        sx, sy = sb[0], sb[1]
        yield from self.travel((sx + w * 0.5, sy + h * 0.5))
        self.tool = "cross"
        self.status_verb = "Copy"
        self.trail.clear()
        dur = clamp(0.16 + math.dist((sx, sy), (x, y)) / (self.show.sw * 1.7), 0.16, 0.34)

        def drag(t):
            cx, cy = blend((sx, sy), (x, y), t)
            self.move_to((cx + w * 0.5, cy + h * 0.5), record_trail=False)
            self.overlay = {"kind": "cad_box", "rect": (cx, cy, cx + w, cy + h), "shape": "rectangle", "dim": "", "show_dim": False, "guides": []}

        yield from self.animate(dur, drag, smooth)
        tok = self.show.enqueue_reveal("commit_sync", a)
        yield from self.wait_for_token_sync(tok)
        self.overlay = None
        self.show.committed_boxes.append((x, y, r, b))
        self.tool = "pointer"
        self.idle_anchor = self.pos

    def do_group(self, a):
        x, y, r, b = self.show.boxes[a["name"]]
        yield from self.travel(((x + r) * 0.5, (y + b) * 0.5))
        self.status_verb = "Group"
        tok = self.show.enqueue_reveal("commit_sync", a)
        yield from self.wait_for_token_sync(tok)
        self.overlay = None
        self.tool = "pointer"
        self.idle_anchor = self.pos

    def run(self, a):
        self.action = a
        self.used = True
        self.active = True
        self.status_verb = VERB_BY_TYPE.get(a["type"], "")
        bx0, by0, bx1, by1 = a["_burst_box"] if a["type"] == "micro_burst" else self.show.boxes[a["name"]]
        self.target_center = ((bx0 + bx1) * 0.5, (by0 + by1) * 0.5)
        seed = int.from_bytes(hashlib.sha256(a["name"].encode()).digest()[:8], "big")
        self.rng.seed(seed)
        self.samples = 0
        self.cue_started = self.show.clock
        self.wall_start = time.perf_counter() - self.show.first
        kind = a["type"]
        if kind == "add_shape":
            yield from self.do_shape(a)
        elif kind == "add_image":
            yield from self.do_image(a)
        elif kind in ("add_line", "freeform"):
            yield from self.do_path(a)
        elif kind == "add_textbox":
            yield from self.do_textbox(a)
        elif kind == "micro_burst":
            yield from self.do_micro_burst(a)
        elif kind == "duplicate_shape":
            yield from self.do_duplicate(a)
        elif kind == "group_shapes":
            yield from self.do_group(a)
        self.active = False
        self.status_verb = ""


class SwarmShow:
    def __init__(self, plan_path, timeout, export_preview=None):
        if not HAS_WIN32:
            raise RuntimeError("Live PowerPoint multi-cursor drawing requires Windows with pywin32 installed.")
        self.plan_path = plan_path
        self.plan = json.loads(plan_path.read_text(encoding="utf-8-sig"))

        def resolve(p):
            return (plan_path.parent / p).resolve()

        self.path = resolve(self.plan["deck"])
        self.output = resolve(self.plan["output"]) if self.plan.get("output") else self.path
        self.export_preview = Path(export_preview).resolve() if export_preview else (
            resolve(self.plan["export_preview"]) if self.plan.get("export_preview") else None
        )
        self.count = int(self.plan.get("cursors", 6))
        self.agent_labels = bool(self.plan.get("agent_labels", True))
        self.show_trails = bool(self.plan.get("motion_trails", True))
        # Static non-flickering pills by default for calm executive look!
        self.show_status_verb = bool(self.plan.get("show_status_verb", False))
        self.auto_fix_text_overflow = bool(self.plan.get("auto_fix_text_overflow", True))
        self.record_scope = str(self.plan.get("record_scope", "window")).strip().lower()
        if self.record_scope not in ("window", "slide"):
            self.record_scope = "window"
        self.show_hud = bool(self.plan.get("hud", self.record_scope == "window"))

        self.worker = ComWorker(
            self.path,
            int(self.plan.get("slide", 1)),
            self.plan.get("background"),
            self.plan.get("canvas"),
        )
        self.worker.start()
        if not self.worker.ready_event.wait(timeout=24):
            raise TimeoutError("PowerPoint COM worker did not initialize within 24 seconds.")
        if self.worker.error:
            raise RuntimeError(self.worker.error)

        info = self.worker.init_info
        self.sw, self.sh = info["sw"], info["sh"]
        self.ppt_hwnd = info["ppt_hwnd"]
        self.slide_offset = info["slide_offset"]
        self.ppt_committed = set(info["existing"])
        self.in_flight_tokens = set()
        self.in_flight_preps = set()
        self.prepared_tokens = set()
        self.completed_tokens = set()
        self.committed_boxes = []
        self._next_token = 1
        self.chroma_key = _pick_chroma_key(self.plan.get("background"))

        self.actions, self.deps, self.boxes, self.units = prepare_plan(self.plan, plan_path.parent, self.ppt_committed)

        user_pace = float(self.plan.get("pace", 1.0))
        target_dur = self.plan.get("target_duration", 24.0)
        est_dur = estimate_plan_duration(self.units, self.count, self.sw)
        if target_dur is not None and float(target_dur) > 0:
            self.pace = clamp((est_dur / float(target_dur)) * user_pace, 0.85, 2.20)
        else:
            self.pace = user_pace
        self.estimated_raw_seconds = est_dur

        self.z_ordered_names = [
            a["name"]
            for _, a in sorted(enumerate(self.actions), key=lambda pair: (float(pair[1].get("z_index", pair[0])), pair[0]))
            if a["type"] != "group_shapes"
        ]
        # Only gate on primary Wave 0 container panels (cards, sidebar, banners) — NOT tiny decorative
        # arcs or individual vertical dividers — so ALL 6 cursors stay continuously active from t=0.0s!
        slide_area = max(1.0, self.sw * self.sh)
        self.wave0_names = {
            a["name"]
            for a in self.actions
            if int(a.get("wave", 2)) == 0
            and a["type"] == "add_shape"
            and (self.boxes[a["name"]][2] - self.boxes[a["name"]][0]) * (self.boxes[a["name"]][3] - self.boxes[a["name"]][1]) >= slide_area * 0.03
        }
        self.wave1_names = {a["name"] for a in self.actions if int(a.get("wave", 2)) == 1}
        self.pending = list(self.units)
        self.timeout = timeout
        self.clock = 0.0
        self.typed = self.frames = self.peak = 0
        self.committed_count = 0
        self.timeline = []
        self.selection_locks = []
        self.first = self.done_at = self.hidden_at = self.error = None
        self.last = time.perf_counter()
        self.next_frame_at = self.last
        self.next_probe = 0.0
        self.slide_rect = info["slide_rect"]
        self.window_rect = info["window_rect"]
        self.overlay_rect = None
        self.capture_started = None
        self.recording_intervals = []
        self.finish_sent = False
        self._font_cache = {}

        # Transparent click-through topmost Tk overlay with background-matched chroma key
        self.root = tk.Tk()
        self.root.protocol("WM_DELETE_WINDOW", self.root.quit)
        self.root.withdraw()
        self.root.overrideredirect(True)
        self.root.configure(bg=self.chroma_key)
        self.root.attributes("-transparentcolor", self.chroma_key)
        self.canvas = tk.Canvas(self.root, bg=self.chroma_key, bd=0, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.root.update_idletasks()
        self.hwnd = gui.GetParent(self.root.winfo_id()) or self.root.winfo_id()
        style = gui.GetWindowLong(self.hwnd, wc.GWL_EXSTYLE)
        gui.SetWindowLong(
            self.hwnd,
            wc.GWL_EXSTYLE,
            style | wc.WS_EX_LAYERED | wc.WS_EX_TRANSPARENT | wc.WS_EX_TOOLWINDOW | 0x08000000,
        )
        self.root.geometry("1x1+0+0")
        self.root.deiconify()
        self.root.update_idletasks()
        gui.ShowWindow(self.hwnd, wc.SW_HIDE)

        profiles = self.plan.get("cursor_profiles") or []
        self.agents = [
            AgentCursor(self, i, self.count, profiles[i] if i < len(profiles) else None)
            for i in range(self.count)
        ]
        self.recorder = None
        if self.plan.get("record"):
            rec_frame_size = (
                (self.window_rect[2], self.window_rect[3])
                if self.record_scope == "window"
                else (self.sw, self.sh)
            )
            self.recorder = Recorder(
                resolve(self.plan["record"]),
                plan_path.parent,
                rec_frame_size,
                overwrite=bool(self.plan.get("overwrite_record", True)),
                fps=int(self.plan.get("record_fps", 60)),
                prefer_ddagrab=bool(self.plan.get("prefer_ddagrab", True)),
            )

    def _measure_tk_char(self, ch, fam, eff_fs_pt, bold, italic, scale_y):
        fs_px = max(8, int(round(eff_fs_pt * scale_y)))
        weight = "bold" if bold else "normal"
        slant = "italic" if italic else "roman"
        key = (fam, fs_px, weight, slant)
        fobj = self._font_cache.get(key)
        if fobj is None and tkfont is not None:
            fobj = tkfont.Font(root=self.root, family=fam, size=-fs_px, weight=weight, slant=slant)
            self._font_cache[key] = fobj
        if fobj is not None:
            return float(fobj.measure(ch))
        return char_em_width(ch, bold) * eff_fs_pt * scale_y

    def _measure_tk_text(self, text, fam, fs_px, bold=True):
        weight = "bold" if bold else "normal"
        key = (fam, fs_px, weight, "roman")
        fobj = self._font_cache.get(key)
        if fobj is None and tkfont is not None:
            fobj = tkfont.Font(root=self.root, family=fam, size=-fs_px, weight=weight, slant="roman")
            self._font_cache[key] = fobj
        if fobj is not None:
            return float(fobj.measure(text))
        return len(text) * fs_px * 0.62

    def build_text_layout(self, action, rect):
        _, _, vw, vh = self.slide_rect
        sx = vw / max(1.0, self.sw)
        sy = vh / max(1.0, self.sh)
        return compute_rich_text_layout(action, rect, scale_x=sx, scale_y=sy, measure_fn=self._measure_tk_char)

    def compute_smart_guides(self, rect):
        return find_smart_guides(rect, self.committed_boxes)

    def enqueue_prepare(self, job_kind, payload):
        token = self._next_token
        self._next_token += 1
        self.in_flight_tokens.add(token)
        self.in_flight_preps.add(token)
        self.worker.submit_lo((job_kind, payload, token))
        return token

    def enqueue_reveal(self, job_kind, payload):
        token = self._next_token
        self._next_token += 1
        self.in_flight_tokens.add(token)
        self.worker.submit_hi((job_kind, payload, token))
        return token

    def drain_done_queue(self):
        while True:
            try:
                tag, payload, token = self.worker.done_queue.get_nowait()
            except queue.Empty:
                break
            if tag == "error":
                raise RuntimeError(payload)
            if token is not None:
                self.in_flight_tokens.discard(token)
                if tag == "prepared":
                    self.prepared_tokens.add(token)
                    self.in_flight_preps.discard(token)
                else:
                    self.completed_tokens.add(token)
            if tag == "ok":
                for nm in payload:
                    if nm not in self.ppt_committed:
                        self.ppt_committed.add(nm)
                        self.committed_count += 1

    def probe_win32(self):
        """Pure Win32 foreground & PowerPoint window/slide geometry check (<0.02ms, ZERO COM!)."""
        fg = gui.GetForegroundWindow()
        if not fg:
            return None
        if self.ppt_hwnd and fg != self.ppt_hwnd:
            if gui.GetClassName(fg) != "PPTFrameClass":
                return None
            self.ppt_hwnd = fg
        elif not self.ppt_hwnd:
            if gui.GetClassName(fg) != "PPTFrameClass":
                return None
            self.ppt_hwnd = fg
        if self.ppt_hwnd:
            wrect = _get_visible_window_rect(self.ppt_hwnd)
            dx, dy, vw, vh = self.slide_offset
            srect = (wrect[0] + dx, wrect[1] + dy, vw, vh)
            return (wrect, srect)
        return (self.worker.init_info["window_rect"], self.worker.init_info["slide_rect"])

    def _unit_deps(self, u):
        if u["type"] == "micro_burst":
            return u["_burst_deps"]
        return self.deps[u["name"]]

    def _unit_box(self, u):
        if u["type"] == "micro_burst":
            return u["_burst_box"]
        return self.boxes[u["name"]]

    def choose(self, agent):
        # Bound background pre-build queue cleanly to active cursor count (+2 for burst lookahead)
        # so Priority-0 reveals always stay instantaneous (<10ms) with zero global queue freezes!
        if len(getattr(self, "in_flight_preps", set())) >= self.count + 2:
            return None

        active_peers = [ag for ag in self.agents if ag is not agent and ag.generator is not None]
        busy_targets = [ag.target_center for ag in active_peers]
        # Elliptical clearance: horizontal 95pt, vertical 55pt (cursor pills are compact ~20pt tall,
        # so adjacent rows 90pt apart draw simultaneously in full 6-cursor harmony without blocking each other!)
        sep_x = self.sw * 0.095
        sep_y = self.sh * 0.105

        def _is_clear(center, scale=1.0):
            sx, sy = max(1.0, sep_x * scale), max(1.0, sep_y * scale)
            return all(
                math.hypot((center[0] - bp[0]) / sx, (center[1] - bp[1]) / sy) >= 1.0
                for bp in busy_targets
            )

        # Full-swarm stage harmony:
        # 1. Wave 0 (skeleton) launches ALL cursors together at t=0.0s and finishes cleanly in ~1.8s.
        # 2. Right after Wave 0 finishes, Wave 1 (Header/Sidebar) AND Wave 2 (Zoned Cards) run
        #    SIMULTANEOUSLY across all 6 cursors so ZERO cursors ever sit parked idle!
        # 3. Wave 4 (Footer & Page Pill) lands cleanly at the end as the finishing signature.
        wave0_done = getattr(self, "wave0_names", set()) <= self.ppt_committed
        has_core_active_or_pending = (
            any(int(u.get("wave", 2)) < 4 for u in self.pending)
            or any(ag.action and int(ag.action.get("wave", 2)) < 4 for ag in active_peers)
        )

        my_lane = agent.index + 1
        # Count pending units PER WAVE per lane (fixes the Wave 0 solo-cursor bug where future Wave 2
        # items prevented idle peers from helping finish Wave 0!)
        cur_wave_lane_counts = {}
        for u in self.pending:
            uw_c = int(u.get("wave", 2))
            ln_c = int(u.get("lane", 1))
            cur_wave_lane_counts[(uw_c, ln_c)] = cur_wave_lane_counts.get((uw_c, ln_c), 0) + 1

        own_clear = []
        steal_clear = []
        fallback_tight = []

        for ordinal, u in enumerate(self.pending):
            uw = int(u.get("wave", 2))
            if not wave0_done and uw > 0:
                continue
            if has_core_active_or_pending and uw >= 4:
                continue
            if not self._unit_deps(u) <= self.ppt_committed:
                continue

            x0, y0, x1, y1 = self._unit_box(u)
            center = ((x0 + x1) * 0.5, (y0 + y1) * 0.5)
            clear_std = _is_clear(center, 0.72 if uw == 0 else 1.0)
            seq_idx = int(u.get("_seq_idx", ordinal))
            u_lane = int(u.get("lane", my_lane))
            is_own = u_lane == my_lane

            # Wave 1 & Wave 2 share the same active phase so Header/Sidebar & Cards build simultaneously!
            eff_stage = 0 if uw == 0 else (1 if uw in (1, 2, 3) else 4)
            # Inside own lane, Wave 1 (title) still precedes Wave 2 cleanly via uw * 50000
            score = eff_stage * 1_000_000.0 + (uw * 50_000.0 if is_own else 0.0) + seq_idx * 100.0 + math.dist(agent.pos, center) * 0.04

            if is_own:
                if clear_std:
                    own_clear.append((score, ordinal, u))
                else:
                    fallback_tight.append((score, ordinal, u))
            else:
                my_rem_in_wave = cur_wave_lane_counts.get((uw, my_lane), 0)
                peer_rem_in_wave = cur_wave_lane_counts.get((uw, u_lane), 0)
                if uw == 0:
                    if my_rem_in_wave == 0 and clear_std:
                        steal_clear.append((score, ordinal, u))
                elif uw == 4:
                    if my_rem_in_wave == 0 and clear_std:
                        steal_clear.append((score, ordinal, u))
                else:
                    my_core_rem = sum(cur_wave_lane_counts.get((w_k, my_lane), 0) for w_k in (1, 2, 3))
                    if my_core_rem == 0 and peer_rem_in_wave >= 3 and clear_std:
                        steal_clear.append((score + 20_000.0, ordinal, u))

        if own_clear:
            return min(own_clear)[2]
        if steal_clear:
            return min(steal_clear)[2]
        if fallback_tight and len(active_peers) < max(2, self.count // 2):
            return min(fallback_tight)[2]
        return None

    def pause_recording(self):
        if self.recorder and self.capture_started is not None:
            stopped = time.perf_counter()
            self.recorder.pause()
            self.recording_intervals.append((self.capture_started, stopped))
            self.capture_started = None

    def _draw_capsule(self, cv, x0, y0, pw, ph, fill_col, border_col=None, border_w=1.2):
        """Draw a razor-sharp stadium capsule in 1-2 Tk primitives (3x faster than oval+oval+rect!)."""
        rad = ph * 0.5
        cy = y0 + rad
        lx = x0 + rad
        rx = max(lx + 1.0, x0 + pw - rad)
        if border_col:
            cv.create_line(lx, cy, rx, cy, fill=border_col, width=ph + 2.0 * border_w, capstyle="round", tags="swarm")
        cv.create_line(lx, cy, rx, cy, fill=fill_col, width=ph, capstyle="round", tags="swarm")

    def _draw_selection_handles(self, cv, x0, y0, x1, y1, col, u, full_eight=False):
        hs = 2.4 * u
        pts = [(x0, y0), (x1, y0), (x0, y1), (x1, y1)]
        if full_eight and abs(x1 - x0) >= 60 * u and abs(y1 - y0) >= 36 * u:
            mx, my = (x0 + x1) * 0.5, (y0 + y1) * 0.5
            pts.extend([(mx, y0), (mx, y1), (x0, my), (x1, my)])
        for hx, hy in pts:
            cv.create_rectangle(hx - hs, hy - hs, hx + hs, hy + hs, fill="white", outline=col, width=1.1 * u, tags="swarm")

    def _render_overlay_item(self, cv, ov, col, tint, px, ow, oh, sx_scale, sy_scale, u):
        kind = ov["kind"]

        for g in ov.get("guides") or ():
            if g[0] == "v":
                gx, gy0 = px((g[1], g[2]))
                _, gy1 = px((g[1], g[3]))
                cv.create_line(gx, gy0, gx, gy1, fill="#EF4444", width=1.0 * u, dash=(4, 3), tags="swarm")
            else:
                gx0, gy = px((g[2], g[1]))
                gx1, _ = px((g[3], g[1]))
                cv.create_line(gx0, gy, gx1, gy, fill="#EF4444", width=1.0 * u, dash=(4, 3), tags="swarm")

        if kind == "cad_box":
            x0, y0 = px(ov["rect"][:2])
            x1, y1 = px(ov["rect"][2:])
            if ov.get("shape") in ("oval", "circle", "ellipse"):
                cv.create_oval(x0, y0, x1, y1, outline=col, width=1.5 * u, tags="swarm")
            else:
                cv.create_rectangle(x0, y0, x1, y1, outline=col, width=1.5 * u, tags="swarm")
            if abs(x1 - x0) >= 18 * u and abs(y1 - y0) >= 14 * u:
                self._draw_selection_handles(cv, x0, y0, x1, y1, col, u, full_eight=ov.get("show_dim", False))
            if ov.get("show_dim") and abs(x1 - x0) > 48 * u and abs(y1 - y0) > 24 * u:
                dim_txt = ov["dim"]
                dim_fs = max(9, int(round(9.2 * u)))
                tw = max(40.0 * u, self._measure_tk_text(dim_txt, "Segoe UI", dim_fs, bold=True) + 12.0 * u)
                th = max(16.0 * u, dim_fs + 5.5 * u)
                cx_d = clamp((x0 + x1) * 0.5, 4.0 + tw * 0.5, ow - 4.0 - tw * 0.5)
                cy_d = clamp(y1 + th * 0.72, 4.0 + th * 0.5, oh - 4.0 - th * 0.5)
                self._draw_capsule(cv, cx_d - tw * 0.5, cy_d - th * 0.5, tw, th, col, border_col="white", border_w=1.0)
                cv.create_text(cx_d, cy_d, text=dim_txt, fill="white", font=("Segoe UI", -dim_fs, "bold"), tags="swarm")

        elif kind == "asset_scan":
            x0, y0 = px(ov["rect"][:2])
            x1, y1 = px(ov["rect"][2:])
            prog = ov["progress"]
            cv.create_rectangle(x0, y0, x1, y1, outline=tint, width=1.1 * u, dash=(3, 3), tags="swarm")
            bl = min(13 * u, abs(x1 - x0) * 0.25, abs(y1 - y0) * 0.25)
            bw = 2.0 * u
            for cx_b, cy_b, sx_s, sy_s in ((x0, y0, 1, 1), (x1, y0, -1, 1), (x0, y1, 1, -1), (x1, y1, -1, -1)):
                cv.create_line(cx_b, cy_b, cx_b + sx_s * bl, cy_b, fill=col, width=bw, tags="swarm")
                cv.create_line(cx_b, cy_b, cx_b, cy_b + sy_s * bl, fill=col, width=bw, tags="swarm")
            if prog < 0.99:
                sy_beam = y0 + (y1 - y0) * prog
                cv.create_line(x0 + 2, sy_beam, x1 - 2, sy_beam, fill=col, width=1.8 * u, tags="swarm")

        elif kind == "vector_guide":
            pts = ov["points"]
            if len(pts) >= 2:
                flat = [z for p in pts for z in px(p)]
                dash_pat = (round(3 * u), round(2 * u)) if ov.get("dash") and ov["dash"] != "solid" else None
                ink = _safe_hex(ov["color"], col.lstrip("#"))
                cv.create_line(*flat, fill=ink, width=max(1.5, ov["width"] * sx_scale), capstyle="round", joinstyle="round", dash=dash_pat, tags="swarm")

        elif kind == "burst_sweep":
            tip = ov.get("tip")
            if tip is not None:
                tx, ty = px(tip)
                rad = 4.2 * u
                cv.create_oval(tx - rad, ty - rad, tx + rad, ty + rad, fill="", outline=tint, width=1.4 * u, tags="swarm")

        elif kind == "live_text":
            rx0, ry0, rx1, ry1 = ov["rect"]
            if ov.get("show_frame", True):
                x0, y0 = px((rx0, ry0))
                x1, y1 = px((rx1, ry1))
                cv.create_rectangle(x0, y0, x1, y1, outline=tint, width=1.0 * u, dash=(3, 2), tags="swarm")

            segments = extract_visible_segments(ov["layout"], ov["shown"])
            for seg_x_pt, seg_y_pt, seg_txt, st in segments:
                fam, eff_fs_pt, bold, italic, tcol, y_shift_pt = st
                spx, spy = px((seg_x_pt, seg_y_pt + y_shift_pt))
                fs_px = max(8, int(round(eff_fs_pt * sy_scale)))
                style_tokens = []
                if bold:
                    style_tokens.append("bold")
                if italic:
                    style_tokens.append("italic")
                tk_font = (fam, -fs_px, " ".join(style_tokens)) if style_tokens else (fam, -fs_px)
                cv.create_text(spx, spy, text=seg_txt, fill=tcol, anchor="w", font=tk_font, tags="swarm")

            caret = ov.get("caret")
            if caret and ov.get("blink", True):
                cx_pt, cy_pt, lh_pt = caret
                cpx, cpy = px((cx_pt, cy_pt))
                ch_half = max(5.0, lh_pt * sy_scale * 0.42)
                cv.create_line(cpx, cpy - ch_half, cpx, cpy + ch_half, fill=col, width=max(1.5, 1.6 * u), tags="swarm")

    def _draw_codex_hud(self, cv, ox, oy, ow, oh, vx, vy, vw, vh, u):
        if not self.show_hud:
            return
        slide_bot_rel = (vy + vh) - oy
        below_space = oh - slide_bot_rel
        bar_h = max(22.0, round(22.0 * u))
        if below_space < bar_h + 6.0:
            return
        cy = slide_bot_rel + min(below_space * 0.5, bar_h * 0.72 + 8.0)
        if cy + bar_h * 0.5 > oh - 3.0:
            cy = oh - 3.0 - bar_h * 0.5

        active_waves = [int(ag.action.get("wave", 2)) for ag in self.agents if ag.active and ag.action]
        if self.done_at is not None:
            stage_txt = "Completed · Editable PowerPoint Ready"
            dot_col = "#22C55E"
        else:
            cur_w = min(active_waves) if active_waves else (min((int(u_item.get("wave", 2)) for u_item in self.pending), default=4))
            stage_txt = STAGE_NAMES.get(cur_w, f"Stage · Wave {cur_w}")
            dot_col = "#38BDF8"

        active_n = sum(1 for ag in self.agents if ag.active)
        total_n = len(self.actions)
        done_n = min(total_n, self.committed_count)
        hud_txt = f"Codex · PowerPoint Live   |   {stage_txt}   |   {done_n}/{total_n} objects   |   {active_n} active"
        font_px = max(10, int(round(10.0 * u)))
        txt_w = self._measure_tk_text(hud_txt, "Segoe UI", font_px, bold=True)
        bar_w = min(vw * 0.92, txt_w + 42.0 * u)
        bx0 = (vx - ox) + (vw - bar_w) * 0.5
        by0 = cy - bar_h * 0.5

        self._draw_capsule(cv, bx0, by0, bar_w, bar_h, "#0F172A", border_col="#334155", border_w=1.1)
        dr = 3.5 * u
        dx = bx0 + 13.0 * u
        cv.create_oval(dx - dr, cy - dr, dx + dr, cy + dr, fill=dot_col, outline="", tags="swarm")
        cv.create_text(
            dx + 9.5 * u,
            cy,
            text=hud_txt,
            fill="#F8FAFC",
            anchor="w",
            font=("Segoe UI", -font_px, "bold"),
            tags="swarm",
        )

    def paint(self):
        cv = self.canvas
        ox, oy, ow, oh = self.overlay_rect
        vx, vy, vw, vh = self.slide_rect
        u = clamp(vw / 1280.0, 0.95, 2.0)
        sx_scale = vw / max(1.0, self.sw)
        sy_scale = vh / max(1.0, self.sh)
        dx_slide = vx - ox
        dy_slide = vy - oy

        def px(p):
            return dx_slide + p[0] * sx_scale, dy_slide + p[1] * sy_scale

        cv.delete("swarm")

        # 1. Brief selection-lock frames (Wave 0 panels & placed images only)
        self.selection_locks = [sl for sl in self.selection_locks if self.clock - sl[0] < 0.14]
        for _, rect, col in self.selection_locks:
            x0, y0 = px(rect[:2])
            x1, y1 = px(rect[2:])
            if abs(x1 - x0) >= 16.0 and abs(y1 - y0) >= 12.0:
                cv.create_rectangle(x0, y0, x1, y1, outline=col, width=1.3 * u, tags="swarm")

        # 2. Active agent overlays (zero stale pinned_handoffs -> zero double-image ghosting!)
        for ag in self.agents:
            if ag.used and ag.overlay:
                self._render_overlay_item(cv, ag.overlay, ag.color, ag.tint, px, ow, oh, sx_scale, sy_scale, u)

        # 4. Cursors + subtle long-travel trails + calm static-width DPI-safe role pills
        placed_pills = []
        if self.record_scope == "slide":
            min_px, min_py = dx_slide + 4.0, dy_slide + 4.0
            max_px, max_py = dx_slide + vw - 4.0, dy_slide + vh - 4.0
        else:
            min_px, min_py = 4.0, 4.0
            max_px, max_py = ow - 4.0, oh - 4.0

        for ag in self.agents:
            if not ag.used and self.done_at is None:
                continue

            if self.done_at is not None:
                exit_t = min(1.0, (time.perf_counter() - self.done_at) / 0.38)
                if exit_t >= 0.95:
                    continue
                cur_pos = blend(ag.idle_anchor, ag.home, smooth(exit_t))
            elif not ag.active:
                cur_pos = ag.idle_anchor
            else:
                cur_pos = ag.pos

            if self.show_trails and ag.active and ag.tool == "pointer" and len(ag.trail) >= 2:
                tpts = [px(p) for p in ag.trail]
                max_seg_px = 32.0 * u
                max_total_px = 34.0 * u
                acc_len = 0.0
                segs = []
                for k in range(len(tpts) - 1, 0, -1):
                    p0, p1 = tpts[k - 1], tpts[k]
                    seg_d = math.hypot(p1[0] - p0[0], p1[1] - p0[1])
                    if seg_d > max_seg_px or seg_d < 0.8:
                        break
                    if acc_len + seg_d > max_total_px:
                        remain = max(0.0, max_total_px - acc_len)
                        ratio = remain / seg_d
                        p0 = (p1[0] + (p0[0] - p1[0]) * ratio, p1[1] + (p0[1] - p1[1]) * ratio)
                        segs.append((p0, p1))
                        break
                    segs.append((p0, p1))
                    acc_len += seg_d
                n_s = len(segs)
                for idx_s, (p0, p1) in enumerate(reversed(segs)):
                    frac = (idx_s + 1) / max(1, n_s)
                    col_t = ag.color if frac > 0.65 else ag.tint
                    cv.create_line(p0[0], p0[1], p1[0], p1[1], fill=col_t, width=max(1.0, 1.8 * u * frac), capstyle="round", tags="swarm")

            x, y = px(cur_pos)
            x = clamp(x, min_px + 2.0, max_px - 2.0)
            y = clamp(y, min_py + 2.0, max_py - 2.0)

            def coords(pts_rel):
                return [v for dx, dy in pts_rel for v in (x + dx * u, y + dy * u)]

            col = ag.color
            if ag.tool == "cross":
                for p, q in (((-6.0, 0), (6.0, 0)), ((0, -6.0), (0, 6.0))):
                    cv.create_line(*coords([p, q]), fill="white", width=2.8 * u, tags="swarm")
                    cv.create_line(*coords([p, q]), fill=col, width=1.4 * u, tags="swarm")
            elif ag.tool == "pen":
                cv.create_polygon(
                    *coords([(0, 0), (1.8, -6.5), (12.2, -16.9), (15.6, -13.5), (5.2, -3.1)]),
                    fill=col, outline="white", width=1.2 * u, tags="swarm",
                )
            elif ag.tool == "ibeam":
                for p, q in (((0, -5.8), (0, 5.8)), ((-2.4, -5.8), (2.4, -5.8)), ((-2.4, 5.8), (2.4, 5.8))):
                    cv.create_line(*coords([p, q]), fill="white", width=2.6 * u, tags="swarm")
                    cv.create_line(*coords([p, q]), fill=col, width=1.3 * u, tags="swarm")
            else:
                cv.create_polygon(
                    *coords([(0, 0), (0, 15.2), (4.2, 11.7), (7.6, 18.0), (10.0, 16.7), (6.8, 10.5), (12.6, 10.5)]),
                    fill=col, outline="white", width=1.2 * u, tags="swarm",
                )

            if self.agent_labels and self.count > 1:
                badge_text = f"{ag.name} · {ag.status_verb}" if (self.show_status_verb and ag.active and ag.status_verb) else ag.name
                font_px = max(9, int(round(9.0 * u)))
                pill_h = max(16.0 * u, font_px + 7.0 * u)
                pill_w = max(46.0 * u, self._measure_tk_text(badge_text, "Segoe UI", font_px, bold=True) + 15.0 * u)

                candidates = [
                    (x + 11.0 * u, y + 11.0 * u),
                    (x - pill_w - 8.0 * u, y + 11.0 * u),
                    (x + 11.0 * u, y - pill_h - 8.0 * u),
                    (x - pill_w - 8.0 * u, y - pill_h - 8.0 * u),
                ]
                lx, ly = clamp(candidates[0][0], min_px, max_px - pill_w), clamp(candidates[0][1], min_py, max_py - pill_h)
                for cx_p, cy_p in candidates:
                    clx = clamp(cx_p, min_px, max_px - pill_w)
                    cly = clamp(cy_p, min_py, max_py - pill_h)
                    pbox = (clx, cly, clx + pill_w, cly + pill_h)
                    if not any(min(pbox[2], pb[2]) > max(pbox[0], pb[0]) and min(pbox[3], pb[3]) > max(pbox[1], pb[1]) for pb in placed_pills):
                        lx, ly = clx, cly
                        break
                placed_pills.append((lx, ly, lx + pill_w, ly + pill_h))

                self._draw_capsule(cv, lx, ly, pill_w, pill_h, col, border_col="white", border_w=1.2 * u)
                cv.create_text(
                    lx + pill_w * 0.5,
                    ly + pill_h * 0.5,
                    text=badge_text,
                    fill="white",
                    font=("Segoe UI", -font_px, "bold"),
                    tags="swarm",
                )

        # 5. Sleek Codex Live Status HUD in empty PowerPoint workspace margin
        self._draw_codex_hud(cv, ox, oy, ow, oh, vx, vy, vw, vh, u)

    def tick(self):
        now = time.perf_counter()
        # Allow up to 65ms step so real-time duration stays crisp (~20-26s) instead of stretching into slow-mo!
        dt = min(0.065, max(0.0, now - self.last))
        self.last = now

        self.drain_done_queue()

        if now >= self.next_probe:
            probed = self.probe_win32()
            self.next_probe = now + 0.22
            if probed is None:
                if self.overlay_rect is not None:
                    self.pause_recording()
                    self.overlay_rect = None
                    gui.ShowWindow(self.hwnd, wc.SW_HIDE)
            else:
                wrect, srect = probed
                # Tight overlay surface over slide+HUD while Recorder captures FULL PowerPoint window!
                target_overlay = compute_tight_overlay_rect(wrect, srect, self.record_scope)
                target_capture = wrect if self.record_scope == "window" else srect
                self.window_rect = wrect
                self.slide_rect = srect
                if target_overlay != self.overlay_rect:
                    self.pause_recording()
                    self.overlay_rect = target_overlay
                    ox, oy, ow, oh = target_overlay
                    self.root.geometry(f"{ow}x{oh}")
                    self.root.update_idletasks()
                    gui.SetWindowPos(self.hwnd, wc.HWND_TOPMOST, ox, oy, ow, oh, wc.SWP_NOACTIVATE | wc.SWP_SHOWWINDOW)
                    if self.recorder:
                        self.recorder.start(target_capture)
                        self.capture_started = time.perf_counter()

        if not self.overlay_rect:
            self.hidden_at = self.hidden_at or now
            if now - self.hidden_at > self.timeout:
                raise TimeoutError("Target PowerPoint slide was not visible; paused drawing.")
            return
        self.hidden_at = None

        if self.first is None:
            self.first = now
            print("READY", flush=True)
        if self.recorder:
            self.recorder.check()
        if now - self.first < 0.20:
            return

        self.clock += dt * self.pace

        for ag in self.agents:
            if ag.generator is None:
                u_item = self.choose(ag)
                if u_item is None:
                    continue
                self.pending.remove(u_item)
                ag.generator = ag.run(u_item)
            try:
                next(ag.generator)
            except StopIteration:
                act = ag.action
                if act["type"] == "micro_burst":
                    for m in act["members"]:
                        self.timeline.append({
                            "name": m["name"],
                            "type": m["type"],
                            "agent": ag.name,
                            "hand": ag.index + 1,
                            "wave": int(m.get("wave", 2)),
                            "burst": act["name"],
                            "wall_start": round(ag.wall_start, 3),
                            "wall_end": round(now - self.first, 3),
                        })
                else:
                    self.timeline.append({
                        "name": act["name"],
                        "type": act["type"],
                        "agent": ag.name,
                        "hand": ag.index + 1,
                        "wave": int(act.get("wave", 2)),
                        "wall_start": round(ag.wall_start, 3),
                        "wall_end": round(now - self.first, 3),
                    })
                if len(self.timeline) % 12 == 0 or len(self.timeline) == len(self.actions):
                    print(f"COMMITTED {len(self.timeline)}/{len(self.actions)}", flush=True)
                ag.generator = None

        self.peak = max(self.peak, sum(ag.generator is not None for ag in self.agents))

        all_dispatched = (not self.pending) and (not any(ag.generator for ag in self.agents))
        all_ppt_done = all_dispatched and (not self.in_flight_tokens)

        if not any(ag.generator for ag in self.agents) and self.pending:
            if not self.in_flight_tokens and not any(self.choose(ag) for ag in self.agents):
                raise RuntimeError("Deadlock: remaining actions cannot satisfy dependencies.")

        if all_ppt_done:
            if not self.finish_sent:
                self.finish_sent = True
                self.done_at = now
                self.worker.submit_lo((
                    "finish",
                    self.output,
                    self.export_preview,
                    self.z_ordered_names,
                    self.boxes,
                    self.actions,
                    self.auto_fix_text_overflow,
                ))
            if self.worker.finished_event.is_set() and (now - self.done_at >= 1.05):
                self.root.quit()
                return

        self.paint()
        self.frames += 1

    def loop(self):
        try:
            self.tick()
        except Exception:
            self.error = traceback.format_exc()
            self.root.quit()
            return
        self.next_frame_at += 1.0 / 60.0
        now = time.perf_counter()
        if self.next_frame_at < now - 0.035:
            self.next_frame_at = now
        delay_ms = max(1, int((self.next_frame_at - now) * 1000.0))
        self.root.after(delay_ms, self.loop)

    def run(self):
        timer_set = False
        if sys.platform == "win32":
            try:
                ctypes.windll.winmm.timeBeginPeriod(1)
                timer_set = True
            except Exception:
                pass
        try:
            self.next_frame_at = time.perf_counter()
            self.root.after(0, self.loop)
            self.root.mainloop()
        finally:
            if timer_set:
                try:
                    ctypes.windll.winmm.timeEndPeriod(1)
                except Exception:
                    pass
            if not self.finish_sent:
                self.worker.submit_hi(("abort",))
            self.worker.join(timeout=20)
            self.pause_recording()
            try:
                self.root.destroy()
            except Exception:
                pass
        if self.error:
            raise RuntimeError(self.error)
        if self.worker.error:
            raise RuntimeError(self.worker.error)
        if len(self.timeline) != len(self.actions):
            raise RuntimeError("Drawing window closed before completion.")

        if self.recorder:
            self.recorder.finish()

        fin = self.worker.finish_info or {"output": str(self.output), "preview_png": None, "audit": None}
        elapsed = max(0.01, self.last - self.first)
        report = {
            "confirmed": len(self.timeline),
            "peak_cursors": self.peak,
            "displayed_frames": self.frames,
            "avg_overlay_fps": round(self.frames / elapsed, 1),
            "typed_characters": self.typed,
            "drawing_seconds": round(elapsed, 2),
            "effective_pace": round(self.pace, 2),
            "record_scope": self.record_scope,
            "record_backend": self.recorder.backend if self.recorder else None,
            "output": fin["output"],
            "preview_png": fin["preview_png"],
            "audit": fin.get("audit"),
            "actions": self.timeline,
        }
        report_path = self.plan_path.with_name(self.plan_path.stem + "-swarm-log.json")
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({k: v for k, v in report.items() if k != "actions"}, ensure_ascii=False), flush=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("plan", type=Path)
    parser.add_argument("--timeout", type=int, default=180)
    parser.add_argument("--export-preview", type=Path, help="Optional PNG path to export Slide after completion.")
    parser.add_argument("--check", action="store_true", help="Validate & lint plan without connecting to PowerPoint.")
    args = parser.parse_args()

    if sys.platform == "win32":
        ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)
    plan_path = args.plan.resolve(strict=True)
    if args.check:
        plan = json.loads(plan_path.read_text(encoding="utf-8-sig"))
        actions, _, boxes, units = prepare_plan(plan, plan_path.parent)
        canvas = plan.get("canvas") or [960, 540]
        sw, sh = float(canvas[0]), float(canvas[1])
        lanes = int(plan.get("cursors", 6))
        est_raw = estimate_plan_duration(units, lanes, sw)
        user_pace = float(plan.get("pace", 1.0))
        target_dur = plan.get("target_duration", 24.0)
        eff_pace = clamp((est_raw / float(target_dur)) * user_pace, 0.85, 2.20) if (target_dur and float(target_dur) > 0) else user_pace
        lint = lint_plan(plan, actions, boxes, sw, sh, units=units)
        by_wave, by_lane = {}, {}
        for a in actions:
            by_wave[a["wave"]] = by_wave.get(a["wave"], 0) + 1
            by_lane[a["lane"]] = by_lane.get(a["lane"], 0) + 1
        print(json.dumps({
            "valid": True,
            "actions": len(actions),
            "execution_units": len(units),
            "cursors": lanes,
            "record_scope": str(plan.get("record_scope", "window")),
            "record_fps": int(plan.get("record_fps", 60)),
            "target_duration": target_dur,
            "estimated_raw_seconds": est_raw,
            "effective_pace": round(eff_pace, 2),
            "estimated_final_seconds": round(est_raw / max(0.25, eff_pace), 1),
            "actions_by_wave": dict(sorted(by_wave.items())),
            "actions_by_lane": dict(sorted(by_lane.items())),
            "lint": lint,
        }, ensure_ascii=False, indent=2))
    else:
        SwarmShow(plan_path, args.timeout, export_preview=args.export_preview).run()


if __name__ == "__main__":
    main()
