"""Buttery-smooth 60fps multi-agent cursor live drawing engine for PowerPoint.

Architecture:
  - Main Thread: 100% COM-free 60fps Tkinter transparent overlay rendering 6+ Figma-style
    studio cursors, short luxury comet trails, translucent CAD anchor boxes, futuristic
    corner-bracket laser scan reveals, glowing vector guides, AND true 60fps character-by-
    character live typing (both standalone textboxes and text inside shapes!) with seamless
    zero-pop handoff to native PowerPoint.
  - Background COM Worker Thread: Dedicated STA COM apartment (`CoInitialize`) executing
    atomic PowerPoint writes asynchronously via FIFO queue so PowerPoint COM IPC latency
    NEVER freezes cursor motion.
  - Overlap-Only Zero-Flash Z-Order: Wave 0 background panels + overlapping painter's order
    prevent mid-show occlusion and end-of-show Z-order wipes.
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
except ImportError:
    tk = None

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
    apply_background,
    apply_shadow,
    apply_z_order,
    auto_assign_waves_and_lanes,
    bounds,
    curve_samples,
    draw_path,
    fill,
    has_cjk,
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

KEY = "#010203"

DEFAULT_PROFILES = [
    {"name": "Layout", "color": "#2563EB", "tint": "#93C5FD"},
    {"name": "Visual-1", "color": "#7C3AED", "tint": "#C4B5FD"},
    {"name": "Visual-2", "color": "#059669", "tint": "#6EE7B7"},
    {"name": "Route", "color": "#EA580C", "tint": "#FDBA74"},
    {"name": "Text-1", "color": "#DB2777", "tint": "#F9A8D4"},
    {"name": "Text-2", "color": "#0891B2", "tint": "#67E8F9"},
    {"name": "Detail", "color": "#4F46E5", "tint": "#A5B4FC"},
    {"name": "Polish", "color": "#D97706", "tint": "#FCD34D"},
]


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
    # Avoid colliding with transparent color key #010203
    if h.lower() == "010203":
        return "#020304"
    return "#" + h


def _extract_plain_text(action):
    if action.get("runs"):
        raw = "".join(str(r.get("text", "")) for r in action["runs"])
    else:
        raw = str(action.get("text", ""))
    return raw.replace("\r\n", "\n").replace("\r", "\n")


def _char_em_width(ch):
    if "\u4e00" <= ch <= "\u9fff" or "\u3000" <= ch <= "\uffef":
        return 1.0
    if ch in "ilIj1.,:;|!`' ":
        return 0.30
    if ch in "mwMW@":
        return 0.80
    if ch.isupper():
        return 0.65
    return 0.52


def _wrap_text_lines(text, max_width_pt, font_size_pt, word_wrap=True):
    """Estimate visual wrapped lines in slide points so the live I-beam caret tracks each line accurately."""
    paragraphs = text.split("\n")
    if not word_wrap or max_width_pt <= 0 or font_size_pt <= 0:
        return paragraphs
    cap = max(1.2, max_width_pt / max(4.0, font_size_pt))
    lines = []
    for para in paragraphs:
        if not para:
            lines.append("")
            continue
        cur = ""
        cur_w = 0.0
        for ch in para:
            cw = _char_em_width(ch)
            if cur and cur_w + cw > cap * 1.02:
                lines.append(cur)
                cur = ch
                cur_w = cw
            else:
                cur += ch
                cur_w += cw
        lines.append(cur)
    return lines or [""]


def _estimate_caret_pos(full_text, shown_count, rect, action):
    """Compute exact slide-point (x, y) for the live typing caret at character `shown_count`."""
    x0, y0, x1, y1 = rect
    ml = float(action.get("margin_left", 0))
    mr = float(action.get("margin_right", 0))
    mt = float(action.get("margin_top", 0))
    mb = float(action.get("margin_bottom", 0))
    ix0, iy0 = x0 + ml, y0 + mt
    ix1, iy1 = max(ix0 + 4.0, x1 - mr), max(iy0 + 4.0, y1 - mb)
    iw, ih = ix1 - ix0, iy1 - iy0

    fs = float(action.get("font_size", 14.0))
    line_sp = float(action.get("line_spacing", 1.0))
    line_h = fs * 1.24 * line_sp
    align = str(action.get("alignment", "left")).lower()
    valign = str(action.get("vertical_alignment", "middle")).lower()
    wrap = bool(action.get("word_wrap", True))

    all_lines = _wrap_text_lines(full_text, iw, fs, wrap)
    total_h = max(line_h, len(all_lines) * line_h)
    if valign == "top":
        top_y = iy0
    elif valign == "bottom":
        top_y = iy1 - total_h
    else:
        top_y = iy0 + (ih - total_h) * 0.5

    prefix = full_text[: max(0, shown_count)]
    typed_lines = _wrap_text_lines(prefix, iw, fs, wrap)
    line_idx = min(max(0, len(typed_lines) - 1), max(0, len(all_lines) - 1))
    cur_line = typed_lines[-1] if typed_lines else ""
    cur_line_w = min(iw, sum(_char_em_width(c) for c in cur_line) * fs)
    full_line_w = min(iw, sum(_char_em_width(c) for c in all_lines[line_idx]) * fs)

    if align in ("center", "centre"):
        # Smoothly advance from centered line start to current typed character
        line_start_x = ix0 + (iw - full_line_w) * 0.5
        cx = line_start_x + cur_line_w
    elif align == "right":
        cx = ix1
    else:
        cx = ix0 + cur_line_w

    cy = top_y + (line_idx + 0.5) * line_h
    return (clamp(cx, ix0, ix1), clamp(cy, iy0, iy1), line_h)


def blend(a, b, t):
    return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)


def ease_out_cubic(t):
    u = 1.0 - max(0.0, min(1.0, t))
    return 1.0 - u * u * u


def smooth(t):
    t = max(0.0, min(1.0, t))
    return t * t * t * (10 + t * (-15 + 6 * t))


def clamp(v, lo, hi):
    return max(lo, min(hi, v))


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
    actions.sort(key=lambda a: (int(a.get("wave", 1)), float(a.get("z_index", 0))))
    deps, boxes = validate(actions, set(existing), lanes)
    return actions, deps, boxes


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


def _commit_native_action(shapes, a, skip_text=False):
    """Execute one atomic PowerPoint COM write on the dedicated COM worker thread."""
    kind = a["type"]
    if kind == "add_shape":
        x, y, r, b = bounds(a)
        w, h = max(0.5, r - x), max(0.5, b - y)
        preset = a.get("shape_type_id") or PRESETS[a["shape"]]
        s = shapes.AddShape(preset, x, y, w, h)
        s.Visible = 0
        s.LockAspectRatio = 0
        fill(s, a)
        stroke(s, a)
        apply_shadow(s, a)
        set_adjustments(s, a, w, h)
        if a.get("rotation"):
            s.Rotation = float(a["rotation"])
        if not skip_text and (a.get("text") or a.get("runs")):
            write_text_fast(s, a)
        s.Name = a["name"]
        s.Visible = -1
        return s.Name

    if kind == "add_image":
        x, y, r, b = bounds(a)
        w, h = max(0.5, r - x), max(0.5, b - y)
        s = shapes.AddPicture(str(a["image_path"]), 0, -1, x, y, w, h)
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
        return s.Name

    if kind in ("add_line", "freeform"):
        segments = list(path_segments(path_commands(a)))
        s = draw_path(shapes, segments)
        s.Visible = 0
        fill(s, a)
        stroke(s, a)
        apply_shadow(s, a)
        s.Name = a["name"]
        s.Visible = -1
        return s.Name

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
        s.Visible = -1
        return s.Name

    if kind == "duplicate_shape":
        x, y, r, b = bounds(a)
        w, h = max(0.5, r - x), max(0.5, b - y)
        src = shapes.Item(a["source"])
        s = src.Duplicate().Item(1)
        s.Left, s.Top, s.Width, s.Height = x, y, w, h
        s.Name = a["name"]
        return s.Name

    if kind == "group_shapes":
        s = shapes.Range(tuple(a["members"])).Group()
        s.Name = a["name"]
        return s.Name

    raise ValueError(f"Unsupported action type: {kind}")


class ComWorker(threading.Thread):
    """Dedicated PowerPoint STA COM thread so COM IPC never blocks the 60fps cursor loop."""

    def __init__(self, deck_path, slide_index, background, expected_canvas):
        super().__init__(daemon=True, name="ppt-com-worker")
        self.deck_path = deck_path
        self.slide_index = slide_index
        self.background = background
        self.expected_canvas = expected_canvas
        self.job_queue = queue.Queue()
        self.done_queue = queue.Queue()
        self.ready_event = threading.Event()
        self.finished_event = threading.Event()
        self.init_info = None
        self.finish_info = None
        self.error = None

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
            slide = deck.Slides.Item(self.slide_index)
            win.View.GotoSlide(slide.SlideIndex)
            if self.background:
                apply_background(slide, self.background)
            win.View.ZoomToFit = True
            win.View.Zoom = win.View.Zoom
            win.Activate()

            ppt_hwnd = _find_ppt_hwnd(win, self.deck_path)
            _bring_to_foreground(ppt_hwnd)
            time.sleep(0.22)
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

            win_rect = gui.GetWindowRect(ppt_hwnd) if ppt_hwnd else (0, 0, 0, 0)
            shapes = slide.Shapes
            existing = {shapes.Item(i).Name for i in range(1, shapes.Count + 1)}

            self.init_info = {
                "sw": sw,
                "sh": sh,
                "ppt_hwnd": ppt_hwnd,
                "viewport": (vx, vy, vw, vh),
                "win_offset": (vx - win_rect[0], vy - win_rect[1], vw, vh),
                "existing": existing,
            }
            self.ready_event.set()

            with quiet_new_shapes(shapes):
                while True:
                    job = self.job_queue.get()
                    kind = job[0]
                    if kind == "abort":
                        break
                    if kind == "commit":
                        action, token = job[1], job[2]
                        committed_name = _commit_native_action(shapes, action, skip_text=False)
                        self.done_queue.put(("ok", committed_name, token))
                    elif kind == "commit_shape_empty":
                        action, token = job[1], job[2]
                        _commit_native_action(shapes, action, skip_text=True)
                        # Do not mark full action dependency done until text is also written,
                        # only signal the intermediate token so Beat 2 live typing can proceed!
                        self.done_queue.put(("sub", action["name"], token))
                    elif kind == "set_text":
                        action, token = job[1], job[2]
                        sh_obj = shapes.Item(action["name"])
                        write_text_fast(sh_obj, action)
                        self.done_queue.put(("ok", action["name"], token))
                    elif kind == "finish":
                        _, out_path, export_preview, z_ordered, boxes = job
                        # Overlap-only zero-flash check (0 COM calls when painter's order holds!)
                        apply_z_order(shapes, z_ordered, boxes=boxes)
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
                        self.finish_info = {"output": str(out_path), "preview_png": exported}
                        self.finished_event.set()
                        break
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
        self.name = prof.get("name") or (base_p["name"] if index < len(DEFAULT_PROFILES) else f"Agent-{index + 1}")
        self.color = prof.get("color") or base_p["color"]
        self.tint = prof.get("tint") or _lighten_hex(self.color, 0.58)
        self.home = (
            clamp(show.sw * (0.14 + 0.72 * index / max(1, count)), show.sw * 0.06, show.sw * 0.94),
            show.sh * (0.05 if index % 2 == 0 else 0.94),
        )
        self.pos = self.home
        self.trail = deque(maxlen=7)
        self.tool = "pointer"
        self.generator = None
        self.action = None
        self.overlay = None
        self.rng = random.Random(index + 17)
        self.used = False
        self.active = False
        self.samples = 0
        self.cue_started = 0.0
        self.wall_start = 0.0
        self.idle_anchor = self.pos

    def move_to(self, pos):
        self.pos = pos
        if not self.trail or math.dist(self.trail[-1], pos) >= 0.5:
            self.trail.append(pos)

    def animate(self, duration, update, easing=smooth):
        start = self.show.clock
        dur = max(0.02, duration)
        while True:
            t = min(1.0, (self.show.clock - start) / dur)
            update(easing(t))
            self.samples += 1
            if t >= 1.0:
                break
            yield

    def wait(self, duration):
        end = self.show.clock + duration
        while self.show.clock < end:
            yield

    def wait_for_token(self, token, max_wait=1.5):
        """Hold current overlay seamlessly until ComWorker confirms native PowerPoint commit."""
        deadline = self.show.clock + max_wait
        while token not in self.show.completed_tokens and self.show.clock < deadline:
            yield

    def travel(self, target):
        start = self.pos
        dist = math.dist(start, target)
        if dist < 0.8:
            self.move_to(target)
            return
        self.tool = "pointer"
        # Start fresh comet trail on each travel so no stale straight lines stretch across slide
        self.trail.clear()
        self.trail.append(start)
        bend = min(dist * 0.075, self.show.sw * 0.025) * self.rng.choice((-1.0, 1.0)) * self.rng.uniform(0.35, 1.0)
        perp = (-(target[1] - start[1]) / dist, (target[0] - start[0]) / dist)
        dur = clamp(0.10 + 0.068 * math.log2(1.0 + dist / 14.0), 0.12, 0.36) * self.rng.uniform(0.92, 1.08)

        def step(t):
            x, y = blend(start, target, t)
            h = math.sin(math.pi * t) * bend
            self.move_to((x + perp[0] * h, y + perp[1] * h))

        yield from self.animate(dur, step, ease_out_cubic)
        self.move_to(target)

    def _live_type_sequence(self, rect, a, job_kind):
        """True 60fps character-by-character live typing preview + seamless native COM handoff."""
        full_text = _extract_plain_text(a)
        n_chars = len(full_text)
        start_x, start_y, line_h = _estimate_caret_pos(full_text, 0, rect, a)
        yield from self.travel((start_x, start_y))
        self.tool = "ibeam"
        self.trail.clear()

        # Crisp natural cadence: short titles ~0.32-0.55s, dense paragraphs ~0.60-1.05s
        if n_chars <= 8:
            dur = clamp(0.24 + n_chars * 0.032, 0.26, 0.52)
        elif n_chars <= 24:
            dur = clamp(0.32 + n_chars * 0.020, 0.36, 0.72)
        else:
            dur = clamp(0.45 + n_chars * 0.011, 0.55, 1.05)

        def type_step(t):
            shown = max(1, min(n_chars, int(round(n_chars * t)))) if n_chars > 0 else 0
            cx, cy, lh = _estimate_caret_pos(full_text, shown, rect, a)
            self.move_to((cx, cy))
            self.overlay = {
                "kind": "live_text",
                "rect": rect,
                "text": full_text[:shown],
                "action": a,
                "caret": (cx, cy, lh),
                "blink": int(self.show.clock * 8.0) % 2 == 0,
            }

        yield from self.animate(dur, type_step, lambda t: t)
        self.show.typed += n_chars

        # Seamless zero-pop handoff: keep full typed preview pinned until PowerPoint finishes writing underneath!
        token = self.show.enqueue_job(job_kind, a)
        yield from self.wait_for_token(token)
        self.overlay = None
        self.tool = "pointer"
        self.idle_anchor = self.pos
        self.show.ripples.append((self.show.clock, self.pos, self.color))
        yield from self.wait(self.rng.uniform(0.02, 0.045))

    def do_shape(self, a):
        x, y, r, b = bounds(a)
        has_text = bool(_extract_plain_text(a).strip())
        yield from self.travel((x, y))
        self.tool = "cross"
        dur = clamp(0.22 + math.dist((x, y), (r, b)) / (self.show.sw * 1.65), 0.22, 0.48)
        fill_preview = a.get("fill_color") or ((a.get("gradient") or [None])[0])

        def drag(t):
            cx, cy = blend((x, y), (r, b), t)
            self.move_to((cx, cy))
            self.overlay = {
                "kind": "cad_box",
                "rect": (x, y, cx, cy),
                "shape": a.get("shape", "rectangle"),
                "dim": f"{round(cx - x)}×{round(cy - y)}",
                "fill": fill_preview,
            }

        yield from self.animate(dur, drag, ease_out_cubic)

        if has_text:
            # Beat 1: commit empty shape so box visibly locks in first
            tok1 = self.show.enqueue_job("commit_shape_empty", a)
            yield from self.wait_for_token(tok1)
            self.overlay = None
            self.show.ripples.append((self.show.clock, ((x + r) * 0.5, (y + b) * 0.5), self.color))
            # Beat 2: cursor hops inside and visibly types text character-by-character!
            yield from self._live_type_sequence((x, y, r, b), a, "set_text")
        else:
            tok = self.show.enqueue_job("commit", a)
            yield from self.wait_for_token(tok)
            self.overlay = None
            self.tool = "pointer"
            self.idle_anchor = self.pos
            self.show.ripples.append((self.show.clock, ((x + r) * 0.5, (y + b) * 0.5), self.color))
            yield from self.wait(self.rng.uniform(0.02, 0.045))

    def do_image(self, a):
        """Futuristic corner-bracket + laser scan reveal for atomic transparent illustrations."""
        x, y, r, b = bounds(a)
        w, h = r - x, b - y
        yield from self.travel((x, y))
        self.tool = "pen"
        dur = clamp(0.36 + (w + h) / (self.show.sw * 2.2), 0.36, 0.60)

        def scan(t):
            scan_y = y + h * t
            wave_x = x + w * (0.15 + 0.70 * (0.5 + 0.5 * math.sin(t * math.pi * 2.0)))
            self.move_to((wave_x, scan_y))
            self.overlay = {"kind": "asset_scan", "rect": (x, y, r, b), "progress": t}

        yield from self.animate(dur, scan, smooth)
        tok = self.show.enqueue_job("commit", a)
        yield from self.wait_for_token(tok)
        self.overlay = None
        self.tool = "pointer"
        self.idle_anchor = self.pos
        self.show.ripples.append((self.show.clock, ((x + r) * 0.5, (y + b) * 0.5), self.color))
        yield from self.wait(self.rng.uniform(0.02, 0.045))

    def do_path(self, a):
        segments = list(path_segments(path_commands(a)))
        yield from self.travel(segments[0][1])
        self.tool = "pen"
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
        dur = clamp(0.18 + total_len / (self.show.sw * 1.45), 0.18, 0.46)

        def trace(t):
            target = total_len * t
            i = max(1, min(len(all_pts) - 1, bisect_left(dists, target)))
            f = (target - dists[i - 1]) / max(1e-8, dists[i] - dists[i - 1])
            tip = blend(all_pts[i - 1], all_pts[i], f)
            self.move_to(tip)
            self.overlay = {
                "kind": "vector_guide",
                "points": all_pts[:i] + [tip],
                "color": ink_col,
                "width": ink_w,
                "dash": a.get("line_dash"),
            }

        yield from self.animate(dur, trace, ease_out_cubic)
        tok = self.show.enqueue_job("commit", a)
        yield from self.wait_for_token(tok)
        self.overlay = None
        self.tool = "pointer"
        self.idle_anchor = self.pos
        self.show.ripples.append((self.show.clock, all_pts[-1], self.color))
        yield from self.wait(self.rng.uniform(0.02, 0.045))

    def do_textbox(self, a):
        rect = bounds(a)
        yield from self._live_type_sequence(rect, a, "commit")

    def do_duplicate(self, a):
        x, y, r, b = bounds(a)
        w, h = r - x, b - y
        sb = self.show.boxes.get(a["source"], (x, y, r, b))
        sx, sy = sb[0], sb[1]
        yield from self.travel((sx + w * 0.5, sy + h * 0.5))
        self.tool = "cross"
        dur = clamp(0.22 + math.dist((sx, sy), (x, y)) / (self.show.sw * 1.4), 0.22, 0.45)

        def drag(t):
            cx, cy = blend((sx, sy), (x, y), t)
            self.move_to((cx + w * 0.5, cy + h * 0.5))
            self.overlay = {"kind": "cad_box", "rect": (cx, cy, cx + w, cy + h), "shape": "rectangle", "dim": "Copy", "fill": None}

        yield from self.animate(dur, drag, ease_out_cubic)
        tok = self.show.enqueue_job("commit", a)
        yield from self.wait_for_token(tok)
        self.overlay = None
        self.tool = "pointer"
        self.idle_anchor = self.pos
        self.show.ripples.append((self.show.clock, (x + w * 0.5, y + h * 0.5), self.color))
        yield from self.wait(0.03)

    def do_group(self, a):
        x, y, r, b = self.show.boxes[a["name"]]
        yield from self.travel(((x + r) * 0.5, (y + b) * 0.5))
        tok = self.show.enqueue_job("commit", a)
        yield from self.wait_for_token(tok)
        self.overlay = None
        self.tool = "pointer"
        self.idle_anchor = self.pos

    def run(self, a):
        self.action = a
        self.used = True
        self.active = True
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
        elif kind == "duplicate_shape":
            yield from self.do_duplicate(a)
        elif kind == "group_shapes":
            yield from self.do_group(a)
        self.active = False


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
        self.pace = float(self.plan.get("pace", 1.0))
        self.count = int(self.plan.get("cursors", 6))
        self.agent_labels = bool(self.plan.get("agent_labels", True))
        self.show_trails = bool(self.plan.get("motion_trails", True))

        # Launch dedicated background COM worker thread
        self.worker = ComWorker(
            self.path,
            int(self.plan.get("slide", 1)),
            self.plan.get("background"),
            self.plan.get("canvas"),
        )
        self.worker.start()
        if not self.worker.ready_event.wait(timeout=22):
            raise TimeoutError("PowerPoint COM worker did not initialize within 22 seconds.")
        if self.worker.error:
            raise RuntimeError(self.worker.error)

        info = self.worker.init_info
        self.sw, self.sh = info["sw"], info["sh"]
        self.ppt_hwnd = info["ppt_hwnd"]
        self.win_offset = info["win_offset"]
        self.ppt_committed = set(info["existing"])
        self.enqueued_commits = set()
        self.completed_tokens = set()
        self._next_token = 1

        self.actions, self.deps, self.boxes = prepare_plan(self.plan, plan_path.parent, self.ppt_committed)
        self.z_ordered_names = [
            a["name"]
            for a in sorted(enumerate(self.actions), key=lambda pair: (float(pair[1].get("z_index", pair[0])), pair[0]))
            if a["type"] != "group_shapes"
        ]
        self.pending = list(self.actions)
        self.timeout = timeout
        self.clock = 0.0
        self.typed = self.frames = self.peak = 0
        self.timeline = []
        self.ripples = []
        self.first = self.done_at = self.hidden_at = self.error = None
        self.last = time.perf_counter()
        self.next_frame_at = self.last
        self.next_probe = 0.0
        self.viewport = None
        self.capture_started = None
        self.recording_intervals = []
        self.finish_sent = False

        # Transparent click-through topmost Tk overlay
        self.root = tk.Tk()
        self.root.protocol("WM_DELETE_WINDOW", self.root.quit)
        self.root.withdraw()
        self.root.overrideredirect(True)
        self.root.configure(bg=KEY)
        self.root.attributes("-transparentcolor", KEY)
        self.canvas = tk.Canvas(self.root, bg=KEY, bd=0, highlightthickness=0)
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
            self.recorder = Recorder(resolve(self.plan["record"]), plan_path.parent, (self.sw, self.sh))

    def enqueue_job(self, job_kind, action):
        token = self._next_token
        self._next_token += 1
        if job_kind in ("commit", "set_text"):
            self.enqueued_commits.add(action["name"])
        self.worker.job_queue.put((job_kind, action, token))
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
                self.completed_tokens.add(token)
            if tag == "ok":
                self.ppt_committed.add(payload)

    def probe_win32(self):
        """Pure Win32 foreground & viewport check (<0.01ms, ZERO PowerPoint COM calls!)."""
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
            rect = gui.GetWindowRect(self.ppt_hwnd)
            dx, dy, vw, vh = self.win_offset
            return (rect[0] + dx, rect[1] + dy, vw, vh)
        return self.worker.init_info["viewport"]

    def choose(self, agent):
        in_flight = len(self.enqueued_commits - self.ppt_committed)
        if in_flight >= max(4, self.count):
            return None

        # Collect active peer positions to prevent multiple cursors clumping in the same spot
        busy_spots = [
            ag.pos for ag in self.agents
            if ag is not agent and ag.generator is not None
        ]
        repel_r = self.sw * 0.14

        own = []
        steal = []
        for ordinal, a in enumerate(self.pending):
            if not self.deps[a["name"]] <= self.ppt_committed:
                continue
            x0, y0, x1, y1 = self.boxes[a["name"]]
            center = ((x0 + x1) * 0.5, (y0 + y1) * 0.5)
            dist = math.dist(agent.pos, center)
            crowd_penalty = sum(
                (repel_r - math.dist(center, bp)) * 1.4
                for bp in busy_spots
                if math.dist(center, bp) < repel_r
            )
            score = int(a.get("wave", 1)) * self.sw * 12.0 + dist + crowd_penalty + ordinal * self.sw * 0.04
            if int(a.get("lane", agent.index + 1)) == agent.index + 1:
                own.append((score, ordinal, a))
            else:
                steal.append((score, ordinal, a))
        if own:
            return min(own)[2]
        if steal:
            return min(steal)[2]
        return None

    def pause_recording(self):
        if self.recorder and self.capture_started is not None:
            stopped = time.perf_counter()
            self.recorder.pause()
            self.recording_intervals.append((self.capture_started, stopped))
            self.capture_started = None

    def paint(self):
        cv = self.canvas
        _, _, w, h = self.viewport
        u = clamp(w / 1280.0, 1.0, 2.3)

        def px(p):
            return p[0] * w / self.sw, p[1] * h / self.sh

        cv.delete("swarm")

        # 1. Double concentric commit ripples
        self.ripples = [rp for rp in self.ripples if self.clock - rp[0] < 0.34]
        for stamp, pos, col in self.ripples:
            age = (self.clock - stamp) / 0.34
            r1 = u * (4.0 + 20.0 * ease_out_cubic(age))
            rx, ry = px(pos)
            cv.create_oval(rx - r1, ry - r1, rx + r1, ry + r1, outline=col, width=max(1.0, 1.8 * u * (1.0 - age)), tags="swarm")
            if age > 0.16:
                r2 = u * (2.0 + 11.0 * ease_out_cubic((age - 0.16) / 0.84))
                cv.create_oval(rx - r2, ry - r2, rx + r2, ry + r2, outline="white", width=max(1.0, 1.0 * u * (1.0 - age)), tags="swarm")

        # 2. Agent overlays, short comet trails, cursors, and razor-sharp Figma pills
        for ag in self.agents:
            if not ag.used and self.done_at is None:
                continue

            if self.done_at is not None:
                exit_t = min(1.0, (time.perf_counter() - self.done_at) / 0.45)
                if exit_t >= 0.98:
                    continue
                cur_pos = blend(ag.idle_anchor, ag.home, ease_out_cubic(exit_t))
            elif not ag.active:
                drift = self.clock * 2.2 + ag.index * 1.3
                cur_pos = (ag.idle_anchor[0] + math.sin(drift) * 1.8, ag.idle_anchor[1] + math.cos(drift * 0.8) * 1.4)
            else:
                cur_pos = ag.pos

            ov = ag.overlay
            if ov:
                kind = ov["kind"]
                if kind == "cad_box":
                    x0, y0 = px(ov["rect"][:2])
                    x1, y1 = px(ov["rect"][2:])
                    fill_hex = _safe_hex(ov.get("fill"), "") if ov.get("fill") else ""
                    if ov.get("shape") in ("oval", "circle", "ellipse"):
                        if fill_hex:
                            cv.create_oval(x0, y0, x1, y1, fill=fill_hex, stipple="gray25", outline="", tags="swarm")
                        cv.create_oval(x0, y0, x1, y1, outline=ag.color, width=1.6 * u, dash=(5, 3), tags="swarm")
                    else:
                        if fill_hex:
                            cv.create_rectangle(x0, y0, x1, y1, fill=fill_hex, stipple="gray25", outline="", tags="swarm")
                        cv.create_rectangle(x0, y0, x1, y1, outline=ag.color, width=1.6 * u, dash=(5, 3), tags="swarm")
                    hs = 2.6 * u
                    for hx, hy in ((x0, y0), (x1, y0), (x0, y1), (x1, y1)):
                        cv.create_rectangle(hx - hs, hy - hs, hx + hs, hy + hs, fill="white", outline=ag.color, width=1.2 * u, tags="swarm")
                    if abs(x1 - x0) > 28 * u and abs(y1 - y0) > 18 * u:
                        dx, dy = (x0 + x1) * 0.5, y1 + 10 * u
                        dim_txt = ov["dim"]
                        tw = max(38 * u, len(dim_txt) * 5.6 * u + 8 * u)
                        cv.create_rectangle(dx - tw * 0.5, dy - 7.5 * u, dx + tw * 0.5, dy + 7.5 * u, fill=ag.color, outline="white", width=1.0 * u, tags="swarm")
                        cv.create_text(dx, dy, text=dim_txt, fill="white", font=("Segoe UI", max(7, int(7 * u)), "bold"), tags="swarm")

                elif kind == "asset_scan":
                    x0, y0 = px(ov["rect"][:2])
                    x1, y1 = px(ov["rect"][2:])
                    prog = ov["progress"]
                    cv.create_rectangle(x0, y0, x1, y1, outline=ag.tint, width=1.2 * u, dash=(3, 3), tags="swarm")
                    bl = min(16 * u, abs(x1 - x0) * 0.28, abs(y1 - y0) * 0.28)
                    bw = 2.5 * u
                    for cx_b, cy_b, sx, sy in ((x0, y0, 1, 1), (x1, y0, -1, 1), (x0, y1, 1, -1), (x1, y1, -1, -1)):
                        cv.create_line(cx_b, cy_b, cx_b + sx * bl, cy_b, fill=ag.color, width=bw, tags="swarm")
                        cv.create_line(cx_b, cy_b, cx_b, cy_b + sy * bl, fill=ag.color, width=bw, tags="swarm")
                    sy_beam = y0 + (y1 - y0) * prog
                    cv.create_line(x0 + 2, sy_beam, x1 - 2, sy_beam, fill=ag.color, width=2.4 * u, tags="swarm")
                    cv.create_line(x0 + 2, sy_beam - 2 * u, x1 - 2, sy_beam - 2 * u, fill=ag.tint, width=1.1 * u, tags="swarm")

                elif kind == "vector_guide":
                    pts = ov["points"]
                    if len(pts) >= 2:
                        flat = [z for p in pts for z in px(p)]
                        dash_pat = (round(3 * u), round(2 * u)) if ov.get("dash") and ov["dash"] != "solid" else None
                        col = _safe_hex(ov["color"], ag.color.lstrip("#"))
                        cv.create_line(*flat, fill=ag.tint, width=max(3.0, ov["width"] * w / self.sw + 2.5 * u), capstyle="round", joinstyle="round", tags="swarm")
                        cv.create_line(*flat, fill=col, width=max(1.5, ov["width"] * w / self.sw), capstyle="round", joinstyle="round", dash=dash_pat, tags="swarm")
                        tx, ty = px(pts[-1])
                        r_tip = 3.5 * u
                        cv.create_oval(tx - r_tip, ty - r_tip, tx + r_tip, ty + r_tip, fill="white", outline=ag.color, width=1.4 * u, tags="swarm")

                elif kind == "live_text":
                    rx0, ry0, rx1, ry1 = ov["rect"]
                    x0, y0 = px((rx0, ry0))
                    x1, y1 = px((rx1, ry1))
                    act = ov["action"]
                    shown_txt = ov["text"]

                    # Subtle active typing bounding guide
                    cv.create_rectangle(x0, y0, x1, y1, outline=ag.tint, width=1.1 * u, dash=(3, 3), tags="swarm")

                    # Render REAL live characters at exact DPI-independent pixel size!
                    ml = float(act.get("margin_left", 0)) * w / self.sw
                    mr = float(act.get("margin_right", 0)) * w / self.sw
                    mt = float(act.get("margin_top", 0)) * h / self.sh
                    mb = float(act.get("margin_bottom", 0)) * h / self.sh
                    ix0, iy0 = x0 + ml, y0 + mt
                    ix1, iy1 = max(ix0 + 6.0, x1 - mr), max(iy0 + 6.0, y1 - mb)

                    align = str(act.get("alignment", "left")).lower()
                    valign = str(act.get("vertical_alignment", "middle")).lower()
                    h_anc = "center" if align in ("center", "centre") else ("e" if align == "right" else "w")
                    v_anc = "n" if valign == "top" else ("s" if valign == "bottom" else "")
                    if h_anc == "center":
                        anchor = v_anc if v_anc else "center"
                        tx = (ix0 + ix1) * 0.5
                        just = "center"
                    else:
                        anchor = (v_anc + h_anc) if v_anc else h_anc
                        tx = ix1 if h_anc == "e" else ix0
                        just = "right" if h_anc == "e" else "left"
                    ty = iy0 if valign == "top" else (iy1 if valign == "bottom" else (iy0 + iy1) * 0.5)

                    fs_pt = float(act.get("font_size", 14.0))
                    # Negative Tk font size = exact pixel height (immune to Windows DPI scaling!)
                    fs_px = max(9, int(round(fs_pt * h / self.sh)))
                    fam = (
                        act.get("font_name_east_asian") or "Microsoft YaHei"
                        if has_cjk(shown_txt)
                        else (act.get("font_name") or "Arial")
                    )
                    style_tokens = []
                    if act.get("bold"):
                        style_tokens.append("bold")
                    if act.get("italic"):
                        style_tokens.append("italic")
                    tk_font = (fam, -fs_px, " ".join(style_tokens)) if style_tokens else (fam, -fs_px)
                    tcol = _safe_hex(act.get("font_color") or act.get("color"), "161616")
                    wrap_px = max(20, int(round(ix1 - ix0))) if act.get("word_wrap", True) else 0

                    if shown_txt:
                        cv.create_text(
                            tx, ty,
                            text=shown_txt,
                            fill=tcol,
                            anchor=anchor,
                            justify=just,
                            width=wrap_px,
                            font=tk_font,
                            tags="swarm",
                        )

                    # Crisp blinking live typing caret right at the tip of the active character
                    cx_pt, cy_pt, lh_pt = ov["caret"]
                    cpx, cpy = px((cx_pt, cy_pt))
                    ch_half = max(5.0, lh_pt * h / self.sh * 0.44)
                    if ov.get("blink", True):
                        cv.create_line(cpx, cpy - ch_half, cpx, cpy + ch_half, fill=ag.color, width=max(1.6, 1.8 * u), tags="swarm")

            # Short, luxury comet tail (strictly capped at <=42px arc length so NO stray lines slash across slide!)
            if self.show_trails and ag.active and len(ag.trail) >= 2:
                tpts = [px(p) for p in ag.trail]
                max_seg_px = 38.0 * u
                max_total_px = 44.0 * u
                acc_len = 0.0
                segs = []
                for k in range(len(tpts) - 1, 0, -1):
                    p0, p1 = tpts[k - 1], tpts[k]
                    seg_d = math.hypot(p1[0] - p0[0], p1[1] - p0[1])
                    if seg_d > max_seg_px or seg_d < 0.5:
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
                    cv.create_line(p0[0], p0[1], p1[0], p1[1], fill=col_t, width=max(1.0, 2.2 * u * frac), capstyle="round", tags="swarm")

            # Crisp vector cursor
            x, y = px(cur_pos)

            def coords(pts_rel):
                return [v for dx, dy in pts_rel for v in (x + dx * u, y + dy * u)]

            col = ag.color
            if ag.tool == "cross":
                for p, q in (((-7, 0), (7, 0)), ((0, -7), (0, 7))):
                    cv.create_line(*coords([p, q]), fill="white", width=3.2 * u, tags="swarm")
                    cv.create_line(*coords([p, q]), fill=col, width=1.5 * u, tags="swarm")
                cv.create_oval(x - 2.2 * u, y - 2.2 * u, x + 2.2 * u, y + 2.2 * u, fill="white", outline=col, width=1.2 * u, tags="swarm")
            elif ag.tool == "pen":
                cv.create_polygon(
                    *coords([(0, 0), (2.0, -7.2), (13.8, -19.0), (17.6, -15.2), (5.8, -3.4)]),
                    fill=col, outline="white", width=1.4 * u, tags="swarm",
                )
                cv.create_polygon(*coords([(0, 0), (2.2, -5.8), (5.2, -2.8)]), fill="white", outline=col, width=0.9 * u, tags="swarm")
            elif ag.tool == "ibeam":
                for p, q in (((0, -6.5), (0, 6.5)), ((-2.8, -6.5), (2.8, -6.5)), ((-2.8, 6.5), (2.8, 6.5))):
                    cv.create_line(*coords([p, q]), fill="white", width=3.0 * u, tags="swarm")
                    cv.create_line(*coords([p, q]), fill=col, width=1.4 * u, tags="swarm")
            else:
                cv.create_polygon(
                    *coords([(0, 0), (0, 16.5), (4.6, 12.6), (8.4, 19.6), (11.0, 18.2), (7.4, 11.4), (13.8, 11.4)]),
                    fill=col, outline="white", width=1.4 * u, tags="swarm",
                )

            # Razor-sharp Figma capsule pill (solid white outer silhouette -> solid colored inner pill -> clean text)
            if self.agent_labels and self.count > 1:
                badge_text = ag.name
                lx = x + 13 * u
                ly = y + 13 * u
                font_px = max(8, int(round(7.8 * u)))
                pill_h = 16.0 * u
                pill_w = max(46 * u, (len(badge_text) * font_px * 0.72) + 18 * u)
                if lx + pill_w > w - 6:
                    lx = max(4, x - pill_w - 10 * u)
                if ly + pill_h > h - 6:
                    ly = max(4, y - pill_h - 10 * u)

                b = 1.4 * u
                oh = pill_h + 2 * b
                ox0, oy0, ox1, oy1 = lx - b, ly - b, lx + pill_w + b, ly + pill_h + b
                orad = oh * 0.5
                cv.create_oval(ox0, oy0, ox0 + oh, oy1, fill="white", outline="", tags="swarm")
                cv.create_oval(ox1 - oh, oy0, ox1, oy1, fill="white", outline="", tags="swarm")
                cv.create_rectangle(ox0 + orad, oy0, ox1 - orad, oy1, fill="white", outline="", tags="swarm")

                irad = pill_h * 0.5
                cv.create_oval(lx, ly, lx + pill_h, ly + pill_h, fill=col, outline="", tags="swarm")
                cv.create_oval(lx + pill_w - pill_h, ly, lx + pill_w, ly + pill_h, fill=col, outline="", tags="swarm")
                cv.create_rectangle(lx + irad, ly, lx + pill_w - irad, ly + pill_h, fill=col, outline="", tags="swarm")

                cv.create_text(
                    lx + pill_w * 0.5, ly + pill_h * 0.5,
                    text=badge_text, fill="white",
                    font=("Segoe UI", -font_px, "bold"), tags="swarm",
                )

    def tick(self):
        now = time.perf_counter()
        dt = min(0.035, max(0.0, now - self.last))
        self.last = now

        self.drain_done_queue()

        if now >= self.next_probe:
            view = self.probe_win32()
            self.next_probe = now + 0.18
            if view != self.viewport:
                self.pause_recording()
                self.viewport = view
                if view:
                    x, y, w, h = view
                    self.root.geometry(f"{w}x{h}")
                    self.root.update_idletasks()
                    gui.SetWindowPos(self.hwnd, wc.HWND_TOPMOST, x, y, w, h, wc.SWP_NOACTIVATE | wc.SWP_SHOWWINDOW)
                    if self.recorder:
                        self.recorder.start(view)
                        self.capture_started = time.perf_counter()
                else:
                    gui.ShowWindow(self.hwnd, wc.SW_HIDE)

        if not self.viewport:
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
        if now - self.first < 0.18:
            return

        self.clock += dt * self.pace

        for ag in self.agents:
            if ag.generator is None:
                a = self.choose(ag)
                if a is None:
                    continue
                self.pending.remove(a)
                ag.generator = ag.run(a)
            try:
                next(ag.generator)
            except StopIteration:
                self.timeline.append({
                    "name": ag.action["name"],
                    "type": ag.action["type"],
                    "agent": ag.name,
                    "hand": ag.index + 1,
                    "wave": int(ag.action.get("wave", 1)),
                    "wall_start": round(ag.wall_start, 3),
                    "wall_end": round(now - self.first, 3),
                })
                if len(self.timeline) % 10 == 0 or len(self.timeline) == len(self.actions):
                    print(f"COMMITTED {len(self.timeline)}/{len(self.actions)}", flush=True)
                ag.generator = None

        self.peak = max(self.peak, sum(ag.generator is not None for ag in self.agents))

        all_dispatched = (not self.pending) and (not any(ag.generator for ag in self.agents))
        all_ppt_done = all_dispatched and (self.enqueued_commits <= self.ppt_committed)

        if not any(ag.generator for ag in self.agents) and self.pending:
            if self.enqueued_commits <= self.ppt_committed and not any(self.choose(ag) for ag in self.agents):
                raise RuntimeError("Deadlock: remaining actions cannot satisfy dependencies.")

        if all_ppt_done:
            if not self.finish_sent:
                self.finish_sent = True
                self.done_at = now
                self.worker.job_queue.put(("finish", self.output, self.export_preview, self.z_ordered_names, self.boxes))
            if self.worker.finished_event.is_set() and (now - self.done_at >= 0.65):
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
        if self.next_frame_at < now - 0.05:
            self.next_frame_at = now
        delay_ms = max(1, int(round((self.next_frame_at - now) * 1000.0)))
        self.root.after(delay_ms, self.loop)

    def run(self):
        try:
            self.next_frame_at = time.perf_counter()
            self.root.after(0, self.loop)
            self.root.mainloop()
        finally:
            if not self.finish_sent:
                self.worker.job_queue.put(("abort",))
            self.worker.join(timeout=15)
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

        fin = self.worker.finish_info or {"output": str(self.output), "preview_png": None}
        report = {
            "confirmed": len(self.timeline),
            "peak_cursors": self.peak,
            "displayed_frames": self.frames,
            "typed_characters": self.typed,
            "drawing_seconds": round(self.last - self.first, 2),
            "output": fin["output"],
            "preview_png": fin["preview_png"],
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
    parser.add_argument("--check", action="store_true", help="Validate plan without connecting to PowerPoint.")
    args = parser.parse_args()

    if sys.platform == "win32":
        ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)
    plan_path = args.plan.resolve(strict=True)
    if args.check:
        plan = json.loads(plan_path.read_text(encoding="utf-8-sig"))
        actions, _, _ = prepare_plan(plan, plan_path.parent)
        by_wave, by_lane = {}, {}
        for a in actions:
            by_wave[a["wave"]] = by_wave.get(a["wave"], 0) + 1
            by_lane[a["lane"]] = by_lane.get(a["lane"], 0) + 1
        print(json.dumps({
            "valid": True,
            "actions": len(actions),
            "cursors": int(plan.get("cursors", 6)),
            "pace": float(plan.get("pace", 1.0)),
            "actions_by_wave": dict(sorted(by_wave.items())),
            "actions_by_lane": dict(sorted(by_lane.items())),
        }, ensure_ascii=False, indent=2))
    else:
        SwarmShow(plan_path, args.timeout, export_preview=args.export_preview).run()


if __name__ == "__main__":
    main()
