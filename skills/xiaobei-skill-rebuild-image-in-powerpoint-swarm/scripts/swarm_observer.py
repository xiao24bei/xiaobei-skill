"""High-end passive multi-agent cursor companion overlay for powerpoint_draw_sequence batches.

Attaches read-only to the active PowerPoint window and renders 6+ studio-colored
Figma-style cursors with motion trails, agent role pills, CAD framing boxes, and
commit ripples synchronized to named shapes arriving on the slide.
"""
import argparse
from collections import deque
import ctypes
import json
import math
from pathlib import Path
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
    import pywintypes
    import win32con as wc
    import win32gui as gui
    from win32com.client import GetActiveObject
    HAS_WIN32 = True
except ImportError:
    pythoncom = pywintypes = wc = gui = GetActiveObject = None
    HAS_WIN32 = False

KEY = "#010203"
AGENTS = [
    ("Layout", "#2563EB", "#93C5FD"),
    ("Visual-1", "#7C3AED", "#C4B5FD"),
    ("Visual-2", "#059669", "#6EE7B7"),
    ("Connector", "#EA580C", "#FDBA74"),
    ("Label-1", "#DB2777", "#F9A8D4"),
    ("Label-2", "#0891B2", "#67E8F9"),
]


def read_plan(path):
    plan = json.loads(path.read_text(encoding="utf-8-sig"))
    deck = (path.parent / plan["deck"]).resolve(strict=True)
    slide = int(plan.get("slide", 1))
    cursors = int(plan.get("cursors", 6))
    cues = plan.get("cues") or []
    if not cues:
        raise ValueError("Observer plan requires non-empty `cues`.")
    return deck, slide, max(1, min(24, cursors)), cues


def observer_thread(deck, slide_idx, cues, state, stop):
    pythoncom.CoInitialize()
    try:
        app = GetActiveObject("PowerPoint.Application")
        pres = app.ActivePresentation
        if Path(pres.FullName).resolve() != deck:
            raise ValueError(f"Activate presentation first: {deck}")
        slide = pres.Slides.Item(slide_idx)
        shapes = slide.Shapes
        width, height = float(pres.PageSetup.SlideWidth), float(pres.PageSetup.SlideHeight)
        names_done = set()
        while not stop.is_set():
            try:
                win = app.ActiveWindow
                if Path(app.ActivePresentation.FullName).resolve() != deck or win.View.Slide.SlideIndex != slide_idx:
                    state["visible"] = False
                else:
                    x, y = win.PointsToScreenPixelsX(0), win.PointsToScreenPixelsY(0)
                    r = win.PointsToScreenPixelsX(round(width))
                    b = win.PointsToScreenPixelsY(round(height))
                    state["viewport"] = (x, y, r - x, b - y, width, height)
                    state["visible"] = True
                present = {shapes.Item(i).Name for i in range(1, shapes.Count + 1)}
                for cue in cues:
                    if cue["name"] in present:
                        names_done.add(cue["name"])
                state["done_names"] = set(names_done)
                state["done"] = len(names_done)
                if len(names_done) >= len(cues):
                    return
            except pywintypes.com_error:
                state["visible"] = False
            stop.wait(0.08)
    except Exception:
        state["error"] = traceback.format_exc()
    finally:
        pythoncom.CoUninitialize()


def ease(t):
    t = max(0.0, min(1.0, t))
    return t * t * (3 - 2 * t)


def walk(points, progress):
    lengths = [math.dist(a, b) for a, b in zip(points, points[1:])]
    rem = sum(lengths) * max(0.0, min(1.0, progress))
    trace = [points[0]]
    for a, b, ln in zip(points, points[1:], lengths):
        if rem <= ln and ln:
            t = rem / ln
            trace.append((a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t))
            break
        trace.append(b)
        rem -= ln
    return trace


class CompanionOverlay:
    def __init__(self, plan_path, timeout):
        if not HAS_WIN32:
            raise RuntimeError("Windows with pywin32 is required.")
        self.deck, self.slide_idx, self.count, self.cues = read_plan(plan_path)
        self.state = {"done": 0, "done_names": set(), "visible": False}
        self.stop = threading.Event()
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.overrideredirect(True)
        self.root.configure(bg=KEY)
        self.root.attributes("-transparentcolor", KEY)
        self.canvas = tk.Canvas(self.root, bg=KEY, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.root.update_idletasks()
        self.hwnd = gui.GetParent(self.root.winfo_id()) or self.root.winfo_id()
        style = gui.GetWindowLong(self.hwnd, wc.GWL_EXSTYLE)
        gui.SetWindowLong(
            self.hwnd, wc.GWL_EXSTYLE,
            style | wc.WS_EX_LAYERED | wc.WS_EX_TRANSPARENT | wc.WS_EX_TOOLWINDOW | 0x08000000,
        )
        self.root.geometry("1x1+0+0")
        self.root.deiconify()
        self.root.update_idletasks()
        gui.ShowWindow(self.hwnd, wc.SW_HIDE)
        self.lanes = {}
        self.origins = {}
        self.trails = {i: deque(maxlen=10) for i in range(self.count)}
        self.ripples = []
        self.timeout = timeout
        self.last_done = 0
        self.last_change = time.monotonic()
        self.mapping = None
        self.release = None
        self.ready = False

    def update(self):
        now = time.monotonic()
        done = self.state["done"]
        if done != self.last_done:
            self.last_done = done
            self.last_change = now
        if now - self.last_change > self.timeout:
            raise TimeoutError(f"Observer timed out after {done}/{len(self.cues)} objects.")
        if self.state.get("error"):
            self.root.quit()
            return
        fg = gui.GetForegroundWindow()
        on_ppt = bool(fg and gui.GetClassName(fg) == "PPTFrameClass")
        if self.state["visible"] and not self.ready:
            self.ready = True
            print("READY", flush=True)
        if not self.state["visible"] or not on_ppt:
            gui.ShowWindow(self.hwnd, wc.SW_HIDE)
        else:
            self.paint(now)
        if done >= len(self.cues):
            if self.release is None:
                self.release = now
            if now - self.release > 0.65:
                self.root.quit()
                return
        self.root.after(16, self.frame)

    def frame(self):
        try:
            self.update()
        except Exception:
            self.state["error"] = traceback.format_exc()
            self.root.quit()

    def paint(self, now):
        x, y, w, h, sw, sh = self.state["viewport"]
        if min(w, h) < 10:
            return
        if self.mapping != (x, y, w, h):
            self.root.geometry(f"{w}x{h}")
            self.root.update_idletasks()
            self.mapping = (x, y, w, h)
        gui.SetWindowPos(self.hwnd, wc.HWND_TOPMOST, x, y, w, h, wc.SWP_NOACTIVATE | wc.SWP_SHOWWINDOW)
        cv = self.canvas
        cv.delete("all")
        u = max(1.0, min(2.3, w / 1280.0))
        done_names = self.state["done_names"]

        busy = set()
        for idx, (lane, started, origin) in list(self.lanes.items()):
            cue = self.cues[idx]
            if cue["name"] in done_names:
                end_pt = cue["points"][-1]
                self.origins[lane] = end_pt
                self.ripples.append((now, lane, end_pt))
                del self.lanes[idx]
            else:
                busy.add(lane)

        for idx, cue in enumerate(self.cues):
            if cue["name"] in done_names or idx in self.lanes:
                continue
            if not set(cue.get("after", [])) <= done_names:
                continue
            lane = next((i for i in range(self.count) if i not in busy), None)
            if lane is None:
                break
            busy.add(lane)
            self.lanes[idx] = (lane, now, self.origins.get(lane, (sw * 0.12, sh * 0.08)))

        self.ripples = [rp for rp in self.ripples if now - rp[0] < 0.35]
        for stamp, lane, (ax, ay) in self.ripples:
            age = (now - stamp) / 0.35
            _, col, _ = AGENTS[lane % len(AGENTS)]
            r = u * (4 + 24 * age)
            px, py = ax * w / sw, ay * h / sh
            cv.create_oval(px - r, py - r, px + r, py + r, outline=col, width=max(1.0, 1.8 * u * (1 - age)))

        for idx, (lane, started, origin) in self.lanes.items():
            cue = self.cues[idx]
            name, col, tint = AGENTS[lane % len(AGENTS)]
            pts = [(a * w / sw, b * h / sh) for a, b in cue["points"]]
            orig = (origin[0] * w / sw, origin[1] * h / sh)
            t = max(0.0, now - started)
            if t < 0.28:
                tr = walk([orig, pts[0]], ease(t / 0.28))
            else:
                tr = walk(pts, ease((t - 0.28) / 0.42))
                ax, ay = pts[0]
                bx, by = tr[-1]
                kind = cue.get("kind", "box")
                if kind == "line":
                    cv.create_line(*[v for p in tr for v in p], fill=col, width=2.0 * u, capstyle="round")
                elif kind == "text":
                    cv.create_line(bx, by - 6 * u, bx, by + 6 * u, fill=col, width=1.6 * u)
                else:
                    cv.create_rectangle(min(ax, bx), min(ay, by), max(ax, bx), max(ay, by), outline=col, width=1.5 * u, dash=(5, 3))
            px, py = tr[-1]
            trail = self.trails[lane]
            trail.append((px, py))
            tlist = list(trail)
            for k in range(len(tlist) - 1):
                frac = (k + 1) / len(tlist)
                cv.create_line(tlist[k][0], tlist[k][1], tlist[k + 1][0], tlist[k + 1][1],
                               fill=col if frac > 0.7 else tint, width=max(1.0, 2.0 * u * frac), capstyle="round")
            arrow = [(0, 0), (0, 16.5), (4.6, 12.6), (8.4, 19.6), (11.0, 18.2), (7.4, 11.4), (13.8, 11.4)]
            cv.create_polygon(*[v for dx, dy in arrow for v in (px + dx * u, py + dy * u)],
                              fill=col, outline="white", width=1.3 * u)
            lx, ly = px + 13 * u, py + 13 * u
            font_px = max(8, int(round(7.8 * u)))
            ph = 16.0 * u
            pw = max(46 * u, (len(name) * font_px * 0.72) + 18 * u)
            if lx + pw > w - 6:
                lx = max(4, px - pw - 10 * u)
            if ly + ph > h - 6:
                ly = max(4, py - ph - 10 * u)
            b = 1.4 * u
            oh = ph + 2 * b
            ox0, oy0, ox1, oy1 = lx - b, ly - b, lx + pw + b, ly + ph + b
            orad = oh * 0.5
            cv.create_oval(ox0, oy0, ox0 + oh, oy1, fill="white", outline="")
            cv.create_oval(ox1 - oh, oy0, ox1, oy1, fill="white", outline="")
            cv.create_rectangle(ox0 + orad, oy0, ox1 - orad, oy1, fill="white", outline="")
            irad = ph * 0.5
            cv.create_oval(lx, ly, lx + ph, ly + ph, fill=col, outline="")
            cv.create_oval(lx + pw - ph, ly, lx + pw, ly + ph, fill=col, outline="")
            cv.create_rectangle(lx + irad, ly, lx + pw - irad, ly + ph, fill=col, outline="")
            cv.create_text(lx + pw * 0.5, ly + ph * 0.5, text=name, fill="white", font=("Segoe UI", font_px, "bold"))

    def run(self):
        threading.Thread(target=observer_thread, args=(self.deck, self.slide_idx, self.cues, self.state, self.stop), daemon=True).start()
        try:
            self.root.after(0, self.frame)
            self.root.mainloop()
        finally:
            self.stop.set()
            self.root.destroy()
        if self.state.get("error"):
            raise RuntimeError(self.state["error"])
        print(json.dumps({"confirmed": self.last_done, "total": len(self.cues)}), flush=True)


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("plan", type=Path)
    p.add_argument("--timeout", type=int, default=120)
    args = p.parse_args()
    if sys.platform == "win32":
        ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)
    CompanionOverlay(args.plan.resolve(strict=True), args.timeout).run()
