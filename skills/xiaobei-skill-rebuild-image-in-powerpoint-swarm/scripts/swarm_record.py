"""High-FPS hardware-accelerated FFmpeg screen recorder (DXGI ddagrab 60fps with gdigrab fallback)."""
import ctypes
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time


def _get_primary_screen_size():
    if sys.platform != "win32":
        return None
    try:
        u32 = ctypes.windll.user32
        w = int(u32.GetSystemMetrics(0))
        h = int(u32.GetSystemMetrics(1))
        if w > 0 and h > 0:
            return (w, h)
    except Exception:
        pass
    return None


class Recorder:
    def __init__(
        self,
        output,
        work,
        frame_size,
        overwrite=True,
        max_width=1920,
        max_height=1080,
        fps=60,
        prefer_ddagrab=True,
    ):
        self.output = Path(output)
        self.ffmpeg = shutil.which("ffmpeg")
        if not self.ffmpeg:
            raise FileNotFoundError("Recording requires ffmpeg (with Windows ddagrab or gdigrab).")
        if self.output.suffix.lower() != ".mp4":
            raise ValueError(f"Recording path must end in .mp4: {output}")
        if self.output.exists() and not overwrite:
            raise ValueError(f"Choose a new .mp4 path or enable overwrite_record: {output}")
        self.output.parent.mkdir(parents=True, exist_ok=True)
        self.folder = Path(tempfile.mkdtemp(prefix="swarm-rec-", dir=work))
        width, height = max(2, float(frame_size[0])), max(2, float(frame_size[1]))
        ratio = min(max_width / width, max_height / height, 1.0 if width >= max_width else max_width / width)
        self.size = tuple(max(2, 2 * round(v * ratio / 2)) for v in (width, height))
        self.fps = max(24, min(60, int(fps)))
        self.has_ddagrab = bool(prefer_ddagrab) and self._probe_ddagrab()
        self.backend = "ddagrab" if self.has_ddagrab else "gdigrab"
        self.files = []
        self.process = self.log = self.capture_rect = None

    def _probe_ddagrab(self):
        try:
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            res = subprocess.run(
                [self.ffmpeg, "-hide_banner", "-filters"],
                capture_output=True,
                text=True,
                timeout=5,
                creationflags=creationflags,
            )
            return "ddagrab" in (res.stdout or "")
        except Exception:
            return False

    def _build_ddagrab_cmd(self, x, y, width, height, path):
        fps = self.fps
        filt = (
            f"ddagrab=output_idx=0:framerate={fps}:draw_mouse=0:"
            f"offset_x={x}:offset_y={y}:video_size={width}x{height},"
            f"hwdownload,format=bgra,"
            f"scale={self.size[0]}:{self.size[1]}:flags=lanczos,setsar=1"
        )
        return [
            self.ffmpeg, "-nostats", "-hide_banner", "-loglevel", "warning",
            "-filter_complex", filt, "-an",
            "-c:v", "libx264", "-threads", "4", "-preset", "veryfast", "-crf", "15",
            "-pix_fmt", "yuv420p", "-r", str(fps), str(path),
        ]

    def _build_gdigrab_cmd(self, x, y, width, height, path):
        fps = min(30, self.fps)
        return [
            self.ffmpeg, "-nostats", "-hide_banner", "-loglevel", "warning",
            "-f", "gdigrab", "-framerate", str(fps), "-draw_mouse", "0",
            "-offset_x", str(x), "-offset_y", str(y),
            "-video_size", f"{width}x{height}", "-i", "desktop", "-an",
            "-vf", f"scale={self.size[0]}:{self.size[1]}:flags=lanczos,setsar=1",
            "-c:v", "libx264", "-threads", "4", "-preset", "ultrafast", "-crf", "15",
            "-pix_fmt", "yuv420p", "-r", str(fps), str(path),
        ]

    def _spawn(self, cmd, path):
        log = path.with_suffix(".log").open("wb")
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            stderr=log,
            creationflags=creationflags,
        )
        return proc, log

    def start(self, capture_rect):
        self.check()
        x, y, width, height = (int(round(v)) for v in capture_rect)
        if x < 0:
            width += x
            x = 0
        if y < 0:
            height += y
            y = 0

        # Clamp capture region cleanly inside primary monitor bounds for DXGI ddagrab safety
        mon = _get_primary_screen_size()
        on_primary = True
        if mon is not None:
            mon_w, mon_h = mon
            if x >= mon_w or y >= mon_h:
                on_primary = False
            else:
                width = min(width, mon_w - x)
                height = min(height, mon_h - y)

        width = max(32, width - (width % 2))
        height = max(32, height - (height % 2))
        norm = (x, y, width, height)
        if self.capture_rect == norm:
            return
        self.pause()

        path = self.folder / f"part-{len(self.files):03d}.mkv"
        use_dda = self.has_ddagrab and on_primary

        if use_dda:
            proc, log = self._spawn(self._build_ddagrab_cmd(x, y, width, height, path), path)
            # Brief 0.22s probe: if DXGI Desktop Duplication fails (RDP/hybrid GPU), fall back cleanly to gdigrab
            time.sleep(0.22)
            if proc.poll() is not None:
                try:
                    proc.stdin.close()
                except Exception:
                    pass
                log.close()
                self.has_ddagrab = False
                self.backend = "gdigrab"
                use_dda = False
            else:
                self.backend = "ddagrab"
                self.process = proc
                self.log = log

        if not use_dda:
            self.backend = "gdigrab"
            self.process, self.log = self._spawn(self._build_gdigrab_cmd(x, y, width, height, path), path)

        self.files.append(path)
        self.capture_rect = norm

    def check(self):
        if self.process and self.process.poll() is not None:
            raise RuntimeError(f"Screen recorder ({self.backend}) exited {self.process.returncode}; see {self.log.name}")

    def pause(self):
        if not self.process:
            return
        proc, self.process = self.process, None
        self.capture_rect = None
        try:
            if proc.poll() is None:
                try:
                    proc.stdin.write(b"q\n")
                    proc.stdin.flush()
                except OSError:
                    pass
            if proc.wait(timeout=15):
                raise RuntimeError(f"Screen recording ({self.backend}) failed; see {self.log.name}")
        finally:
            if proc.poll() is None:
                proc.kill()
                proc.wait()
            try:
                proc.stdin.close()
            except Exception:
                pass
            if self.log:
                self.log.close()

    def finish(self):
        self.pause()
        if not self.files:
            raise RuntimeError("No video segments were captured.")
        listing = self.folder / "segments.txt"
        listing.write_text("".join(f"file '{p.name}'\n" for p in self.files), encoding="ascii")
        temp = self.folder / "joined.mp4"
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.run(
            [self.ffmpeg, "-hide_banner", "-loglevel", "error", "-y", "-f", "concat", "-safe",
             "1", "-i", str(listing), "-c", "copy", "-movflags", "+faststart", str(temp)],
            check=True, creationflags=creationflags,
        )
        if temp.drive.casefold() == self.output.drive.casefold():
            os.replace(temp, self.output)
        else:
            with temp.open("rb") as src, self.output.open("wb") as dst:
                shutil.copyfileobj(src, dst)
        print(f"RECORDING {self.output} ({self.backend}@{self.fps if self.backend == 'ddagrab' else min(30, self.fps)}fps)", flush=True)
