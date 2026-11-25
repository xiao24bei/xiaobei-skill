"""Optional FFmpeg gdigrab screen recorder for the composited slide + multi-cursor overlay."""
from pathlib import Path
import shutil
import subprocess
import tempfile


class Recorder:
    def __init__(self, output, work, slide_size):
        self.output = Path(output)
        self.ffmpeg = shutil.which("ffmpeg")
        if not self.ffmpeg:
            raise FileNotFoundError("Recording requires ffmpeg with the Windows gdigrab input.")
        if self.output.suffix.lower() != ".mp4" or self.output.exists():
            raise ValueError(f"Choose a new .mp4 path: {output}")
        self.output.parent.mkdir(parents=True, exist_ok=True)
        self.folder = Path(tempfile.mkdtemp(prefix="swarm-rec-", dir=work))
        width, height = slide_size
        ratio = min(1920 / width, 1080 / height)
        self.size = tuple(max(2, 2 * round(v * ratio / 2)) for v in (width, height))
        self.files = []
        self.process = self.log = self.viewport = None

    def start(self, viewport):
        self.check()
        if self.viewport == viewport:
            return
        self.pause()
        x, y, width, height = viewport
        path = self.folder / f"part-{len(self.files):03d}.mkv"
        self.log = path.with_suffix(".log").open("wb")
        cmd = [
            self.ffmpeg, "-nostats", "-hide_banner", "-loglevel", "warning",
            "-f", "gdigrab", "-framerate", "30", "-draw_mouse", "0",
            "-offset_x", str(x), "-offset_y", str(y),
            "-video_size", f"{width}x{height}", "-i", "desktop", "-an",
            "-vf", f"scale={self.size[0]}:{self.size[1]}:flags=bicubic,setsar=1",
            "-c:v", "libx264", "-threads", "4", "-preset", "ultrafast", "-crf", "16",
            "-pix_fmt", "yuv420p", "-r", "30", str(path),
        ]
        self.process = subprocess.Popen(
            cmd, stdin=subprocess.PIPE, stdout=subprocess.DEVNULL,
            stderr=self.log, creationflags=subprocess.CREATE_NO_WINDOW,
        )
        self.files.append(path)
        self.viewport = viewport

    def check(self):
        if self.process and self.process.poll() is not None:
            raise RuntimeError(f"Screen recorder exited {self.process.returncode}; see {self.log.name}")

    def pause(self):
        if not self.process:
            return
        proc, self.process = self.process, None
        self.viewport = None
        try:
            if proc.poll() is None:
                proc.stdin.write(b"q\n")
                proc.stdin.flush()
            if proc.wait(timeout=15):
                raise RuntimeError(f"Screen recording failed; see {self.log.name}")
        finally:
            if proc.poll() is None:
                proc.kill()
                proc.wait()
            proc.stdin.close()
            self.log.close()

    def finish(self):
        self.pause()
        if not self.files:
            raise RuntimeError("No video segments were captured.")
        listing = self.folder / "segments.txt"
        listing.write_text("".join(f"file '{p.name}'\n" for p in self.files), encoding="ascii")
        temp = self.folder / "joined.mp4"
        subprocess.run(
            [self.ffmpeg, "-hide_banner", "-loglevel", "error", "-f", "concat", "-safe",
             "1", "-i", str(listing), "-c", "copy", "-movflags", "+faststart", str(temp)],
            check=True, creationflags=subprocess.CREATE_NO_WINDOW,
        )
        if temp.drive.casefold() == self.output.drive.casefold():
            temp.rename(self.output)
        else:
            with temp.open("rb") as src, self.output.open("xb") as dst:
                shutil.copyfileobj(src, dst)
        print(f"RECORDING {self.output}", flush=True)
