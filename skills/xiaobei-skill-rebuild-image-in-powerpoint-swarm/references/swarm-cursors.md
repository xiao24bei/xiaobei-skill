# Silky 60fps Codex-in-PowerPoint Choreography (Windows Live)

Use `scripts/start-swarm-draw.ps1` (`scripts/swarm_live_draw.py`) to turn the measured `scene-map.json` + extracted atomic assets into a **buttery-smooth 60fps, calm, spatially-zoned Codex-in-PowerPoint live drawing show** directly inside visible Windows PowerPoint.

## Why v5 feels silky-smooth, zero-ghosting, and strictly synchronized

- **Hardware 60fps DXGI `ddagrab` (`lanczos` / `crf 15`) + Win32 1ms timer (`timeBeginPeriod(1)`) + tight `WS_EX_LAYERED` overlay**:
  - `swarm_record.py` uses **FFmpeg DXGI Desktop Duplication (`ddagrab` at 60fps)** directly from GPU memory (`prefer_ddagrab: true`, auto-fallback to `gdigrab` only on unsupported RDP/virtual displays), while `swarm_live_draw.py` activates 1ms Windows multimedia timer precision so the Tkinter overlay updates at true 58–60fps (instead of 31fps under Windows's default 15.6ms timer tick) over a tightly cropped slide + bottom HUD strip.
- **Hidden Pre-Build (`Visible=0`) + Priority-0 Reveal Lockstep (ZERO ghosting AND ZERO cursor lag!)**:
  - The instant a cursor claims an object (before travel/drag/type/sweep), the background STA COM worker builds the native PowerPoint object **invisibly (`Visible = 0`)** in parallel with the cursor's smooth 60fps motion (images are created off-slide at `(-2400, -2400)` and hidden before positioning so zero premature flash ever occurs).
  - The exact frame the cursor completes its drag, trace, scan, typed word, or `micro_burst` chunk, a **Priority-0 `reveal`** jumps ahead of background pre-builds, flips `Visible = -1` in **~2–8ms** (followed by an 8ms paint breather so PowerPoint flushes `WM_PAINT` within 1 frame), and clears the Tk preview **immediately (`hold_s = 0`)**.
  - Result: **ZERO double-image ghosting ("重影")** (Tk text/lines and PowerPoint native objects are NEVER stacked simultaneously), **ZERO cursor-left-before-object-appeared lag** (every shape, label, and `micro_burst` chunk appears directly under the cursor tip before the cursor moves on), and **ZERO dead-stop freezes** (elastic follow-through gently eases the final 12% of a stroke if COM is momentarily busy).
- **Background-matched chroma key (zero dark ClearType halo!)**:
  - On light slides (`#FFFFFF`, luminance ≥ 128), the transparent Tk overlay uses `#FDFDFE` near-white chroma key instead of near-black `#010203`, eliminating dark anti-aliasing halos around live-typed CJK/Latin text and cursors; on dark slides it switches automatically to `#010203`.
- **Full-swarm 6-cursor harmony from frame 1 (zero solo-cursor start, zero chaotic jumping!)**:
  - Default **`cursors: 6`** (`Codex-1`..`Codex-6`, static non-flickering pill widths with `show_status_verb: false`).
  - **Opening (Wave 0, ~1.5s)**: ALL 6 cursors launch simultaneously at `t=0.0s`, evenly sharing the sidebar, header band, cards, and dividers (with subtle red Smart Alignment Guides ONLY during Wave 0) so ZERO cursors sit parked idle.
  - **Parallel Build (Waves 1+2)**: Right after Wave 0 finishes, ALL 6 cursors seamlessly continue together — primary cursors handle header/sidebar (`汇 → 报 → 提 → 纲` top-to-bottom) and glide smoothly left → right through each card's `Badge → Heading → Diagram/Chart`, while companion cursors build the right-column `WHY / 边界` insights cleanly in parallel!
  - **Finale (Wave 4)**: Footer citation & page pill (`02/15`) → smooth exit + 1.0s clean final hold.
- **No red spellcheck squigglies, safe slanted badges & chunk-lockstep sweeps**:
  - `write_text_fast` sets `LanguageID=2052` on CJK runs and `LanguageID=1024` (`msoLanguageIDNoProofing`) on Latin/math runs (`Kb`, `Unb`, `γτ`, `kσ`) while trimming redundant default COM calls.
  - Symmetric pills (`WHY`, `边界`) embed text inside `rounded_rectangle`; slanted/irregular trapezoid badges (`01`–`04`, `02/15`) use a `freeform` polygon + coupled centered `add_textbox` (auto-bound to the same cursor via badge-text coupling; COM audit checks `BoundLeft`/`BoundTop` so no badge text can clip outside its shape!).
  - Both micro-geometry (ticks, error bars, hatches, spheres) AND numeric axis tick labels (`0.1`..`0.5`, `10¹`..`10⁻²`) sweep cleanly in 1-chunk-lookahead lockstep `micro_burst` passes with a subtle pen-tip ring (no dashed box colliding with chart axes).

## Plan format (`swarm-plan.json`)

```json
{
  "deck": "editable-figure.pptx",
  "output": "editable-figure.pptx",
  "export_preview": "preview.png",
  "record": "drawing-process.mp4",
  "record_scope": "window",
  "record_fps": 60,
  "prefer_ddagrab": true,
  "hud": true,
  "slide": 1,
  "canvas": [960.0, 540.0],
  "background": {"fill_color": "FFFFFF"},
  "cursors": 6,
  "target_duration": 24.0,
  "pace": 1.0,
  "agent_labels": true,
  "show_status_verb": false,
  "motion_trails": true,
  "micro_bursts": true,
  "auto_fix_text_overflow": true,
  "actions": [
    {
      "type": "add_shape", "name": "VSS_card_01", "shape": "rounded_rectangle",
      "left": 146, "top": 122, "width": 798, "height": 82,
      "fill_color": "FFFFFF", "line_color": "CBD5E1", "line_width": 1.0, "radius": 6,
      "stage": "structure", "z_index": 110
    },
    {
      "type": "add_textbox", "name": "VSS_main_title",
      "left": 152, "top": 10, "width": 260, "height": 36,
      "text": "汇报提纲", "font_name": "Arial", "font_name_east_asian": "Microsoft YaHei",
      "font_size": 26, "font_color": "0B3D7A", "bold": true, "word_wrap": false,
      "stage": "headers", "z_index": 510
    },
    {
      "type": "freeform", "name": "VSS_badge_01_bg",
      "path": [["M", 154, 134], ["L", 210, 134], ["L", 202, 160], ["L", 154, 160], ["Z"]],
      "fill_color": "0B4F94", "line_width": 0, "z_index": 220
    },
    {
      "type": "add_textbox", "name": "VSS_badge_01_txt",
      "left": 156, "top": 136, "width": 46, "height": 22,
      "text": "01", "font_name": "Arial", "font_size": 16, "font_color": "FFFFFF",
      "bold": true, "alignment": "center", "vertical_alignment": "middle", "word_wrap": false,
      "z_index": 520
    },
    {
      "type": "add_textbox", "name": "VSS_r3_boundary_line2",
      "left": 756, "top": 372, "width": 182, "height": 18,
      "font_name": "Arial", "font_name_east_asian": "Microsoft YaHei",
      "font_size": 10.0, "font_color": "1E293B", "word_wrap": false,
      "runs": [
        {"text": "当 "},
        {"text": "k", "font_name": "Times New Roman", "italic": true, "bold": true, "font_color": "B91C1C"},
        {"text": "*", "font_name": "Times New Roman", "superscript": true, "bold": true, "font_color": "B91C1C"},
        {"text": " 稳定且峰强显著增强时判定为分层。"}
      ],
      "z_index": 530
    }
  ]
}
```

### Supported plan options

- `record`: Optional `.mp4` output path.
- `record_scope`: `"window"` (default — full maximized PowerPoint window) or `"slide"` (bare slide canvas).
- `record_fps`: `24–60` (default `60` via DXGI `ddagrab`, caps cleanly to `30` on `gdigrab` fallback).
- `prefer_ddagrab`: Use hardware DXGI Desktop Duplication capture (default `true`; auto-falls back to `gdigrab` if unsupported).
- `hud`: Show non-overlapping Codex live status pill in the empty grey PowerPoint workspace below the slide (default `true` in `"window"` mode, `false` in `"slide"` mode).
- `cursors`: `1–24` (default `6` synchronized spatially-zoned cursors `Codex-1`..`Codex-6`; custom `cursor_profiles` supported).
- `show_status_verb`: Default `false` for calm, static-width pills without width-resizing flicker.
- `target_duration`: Target wall-clock duration in seconds (default `24.0`; set `null` or `0` to disable auto-budgeting).
- `micro_bursts`: Automatically sweep clusters of >=4 tiny primitives OR numeric axis tick labels with one fluid pen stroke (default `true`).
- `auto_fix_text_overflow`: Automatically widen safe single-line `add_textbox` shapes if PowerPoint COM `BoundWidth` slightly exceeds `Width` (default `true`).
- `overwrite_record`: Replace existing `.mp4` path cleanly on re-record (default `true`).

## Running the multi-cursor show

1. **Preflight validation & quality lint**:
   ```powershell
   & '<skill目录>/scripts/start-swarm-draw.ps1' -Plan '<工作目录>/swarm-plan.json' -CheckOnly
   ```
2. **Launch visible PowerPoint and draw live**:
   ```powershell
   & '<skill目录>/scripts/start-swarm-draw.ps1' -Plan '<工作目录>/swarm-plan.json' -CreateBlank -Wait -ExportPreview '<工作目录>/preview.png'
   ```
   Check `summary.record_backend` (`ddagrab`), `summary.avg_overlay_fps`, `summary.drawing_seconds`, and `summary.audit_hard_findings` (`0`), then inspect `preview.png` under `references/self-correction.md`.
