# Windows live drawing (Silky 60fps Codex Swarm in PowerPoint)

This guide is used when a visible native Windows PowerPoint session is available.

## Default live route: Codex Swarm 60fps Engine

To deliver BOTH `reference_lock` restoration fidelity AND a silky-smooth, zero-ghosting, strictly synchronized "Codex operating PowerPoint" live drawing show (60fps DXGI `ddagrab` full-window capture, Win32 1ms timer, Hidden Pre-Build + Priority-0 Reveal lockstep, and spatially-zoned 6-cursor parallelism):

1. Finish the measured `scene-map.json` and pass `scripts/validate_scene_map.ps1 -Phase planning`.
2. If `mode: hybrid` is active, run `scripts/extract_visual_assets.ps1` and inspect the raw-versus-processed contact sheet at 100% before drawing.
3. Compile the measured scene-map objects and extracted atomic assets into `<工作目录>/swarm-plan.json` (see [references/swarm-cursors.md](swarm-cursors.md)):
   - Keep default `cursors: 6`, `target_duration: 24`, `record_scope: "window"`, `record_fps: 60`, `prefer_ddagrab: true`, `show_status_verb: false`.
   - Emit vertical sidebar titles (`汇\n报\n提\n纲`) as ONE `add_textbox`, symmetric pill badges (`WHY`, `边界`) as ONE `rounded_rectangle` with embedded text, slanted/irregular polygon badges (`01`–`04`, `02/15`) as ONE exact-vertex `freeform` + coupled centered `add_textbox`, and mixed-style sentences as ONE `add_textbox` with `runs` (+10–15% CJK width headroom; never `"justify"` on single-line CJK).
   - Leave card/panel internal actions without manual cross-lane overrides so automatic spatial zone ownership lets each cursor smoothly sweep its own card left → right with zero erratic cross-slide jumping.
4. Run the deterministic preflight check & quality linter:
   ```powershell
   & '<skill目录>/scripts/start-swarm-draw.ps1' -Plan '<工作目录>/swarm-plan.json' -CheckOnly
   ```
   Fix any `cjk_width_risk`, `split_inline_text_risk`, `non_rect_embedded_text_risk`, `late_title_order`, or `out_of_canvas` lint warnings before drawing.
5. Execute the live visible show (use `-CreateBlank` to create, maximize, collapse the Notes pane, and zoom-to-fit a matched-canvas blank deck automatically):
   ```powershell
   & '<skill目录>/scripts/start-swarm-draw.ps1' -Plan '<工作目录>/swarm-plan.json' -CreateBlank -Wait -ExportPreview '<工作目录>/preview.png'
   ```
   - `swarm_live_draw.py` runs inside one pinned PowerPoint COM session with Hidden Pre-Build + Priority-0 Reveal lockstep (zero double-image ghosting, zero cursor-to-object lag), CJK-safe no-squiggle typography (`LanguageID=2052` CJK / `1024` math), built-in post-draw COM geometry/overflow/badge-bounds audit, saves the editable `.pptx`, exports `preview.png`, and **leaves PowerPoint visibly open**.
   - **Clean recording hygiene**: When recording (`"record": "drawing-process.mp4"`), ensure Windows Do Not Disturb / Focus Assist is enabled so third-party desktop notifications do not intrude over PowerPoint.

## Targeted repairs & Scientific Illustrator bridge fallback

- After the initial pass, inspect BOTH the automated `audit` report in `<plan>-swarm-log.json` AND `preview.png` side-by-side with the reference under [references/self-correction.md](references/self-correction.md).
- For small targeted corrections (tweaking 1–5 named `VSS_...` shapes, rerouting a connector, fixing text fit, or inserting native PowerPoint charts), apply minimal in-place updates via live PowerPoint COM tools / `powerpoint_draw_sequence` without redrawing unrelated objects. If a clean recording (`drawing-process.mp4`) was requested and visual repairs were needed, also update `swarm-plan.json` so any re-recorded video reflects the verified layout.
- If Python/`pywin32` is unavailable on the host, fall back cleanly to the Scientific Illustrator `powerpoint_draw_sequence` route:
  1. Run `scripts/validate_live_powerpoint_sequence.ps1` and confirm `valid: true`.
  2. Open/activate the visible deck via `powerpoint_launch` (with task-persistent COM keeper).
  3. Submit observable region batches via `powerpoint_draw_sequence` (120–260 ms step delay). When Python/`pywin32` is present alongside `powerpoint_draw_sequence`, you may also launch `scripts/start-swarm-draw.ps1` with a `cues` plan (`swarm_observer.py`).

## Native-versus-raster rule

For every inserted picture, supply an absolute/relative local path, alt text, a reason that the texture cannot be reconstructed economically, and a note describing what surrounding content was rebuilt natively. Keep the crop tight and preserve its aspect ratio. Do not add the original full-page image as a background or hidden fidelity layer.

## Finish without hiding the process

Apply the final-save integrity gate from `self-correction.md`: save after the last edit, verify the saved state, export a fresh preview from that state, and run `scripts/validate_scene_map.ps1 -SceneMapPath <scene-map.json> -Phase final`. Leave the visible PowerPoint window open. Do not quit or close PowerPoint at the end.
