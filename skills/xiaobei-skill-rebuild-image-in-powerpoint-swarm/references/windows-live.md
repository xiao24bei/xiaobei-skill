# Windows live drawing (with Multi-Agent Swarm Cursors)

This guide is used when a visible native Windows PowerPoint session is available.

## Default live route: Multi-Agent Swarm Cursor Engine

To deliver BOTH `reference_lock` restoration fidelity AND a sleek, fast, high-end multi-cursor collaborative drawing experience (6 simultaneous Figma-style agent cursors gliding with motion trails, role pills, CAD framing boxes, and laser-scan asset reveals):

1. Finish the measured `scene-map.json` and pass `scripts/validate_scene_map.ps1 -Phase planning`.
2. If `mode: hybrid` is active, run `scripts/extract_visual_assets.ps1` and inspect the raw-versus-processed contact sheet at 100% before drawing.
3. Compile the measured scene-map objects and extracted atomic assets into `<工作目录>/swarm-plan.json` (see [references/swarm-cursors.md](swarm-cursors.md)). Use `cursors: 6` (or the user's requested count) and `pace: 1.0` (raise to `1.5` for 100+ object slides) so a dense mechanism figure finishes smoothly in ~10–22 seconds with multiple agents visibly active at all times.
4. Run the deterministic preflight check:
   ```powershell
   & '<skill目录>/scripts/start-swarm-draw.ps1' -Plan '<工作目录>/swarm-plan.json' -CheckOnly
   ```
5. Execute the live visible multi-cursor show (use `-CreateBlank` to create and maximize a matched-canvas blank deck automatically, or attach to an already open presentation):
   ```powershell
   & '<skill目录>/scripts/start-swarm-draw.ps1' -Plan '<工作目录>/swarm-plan.json' -CreateBlank -Wait -ExportPreview '<工作目录>/preview.png'
   ```
   - `swarm_live_draw.py` runs inside one pinned PowerPoint COM session, writes real native editable shapes/text/connectors/pictures atomically (zero COM reconnect churn, zero slow per-character `BoundLeft` lag), syncs exact `z_index`, saves the editable `.pptx`, exports `preview.png`, and **leaves PowerPoint visibly open**.
   - If the user also wants a recorded video of the multi-cursor process, set `"record": "drawing-process.mp4"` in `swarm-plan.json`.

## Targeted repairs & Scientific Illustrator bridge fallback

- After the initial multi-cursor pass, inspect `preview.png` side-by-side with the reference under [references/self-correction.md](self-correction.md).
- For small targeted corrections (tweaking 1–5 named `VSS_...` shapes, rerouting a connector, fixing text fit, or inserting native PowerPoint charts), apply minimal in-place updates via the live PowerPoint COM tools / `powerpoint_draw_sequence` without redrawing unrelated objects.
- If Python/`pywin32` is unavailable on the host, fall back cleanly to the Scientific Illustrator `powerpoint_draw_sequence` route:
  1. Run `scripts/validate_live_powerpoint_sequence.ps1` and confirm `valid: true`.
  2. Open/activate the visible deck via `powerpoint_launch` (with task-persistent COM keeper).
  3. Submit observable region batches via `powerpoint_draw_sequence` (120–260 ms step delay). When Python/`pywin32` is present alongside `powerpoint_draw_sequence`, you may also launch `scripts/start-swarm-draw.ps1` with a `cues` plan (`swarm_observer.py`) to keep the multi-agent cursors visible during sequence batches.

## Native-versus-raster rule

For every inserted picture, supply an absolute/relative local path, alt text, a reason that the texture cannot be reconstructed economically, and a note describing what surrounding content was rebuilt natively. Keep the crop tight and preserve its aspect ratio. Do not add the original full-page image as a background or hidden fidelity layer.

## Finish without hiding the process

Apply the final-save integrity gate from `self-correction.md`: save after the last edit, verify the saved state, export a fresh preview from that state, and run `scripts/validate_scene_map.ps1 -Phase final`. Leave the visible PowerPoint window open. Do not quit or close PowerPoint at the end.
