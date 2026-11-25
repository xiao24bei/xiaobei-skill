# Multi-Agent Swarm Cursor Choreography (Windows Live)

Use `scripts/start-swarm-draw.ps1` (`scripts/swarm_live_draw.py`) to turn the measured `scene-map.json` + extracted atomic assets into a **fast, 60fps multi-agent collaborative drawing show** directly inside visible Windows PowerPoint.

## Why this route is both fast and visually premium

- **Preserves 100% of `rebuild-image-in-powerpoint`'s restoration fidelity**:
  - Complex biological/microscopy visual cores stay crisp atomic transparent PNGs extracted by `scripts/extract_visual_assets.ps1` — never downgraded into crude hand-guessed curves or slow 200-stroke vector blobs.
  - All panels, shapes, arrows/connectors, legends, and text remain 100% native editable PowerPoint objects.
- **High-end multi-agent visual feel**:
  - Default **6 simultaneous studio-colored cursors** (`Layout`, `Visual-1`, `Visual-2`, `Connector`, `Label-1`, `Label-2`) with smooth tapered motion trails, Figma-style rounded role pills + live micro-status (`Framing` / `Rendering` / `Routing` / `Labeling`), CAD corner anchor handles + live size tags while dragging panels, futuristic corner-bracket + laser scan reveal when locking atomic biological assets, glowing vector guides for arrows, and concentric commit ripples.
- **Fast end-to-end execution (~12–25 seconds per dense slide)**:
  - Wave-parallel execution + dynamic work-stealing keeps 4–6 cursors gliding continuously across the canvas with zero dead air.
  - Every native PowerPoint shape/textbox/picture commits in **one atomic COM call** (no slow per-character `BoundLeft` loops, no per-frame shape delete thrash).
  - Exact `z_index` synchronization runs automatically after each wave and at finale.
  - Lightweight runtime: requires only Python + `pywin32` (no heavy CV/OCR/vtracer dependencies needed).

## Plan format (`swarm-plan.json`)

```json
{
  "deck": "editable-figure.pptx",
  "output": "editable-figure.pptx",
  "export_preview": "preview.png",
  "slide": 1,
  "canvas": [720.0, 540.0],
  "background": {"fill_color": "FFFFFF"},
  "cursors": 6,
  "pace": 1.0,
  "agent_labels": true,
  "motion_trails": true,
  "cursor_profiles": [
    {"name": "Layout", "color": "#2563EB"},
    {"name": "Visual-1", "color": "#7C3AED"},
    {"name": "Visual-2", "color": "#059669"},
    {"name": "Connector", "color": "#EA580C"},
    {"name": "Label-1", "color": "#DB2777"},
    {"name": "Label-2", "color": "#0891B2"}
  ],
  "actions": [
    {
      "type": "add_shape", "name": "VSS_A_panel", "shape": "rounded_rectangle",
      "left": 24, "top": 60, "width": 220, "height": 340,
      "fill_color": "F8FAFC", "line_color": "CBD5E1", "line_width": 1.2, "radius": 10,
      "wave": 1, "z_index": 10
    },
    {
      "type": "add_image", "name": "VSS_A_mouse_asset",
      "image_path": "assets/mouse_cutout.png", "visual_role": "GBM-bearing mouse",
      "left": 44, "top": 120, "width": 180, "height": 118,
      "wave": 2, "z_index": 20
    },
    {
      "type": "add_line", "name": "VSS_A_route_01",
      "begin_x": 248, "begin_y": 180, "end_x": 310, "end_y": 180,
      "line_color": "2563EB", "line_width": 2.0, "end_arrow": "triangle",
      "wave": 3, "z_index": 30
    },
    {
      "type": "add_textbox", "name": "VSS_A_title",
      "left": 36, "top": 72, "width": 196, "height": 28,
      "text": "Panel A: Mechanism", "font_name": "Arial", "font_name_east_asian": "Microsoft YaHei",
      "font_size": 14, "font_color": "0F172A", "bold": true, "alignment": "center",
      "wave": 4, "z_index": 40
    }
  ]
}
```

### Supported action types & shortcuts

- `add_shape`: 35+ presets (`rectangle`, `rounded_rectangle`, `oval`, `diamond`, `triangle`, `right_triangle`, `parallelogram`, `trapezoid`, `pentagon`, `hexagon`, `octagon`, `chevron`, `cylinder`, `cube`, `donut`, `right_arrow`, `left_arrow`, `up_arrow`, `down_arrow`, `left_right_arrow`, `up_down_arrow`, `bent_arrow`, `curved_right_arrow`, `callout_rect`, `callout_rounded`, `cloud`, `star5`, `plus`) or `shape_type_id`. Supports `fill_color`, `gradient` + `gradient_angle`, `gradient_stops`, `fill_transparency`, `line_color`, `line_width`, `line_dash`, `radius` / `adjustments`, `shadow`, `rotation`, and optional embedded `text` / `runs`.
- `add_image`: Atomic transparent PNGs or evidence tiles (`image_path` resolved relative to plan JSON). Animated with glowing corner brackets + laser scan sweep before atomic native placement.
- `add_line`: Supports explicit `begin_x, begin_y, end_x, end_y` OR `from` / `to` named object IDs (endpoints auto-computed from object boundaries). Supports `start_arrow`, `end_arrow`, `arrow_width`, `arrow_length`, `line_dash`.
- `freeform`: Custom `M/L/C/Z` vector curve/polygon with `fill_color` / `gradient` / `line_color` / `line_width` / `shadow`.
- `add_textbox`: Fast atomic styled text with `font_name`, `font_name_east_asian`, `font_size`, `font_color`, `bold`, `italic`, `underline`, `alignment`, `vertical_alignment`, `word_wrap`, `line_spacing`, and rich mixed-format `runs` (`[{"text": "H"}, {"text": "2", "subscript": true}, {"text": "O"}]`).
- `duplicate_shape`, `group_shapes`: Native copy & grouping.
- **Auto wave & lane balancing**: `wave` and `lane` are optional — if omitted, the engine automatically assigns waves (`1` shapes → `2` atomic assets → `3` arrows/lines → `4` text labels) and distributes spatially across all `cursors` with dynamic work-stealing.
- **Optional MP4 recording**: Add `"record": "drawing-process.mp4"` (requires `ffmpeg` on PATH) to capture the live 30fps composited PowerPoint + multi-cursor video automatically.

## Running the multi-cursor show

1. **Preflight check** (validates geometry, image paths, presets, and dependencies without touching PowerPoint):

```powershell
& '<skill目录>/scripts/start-swarm-draw.ps1' -Plan '<工作目录>/swarm-plan.json' -CheckOnly
```

2. **Launch visible PowerPoint and draw with 6 synchronized agent cursors** (pass `-CreateBlank` to automatically create and maximize a matched-canvas blank deck, or omit `-CreateBlank` to draw onto the currently active open deck):

```powershell
& '<skill目录>/scripts/start-swarm-draw.ps1' -Plan '<工作目录>/swarm-plan.json' -CreateBlank -Wait -ExportPreview '<工作目录>/preview.png'
```

PowerPoint stays visibly open on screen when done, with the editable `.pptx` saved and `preview.png` exported for immediate visual review under `references/self-correction.md`.

## Passive companion mode (fallback for `powerpoint_draw_sequence`)

If a batch must be emitted via `powerpoint_draw_sequence` (for example native charts or a small targeted repair batch), you can launch `scripts/start-swarm-draw.ps1` with a `cues` plan (`{"deck": "...", "slide": 1, "cursors": 6, "cues": [{"name": "VSS_...", "kind": "box|line|text", "points": [[x1,y1],[x2,y2]]}]}`), which runs `scripts/swarm_observer.py` read-only over the PowerPoint window with the same Figma-style multi-agent cursors.
