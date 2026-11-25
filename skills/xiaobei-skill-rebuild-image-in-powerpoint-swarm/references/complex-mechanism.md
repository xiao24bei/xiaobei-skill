# Complex mechanism mode

This branch is only for dense scientific mechanism figures whose visual identity depends on irregular biological illustrations, photographic evidence, or many interacting panels. It is intentionally conditional so ordinary PPT redraws and technical-roadmap diagrams keep the original native-first behavior.

When a reference image is supplied, hybrid mode remains under `fidelity_profile: reference_lock`. Hybrid means selective raster preservation, not permission to summarize, redesign, or reduce the source's information density.

## Activation contract

Set `mode: hybrid` only when the user explicitly asks for a complex mechanism redraw or the shared complexity score reaches the threshold in `SKILL.md`. Record the reason in the scene map. Never use this branch as a shortcut for a simple flowchart, a slide screenshot with ordinary icons, or a technical route made mostly from boxes and arrows.

## Inventory and classify before drawing

First enumerate every visible item by region: panel/background, complex visual core, native label, legend row, route/arrow, inset/callout, evidence tile, chart/series, and caption. Record the inventory in `scene-map.md` and pass the coverage gate before extracting or drawing anything.

For every visible region, assign one class:

| Class | Examples | PowerPoint treatment |
|---|---|---|
| `semantic_native` | headings, labels, legends, borders, arrows, connectors, axes, regular tables, simple symbols | Native text/shape/line/table/chart. |
| `complex_visual_core` | mouse, animal, organ, cell cross-section, mitochondrion, tissue illustration, dense molecular artwork | Atomic image asset, preferably transparent-background PNG; keep the surrounding semantics native. |
| `evidence_raster` | microscopy field, thermal image, histology, brain photograph, heatmap | Atomic image asset; split montages and channel stacks into individual tiles. |
| `reconstructable_plot` | survival curve, bar chart, regular scatter/line plot | Native editable chart or native axes/series. |
| `composite_exception` | A visual core whose internal labels cannot be separated without destroying the source | Smallest possible atomic crop, documented as an exception; do not use it for a whole panel. |

Do not redraw a `complex_visual_core` as a generic oval, rounded rectangle, stick figure, or emoji-like symbol when an isolated source crop can be made. The purpose of hybrid mode is to preserve the visual information that native primitives cannot express economically.

## Asset-first workflow

1. Inspect the source at native resolution and mark source rectangles for every complex visual core and evidence tile. Use pixel rectangles, not guessed slide coordinates.
2. Extract each object at native resolution. Preserve both the raw source crop and the processed asset. Prefer a transparent cutout when the object has a separable silhouette; otherwise use a tight crop whose background is sampled from and matches the destination panel.
3. Remove neighboring labels, arrowheads, leader lines, borders, and unrelated objects from the crop. If they cannot be removed without damaging the object, mark the asset `composite_exception`, document the exact retained content, and rebuild the surrounding route so no duplicate text or arrow is visible.
4. Preserve aspect ratio and do not upscale a small crop beyond what remains crisp at the final slide size. Keep both the raw crop and the placed asset path when practical.
5. Build a contact sheet showing each raw crop beside its processed asset, with asset names, source rectangles, expected content, forbidden content, and intended placement. Inspect it at 100% before opening the PowerPoint drawing pass.
6. Record each asset in the scene map and QA log with: source rectangle, output path, alpha/background method, visual role, raster reason, and the native objects rebuilt around it.
7. Place the assets before their local labels and connectors. Use stable names such as `VSS_A_mouse_asset`, `VSS_B_cell_asset`, or `VSS_C_brain_tile_03`.

On Windows, prefer the bundled deterministic extractor after preparing the manifest:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/extract_visual_assets.ps1 -ManifestPath <manifest.json> -Overwrite
```

It validates source rectangles and raster scope, creates PNG assets, optionally removes only edge-connected background pixels, optionally keeps the largest connected foreground object, and generates a labeled raw-versus-processed contact sheet plus an extraction report. Use `alpha_method: edge-connected` only when the object has a visible boundary; use `none` or `matched-background` when a pale object could merge into the background.

Default to `component_filter: none`. Use `component_filter: largest` only when the intended visual is known to be one connected silhouette. Never use it for multi-part cell structures, endoplasmic-reticulum/nucleus groups, membrane assemblies, molecular clusters, or evidence containing multiple legitimate disconnected components. Under strict validation, `largest` requires `expected_connected_parts: 1`, a `component_filter_reason`, and explicit approval of any meaningful component loss. A processed asset that discards a legitimate source component is a hard failure even when the remaining component looks clean.

## Asset manifest contract

Each asset entry must contain at least:

```json
{
  "id": "VSS_A_mouse_asset",
  "kind": "picture",
  "visual_role": "GBM-bearing mouse",
  "source_rect_px": [120, 930, 520, 340],
  "asset": "assets/mouse_cutout.png",
  "alpha_method": "edge-connected|none|matched-background",
  "editability": "atomic-raster",
  "raster_reason": "irregular photographic/biological silhouette",
  "native_surroundings": ["VSS_A_mouse_label", "VSS_A_injection_arrow"],
  "fidelity_priority": "high"
}
```

The extractor manifest uses the same fields at the top level:

```json
{
  "source": "source.png",
  "output_dir": "assets",
  "contact_sheet": "asset-contact-sheet.png",
  "report": "asset-extraction-report.json",
  "strict_validation": true,
  "save_raw_crops": true,
  "max_asset_area_ratio": 0.35,
  "assets": [
    {
      "id": "VSS_A_mouse_asset",
      "visual_role": "GBM-bearing mouse",
      "source_rect_px": [120, 930, 520, 340],
      "output": "mouse_cutout.png",
      "alpha_method": "edge-connected",
      "component_filter": "none",
      "background_color": "#FFFFFF",
      "destination_background_color": "#FFFFFF",
      "tolerance": 28,
      "expected_content": ["mouse silhouette"],
      "forbidden_content": ["caption", "syringe", "arrow"],
      "raster_reason": "irregular biological silhouette",
      "native_surroundings": ["VSS_A_mouse_label", "VSS_A_injection_arrow"]
    }
  ]
}
```

The manifest is not permission to paste the complete reference image. Every picture must be one isolated visual unit, and every neighboring label, line, arrow, legend, frame, axis, and chart must remain separately editable.

Reject an asset before placement when the raw-versus-processed comparison shows any of the following: a valid component was removed; an important silhouette touches the crop edge; a leader line, text fragment, arrowhead, or neighboring object remains; the background creates a visible rectangular seam; or the asset's visible bounds/scale no longer correspond to the source.

## Composition order

Use this z-order within each local mechanism region:

1. panel/background fill;
2. atomic complex visual assets and evidence tiles;
3. native connectors and arrows, with reserved lanes and endpoint clearance;
4. native labels, legends, axes, and callouts;
5. highlights or selection outlines.

Keep connectors out of asset interiors unless the source clearly shows an arrow entering the object. In that case, attach the connector to the asset boundary or use a short, explicitly documented overlay segment. Do not hide a wrong connector under a white patch.

## Fidelity rules for dense figures

- Match the source's dominant visual hierarchy first: panel proportions, large visual cores, arrow direction, and outcome boxes.
- Give complex visual cores enough area to remain recognizable at the final slide size. Do not shrink an extracted mouse, cell, or organ to the point that a native placeholder would appear equivalent.
- Use a native chart for regular quantitative plots. If the source chart contains a dense photographic background, separate the background crop from the native axes, labels, and series.
- Split comparison montages into individual pictures. A montage, prediction grid, or channel stack is not an atomic asset.
- Preserve source labels exactly when legible; use OCR/vision as an aid, then visually verify every line break and symbol.
- Keep text and arrows outside the safe visual area of a crop. Reserve clear lanes before placing them.
- If a complex object is partly occluded by a source label, crop the cleanest visible region and rebuild the label natively rather than duplicating the source text.
- Preserve source-region occupancy. Do not shrink the mouse, tumor, cell, nucleus, evidence strip, or chart to create aesthetically convenient blank space that is absent from the reference.
- Reconstruct all visible traces of a regular plot. Placeholder axes, short sample segments, or a text-only legend are not an acceptable substitute.
- Preserve mechanism topology. Every release, translocation, regulatory, and outcome arrow must retain its source, target, and direction; do not replace a path with prose.

## Hybrid review gates

After each asset group and again on the whole slide, verify:

1. **Region completeness:** every required scene-map item in the region is present; missing legends, insets, evidence tiles, chart traces, labels, or outcomes are hard failures.
2. **Asset isolation:** no crop contains unrelated labels, arrowheads, leader lines, borders, or neighboring panels; no important silhouette is clipped.
3. **Raw/processed integrity:** the processed asset preserves every intended component from the raw crop and does not create a visible rectangular seam.
4. **Asset fidelity:** the mouse/organ/cell/photograph is recognizable and visually closer to the source than a generic primitive.
5. **Semantic and topology coverage:** all text, arrows, legends, outcomes, chart semantics, endpoints, and arrow directions remain present and editable.
6. **Density and scale:** normalized panel bounds, major visual-core size, and occupied-area balance correspond to the source; whitespace does not replace omitted content.
7. **Layering and routing:** connectors are visible, arrows point to the intended object, and labels do not overlap assets or other labels.
8. **Raster scope:** no single picture is a full slide or a reconstructable panel; every picture has a manifest entry and a reason.

If a gate fails, repair the asset or its local geometry first. Do not compensate for a bad crop by adding a second raster overlay or by flattening the panel.

## Fallbacks

If reliable segmentation is unavailable, use a tight crop with a matched background and mark `alpha: matched-background`. If even a tight crop would contain unrelated reconstructable content, split the region further or rebuild the object natively only when its visual complexity is genuinely low. Report the exact residual in the QA file.

## Non-regression check

At handoff, state whether `hybrid` was activated and why. If the score was below threshold, confirm that no complex-mode asset extraction was introduced. This keeps ordinary PPT image redraws, academic technical routes, and simple diagrams on the original workflow.
