# Review and repair rules

The quality pass is a correction task, not a ceremonial checklist. Review a fresh render, the scene-map inventory, and the object inventory together. In `reference_lock`, also inspect the full reference and matching region crops at readable resolution.

## Gates

| Gate | Inspect | Hard failure examples |
|---|---|---|
| Coverage | required source inventory versus verified objects | missing legend row, inset, evidence tile, plot series, label, or outcome |
| Semantics | exact text, symbols, chart meaning, arrow direction | paraphrased label, placeholder curve, wrong regulatory direction |
| Canvas | bounds, aspect ratio, margins | object clipped or shifted outside the frame |
| Text | fit, line breaks, font fallback, contrast | truncation, unintended wrap, unreadable size |
| Geometry | repeated widths, centers, spacing | cards drift, uneven columns, misaligned bands |
| Routing | arrow direction, clearances, attachment | line crosses a label, arrowhead enters a box, wrong target |
| Layering | occlusion and draw order | connector hidden, label behind fill, crop covers text |
| Fidelity | panel proportions, object scale, occupied-area density, dominant color blocks | large panel misplaced; evidence band compressed; blank space replaces source content |
| Editability | native object inventory and picture scope | full-slide raster, undocumented composite image |

When `mode: hybrid` is active, add these gates:

| Hybrid gate | Inspect | Hard failure examples |
|---|---|---|
| Asset isolation | crop bounds, alpha/background, neighboring pixels | mouse crop contains a caption or arrow fragment; silhouette clipped |
| Component integrity | raw crop versus processed asset | legitimate ER arcs, molecular components, or membrane parts removed by filtering |
| Background integration | crop edge and destination fill | visible white/pink rectangle or color seam around an asset |
| Asset fidelity | visual recognizability at final size | complex cell/organ/mouse replaced by a generic primitive |
| Semantic coverage | labels, arrows, chart data, outcomes | asset crop silently carries text that is no longer editable, or a source label is missing |
| Asset scale | visual-core size versus source | the extracted object is too small to recognize, while whitespace dominates |

## Evidence collection

For each completed region and the final slide:

1. export a new render from the current saved state;
2. compare it side by side with the reference at the same aspect ratio;
3. inspect a crop of the corresponding region at 100% or another readable scale;
4. reconcile required scene-map IDs with the named-shape inventory;
5. in hybrid mode, inspect raw and processed asset previews together;
6. record concrete findings before making corrections.

Do not approve from a stale render, object-count claim, or successful API call.

## Repair order

1. Restore missing or incorrect semantics, evidence, chart traces, and topology.
2. Correct the canvas, panel anchors, major rectangles, and visual-core scale.
3. Re-extract contaminated, clipped, or component-losing assets.
4. Reroute or trim connectors and arrows.
5. Fix text boxes, exact line breaks, and font sizes.
6. Adjust z-order and local spacing.
7. Tune colors or minor decorative details.

In hybrid mode, repair a failed crop or asset placement before tuning colors. Re-extract from a tighter source rectangle, correct the mask/background, or enlarge the visual core; never hide a defective crop with a second picture or a white patch.

Every repair should name the affected object(s), change one or a few measurable properties, and produce a new render. Prefer a named update over rebuilding unrelated regions. If a connector is wrong, do not cover it with a white shape; redraw or remove the connector.

## Pass thresholds and stop conditions

Score the whole figure and affected regions from 0 to 1 for semantic/text accuracy, editability coverage, geometry/alignment, spacing/whitespace, connector clarity, typography/color consistency, clipping/overlap safety, and reference correspondence.

Pass only when:

- semantic/text accuracy, required-source coverage, reconstructable editability, and clipping/overlap safety are 1.00;
- geometry/alignment and connector clarity are at least 0.95;
- reference correspondence is at least 0.90;
- the deterministic audit has zero hard findings;
- no warning remains except an explicitly recorded source ambiguity.

Allow at most three correction rounds by default. If the cap is reached with a hard finding, save the work as an incomplete draft and report the blocker. A missing font, renderer limitation, difficult crop, or irreducible texture may explain a residual but does not turn a failed gate into a pass. Never claim pixel identity without evidence.

## Final-save integrity gate

After the final correction:

1. save the editable deck;
2. verify the saved file reflects the latest edit state;
3. export a new preview from that saved state;
4. run the whole-slide audit on the same state;
5. run `scripts/validate_scene_map.ps1 -SceneMapPath <scene-map.json> -Phase final`;
6. ensure the preview is not newer in content than the delivered deck.

If any change is made after saving or auditing, repeat this gate.

## QA record

Keep a short record such as:

```text
route: windows-live
fidelity_profile: reference_lock
mode: native|hybrid
source required items: 184
verified required items: 184
round 0: missing legend row and right-card arrows entered the middle card
round 1: restored legend; moved endpoints to reserved outer lanes; audit hard findings 4 -> 0
native objects: 149
atomic raster objects: 24
hybrid asset gates: isolation / component integrity / background integration / fidelity / semantic coverage / scale
final hard findings: 0
remaining: none
```

The record is for reproducibility and user trust; it should not expose private tool output or include a full copy of the reference image.
