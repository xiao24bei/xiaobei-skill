# Scene-map contract

The scene map is the measured reconstruction contract shared by the live and Office CLI routes. It must be complete enough to detect an omitted object before drawing and to repair a named object without guessing.

## Required top-level structure

```json
{
  "source": {"path": "absolute/source.png", "width_px": 2400, "height_px": 1790},
  "fidelity_profile": "reference_lock",
  "mode": "native|hybrid",
  "complexity": {"score": 0, "reason": ""},
  "canvas": {"width_pt": 720.0, "height_pt": 537.0},
  "reference_inventory": {
    "complete": true,
    "regions": ["A", "B", "C", "caption"],
    "counts": {
      "panels": 3,
      "text_items": 0,
      "visual_cores": 0,
      "evidence_tiles": 0,
      "plots": 0,
      "native_symbols": 0,
      "insets": 0,
      "captions": 0,
      "links": 0
    },
    "required_ids": [],
    "unresolved_ambiguities": [],
    "authorized_omissions": []
  },
  "objects": [],
  "links": []
}
```

`reference_inventory.complete` may be true only after a panel-by-panel inspection at readable resolution. Counts are source counts, not generated-object counts. Every `required_id` must resolve to exactly one object or link. A convenient-looking subset is not a complete inventory.

## Required object fields

Every visible source item needs an entry with:

- `id`, `parent_region`, `kind`, `source_category`, `required`, and `status` (`planned`, `drawn`, or `verified`); use `panel`, `text_item`, `visual_core`, `evidence_tile`, `plot`, `native_symbol`, `inset`, or `caption` for `source_category`;
- measured `source_rect_px: [x, y, width, height]`;
- `mapping: scaled` unless a small documented adjustment is necessary;
- `layer`, `editability`, and reconstruction treatment;
- for text: exact `text`, `line_breaks: exact`, alignment, font family/weight, color, and measured source rectangle;
- for regular plots: axes, tick labels, legend, and every visible series/step trace as separate native objects or a native chart plan;
- for evidence grids: one entry per irreducible field or photograph, plus native grid/frame/labels;
- for pictures: `asset`, `visual_role`, `raster_reason`, `alpha_method`, `expected_content`, `forbidden_content`, `native_surroundings`, and `fidelity_priority`.

Example:

```json
{
  "id": "VSS_A_mouse_asset",
  "parent_region": "A",
  "kind": "picture",
  "source_category": "visual_core",
  "required": true,
  "status": "planned",
  "source_rect_px": [120, 930, 520, 340],
  "mapping": "scaled",
  "asset": "assets/mouse_cutout.png",
  "editability": "atomic-raster",
  "visual_role": "GBM-bearing mouse",
  "raster_reason": "irregular biological silhouette",
  "alpha_method": "edge-connected",
  "component_filter": "none",
  "expected_content": ["mouse silhouette", "implanted tumor marker"],
  "forbidden_content": ["caption", "syringe", "route arrow"],
  "native_surroundings": ["VSS_A_mouse_label", "VSS_A_injection_arrow"],
  "fidelity_priority": "high",
  "layer": 10
}
```

## Required link fields

Do not reduce a mechanism route to decorative lines. Every arrow or connector entry must include:

- `id`, `parent_region`, `required`, and `status`;
- `from`, `to`, and `direction`;
- `source_route_px` or measured waypoints when the source route is not straight;
- arrowhead style, line style, endpoint clearance, and any attached label ID;
- `topology_verified: false` until the fresh render confirms the intended source and target.

Connector endpoints should reference named object IDs. An arrow pointing into empty space, a disconnected continuation, or a text-only substitute does not satisfy the link entry.

## Coordinate rule

For a source of `W × H` pixels and a target canvas of `Tw × Th` points, use one scale:

```text
s = min(Tw / W, Th / H)
offset_x = (Tw - W*s) / 2
offset_y = (Th - H*s) / 2
left = offset_x + x*s
top = offset_y + y*s
width = w*s
height = h*s
```

Keep the original pixel rectangle in the map. A target rectangle derived from it may be cached, but the source rectangle remains authoritative. Do not independently resize or reposition objects until the measured first-pass mapping has been rendered. Any adjustment larger than 1.5% of the slide dimension for a panel/major visual core or 2.5% for a secondary object needs a `mapping_adjustment_reason` and a fresh regional comparison.

## Coverage gate

Before drawing, verify:

1. every source panel has entries for its background, major visual anchors, labels, legends, routes, insets, evidence, plots, and outcomes;
2. all `required_ids` exist and all text has been transcribed;
3. links reference existing endpoints and preserve arrow direction;
4. regular plots are not represented by placeholder axes or sample lines;
5. no asset entry covers a whole reconstructable panel or contains editable labels/arrows;
6. unresolved ambiguities and authorized omissions have been surfaced rather than silently ignored.

At handoff, every required entry must have `status: verified`. Reconcile source counts with verified counts in the QA report.

Run `scripts/validate_scene_map.ps1 -SceneMapPath <scene-map.json> -Phase planning` before drawing and rerun it with `-Phase final` after verification. The validator checks required IDs, count reconciliation, source bounds, picture metadata, plot plans, link endpoints, and final statuses. It cannot determine whether the visual inventory itself is semantically complete, so the native-resolution source review remains mandatory.
