# Windows live drawing

This guide is used only when a visible native PowerPoint session is available.

## Establish the session

1. Run `scripts/validate_live_powerpoint_sequence.ps1` from the skill directory. Stop before mutation unless it returns `valid: true`; an invalid result means the installed live bridge may reconnect once per object.
2. Query the PowerPoint status and capabilities before mutating anything.
3. Open the requested deck with `powerpoint_launch`, or create a dedicated blank presentation with the live PowerPoint tool, with visibility and maximization enabled. Confirm that the launch result reports the task-persistent COM keeper. Do not shell-launch `POWERPNT.EXE`; `/x` is not a PowerPoint switch.
4. Inspect the connected presentation and confirm its path, slide count, and dimensions before drawing. A running process without an active presentation is not a usable session.
5. Match the slide aspect ratio to the source. Activate the working slide before every region batch.
6. Keep a stable output path from the beginning. Saving in place is safer than replacing an open file through a second process.

If `mode: hybrid` is active, finish the source asset manifest and contact-sheet inspection before stepwise PowerPoint drawing. The asset pass is preparation, not a substitute for the visible native drawing route.

Do not establish the drawing session until the scene-map coverage gate passes. In `reference_lock`, keep the reference available for full-slide and region-crop comparison throughout the build.

The available live operations are the native PowerPoint actions for adding a slide, shape, text box, image, line, connector, updating a named shape, setting z-order, exporting a slide image, inspecting, auditing, and saving. Prefer the paced sequence operation so each object addition is visible to the user.

The live task must keep one background COM keeper attached to the intended PowerPoint process and presentation across region boundaries. Each paced sequence must additionally pin that application, presentation, and presentation window for its entire batch. Submit one logical region with one `powerpoint_draw_sequence` call. Do not reproduce it with a client-side loop, shell loop, or one short-lived PowerShell process per object; that can release or lose the COM server between two otherwise valid objects.

## Draw in observable batches

Use several small batches rather than one monolithic call. A useful rhythm is:

1. canvas and measured panel anchors;
2. in `hybrid` mode, place the verified isolated complex visual cores and evidence tiles at their measured source-mapped rectangles;
3. complete one source region's labels, legends, insets, charts, and outcomes;
4. add that region's arrows/connectors from the topology plan;
5. render and reconcile the region's required IDs before moving to the next region;
6. correction operations and whole-slide reconciliation.

Use a modest non-zero step delay (normally about 120–260 ms) and a short pause after a region. Keep object names in the scene map and in the operation log. A batch may contain images, but an image must represent one irreducible crop, never an entire panel containing text and arrows.

Save a checkpoint after each complete logical region. When a batch reports that PowerPoint or the presentation is no longer available, stop that batch, relaunch the latest checkpoint, inspect it, and replay the same batch once. Treat it as an object-level defect only if the same named object fails again inside a pinned session; otherwise record it as a recovered session-lifecycle failure.

## Native-versus-raster rule

For every inserted picture, supply an absolute local path, alt text, a reason that the texture cannot be reconstructed economically, and a note describing what surrounding content was rebuilt natively. Keep the crop tight and preserve its aspect ratio. Do not add the original full-page image as a background or hidden fidelity layer.

## Review while the window is visible

After each major batch:

- export the active slide to a temporary PNG;
- inspect the rendered region and the named-shape inventory;
- run the PowerPoint figure audit when available;
- repair named objects with update operations, or remove and redraw only the faulty line/shape.

Use lines with explicit clearances when a free arrow is needed. Use named connectors when an edge must remain attached to two shapes. If an audit reports a false positive caused by a label overlay, move the endpoint outside the label bounds and rerun the audit rather than ignoring the finding.

For `hybrid` mode, also inspect the raw-versus-processed contact sheet and each placed asset at 100%: the crop must not contain a neighboring label, leader line, arrowhead, border, visible background rectangle, or clipped silhouette, and processing must not remove a valid disconnected component. If a complex object can be isolated, do not replace it with a generic primitive merely to reduce object count.

## Finish without hiding the process

Apply the final-save integrity gate from `self-correction.md`: save after the last edit, verify the saved state, export a fresh preview from that state, and run the whole-slide audit again. Leave the visible PowerPoint window open. Do not call a quit/close action merely to release the file. If the user is recording, the saved deck, final preview, and final visible state should be the same presentation.
