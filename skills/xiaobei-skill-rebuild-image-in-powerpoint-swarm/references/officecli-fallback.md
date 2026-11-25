# Office CLI fallback

Use this route on macOS, on Windows when native PowerPoint cannot be connected, or when the user explicitly asks for an offline build.

## Build from a scene map

1. Run the local Office CLI preflight and load its PowerPoint capability definitions.
2. Ask the CLI for the schema of any shape, text, picture, or connector property that is uncertain; do not guess enum names.
3. Pass the scene-map coverage gate before emitting slide operations.
4. Create the custom slide canvas using the aspect-preserving scale in `scene-map.md`.
5. If `mode: hybrid` is active, extract and inspect the atomic asset manifest/raw-versus-processed contact sheet before emitting slide operations.
6. Emit one complete source region at a time in layer order, then emit connectors after their named endpoints exist. Reconcile all required region IDs before continuing. Keep the operation data reproducible in a JSON draw log.
7. Save the PPTX, render a PNG from the saved artifact, and inspect the rendered output before calling it final.

## What remains editable

Use Office shapes for text, cards, borders, arrows, lines, legends, tables, and simple icons. In `hybrid` mode, use small, isolated pictures for irregular biological visual cores (mouse, organ, cell, tissue) as well as photos, microscopy, and dense texture when reconstructing them natively would destroy visual fidelity. Document each asset in the scene map and QA report. Do not use a full panel or a montage as one picture.

## Offline review loop

Run the CLI's structural validation, text/issue inspection, and slide screenshot commands. Compare the whole slide and matching region crops with the source after an aspect-preserving resize. Reconcile the required-source inventory with generated objects. Convert every actionable difference into a scene-map edit (object ID, property, old value, new value), rebuild a new iteration, and render again. Do not repeat an unchanged build. Use a finite repair cap; if a hard finding remains, hand off only an explicitly incomplete draft rather than treating a residual report as a pass.

## Expected handoff

Keep user-facing files together: editable PPTX, scene map, draw log or batch files, render PNG, and a short QA report. A PNG or PDF is evidence for review only; it is not a substitute for the editable deck.
