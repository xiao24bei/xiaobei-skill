# Output Template

Use this response structure when generating reconstruction VBA. For presentation targets, do not return only a PNG or preview image; the response must account for editable VBA/deck artifacts and the detected Office/WPS runtime.

## Short Explanation

State the selected Office host, canvas size, and reconstruction strategy in 2-4 sentences. Mention whether the result is pure editable Shapes, hybrid preserved-element reconstruction, or background-image-assisted. Say whether Microsoft PowerPoint automation, WPS-compatible manual flow, or a no-local-runtime fallback was used.

## Environment

List what was detected or inferred:

- OS:
- Presentation app:
- Materialization path: Windows PowerPoint COM / macOS PowerPoint AppleScript / WPS-compatible manual path / no local runtime.
- Automation status:

## Generated Artifacts

List the local artifacts that were created or attempted:

- VBA source:
- Editable deck:
- Preserved raster assets, if any:
- Opened in app: yes/no, with blocker if no.
- Automation runner, if created or used:
- Preview screenshot, if any: optional diagnostic only; not the editable deliverable.

## Editability Map

Summarize what is editable and what is preserved:

- Editable Shapes:
- Preserved raster elements:
- Full-background image used: yes/no.

For each preserved raster element, include its asset path, approximate source region, slide placement, and reason for preservation.

## Complete VBA Code

Return one complete code block:

```vb
' Paste into the VBA editor and run ReconstructFromImage
Option Explicit

Sub ReconstructFromImage()
    ' Complete runnable macro here.
End Sub
```

The code must include cleanup, canvas setup, background, each object, shape names, color settings, and visual-region comments. Use a stable generated-object prefix such as `AITVBA_`. In hybrid mode, preserved raster crops must be inserted with `Shapes.AddPicture` and named with the same prefix.

## How to Run

Give brief host-specific steps and prefer the already-created local artifact when available:

- Windows PowerPoint: open the generated deck if it is not already open. If automation failed, open VBA editor, import or paste the generated `.bas` code, and run the macro. If VBA import was blocked, mention Trust Center macro/VBA project access.
- macOS PowerPoint: open the generated deck if it is not already open. If automation failed, open VBA editor, import or paste the generated `.bas` code, and run the macro. If AppleScript or macro security blocked execution, say so.
- WPS Presentation: open WPS Presentation, open or create a deck, open its macro/VBA editor if the installed WPS version supports it, import or paste the generated `.bas` code, and run the macro. If macro support is unavailable, use the editable `.pptx` fallback when provided.
- Excel: open VBA editor, insert module, paste code, select/create target sheet, run the macro.
- Word: open VBA editor, insert module, paste code, run the macro.

## Self-Check Results

Include the required checks:

- Structure check:
- Editability check:
- Preservation check:
- Code check:
- Materialization check:
- Visual check:

Each item should state pass/revised/remaining concern. Preservation check should state "not applicable" for pure editable reconstructions. Materialization check must mention the detected app and whether the macro actually ran. Visual check must say "manual estimate" unless a rendered screenshot was actually compared.

## Possible Errors

List expected differences such as approximate gradients, substituted fonts, unreadable text placeholders, simplified icons, preserved raster regions that are not editable Shapes, photo-like areas that cannot be fully recreated with Shapes, PowerPoint automation limitations, WPS compatibility limits, or absent local office runtime.

## Next Optimization

Suggest one concrete next step, such as providing a screenshot of the VBA render for image comparison or specifying exact slide size/font constraints.
