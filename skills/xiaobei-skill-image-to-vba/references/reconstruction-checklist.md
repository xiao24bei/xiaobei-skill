# Reconstruction Checklist

Use this checklist before returning VBA code. If an item fails, revise the code or explicitly document the remaining uncertainty.

## Canvas

- Target host is PowerPoint, Excel, or Word.
- For presentation targets, the local runtime is identified as Microsoft PowerPoint, WPS Presentation/WPS Office, none, or unknown.
- Reconstruction mode is identified as pure editable, hybrid preserved-element, or full-background-assisted.
- Canvas size is set or clearly controlled.
- Aspect ratio matches the source image.
- If a standard slide ratio is required, the source image is uniformly fitted with offsets instead of stretched.
- Coordinate system is in points unless the user asked otherwise.
- Scale factor is consistent for all objects.

## Background

- Background color or gradient is recreated.
- Page/slide edges and margins match the source.
- Textures are approximated honestly, not claimed as exact.
- If a source image is used as a background, that choice is explicitly labeled.

## Main Objects

- Every major visible object is represented.
- Every major object is classified as editable Shape, preserved raster element, or background/base.
- Repeated elements use consistent dimensions and spacing.
- Icons are built from simple shapes, lines, or freeforms where practical.
- Tables, charts, and UI panels retain their hierarchy.
- Text, arrows, boxes, callouts, and connectors are editable unless explicitly included in a preserved raster crop for a stated reason.
- Preserved raster crops are limited to the smallest useful region around complex artwork.

## Margins and Alignment

- Outer margins match the reference.
- Left/right/top/bottom alignments are intentional.
- Centers, baselines, gutters, and repeated spacing are consistent.
- Objects that should touch or overlap do so without accidental gaps.

## Color

- Key colors are expressed as `RGB(r, g, b)`.
- Palette is consistent across repeated components.
- Transparency values match glass, overlay, disabled, or muted elements.
- Gradients are approximated with Office gradient APIs or layered translucent shapes.

## Typography

- Legible text is converted to TextBoxes.
- Unreadable text uses placeholders with comments.
- Font family, size, weight, case, and color approximate the source.
- Text box margins do not shift text away from the intended position.
- Text does not overflow unless the original does.

## Lines and Borders

- Stroke colors, weights, dashes, and arrowheads match the source.
- Border visibility is correct for each shape.
- Rounded corners are approximated using shape adjustments where possible.

## Layers

- Objects are created back-to-front or explicitly ordered with `ZOrder`.
- Shadows sit behind their objects.
- Overlays, masks, highlights, and callouts appear above the correct base layers.
- Preserved raster elements are placed at the correct z-order relative to editable overlays.

## Shadows and Effects

- Shadow offset, blur, color, and transparency are approximated.
- Soft edges, glow, or blur-like effects are represented with Office effects or layered shapes.
- Effects are omitted only when the Office API cannot reproduce them reasonably.

## Code Runability

- Code contains `Sub ... End Sub`.
- Variables are declared or intentionally simple.
- Cleanup code avoids deleting unrelated user content when a safer name-prefix cleanup is required.
- Host-specific object models are valid.
- WPS-only mode avoids advanced PowerPoint-only APIs unless local WPS support was verified.
- Shape names are assigned.
- Generated shape names use a stable prefix such as `AITVBA_`.
- Preserved raster shapes use the generated prefix and descriptive names.
- `Shapes.AddPicture` uses local asset paths and embeds images when the host supports it.
- Comments label visual regions.
- Generated code can be pasted into the VBA editor and run.

## Hybrid Preservation

- Hybrid mode is used only when requested or when pure shape reconstruction would be noisy/low-fidelity.
- Preserved elements have asset files, source crop coordinates, slide placement coordinates, and preservation reasons.
- Preserved assets are stored next to the generated VBA/deck, preferably in an `assets/` folder.
- Editable overlays do not get flattened into the preserved crop unless necessary and documented.
- The final response tells the user which elements are not editable because they are preserved raster assets.

## Materialization

- VBA source was saved to a local `.bas`, `.vba`, or `.txt` file.
- Preserved raster assets were saved locally when hybrid mode is used.
- The local presentation runtime was detected or explicitly marked unknown.
- Windows Microsoft PowerPoint uses the Windows COM path when available.
- macOS Microsoft PowerPoint uses the AppleScript path when available.
- WPS-only environments are treated as compatibility/manual unless local macro automation is actually verified.
- The editable deck was created/opened, or the blocker is stated plainly.
- If no local presentation app exists, a `.bas` is still delivered and an editable `.pptx` fallback is generated when feasible.
- PNG/JPEG/PDF previews are labeled as diagnostics only and are not treated as the final editable deliverable.

## Difference Statement

- Remaining visual differences are listed plainly.
- Complex photo, texture, or person limitations are acknowledged.
- Hybrid preserved raster regions are identified as non-editable image assets.
- The next iteration request is specific: ask for a rendered screenshot, target host, or higher-resolution source only when useful.
