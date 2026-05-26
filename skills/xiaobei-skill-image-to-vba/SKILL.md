---
name: xiaobei-skill-image-to-vba
description: use when users want XiaoBei skill / xiaobei-skill / 小北在读研 style academic image-to-vba reconstruction: convert academic figures, scientific diagrams, slides, screenshots, or other images into vba, office drawing code, powerpoint/excel/word shapes, editable shape reconstruction, 1:1 recreation, pixel-like approximation, or hybrid editable Office shape reconstruction.
---

# XiaoBei Image to VBA

小北在读研出品的 academic image-to-vba reconstruction skill. The workflow focuses on editable Office Shapes first, with carefully labeled raster crops only when complex scientific artwork should be preserved.

## Output Language

All analysis, manifest tables, self-check reports, run reports, and conversational responses produced by this skill MUST be written in Simplified Chinese (简体中文). This includes:

- 视觉拆解、元素清单、坐标说明、自检结论、运行报告、给用户的最终回复。
- VBA 代码内的注释也使用中文（例如 `' 标题区域背景`），便于用户后续维护。

仅以下内容保留英文原样，不要翻译：

- VBA 关键字、API 名、属性名（`Sub`、`Shapes.AddShape`、`Fill.ForeColor.RGB` 等）。
- 形状名前缀 `AITVBA_` 与脚本路径、文件名。
- 颜色十六进制值与 `RGB(...)` 调用本身。

如果用户在对话里明确要求换成英文或其它语言，则按用户指定的语言输出，但 manifest 字段名仍保留英文 key 以便机器解析。

## Purpose

Convert a reference image into a complete, runnable Office VBA macro that reconstructs the visual using editable Shapes, TextBoxes, Lines, Freeforms, fills, gradients, transparency, shadows, and ZOrder. When the image contains complex, distinctive graphic elements that would become noisy or low-value if rebuilt as hundreds of shapes, preserve those elements as cropped local image assets and reconstruct the surrounding structure with editable VBA Shapes. The core loop is: image -> environment detection -> visual decomposition -> editable-vs-preserved element plan -> Office coordinate model -> VBA Shapes code and cropped assets -> run or prepare the macro in the available presentation app -> self-check -> iterative correction.

Prefer editable Office Shapes for structure, labels, arrows, boxes, chart scaffolds, and readable text. Preserve complex graphic elements only as local cropped assets when that improves fidelity and keeps the VBA maintainable. Do not embed the full source image as a shortcut unless the user explicitly permits it or chooses the background-image-assisted option for complex photos. A PNG/JPEG/PDF preview is never the final deliverable for this skill; it is only allowed as a diagnostic screenshot after an editable Office reconstruction has been produced or attempted.

## Default Deliverable Contract

When the user provides an image and asks for reconstruction, the default target is a presentation deck unless another Office host is explicit. Prefer Microsoft PowerPoint automation when available, but support WPS users and no-local-office users with a truthful fallback path. The agent must produce:

- A complete VBA macro file or code block that reconstructs the image with editable Office Shapes.
- Any preserved local image assets needed for hybrid reconstruction, with clear labels showing they are not editable Shapes.
- A presentation file created from that macro when local automation is available, or a clear statement that local presentation automation was blocked or unsupported.
- The generated deck opened for the user when local PowerPoint/WPS opening is available.
- A short run report that says whether the macro was executed, whether the deck opened, and where the generated `.bas` / `.pptx` / `.pptm` files are located.

Do not satisfy an image-to-VBA request by generating only a PNG image, raster preview, static screenshot, or visual mockup. If a preview image is useful, create it only in addition to the editable VBA/PPT deliverables.

## Hybrid Preservation Mode

Use hybrid preservation when any of the hard triggers below fire. Do not rely on subjective judgment about whether the result "looks cluttered" — model self-assessment is biased toward editable reconstruction, which causes complex artwork to be redrawn poorly. The goal is to preserve high-complexity artwork while keeping the diagram's editable structure editable.

Hard preservation triggers (hit any one → the element MUST be preserved as a raster crop):

- The element is a microscopy image, 3D render, real photograph, or photo-realistic illustration.
- The element contains photo-grade gradients, textures, noise, or organic shading that cannot be matched by a few solid fills or simple gradients.
- The element is a logo, trademarked mark, or branded artwork.
- A faithful editable reconstruction would require more than ~15 Shapes for that single element.
- The element is hand-drawn, painterly, or has stroke-width variation that cannot be matched by simple Line/Freeform.
- The user explicitly named the element as "keep as image", "保留", "不拆", or equivalent.

If none of the triggers fire, treat the element as editable. If a trigger fires but the user demanded "all editable", explain the fidelity tradeoff first and get permission before redrawing.

Classify every visible element into one of three buckets:

- Editable Shapes: text, arrows, boxes, connectors, axes, callouts, tables, UI panels, simple icons, repeated markers, and diagram scaffolds.
- Preserved raster elements: distinctive illustrations, logos, scientific renderings, microscopy/photo crops, 3D objects, complex texture patches, dense molecular/mesh graphics, screenshots inside device frames, and hand-drawn/artistic elements whose shape-by-shape reconstruction would be noisy.
- Background/base: plain solid/gradient backgrounds should be rebuilt as shapes; full-image background insertion is allowed only when the user explicitly asks for it or the source is photo-like.

Hybrid rules:

- Preserve only the smallest useful crop around each complex element; do not insert the entire source image just to preserve a small object.
- Put preserved assets behind or between editable layers according to the original z-order.
- Recreate nearby labels, arrows, boxes, highlights, and callouts as editable VBA objects even when they overlap preserved artwork.
- Name preserved image shapes with the generated prefix, for example `AITVBA_Raster_Microscope_01`.
- Save cropped assets next to the generated VBA/deck, preferably under an `assets/` subdirectory.
- In the final response, list preserved elements separately from editable elements so the user knows which parts remain raster.
- If the user requests "all editable", do not use preserved raster elements unless you first explain the fidelity tradeoff and get permission.

### 保留元素的几何契约 (Preservation Geometry Contract)

保留元素出现"裁错 / 变形 / 缩放比例不对"是 hybrid 模式最常见的失败。下面这套契约是硬性的,违反任何一条都视为本轮 render-verify 不通过,必须回到 manifest 修正。

1. **裁剪 bbox 必须等于 manifest bbox_px,不准 padding,不准外扩,不准内缩**
   - 调用 `scripts/crop_preserved_elements.py` 时,传入的坐标必须与 manifest 中该元素的 `bbox_px` 完全一致(整数像素)。
   - 不要为了"美观"或"留出阴影"自行扩边。如果原图边界外有阴影/光晕需要保留,那应在 manifest 里就把 `bbox_px` 扩到包含阴影的范围,而不是裁的时候临时加。
   - 裁出来的 PNG 必须保留原图 alpha 通道(若有),不要预先合成到白底。

2. **插入时必须保持裁剪图自身的宽高比**
   - 在 VBA 里写 `Shapes.AddPicture` 后,只能用 ratio-safe 缩放:已知裁剪图原始像素宽高 `cw_px, ch_px`,目标插入区域宽高 `tw_pt, th_pt`,则:
     - `s = min(tw_pt / cw_px, ch_px > 0 ? th_pt / ch_px : 1)`
     - 实际插入的形状宽 = `cw_px * s`,高 = `ch_px * s`
   - 严禁直接用 `Width = X : Height = Y` 把图片强行拉到 manifest 的目标尺寸,除非目标尺寸恰好与裁剪图等比。一旦目标 bbox 与裁剪源比例不一致,要么调整目标 bbox 使之与源等比,要么在目标区域内居中留白。
   - 不允许给 `AddPicture` 之后再设置 `LockAspectRatio = msoFalse` 然后独立改宽高。

3. **裁剪图与目标 bbox 的比例必须一致**
   - manifest 中保留元素的 `bbox_px` 既描述源图坐标,也决定目标 bbox 的宽高比。
   - 如果裁出来的 PNG 实际宽高比与 manifest 的 `bbox_px` 比例偏差超过 ±2%,说明裁剪坐标错了,必须重新裁,而不是在插入时拉伸去"对齐"。

4. **不允许旋转 / 翻转 / 任意 freeform 蒙版**
   - 默认 `Rotation = 0`,`FlipH = msoFalse`,`FlipV = msoFalse`。除非原图本来就是斜的,且 manifest 里该行显式给出了 `rotation_deg` 字段,否则不许动。
   - 不要给保留图片套圆角矩形蒙版、椭圆蒙版、freeform 蒙版,这会改变它的视觉形状。需要圆角的,只能在 manifest 里把元素本身标为 editable,而不是把矩形裁剪强行变形。

5. **插入位置使用 ratio-safe 坐标转换**
   - 用第 3 节的 `scale / offset_x_pt / offset_y_pt` 公式把 `bbox_px` 的 (x, y) 转成插入点 (Left, Top)。
   - 不要用独立 X/Y 拉伸的版本,即使画布与原图同比也优先用 ratio-safe 形式,避免一处偷懒导致其它元素错位。

6. **裁剪资产的命名与可追溯**
   - 文件名 = `assets/<id>_<语义短名>.png`,例如 `assets/R03_microscope_left.png`。
   - 在 manifest 中保留该文件的 SHA-1 前 8 位或文件大小,render-verify 闭环里若某个保留元素差异大,先核对这一行哈希是否变了,排除"裁了又被覆盖"的情况。

### 保留元素自检 (额外的第 3.5 项 self-check)

在第 6 节 Mandatory Self-Checks 的"Preservation check"基础上,每一轮 render-verify 都额外运行下面四问,任何一问回答不出"是",该元素就不算合格:

- 这张裁剪图的宽高比 (cw_px / ch_px) 与 manifest bbox_px 的宽高比之差 < 2% 吗?
- VBA 中该 `AddPicture` 之后的 `Shape.Width / Shape.Height` 比值与裁剪图原始比值之差 < 2% 吗?
- 该形状的 Rotation/FlipH/FlipV 是否与 manifest 显式声明一致(默认 0/false/false)?
- 渲染图中该形状的边界与原图中对应区域的边界,在 ratio-safe 映射后误差是否 < 画布短边的 2%?

不达标时,优先调整 manifest 的 `bbox_px` 与 `crop_path`,而不是在 VBA 里硬拉尺寸。

## Environment Compatibility

Before trying to materialize a deck, identify the user's local runtime. Use `scripts/detect_office_environment.py` when local shell access is available; otherwise infer from the user's stated OS and installed office suite.

Choose the delivery path in this order:

1. Windows + Microsoft PowerPoint: generate `.bas`, run `scripts/vba_lint.py`, then try `scripts/run_powerpoint_vba_windows.ps1`. Report if PowerPoint Trust Center blocks VBA import.
2. macOS + Microsoft PowerPoint: generate `.bas`, run `scripts/vba_lint.py`, then try `scripts/run_powerpoint_vba_macos.py`. Report if AppleScript automation or Office macro security blocks execution.
3. WPS Presentation / WPS Office only: generate conservative PowerPoint-compatible VBA and a `.bas` file. Do not assume automatic VBA import/run is available. If WPS can be opened locally, open it and provide WPS-specific manual import/run steps. If WPS macro support is absent or unverified, say that plainly.
4. No local presentation app or unsupported platform: still generate the `.bas` VBA source. When feasible, also generate an editable `.pptx` fallback from the same shape plan, clearly labeled as a non-VBA fallback. Do not substitute a PNG as the final output.

Use WPS-compatible mode when WPS is the only available app:

- Prefer basic Shapes, Lines, TextBoxes, solid fills, RGB colors, simple transparency, and simple shadows.
- Avoid depending on advanced PowerPoint-only effects, complex `TextFrame2` behavior, gradient stops, and macro security settings unless verified locally.
- Save outputs in Microsoft-compatible formats (`.bas`, `.pptx`, `.pptm`) rather than WPS-only formats unless the user asks otherwise.
- Label WPS results as compatibility-targeted, not fully PowerPoint-verified, unless they were actually opened and run in WPS.

## Workflow

### 1. Input Recognition

Identify the target Office host:

- Use PowerPoint when the user does not specify a host.
- Use Excel when the user asks for a worksheet, dashboard, spreadsheet, chart sheet, grid, or Excel Shapes.
- Use Word when the user asks for a document page, report, Word canvas, or inline/anchored drawing.

Ask a concise clarification only when the target host materially changes the solution and cannot be inferred. Otherwise proceed with the default presentation target, detect the available runtime, and attempt the best supported create/open path.

Classify the source image:

- Vector/flat graphic: icons, diagrams, simple logos, flowcharts, infographics. Aim for high-fidelity editable shape reconstruction.
- UI screenshot/poster: cards, bars, text, panels, tables, app screens. Aim for detailed reconstruction with editable text and shape groups.
- Mixed scientific/academic figure with complex artwork plus labels: default to hybrid preservation for the complex artwork and editable reconstruction for text, arrows, boxes, connectors, legends, and callouts.
- Complex photo/texture/person: explain limits. Offer:
  - A. Pure VBA Shapes approximate reconstruction.
  - B. Hybrid preservation of local crops plus editable VBA overlays.
  - C. Use the image as a full background and add editable VBA overlays.

### 2. Image Parsing

Inspect the image before coding. Record:

- Canvas ratio, orientation, approximate pixel dimensions, and target Office canvas size.
- Background: solid, gradient, image-like, texture, transparency, page margins.
- Major objects: shapes, panels, icons, chart elements, tables, arrows, connectors, callouts.
- Reconstruction bucket for each major object: editable shape, preserved raster crop, or background/base.
- Preserved raster crop coordinates and asset names when hybrid mode is used.
- Text: exact words when legible; use placeholder text with a comment when uncertain.
- Colors: approximate RGB values; keep a palette list for reuse.
- Lines: stroke width, caps, joins, dash style, arrowheads.
- Effects: shadows, blur-like approximations, gradients, transparency, glow, bevel-like effects.
- Layers: back-to-front order and overlaps.
- Alignment: margins, gutters, repeated spacing, baseline grids, visual centers.

For detailed pattern examples, load `references/vba-shape-patterns.md` only when needed.

### 2.5 Element Manifest (mandatory)

Before writing any VBA, produce an explicit element manifest. This step is mandatory — do not skip to coordinate modeling or code generation without it. Skipping the manifest is the root cause of three frequent failures: missed preservation, drifted layout, and arrows pointing to wrong endpoints.

Output the manifest as a Markdown table or JSON list with these fields per element:

- `id`: stable short id, for example `E01`, `R01`, `L01`. Use `E*` for editable, `R*` for preserved raster, `L*` for lines/arrows/connectors, `B*` for background.
- `type`: rect, ellipse, textbox, arrow, line, freeform, group, raster_crop, background, etc.
- `bucket`: `editable`, `preserved`, or `background`. Apply the hard preservation triggers from "Hybrid Preservation Mode" — do not guess.
- `bbox_px`: `(x, y, w, h)` in source-image pixels. Required for every element except pure connectors. Estimate visually but commit to numbers; do not write `?` or "approx".
- `style`: fill color (hex or RGB), stroke color, stroke width, font size, font weight, opacity, shadow, gradient endpoints, etc. Use `null` only when the element has no such property.
- `text`: exact text content for textboxes; `null` otherwise. Mark uncertain readings with a trailing `(?)`.
- `endpoints`: for arrows/lines/connectors only. **Mandatory and strict**. Format must be `from:(x1_px,y1_px) -> to:(x2_px,y2_px) [anchor:<id>.<side>=<id>.<side>]`. Both absolute pixel coordinates AND the semantic anchor (which element's which side) are required when the line connects two elements. Reasons:
  - 仅靠"靠近 A"或"指向 B"会导致箭头乱指 — 模型会按印象画。
  - 必须在原图上目测出起点和终点的像素坐标，并写明粘连到哪个元素的哪条边（top/right/bottom/left/center）。
  - 如果箭头中间有折线（elbow / orthogonal），用 `via:(x,y);(x,y)` 字段补充中间拐点像素坐标。
  - 如果起点或终点是悬空的（不连接任何元素），写 `anchor:free`。
  - 生成 VBA 时，连接两元素的箭头必须用 `Shapes.AddConnector` 并 `BeginConnect` / `EndConnect` 到具体形状，而不是用绝对坐标画一根 `AddLine`。这样箭头会随被连接元素位置自动跟随,不会乱指。
- `z_order`: integer, lower draws first.
- `preserve_reason`: required when `bucket = preserved`. State which hard trigger fired.
- `crop_path`: required when `bucket = preserved`. Use `assets/<id>_<short_name>.png`.

Example minimal row:

```
| id  | type      | bucket    | bbox_px            | style                       | text       | endpoints                    | z | preserve_reason | crop_path                  |
|-----|-----------|-----------|--------------------|-----------------------------|------------|------------------------------|---|-----------------|----------------------------|
| E01 | rect      | editable  | 120,80,200,60      | fill=#FF8800,stroke=#333,2pt| "Input"    | —                            | 2 | —               | —                          |
| R01 | raster    | preserved | 400,200,180,180    | —                           | —          | —                            | 1 | microscopy image| assets/R01_microscopy.png  |
| L01 | arrow     | editable  | —                  | stroke=#000,2pt,arrow_end   | —          | from:E01.right -> to:R01.left| 3 | —               | —                          |
```

Manifest gate: do not generate VBA until every visible element appears in the manifest. The number of rows in the manifest is the lower bound on the number of Shapes the macro will create (preserved elements count as one `AddPicture` each). If a region of the image has no manifest row, you have not parsed that region — go back to step 2.

### 3. Coordinate Modeling

Model the image in Office points by default.

- Preserve the original aspect ratio.
- For PowerPoint, prefer a proportional custom slide size based on the image ratio. Use `960 pt` as the long side unless the user requests a specific slide size.
- If the user requires a standard slide size such as 16:9, fit the image uniformly within that canvas and add `offset_x_pt` / `offset_y_pt`; do not stretch the image independently on X and Y.
- Only use independent X/Y scaling when the target canvas exactly matches the source aspect ratio or when the user explicitly accepts distortion.
- Convert each image coordinate using the ratio-safe method:
  - `scale = min(target_width_pt / image_width_px, target_height_pt / image_height_px)`
  - `offset_x_pt = (target_width_pt - image_width_px * scale) / 2`
  - `offset_y_pt = (target_height_pt - image_height_px * scale) / 2`
  - `x_pt = offset_x_pt + x_px * scale`
  - `y_pt = offset_y_pt + y_px * scale`
  - `w_pt = w_px * scale`
  - `h_pt = h_px * scale`
- Use stretched conversion only for a ratio-matched canvas:
  - `x_pt = x_px * target_width_pt / image_width_px`
  - `y_pt = y_px * target_height_pt / image_height_px`
  - `w_pt = w_px * target_width_pt / image_width_px`
  - `h_pt = h_px * target_height_pt / image_height_px`
- When exact image dimensions are known, use `scripts/coordinate_helper.py` to calculate scale values and examples.
- Use named constants for canvas width, canvas height, and repeated colors.
- Keep related objects grouped in code sections matching visual regions.

### 4. Code Generation

Generate a complete runnable VBA macro, not a fragment. For the default presentation target, also save the macro as a `.bas` or `.vba` file in the working directory so it can be run or pasted without copying from chat.

Use a two-pass generation strategy. The two-pass split exists because layout errors are cheapest to catch before any styling is applied; do not skip Pass 1 to save time.

- Pass 1 — Skeleton: emit one macro that draws every manifest row as a flat gray rectangle (or thin gray line for connectors) at its mapped point coordinates, with the element `id` printed inside as a label. No fills, no gradients, no shadows, no fonts beyond a default 10pt. The skeleton's only job is to confirm bbox positions, sizes, and arrow endpoints match the source. If the skeleton is visibly wrong, fix the manifest first, not the styling.
- Pass 2 — Styled reconstruction: only after the skeleton matches the source, regenerate the macro with real fills, strokes, fonts, gradients, shadows, transparency, and ZOrder. Preserved raster crops enter at this pass via `Shapes.AddPicture`.

You may emit both passes as two separate `Sub` procedures inside the same `.bas` file (for example `Sub BuildSkeleton()` and `Sub BuildFinal()`), so the user can run the skeleton first and the styled version second without regenerating code.

The macro must include:

- A `Sub ... End Sub` wrapper.
- Host-appropriate setup:
  - PowerPoint: create a new presentation or a new blank slide by default, set slide size, and clear only generated shapes if working inside an existing deck.
  - Excel: create or use a worksheet drawing area, clear shapes in the target area, optionally adjust cells behind it.
  - Word: create or use a document page/canvas, clear generated shapes by name prefix where possible.
- Canvas cleanup code.
- Canvas size setup code.
- Background creation code.
- Code for every visible object, in back-to-front order.
- `Shapes.AddPicture` code for each preserved local crop in hybrid mode, with `LinkToFile:=msoFalse` and `SaveWithDocument:=msoTrue` where the host supports it.
- Named objects, for example `shp.Name = "Header_Background"`.
- Shared color helpers using `RGB(...)` or small helper functions.
- Comments that label each visual region.
- TextBox creation for legible image text.
- Z-ordering when overlap matters, using `ZOrder` or creation order.
- A stable generated-object prefix such as `AITVBA_` for shape names and cleanup.

Use Office APIs conservatively:

- `Shapes.AddShape`, `Shapes.AddTextbox`, `Shapes.AddLine`, `Shapes.BuildFreeform`, `FreeformBuilder`.
- `Shapes.AddPicture` for preserved raster crops in hybrid mode.
- `Fill.ForeColor.RGB`, `Fill.Transparency`, `Line.ForeColor.RGB`, `Line.Weight`.
- `TextFrame` / `TextFrame2` font settings as appropriate for the host.
- `Shadow`, `GradientStops`, `PresetGradient`, or layered translucent shapes to approximate effects.

### 5. Presentation Materialization

For the default presentation target, after generating the VBA:

1. Save the generated code to a local `.bas` or `.vba` file.
2. If hybrid mode is used, save preserved crop assets locally. Use `scripts/crop_preserved_elements.py` when crop coordinates are known.
3. Run `scripts/vba_lint.py` against the generated VBA file.
4. Detect the available local presentation environment:
   - Use `scripts/detect_office_environment.py` when possible.
5. Try to create/open the editable deck automatically when a supported path exists:
   - On Windows with Microsoft PowerPoint installed, use `scripts/run_powerpoint_vba_windows.ps1`.
   - On macOS with PowerPoint installed, use `scripts/run_powerpoint_vba_macos.py` when applicable.
   - With WPS only, do not assume automatic macro injection. Open WPS when possible, provide the `.bas` path, and give WPS-specific manual run steps.
   - If automation is blocked by Office macro security, AppleScript permissions, PowerPoint Trust Center, unsupported VBA injection, or WPS macro limitations, open the best available presentation app when possible and provide the exact macro file path plus concise manual steps.
6. Confirm in the final response which app was detected, which app was opened, whether the macro executed, which preserved assets were used, and whether a deck file was created. If it is not open, state the blocker directly.

The final answer must not imply that a PNG preview is the editable deliverable. It must distinguish between source image, preview screenshot, VBA source file, editable deck file, WPS-compatible file, and any non-VBA PPTX fallback.

### 6. Mandatory Self-Checks

Perform these self-check rounds before final output:

1. Structure check: verify no major object, region, text block, connector, or layer is omitted; verify back-to-front order.
2. Editability check: verify that text, arrows, boxes, connectors, and other structural elements were recreated as editable Shapes unless explicitly placed in a preserved crop.
3. Preservation check: if hybrid mode is used, verify that every preserved raster element has a crop asset, asset path, `AddPicture` placement, shape name, and reason for preservation.
4. Code check: verify VBA syntax, variable declarations, object references, Office host APIs, cleanup logic, shape names, generated-file path, and preserved asset paths. Use `scripts/vba_lint.py` as a lightweight smoke check when code is available as a file or can be copied into a temporary file; do not overstate it as a full VBA compiler.
5. Materialization check: verify that the macro was saved, that preserved assets were saved when needed, that the available office runtime was detected, that the appropriate PowerPoint/WPS path was attempted, and that the output deck is open or a concrete blocker is reported.
6. Visual check (render-verify 强制闭环): 不允许停留在"人工目测"。这是保证还原度的关键环节，必须按下面步骤反复迭代直到达标或触发硬性上限。

   每一轮的固定动作：
   1. 运行最新的 VBA 宏，生成 / 更新本地 deck。
   2. 把当前幻灯片导出成 PNG（Windows 用 `Slide.Export`；macOS 用 AppleScript / PowerPoint 导出）。
   3. 调用 `scripts/compare_images.py` 与原图比对，输出整体 SSIM、整体像素差比例，以及差异最大的若干局部区域（按 bbox 列出）。
   4. 把每个高差异区域映射回 manifest 行，明确标注 "该区域对应 E03/R01/L05 的位置/颜色/尺寸/箭头终点偏差"，再针对性修改 manifest（不要漫无目的地全部重写）。
   5. 重新生成 Pass 2 VBA，回到第 1 步。

   通过条件（同时满足才能交付）：
   - 整体 SSIM ≥ 0.90，且
   - 没有任何高差异区域占画布面积 > 5%，且
   - manifest 里所有 `endpoints` 字段对应的箭头在渲染图里目测连接到了正确的元素边界。

   迭代上限：最多 6 轮。每轮必须在最终报告里附一行"第 N 轮：SSIM=0.xx，最大差异区域=…，本轮修改了 …"。

   终止策略：
   - 若 6 轮内达标，停止迭代，输出 deck 和报告。
   - 若 6 轮仍未达标，但每轮 SSIM 持续上升（趋势明显），允许再扩 3 轮（共 9 轮上限),并在报告中说明仍在收敛。
   - 若连续 2 轮 SSIM 不再上升或反而下降,立即停止,在报告里列出仍不达标的具体 manifest 行,而不是默默交付一个差的版本。
   - 用户在对话中显式说"够了 / 先这样 / 不用再迭代"时，提前停止并说明当前最终 SSIM。

   报告语言：每轮的差异分析、区域定位、修改说明都用简体中文（遵循文件顶部 Output Language 规则）。

   仅在本机确实无法跑宏时（无 Office、Trust Center 拦截、沙箱环境），才允许跳过该闭环并降级为标注清楚的"人工目测估计"，且必须在报告中说明为什么没跑、用户在哪里能自行运行该闭环。

If a check finds a gap, revise the VBA before responding. Mention unresolved uncertainty only after revising what can be revised.

### 7. Output Requirements

Use the structure in `references/output-template.md`:

- Short explanation of the reconstruction strategy.
- Artifact paths for the VBA file and deck file/open state.
- Preserved raster asset paths and a short editable-vs-preserved element summary when hybrid mode is used.
- Complete VBA code block.
- Brief steps for running the VBA.
- Self-check results.
- Possible visual errors and why they remain.
- Next-round optimization suggestions, especially if the user can provide a rendered screenshot.

Always include a short "How to run" section. Keep it host-specific and practical.

### 8. Error and Fidelity Policy

Do not promise perfect pixel-level reconstruction for complex photos, textures, real people, dense raster effects, or highly complex gradients. Explain that Office Shapes can approximate these but cannot fully reproduce raster detail.

Unless the user explicitly allows full-image insertion, do not place the original image on the slide/sheet/document and claim it is reconstructed. This restriction does not prohibit small local preserved crops in hybrid mode, as long as they are clearly labeled. If using a full background image is appropriate, label it clearly as the background-assisted option and keep editable overlays separate.

Never replace the editable deliverable with a PNG, screenshot, or generated raster image. A raster image may be used only for source analysis, validation, optional preview, or explicitly labeled local preserved elements in hybrid mode.

For "1:1" requests:

- Increase detail and object count.
- Preserve aspect ratio, palette, and layer order.
- Avoid inventing hidden or unreadable details.
- Mark uncertain text or shapes with comments in the VBA.

## Bundled Resources

- `scripts/vba_lint.py`: lightweight checks for generated VBA.
- `scripts/detect_office_environment.py`: detect likely Microsoft PowerPoint/WPS runtimes and recommend a materialization path.
- `scripts/crop_preserved_elements.py`: crop local preserved raster elements from the source image for hybrid reconstruction.
- `scripts/run_powerpoint_vba_macos.py`: attempt to run generated PowerPoint VBA through macOS PowerPoint automation and bring PowerPoint forward.
- `scripts/run_powerpoint_vba_windows.ps1`: attempt to import and run generated VBA through Windows Microsoft PowerPoint COM automation.
- `scripts/compare_images.py`: compare source image and rendered VBA screenshot.
- `scripts/coordinate_helper.py`: convert image pixels to Office point coordinates.
- `references/vba-shape-patterns.md`: common Office Shapes patterns.
- `references/reconstruction-checklist.md`: detailed review checklist.
- `references/output-template.md`: final response template.
