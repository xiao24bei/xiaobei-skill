---
name: xiaobei-skill-academic-paper-to-ppt
description: "Use for 小北在读研 / XiaoBei academic paper-to-PPT workflows: generate thesis-defense or project-report decks from uploaded papers, literature, reports, technical documents, and optional PPTX visual templates. Produces a source-grounded outline, an imagegen visual scheme, one complete bitmap per slide, an imagegen provenance manifest, and an image-only PPTX. Do not use when the user needs editable slide text or native editable charts."
---

# 小北在读研 · Academic Paper to PPT

## Core Behavior

When this skill is triggered, execute the workflow in this exact order:

1. Analyze the uploaded source document
2. If the user provides a PPT/PPTX template, analyze the template's visual style and page structure
3. Create `ppt_outline.md`
4. Generate one complete visual preview scheme based strictly on `ppt_outline.md` using `$imagegen`, then immediately display that image inline in the chat
5. Generate every slide page as a full-page image using `$imagegen`; immediately display each completed page image inline in the chat before waiting for the whole deck
6. Create `imagegen_manifest.json` recording the selected `$imagegen` generation method and every generated asset
7. Assemble the generated slide images into a `.pptx` file where each image fills one complete slide
8. Validate that required output files exist and match the imagegen manifest

Never skip stages unless the user explicitly requests only one stage.

Strict image-only contract:
- A final slide image means the whole finished PPT page is baked into one bitmap, including Chinese title, body text, labels, diagrams, tables, and page visual design.
- The final `.pptx` must contain one full-slide image per slide and no editable visible text boxes.
- Background-only slide images are invalid. Do not generate text-free templates, blank layout backgrounds, placeholder panels, blank areas for later text, or editable-text overlays.
- Prompts for final slide images must not include phrases such as `text-free`, `no readable words`, `background template`, `blank areas`, `editable Chinese text`, `text to be added later`, or equivalent wording.
- If `$imagegen` cannot generate complete page images with the required text, retry with shorter full-page prompts or switch to the other allowed `$imagegen` mode. If complete page images still cannot be produced, stop at image generation and report the failed stage. Do not assemble a fallback PPT.

Chat-visible image contract:
- Apply this contract to both built-in `image_gen` and CLI generation. Changing the image backend must not make generated images invisible to the user.
- Stage 2 produces one visual-scheme image. As soon as it exists, show it in the chat with local-image Markdown using its absolute path, for example `![PPT 视觉方案](/absolute/path/visual_preview.png)`.
- During Stage 3, show each completed slide image individually as soon as it is available. Label it briefly as `第 X/N 页`, then render that page with its absolute local path. A text-only progress message such as “已生成 6 页” does not satisfy this requirement.
- Do not wait until all pages finish before showing them. Do not replace per-page visibility with only a contact sheet, folder path, file count, or final PPTX link.
- In CLI mode, run generation in a monitorable process and poll for newly completed files. After each poll, display every newly completed `slide_XX.png` individually before continuing to wait. Limited concurrency is allowed, but the chat must expose each completed page rather than hiding the batch until completion.
- In built-in mode, the tool result may render automatically; still add the brief `第 X/N 页` label and preserve/copy the selected image into the workspace for manifesting and PPT assembly.
- After all pages are visible, optionally add a complete contact sheet for convenient overview, then provide the final PPTX.

Required image-generation dependency:
- The visual preview and full-slide generation stages must load and follow the installed `$imagegen` skill.
- Use built-in `image_gen` first when the user does not specify a generation mode.
- Use CLI mode when the user explicitly requests it or when a real built-in attempt is unavailable or fails, but only through the runtime, endpoint, and credential mechanism authorized in the current user's `$imagegen` workflow.
- Do not hardcode a maintainer's workstation path, account name, private endpoint, or credential source. Never display, log, or persist credentials.
- In CLI mode, write generated image files under `outputs/imagegen/` and record method `imagegen_cli_fallback` in `imagegen_manifest.json`.
- Use a Python environment with the packages in `requirements.txt` for PowerPoint assembly and validation.
- Do not replace `$imagegen` with textual visual suggestions, SVG/vector/code-native drawings, HTML/CSS/canvas mockups, manual diagrams, or a generic image-generation placeholder.
- When built-in mode is used, copy the selected generated files from `$CODEX_HOME/generated_images/...` or the workspace into `outputs/imagegen/generated_slides/`, and record method `built_in_image_gen` in `imagegen_manifest.json`.
- If an explicitly selected CLI attempt fails to generate real files, try built-in `image_gen` unless the user asked for CLI only. If the default built-in attempt fails, use CLI only when the current environment authorizes and supports that fallback.
- If neither CLI fallback nor built-in `$imagegen` can generate real image files, stop at that stage and report the incomplete stage. Do not assemble a PPT from non-imagegen substitutes.
- A visual preview generated with `$imagegen` is not enough. Every final slide image recorded in `imagegen_manifest.json` must also be generated with `$imagegen`.
- If CLI fallback fails only for long Chinese prompts, that is not permission to generate text-free backgrounds. First shorten each prompt while preserving the requirement that the final image contains the slide title and key visible text. Then try built-in `image_gen` if needed. Stop rather than replacing text with editable PowerPoint objects.

Do not expand the full Markdown outline in the chat.
Do not output long analysis.
Do not invent information that is not in the uploaded document.
Do not cite or use external sources unless the user explicitly asks for external supplementation.

## Visual Style Priority

Use this priority order for the deck's visual style:

1. Explicit style instructions from the user in the current request.
2. A user-provided PPT/PPTX template, when available.
3. The default academic visual preset in `references/visual-style-spec.md`.

If a template is provided, treat it as a visual and structural reference only:
- Extract its slide size, color palette, typography, title hierarchy, section-page design, content-page design, chart/table treatment, spacing, and visual motifs.
- Prefer the template's cover, directory, section divider, body, chart/table, summary, and closing-page patterns when generating the visual preview and full-slide images.
- Use the template's topic distribution to decide the outline's page grouping and content density, but do not copy its placeholder text as source content.
- Do not force the generated PPT to use every page from the template. Compress or expand the template's page families into a source-appropriate academic or project-report narrative.
- Do not add market analysis, team information, intellectual property, or implementation-condition pages unless those topics are present in the uploaded source document or explicitly requested by the user.
- If the template conflicts with source-document fidelity, source-document fidelity wins.

If no template is provided, use the default academic preset from `references/visual-style-spec.md`.

## Stage 1: Document Parsing and Markdown PPT Outline

Input:
- A user-uploaded literature paper, academic article, research report, technical document, PDF, DOCX, Markdown file, or plain text document.
- Optional: a user-provided PPT/PPTX template used only for visual style, page-family structure, and outline density guidance.

Output:
- Create or update a file named exactly:

`ppt_outline.md`

The file must contain a complete Markdown PPT outline of about 12-18 pages by default. Use about 10 pages only when the source document is short or the user explicitly asks for a compact deck.

Strict rules:
1. All content must be strictly based on the uploaded document.
2. Do not fabricate information.
3. If information is not explicitly stated in the original document, write: `原文未明确说明`.
4. Do not use external sources.
5. Keep the structure suitable for academic presentation, course presentation, thesis sharing, or project introduction.
6. Each slide must include:
   - 页面标题
   - 页面类型
   - 核心内容
   - 内容展开
   - 建议展示方式
   - 视觉生成要点
   - 原文依据
7. Each slide must contain enough detail to support later PPT generation, especially enough visual and content cues for `$imagegen`.
8. Prioritize extracting important data, experiments, methods, charts, formulas, models, algorithms, case studies, results, limitations, and conclusion evidence from the original document.
9. If the uploaded document is long, reorganize the content into a presentation logic instead of copying the original table of contents.
10. If a PPT template is provided, use its section rhythm and page families to guide page distribution and content density, while keeping all factual content grounded in the source document.
11. The output must be usable by AI PPT tools, Codex, or visual design tools.

Use the structure in `references/outline-template.md`.

After Stage 1, reply briefly only:
- 已生成 ppt_outline.md
- 文件中包含约 12-18 页 PPT 大纲
- 内容严格基于上传文档

Do not paste the full outline into chat.

## Stage 2: Visual Style Generation

Input:
- `ppt_outline.md`

Output:
- One full visual preview image or multi-page thumbnail contact sheet representing the complete PPT visual scheme, saved as `outputs/imagegen/visual_preview.png` for either generation mode.

Must use `$imagegen`.

Execution rules:
1. Load and follow `$imagegen` before generating the visual preview.
2. Create `outputs/imagegen/` before generation. Use built-in `image_gen` first unless the user explicitly selected CLI.
3. In built-in mode, generate the preview, copy the selected result to `outputs/imagegen/visual_preview.png`, and immediately display that image inline in the chat.
4. If the built-in tool is unavailable or a real call fails to return a valid image, follow the installed `$imagegen` skill's authorized CLI fallback, save the output to `outputs/imagegen/visual_preview.png`, and immediately display that image inline in the chat.
5. If the user explicitly selected CLI, start with the CLI path above. If it fails and the request was not CLI-only, try built-in mode.
6. Treat the preview as a generated bitmap asset, not as text, Markdown, code, SVG, or HTML.

Visual requirements:
- Follow the Visual Style Priority above.
- If a PPT/PPTX template is provided, reproduce its overall visual language in a new deck: palette, typography feel, section-page composition, content-page rhythm, spacing, title treatment, chart/table treatment, and recurring visual motifs.
- If no template is provided, use the default academic visual preset in `references/visual-style-spec.md`.
- Keep the result formal, professional, academic, clean, stable, and suitable for thesis defense, science and engineering papers, or project reporting.
- Use a unified layout system.
- Keep content dense but readable.
- Maintain clear hierarchy.
- Maintain consistent chart, process diagram, table, and framework diagram style.
- Cover pages including:
  - 封面
  - 目录
  - 背景
  - 方法
  - 实验
  - 结果
  - 总结
  - 展望
  - 致谢

Important:
- Do not output text suggestions.
- Do not output analysis.
- Do not output code.
- Do not output Markdown descriptions as a substitute for the generated visual.
- After the bitmap exists, use Markdown image embedding only to make that visual result visible in the chat.
- The result should look like a full PPT page preview board / multi-page thumbnail grid.

## Stage 3: Full PPT Generation

Input:
- `ppt_outline.md`
- Visual scheme from Stage 2 generated by `$imagegen`
- Any generated slide images

Output:
- Full-slide images in `outputs/imagegen/generated_slides/` for either generation mode
- `imagegen_manifest.json`
- A complete PowerPoint file, for example:

`defense_presentation.pptx`

Generation rules:
1. Generate every slide as a full-page 16:9 image.
   - Every final bitmap must already contain the slide's visible Chinese title, core content, labels, and layout.
   - The PPT assembly step may not add or edit visible slide text.
2. Use `$imagegen` and the selected method to generate each slide page: one built-in `image_gen` call per page in default built-in mode, or a monitorable `scripts/image_gen.py generate-batch` process in CLI mode.
3. Each generated image must represent one complete PPT slide.
4. The final PPT must use generated images as full-slide backgrounds/pages.
5. Do not rely on editable text boxes.
6. Do not worry about text editability.
7. Do not worry about minor OCR-like issues in generated image text unless the user specifically asks for text-perfect slides.
8. Insert each generated image into a PowerPoint slide so that it fully covers the slide canvas.
9. Maintain the same formal academic visual style across all pages.
10. The PPT should follow the page structure in `ppt_outline.md` and the Visual Style Priority above.
11. Do not create slide images through HTML/CSS screenshots, SVG export, canvas drawing, manually composed placeholders, or any route that bypasses `$imagegen`.
12. If `$imagegen` cannot generate the slide images, stop and report the incomplete stage instead of assembling a PPT from non-imagegen placeholders.
13. Generate one final page image per slide in `ppt_outline.md`; the number of image files recorded in `imagegen_manifest.json` must equal the number of outline slides.
14. Name images with zero-padded slide numbers, such as `slide_01.png`, `slide_02.png`, and so on.
15. Do not accept or manifest final prompts that ask for background-only or text-free pages. Any final prompt containing `text-free`, `no readable words`, `background template`, `blank areas`, `editable text`, `editable Chinese text`, or `text to be added later` must be treated as a failed workflow.
16. Immediately after each page image becomes available, display that individual image in the chat with `第 X/N 页` and an absolute-path Markdown image. Do not wait for the full deck, and do not substitute only a contact sheet.

For CLI slide generation, create `outputs/imagegen/prompts.jsonl` with one JSON object per slide prompt, then use the installed `$imagegen` skill's supported batch command and credential mechanism. Write results to `outputs/imagegen/generated_slides/`, keep the process monitorable, poll for new `slide_XX.png` files, and display each new page inline before continuing to wait. Do not copy a maintainer-specific CLI path or credential configuration into the project.

Create `imagegen_manifest.json` after generating the preview and slide images. It must include:

```json
{
  "generation_method": "imagegen_cli_fallback",
  "visual_preview": "outputs/imagegen/visual_preview.png",
  "image_output_dir": "outputs/imagegen/generated_slides",
  "model": "gpt-image-2",
  "quality": "medium",
  "size": "1536x864",
  "slides": [
    {
      "slide": 1,
      "file": "outputs/imagegen/generated_slides/slide_01.png",
      "method": "imagegen_cli_fallback",
      "prompt": "full slide prompt used for this page"
    }
  ]
}
```

Use `scripts/build_ppt_from_images.py` to assemble the images into a `.pptx`.

The script should:
- Accept an input directory containing slide images
- Sort images by slide number
- Create a 16:9 PowerPoint deck
- Add each image as a full-slide image
- Save the output as `defense_presentation.pptx`
- Preserve slide order
- Avoid margins or cropping unless needed to fill the slide

Example command:

```bash
python scripts/build_ppt_from_images.py \
  --image-dir ./outputs/imagegen/generated_slides \
  --output ./defense_presentation.pptx
```

Use `--image-dir ./outputs/imagegen/generated_slides` for both generation modes.

Validate final outputs with:

```bash
python scripts/validate_outputs.py --base-dir .
```

The validation must fail if:
- The visual preview file recorded in `imagegen_manifest.json` is missing.
- `imagegen_manifest.json` is missing.
- Any manifest entry does not use an allowed `$imagegen` method: `built_in_image_gen` or `imagegen_cli_fallback`.
- The number of slide images does not equal the number of slides in `ppt_outline.md`.
- Any expected `slide_XX.*` image recorded in `imagegen_manifest.json` is missing.
- Any manifest prompt describes a text-free/background/template page rather than a finished slide.
- The final PPTX contains editable visible text, multiple picture objects on a slide, or anything other than one full-slide image per slide.

## Final Response Rules

After completing the full workflow, reply briefly only with:
- 已生成 ppt_outline.md
- 已通过 imagegen 来源校验，并注明使用 `built_in_image_gen` 或 `imagegen_cli_fallback`
- 已生成整套 PPT 页面视觉图
- 已生成 defense_presentation.pptx
- PPT 使用生成图片作为完整页面
- 直接展示整套联系表（若已生成），并提供最终 PPTX 的可点击链接

Do not output:
- Full outline content
- Long explanation
- Internal reasoning
- Irrelevant suggestions
