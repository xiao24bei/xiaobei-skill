# Workflow Rules

## End-to-end mode

When the user uploads a source document and asks to generate the full PPT, run:

1. Read and analyze source document
2. If a PPT/PPTX template is provided, analyze its visual style and page-family structure
3. Generate `ppt_outline.md`
4. Load and follow `$imagegen`
5. Generate one visual preview board with `$imagegen` and immediately display it inline in the chat
6. Generate full-slide images with `$imagegen` and immediately display each completed page image individually in the chat
7. Assemble full-slide images into `defense_presentation.pptx`
8. Validate outputs
9. Reply briefly with file status

Do not replace `$imagegen` with textual suggestions, SVG/vector/code-native drawings, HTML/CSS/canvas mockups, screenshots, manually composed placeholder slides, or a generic image-generation tool. Stop and report the failed stage only if both the default built-in path and any CLI fallback authorized in the current environment cannot produce real images.

Use built-in `image_gen` first by default. Use the configured CLI when the user explicitly requests it or when a real built-in call is unavailable or fails to return a valid image. CLI changes the generation backend only; it must preserve the same chat-visible preview experience.

For every mode, show the single visual-scheme image as soon as it exists. During slide generation, label and show each page as `第 X/N 页` with an absolute-path Markdown image immediately after that file completes. Do not replace this with a text-only progress count, folder path, contact sheet, or final PPTX link. A contact sheet may be added only after individual page visibility has been satisfied.

Final slide images must be complete finished PPT pages. Each bitmap must already contain all visible slide text, labels, diagrams, tables, and design elements. Never generate text-free backgrounds, blank page templates, placeholder panels, or images that leave space for editable PowerPoint text to be added later. If complete page images cannot be generated with `$imagegen`, stop and report the image-generation failure instead of assembling a fallback PPT.

## Visual style priority

Use this priority order:

1. Explicit style instructions from the user in the current request.
2. User-provided PPT/PPTX template.
3. Default academic style in `visual-style-spec.md`.

When a template is provided:
- Use it as a visual and structural reference, not a factual source.
- Extract palette, typography, slide size, section dividers, body-page layout, chart/table treatment, spacing, and recurring visual motifs.
- Use its topic distribution to guide outline density and section rhythm, but do not copy unrelated template topics into a paper deck.
- Do not force market, team, intellectual-property, or implementation pages unless the source document contains those topics or the user explicitly requests them.

When no template is provided, use the default academic style in `visual-style-spec.md`.

## Stage-only mode

If the user explicitly asks only for the outline:
- Analyze any provided PPT/PPTX template first, then stop after creating `ppt_outline.md`

If the user explicitly asks only for the visual scheme:
- Require or use an existing PPT outline
- Analyze any provided PPT/PPTX template before style generation
- Load and follow `$imagegen`
- Generate only the visual preview board with `$imagegen`
- Immediately display that preview image inline in the chat

If the user explicitly asks only for PPT assembly:
- Require existing slide images
- Use `build_ppt_from_images.py`

If the user explicitly asks to generate slide images:
- Load and follow `$imagegen`
- Generate each full-slide bitmap image with `$imagegen`
- Immediately display each completed page image individually in the chat before waiting for all pages
- Save the generated images under `outputs/imagegen/generated_slides/`
- Do not generate text-free backgrounds or blank templates. The saved images must already be finished slide pages with all visible text baked into the image.

## Imagegen provenance

Create `imagegen_manifest.json` for every full workflow or slide-image generation workflow. It must record:
- `generation_method`: `built_in_image_gen` or `imagegen_cli_fallback`.
- `visual_preview`: path to the generated preview image.
- `image_output_dir`: the directory containing final slide images.
- `slides`: one object per slide, with slide number, file path, `method`, and the prompt used for that page.

This manifest is a required execution record. Do not create it for non-imagegen assets. If any slide image was not produced through `$imagegen` built-in mode or `$imagegen` CLI fallback, stop and report the failed stage instead of assembling the PPT.

Manifest prompts for final slide images must describe finished slides, not backgrounds. Treat prompts containing `text-free`, `no readable words`, `background template`, `blank areas`, `editable text`, `editable Chinese text`, or `text to be added later` as invalid final-slide prompts.

## Source fidelity

Always distinguish:
- 原文直接说明
- 基于原文归纳
- 原文未明确说明

Never add external facts unless the user explicitly requests external supplementation.

## Output discipline

Never paste the complete `ppt_outline.md` in chat unless the user explicitly asks to view it.
