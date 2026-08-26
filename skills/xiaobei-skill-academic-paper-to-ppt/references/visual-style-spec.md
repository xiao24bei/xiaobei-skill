# Visual Style Specification

## Style selection rule

Use this style priority:

1. Current user instructions.
2. User-provided PPT/PPTX template.
3. Default academic preset below.

When a PPT/PPTX template is provided, analyze it first and use it as the main visual reference for all `$imagegen` visual preview and slide-page generation. Extract:

- Slide aspect ratio and page families
- Main colors and accent colors
- Font feel and title hierarchy
- Cover, directory, section divider, content, chart/table, summary, and closing-page patterns
- Spacing, margins, line systems, color blocks, image treatments, and recurring motifs
- Typical page density and the way the template separates section pages from content pages

Do not copy placeholder text from the template as factual content. The template controls visual style and structure only; the uploaded source document controls factual content.

If no PPT/PPTX template is provided, use this default academic preset for all pages generated with `$imagegen`:

- Theme: formal academic thesis defense
- Target field: science and engineering paper presentation
- Tone: professional, stable, clean, concise
- Main colors:
  - light gray background
  - dark navy blue headings
  - cyan blue accents
  - white content cards
- Layout:
  - 16:9 widescreen
  - strong page title hierarchy
  - consistent header/footer system
  - generous whitespace
  - content-dense but not cluttered
- Visual components:
  - academic charts
  - clean tables
  - modular process diagrams
  - technical route diagrams
  - model architecture blocks
  - experiment result dashboards
  - conclusion cards
- Typography:
  - clean sans-serif
  - strong title/subtitle contrast
  - clear Chinese text hierarchy
- Avoid:
  - cartoon style
  - overly colorful design
  - decorative backgrounds
  - low-density business pitch style
  - casual or playful elements

For the visual preview stage, use `$imagegen` to generate a single multi-page thumbnail board showing the full deck style.

For the final PPT stage, use `$imagegen` to generate one full-slide image per page and keep visual consistency.

Do not produce this visual style as text-only design advice, SVG/vector art, HTML/CSS/canvas output, screenshots, or manually composed placeholders. The required output for Stage 2 and Stage 3 is bitmap imagery produced through `$imagegen`.
