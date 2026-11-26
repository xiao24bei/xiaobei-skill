---
name: xiaobei-skill-academic-paper-to-ppt
description: "Convert academic papers (.pdf/.docx), theses, grant proposals, and research reports into professor-convincing, publication-grade 16:9 Chinese Navy-White-Crimson (学院蓝白红) strict image-only PPT decks for lab group meetings (组会汇报/文献精读), thesis proposals (开题答辩), grant reports (基金结题), and graduation defenses (毕业答辩). Follows a strict 3-stage pipeline with a Stage-1 Advisor Rigor Gate: (1) Full-fidelity paper/OMML-equation/table parsing, real figure cropping/stitching & 4-ledger high-density outline (ppt_outline.md), (2) Locked visual scheme preview (visual_preview.png), and (3) Full-slide $imagegen generation with lossless real paper figure embedding, complete paper equations, multi-baseline trade-off attribution, Speaker Notes, and automated content+visual validation."
metadata:
  version: "5.0"
---

# Academic Paper To PPT (v5.0 — 导师五连问闭环 · 四大前置台账 · 拒绝口号空壳版)

Turn research papers (`.pdf` or `.docx`), proposals, and theses into high-rigor, professor-convincing 16:9 Chinese academic PPT decks (`defense_presentation.pptx`) that **thoroughly explain the paper's problem positioning, mathematical mechanisms, training/inference protocols, multi-baseline evidence, and causal trade-offs**—eliminating shallow 40-word slogan flashcards, dropped equations, and generic AI filler.

## Core Non-Negotiable Contracts

1. **Clear 3-Stage Pipeline + Stage-1 Depth Gate (Strict Execution Order)**:
   - **Stage 1（全量解析、四大证据台账与高密度大纲阶段）**：运行 `scripts/extract_paper_assets.py` 提取全文、OMML/LaTeX 编号公式、结构化数据表及高清原图/拼合组图 → 建立**四大前置证据台账**（原图台账 + 编号公式上屏台账 + 既有基线卡点与训练/推理协议台账 + 多维实证与权衡台账）→ 编写高含金量 `outputs/ppt_outline.md`，并**强制通过 `scripts/validate_outputs.py --base-dir ./outputs --outline-only` 深度门禁后**再进入 Stage 2！
   - **Stage 2（视觉方案阶段）**：锁定「学院蓝白红」`Style Bible`（体系 A 或 体系 B，或用户模板）→ 用 `$imagegen` 生成并审校全套缩略图视觉方案 `outputs/imagegen/visual_preview.png`；
   - **Stage 3（高密度生图、真图精嵌与双重质检阶段）**：逐页用 `$imagegen` 生成完整高清页图 → 对引用论文原图的页面用 `scripts/embed_paper_figure.py` **像素级无损精嵌论文真图（支持单图 `L11a/L11b` 与双图并列 `L11c`）** → 8 维导师级质检 → 装配带演讲备注的严格全图 `outputs/defense_presentation.pptx` 并通过 `validate_outputs.py`。
   - Never skip stages unless the user explicitly requests a single stage (or asks to pause after each stage for manual review).

2. **Pass the "Advisor's 5-Question Interrogation" (导师五连问硬核内容标准 — 拒绝提词卡式空壳！)**:
   Every deck MUST thoroughly answer the 5 questions a thesis advisor or reviewer immediately asks:
   - **Q1 场景约束与物理/数学问题本质**：具体应用工况有什么特殊约束？核心退化/矛盾机理是什么？问题如何形式化建模（必须给出完整建模方程、各符号定义与映射实现，如 $\Psi(\cdot)$ 的具体网络结构）？
   - **Q2 既有路线谱系与本文创新定位（完整论文必设独立 `1.2` 页！）**：**严禁跳过相关工作直接讲自身网络！** 必须具名梳理 3–4 条代表性技术路线及关键基线（如确定性回归 `Retinexformer`、像素扩散 `Diff-Retinex`、非扩散细节通路 `DiffLight`、二元潜在扩散 `Reti-Diff`），明确各自做对了什么、**卡在哪个结构瓶颈**，从而导出本文不可替代的破局闭环。
   - **Q3 模块张量流、完整编号公式与训练/推理边界**：每个方法模块必须讲清**输入/输出张量及维度变换、核心算子与论文编号公式本体、参数冻结 vs 联合更新策略、阶段损失函数各项物理分工与权重系数（如 $\lambda_2=1.5$）、以及测试推理如何脱离真值参考图独立运行**。
   - **Q4 多路线实证全貌、指标权衡与消融分工归因**：
     - 对比实验必须展示**跨技术路线的代表基线矩阵（≥ 4–6 个代表方法）与全部核心互补指标**（**严禁把十几个对比方法删剩 2 个点连线！**）；
     - 当不同指标或不同数据集出现分化（如感知指标 FID/BIQE 最优而逐像素 PSNR 存在取舍，或工业实测集全指标反超）时，必须给出**物理/统计层面的深度归因**；
     - 消融实验必须给出**各变体移除后的多指标退化幅度与物理崩溃原因**，以及逐步叠加模块时**哪一步拉升结构指标（如 SSIM）、哪一步优化感知/边缘指标**（**严禁只画单列指标配“提升特征一致性 / 增强全局建模能力 / 进一步提升重建精度”等贴到任何论文都能用的 AI 万能套话！**）。
   - **Q5 真实工程边界与可验证展望**：交代推理延迟/算力开销、极端失效场景（如强镜面反光）、下游任务闭环验证方案及论文明确提出的技术演进路径（如蒸馏加速、偏振/跨模态融合）。

3. **Formula Integrity, Substantive Density & Zero-Echo Law (公式必上屏 · 高密度干货 · 零复读零注水)**:
   - **Cited Equations MUST Appear On-Screen**: Whenever a slide cites paper equations (`式(1)`, `式(2)–(5)`, `式(6)–(8)`, `式(9)–(13)`) in its rationale or footer, **those exact paper equations MUST be written explicitly on the slide canvas**! Never put `式(6)-(8)` in the footer while rendering zero equations on the slide!
   - **STRICT BAN on Textbook Metric Formula Padding**: **NEVER** pad experimental or ablation slides with textbook definitions of generic evaluation metrics (`PSNR = 10 log10(MAX^2/MSE)`, `FID = ||μ_r - μ_g||^2 + Tr(...)`, `SSIM = ...`) unless proposing a new metric is the paper's actual contribution! Formula space belongs strictly to the paper's own modeling, operator, and loss equations.
   - **Substantive Per-Slide Density (140–260 Chinese chars + Formulas/Data)**: Every core content slide (`3 .. N-1`, excluding Cover/TOC/Divider/Thanks) MUST carry **140–260 Chinese characters of structured academic substance** in `### 上屏文案` (organized cleanly into `1 句顶栏定调结论 + 2–3 张各司其职的结构化剖析卡（每卡 2–3 条紧凑机制/公式/数据要点）+ 1 条底部边界/启示总结条`). **Slides with < 100 chars of bare 12-word slogans are strictly rejected!**
   - **Zero Verbatim Echo**: Slide title, top thesis banner, body cards, and bottom takeaway bar MUST each carry distinct information (`论点标题 → 全局机制横幅 → 具体公式/张量/数据拆解卡 → 适用边界底栏`). **NEVER repeat the banner sentence inside a card, and NEVER render the same formula twice on one slide!**

4. **Real Paper Figure Integration (No Toy Hallucinations!)**:
   - Extract/crop real architecture diagrams, module diagrams, qualitative comparisons, and multi-panel plots into `outputs/source_assets/figures/` in Stage 1 (for `.docx` manuscripts, `extract_paper_assets.py` automatically extracts embedded figures AND stitches multi-panel `(a)..(h)` comparison tables into `composite_panel_XX.png`).
   - Use `L11a` (left figure + right 3-card deep analysis), `L11b` (top wide figure + bottom 3-column deep analysis), or `L11c` (top dual side-by-side figures + bottom 3-column architecture/formula/mechanism breakdown), compositing real figures losslessly via `scripts/embed_paper_figure.py`.

5. **Strict Image-Only Slide Canvas, Provenance & Zero Style Cross-Contamination**:
   - Every final slide in `defense_presentation.pptx` contains **exactly one unified 16:9 bitmap** and **zero editable visible text boxes**, with oral scripts (`演讲备注`, 130–220 chars/slide) attached to PPTX Speaker Notes (`ppt/notesSlides/`).
   - Load `$CODEX_HOME/skills/.system/imagegen/SKILL.md` before Stage 2 and Stage 3 when available (`built_in_image_gen` primary; `imagegen_cli_fallback` via `gpt-image-2` on transient/bounded failure; never print `OPENAI_API_KEY`). Never use HTML/SVG/python-pptx text shapes as slide substitutes.
   - **Zero Style Leaks**: Never copy sample institution names (`南京审计大学`, `北京交通大学`, `NIH/NIDCD`), cross-discipline motifs (brain/neurons), or Style A campus-building footer line-art onto a Style B deck; keep bottom-right page badge color/geometry 100% identical across ALL slides.

---

## Visual Style System (默认「学院蓝白红」双核体系)

Priority: (1) Current user instructions → (2) User-provided PPT/PPTX template (visual system only) → (3) Default **学院蓝白红（Navy-White-Crimson）** in `references/visual-style-spec.md`:
- **体系 A「学院深蓝红·答辩版」**（开题/毕业/基金答辩、社科/经管/工程综合报告默认；锚点 `style-a-*.png`）：深海蓝 `#0B2954` 顶底通栏 + 实心深蓝小标题条白底卡片 + 右置红顶「结论」锚定卡 + 大数字章节过渡页 + 居中白字红线页码。
- **体系 B「白蓝红科研·理工生医版」**（AI/计算机/生医/物理/工程组会精读与技术答辩默认；锚点 `style-b-*.png`, `style-c-engineering-pipeline.png`）：纯白底 + 编号宝蓝 `#105091` 断言式主标题与细蓝线 + 浅冰蓝底红蓝双色核心结论横幅 + 实心宝蓝/绛红头微影结构化卡片 + 论文无损原图框/Origin 高精多维图表 + 统一深蓝斜角梯形页码 `XX/NN`（严禁混入体系 A 校园建筑线稿或末页突变红底角标）。
- Curated front & closing anchors: `front-cover-anchor.png`, `front-toc-anchor.png`, `closing-synthesis-anchor.png`, `closing-thanks-anchor.png`.

Before Stage 1, read:
- `references/academic-narrative-guide.md`
- `references/layout-library.md`
- `references/visual-style-spec.md`
- `references/slide-prompt-guide.md`
- `references/outline-template.md`

---

## Stage 1: Full-Fidelity Asset/Equation Extraction, 4 Evidence Ledgers & Deep Outline (`outputs/ppt_outline.md`)

1. **Extract Full Paper Text, OMML/PDF Equations, Structured Tables & Real Figures**:
   ```bash
   PYTHON_BIN="$(command -v python3)"
   [ -x "$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3" ] && \
     PYTHON_BIN="$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
   "$PYTHON_BIN" scripts/extract_paper_assets.py --input <source_file> --out-dir ./outputs/source_assets
   ```
   - Read BOTH `outputs/source_assets/extracted_text.md` AND `outputs/source_assets/figure_guide.md` thoroughly (including all extracted numbered equations `式(1)..式(K)`, structured data tables `表1..表M`, and stitched/embedded figures).
   - For PDF inputs, visually inspect `outputs/source_assets/pages/page_XX.png` and crop every high-value figure cleanly at high resolution into `outputs/source_assets/figures/` via `extract_paper_assets.py crop`. For `.docx` inputs, inspect all `figures/image*.png` and auto-stitched `figures/composite_panel_*.png` visually to verify crispness and exact figure-number mapping.

2. **Build the 4 Mandatory Pre-Outline Ledgers & Persuasive Outline (`outputs/ppt_outline.md`)**:
   - **Page Budget for Full Papers** (give methods, related-work positioning, equations, and experiments enough room to breathe!):
     - **组会深度文献精读**：**默认 16–18 页（完整论文硬下限 ≥ 15 页）**
     - **开题答辩 / 基金申请**：**默认 16–20 页**
     - **毕业 / 结题答辩**：**默认 18–22 页**
     - Short notes / single-figure briefs: 5–8 pages.
   - **Enforce the 8-Module Advisor-Convincing Structure**:
     1. `L01` 庄重学术封面（含底部深蓝四大模块通栏，1 页）
     2. `L02a` 深蓝竖栏汇报提纲（每章必须写出**核心科学子问题 + 方法/实验抓手 + 页码区间 `P03–P04`**，严禁只写 8 字裸标题！1 页）
     3. `L03` + `L03b` **研究背景、物理/数学问题建模 & 既有技术路线卡点对比（必设 2 页！）**：
        - `1.1` 工业/科学场景痛点、退化混叠机理与本文核心建模方程（如二元 vs 三元方程 + 各符号与非线性映射具体结构）；
        - `1.2` 既有主流路线对比矩阵（如表 1 四类代表方法：确定性回归 / 像素扩散 / 结构细节分支 / 二元潜在扩散各自局限）与本文破局定位闭环。
     4. `L11a/b/c` + `L04/L05` **核心方法、编号公式与训练/推理协议深度拆解（4–5 页）**：覆盖总体两阶段/多模块训练与推理闭环（标清冻结 vs 更新模块）、各子网络原图精嵌 + 输入输出张量维度 + **完整编号公式组与损失权重分工**。
     5. `L06/L08` **实验数据集构建、防泄漏划分协议与多维评价体系（1 页）**：交代采集/对齐质控、训练/验证/测试严格隔离、公开基准定位、硬件与阶段超参、互补评价指标含义。
     6. `L06/L07` + `L11a/b` **多路线定量实证、取舍归因与真实定性对比（4–6 页）**：公开集多路线代表基线多指标对比（严禁砍剩 2 点连线！）+ 像素保真 vs 感知质量权衡机理 + 原文放大细节对比读图 + 领域自建集重训全指标优势与场景归因。
     7. `L11a/b` + `L08` **分量消融、逐步模块叠加（全指标）与超参敏感性剖析（2 页）**：展示移除各分量的物理崩溃表现（全四列指标）+ 逐步引入各模块的边际增益分工 + 关键损失权重寻优（如 $\lambda=1.5$）。
     8. `L05/L09/L10a` + `L10b` **真实部署边界、全文三维机制收束与庄重致谢讨论（2–3 页）**。
   - **Run Mandatory Stage-1 Depth Validation BEFORE Stage 2**:
     ```bash
     "$PYTHON_BIN" scripts/validate_outputs.py --base-dir ./outputs --outline-only
     ```
     Fix any flagged shallow slide (`<110` chars `上屏文案`), missing cited equation, verbatim banner echo, or generic cliché before generating visuals!

---

## Stage 2: Visual Scheme Generation & Lock (`outputs/imagegen/visual_preview.png`)

1. Inspect `front-cover-anchor.png`, `front-toc-anchor.png`, `closing-synthesis-anchor.png`, `closing-thanks-anchor.png`, and matching sub-system anchors (`style-a-*.png` or `style-b-*.png`).
2. Generate `outputs/imagegen/visual_preview.png` via `$imagegen` reflecting exact slide count, locked `Style Bible`, diverse layout sequence, real-figure slots (`L11a/b/c`), and rich multi-card/multi-metric visual density.
3. **Internal Visual Gate**: verify stately Front-3 gravitas, domain-neutral TOC sidebar, rich academic density (zero giant-font 3-line flashcards!), zero mint-green/3D slop, and zero school logo / cross-style building leaks.

---

## Stage 3: High-Density Slide Generation, Lossless Figure Compositing, QA & Assembly

1. **Generate Complete Slide Bitmaps (`slide_01.png` ... `slide_NN.png`)**:
   - Build self-contained prompts from `references/slide-prompt-guide.md`, prepending verbatim `LOCKED STYLE BIBLE` and explicit math symbol/subscript spelling locks.
   - **Mandatory Reference Routing**:
     - Slide 1 → `front-cover-anchor.png`; Slide 2 → `front-toc-anchor.png`;
     - Slides 3..N-2 → reviewed `visual_preview.png` + matching layout anchor;
     - Slide N-1 → `closing-synthesis-anchor.png` (pure Style B footer if Style B—DO NOT copy Style A's bottom campus building sketch!);
     - Slide N → `closing-thanks-anchor.png` (centered compact discussion card + uniform navy bottom ribbon & navy page badge matching Slide 1).
   - For **`L11a` / `L11b` / `L11c` 原文真图精嵌剖析页**: render ALL Chinese titles, thesis banners, structured multi-point analysis cards with exact paper equations, bottom takeaway bar, and page badge via `$imagegen`, leaving clean pure-white framed `原文图展示框` at exact target coordinates (`--slot`).

2. **Pixel-Exact Lossless Paper Figure Compositing (`embed_paper_figure.py`)**:
   - Composite single (`L11a/b`) or dual (`L11c`, run twice for left & right slots) real paper figures cleanly:
     ```bash
     "$PYTHON_BIN" scripts/embed_paper_figure.py \
       --slide ./outputs/imagegen/generated_slides/slide_05.png \
       --figure ./outputs/source_assets/figures/image1.png \
       --slot 0.035,0.25,0.57,0.85 \
       --backup-dir ./outputs/imagegen/raw_before_composite \
       --output ./outputs/imagegen/generated_slides/slide_05.png
     ```
   - Record every slide's `method`, `layout`, `embedded_source_figures`, and exact `prompt` in `outputs/imagegen_manifest.json`.

3. **Assemble PPTX + 8-Point Advisor Visual/Content Audit + Automated Validation**:
   ```bash
   "$PYTHON_BIN" scripts/build_ppt_from_images.py \
     --image-dir ./outputs/imagegen/generated_slides \
     --outline ./outputs/ppt_outline.md \
     --contact-sheet ./outputs/imagegen/final_contact_sheet.png \
     --output ./outputs/defense_presentation.pptx

   "$PYTHON_BIN" scripts/validate_outputs.py --base-dir ./outputs
   ```
   - Visually inspect `final_contact_sheet.png` AND every equation / real-figure / quantitative slide at full resolution: verify zero missing equations, zero textbook PSNR/FID filler formulas, zero banner-card verbatim echoes, crisp embedded figures, accurate multi-baseline numbers, and 100% uniform footer geometry. Regenerate/re-composite any defective slide (max 2 passes) and re-run validation.

## Final Response Rules

After completing all 3 stages, reply concisely with:
- 汇报场景、选定视觉体系与总页数（覆盖 8 大核心模块）
- 已生成并通过深度门禁的 `outputs/ppt_outline.md`（含四大前置台账与逐页口播备注）
- 已无损嵌入的论文真图清单及已完整上屏的核心编号公式清单
- 已覆盖的代表性对比基线谱系、核心定量权衡结论与消融归因亮点
- 质检与溯源结果，以及审核总览图 `outputs/imagegen/final_contact_sheet.png` 与成品 `outputs/defense_presentation.pptx`
