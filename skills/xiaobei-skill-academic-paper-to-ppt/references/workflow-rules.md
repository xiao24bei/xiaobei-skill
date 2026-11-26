# Three-Stage Workflow, Stage-1 Depth Gate & Provenance Rules (v5.0)

本技能采用带有**阶段一导师级内容深度门禁（Stage-1 Depth Gate）**的三阶段科研汇报制作流水线（阶段 1 全量解析与高密度大纲 → 阶段 2 视觉方案锁定 → 阶段 3 高密度生图、真图精嵌与双重校验装配）：

## 1. End-to-End 3-Stage Pipeline

### 阶段 1：全文/公式/表格全量解析、原图裁剪/拼合与四大台账大纲
1. 运行 `scripts/extract_paper_assets.py --input <source_file> --out-dir ./outputs/source_assets`：
   - 对 `.pdf`：渲染 220 DPI 页面、提取内嵌图像、识别图注表注与编号公式；
   - 对 `.docx`：完整转换全部 Word OMML 数学公式（`<m:oMath>`）、识别双栏编号公式表生成 `Numbered Equations Catalog`、保留结构化 Markdown 数据表、并自动将多子图对比表格拼合为高清 `figures/composite_panel_XX.png`。
2. 审阅 `extracted_text.md` 与 `figure_guide.md`（对 PDF 使用 `extract_paper_assets.py crop` 精准裁剪矢量/多子图原图）。
3. 编写 `outputs/ppt_outline.md`：必须先完成**四大前置证据台账**（原图台账、编号公式上屏映射台账、既有基线卡点与训练/推理冻结协议台账、多维实验对比与权衡台账），再逐页编写具备 **140–260 字实质干货 + 完整论文公式 + 深度 WHY 归因 + 130–220 字口播备注** 的页面大纲（完整论文默认 16–18 页，硬下限 ≥ 15 页，含独立 `1.2 既有路线对比与本文定位` 页）。
4. **【强制深度门禁】在进入阶段 2 前，必须运行**：
   ```bash
   python3 scripts/validate_outputs.py --base-dir ./outputs --outline-only
   ```
   若检出核心正文页字数过少（提词卡化）、引用公式未上屏、横幅与卡片原句复读、塞入教科书 `PSNR/FID` 定义公式或空洞 AI 套话，**必须先在 `ppt_outline.md` 中补齐修正并通过校验，严禁带着浅薄大纲进入生图阶段！**

### 阶段 2：全套视觉方案生成与锁定
1. 加载 `$CODEX_HOME/skills/.system/imagegen/SKILL.md`（若在当前环境中可用），审阅 `front-cover-anchor.png`、`front-toc-anchor.png`、`closing-synthesis-anchor.png`、`closing-thanks-anchor.png` 与对应子体系锚点。
2. 调用 `$imagegen` 生成多页缩略图视觉总览 `outputs/imagegen/visual_preview.png`，确保与大纲页数、版式多样性、原图展示框（`L11a/b/c`）与高信息密度卡片结构一致。
3. 内部视觉把关：确认无巨字大白框空壳、无跨学科底纹串台、无跨体系建筑线稿污染、无示例校名抄袭。

### 阶段 3：高密度逐页生图、论文原图无损精嵌、8 维质检与装配
1. 逐页调用 `$imagegen` 生成 `outputs/imagegen/generated_slides/slide_01.png` ... `slide_NN.png`（完整绘入中文标题、顶栏结论、多条紧凑结构化剖析卡片、论文真实编号公式、多基线多指标图表与锁定页码）。
2. 对大纲标记为 `embed_real_figure`（`L11a` / `L11b` 单图或 `L11c` 左右双图）的页面，立刻运行 `scripts/embed_paper_figure.py` 将高清论文原图或自动拼合组图**像素级无损精嵌**进预留纯白框内（双图页依次对左右两个 `--slot` 执行两次嵌入）。
3. 写入完整溯源记录 `outputs/imagegen_manifest.json`。
4. 运行 `scripts/build_ppt_from_images.py` 生成带演讲备注的严格全图 `outputs/defense_presentation.pptx` 与审核总览图 `outputs/imagegen/final_contact_sheet.png`。
5. 视觉逐页核查 `final_contact_sheet.png` 及所有公式页/原图页/定量页，定向返修任何漏公式、复读、裁切或页脚不一致页面（最多 2 轮），最后运行完整校验 `scripts/validate_outputs.py --base-dir ./outputs` 全部通过后交付。

## 2. Stage-Only & Interactive Step-by-Step Mode

- 当用户显式要求仅执行某一阶段，或要求“按三步法每步停下确认后再继续”时：
  - **Step 1 停止点**：产出 `outputs/source_assets/` 与通过 `--outline-only` 门禁的 `outputs/ppt_outline.md` 后汇报大纲结构与四大台账摘要，等待用户确认；
  - **Step 2 停止点**：产出 `outputs/imagegen/visual_preview.png` 后展示视觉方案，等待用户确认；
  - **Step 3 完成点**：生成全部高清页面、无损嵌图、组装 PPTX 并通过完整校验。

## 3. `$imagegen` Mode & Authorized Fallback Policy

- **Default primary mode (`built_in_image_gen`)**：优先使用内置生图工具逐页生成并保存至 `outputs/imagegen/`。
- **Authorized fallback (`imagegen_cli_fallback`)**：仅当内置模式不可用、瞬时重试 1 次仍失败、或经有界视觉修正仍不合格时，自动回退至 `image_gen.py`（默认 `gpt-image-2`，依赖环境变量 `OPENAI_API_KEY`，严禁打印密钥值）。
- **Hard Blocker Rule**：严禁用 HTML/CSS 截图、SVG/Canvas 或 python-pptx 文本框替代 `$imagegen`；仅允许在 `$imagegen` 已生成的完整学术页上通过 `embed_paper_figure.py` 无损精嵌论文真实原图。

## 4. `outputs/imagegen_manifest.json` Schema

所有路径相对 `outputs/`：

```json
{
  "generation_method": "built_in_image_gen",
  "visual_system": "style_b_science_blue_crimson",
  "visual_preview": "imagegen/visual_preview.png",
  "visual_preview_method": "built_in_image_gen",
  "image_output_dir": "imagegen/generated_slides",
  "slides": [
    {
      "slide": 7,
      "file": "imagegen/generated_slides/slide_07.png",
      "method": "built_in_image_gen",
      "layout": "L11c",
      "embedded_source_figures": [
        "source_assets/figures/image3.png",
        "source_assets/figures/image4.png"
      ],
      "prompt": "full self-contained slide prompt used for this page"
    }
  ]
}
```
