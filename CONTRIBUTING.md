# Contributing

欢迎改进 XiaoBei Image to VBA。

## 提交前检查

- 不要提交 `__pycache__/`、`.pyc`、生成的 `.pptx`、`.pptm`、`.bas`、截图或用户原始图片。
- 不要提交未授权论文图、商标、截图、显微图或商业素材。
- 修改 `SKILL.md` 时，保持 workflow 清晰，不要把大量示例堆进主文件；复杂说明放到 `references/`。
- 修改脚本后，至少运行一次对应脚本的 `--help` 或最小 smoke test。
- 如果改动了依赖，更新 `requirements.txt`。

## Pull Request 建议

请说明：

- 改了什么；
- 为什么需要改；
- 是否影响 PowerPoint、WPS、macOS 或 Windows 路径；
- 是否引入新的依赖或安全风险。

## 品牌说明

你可以 fork、修改和分发代码，但请不要把修改版描述成“小北在读研官方版本”，也不要使用“小北在读研”的个人品牌来暗示作者背书。
