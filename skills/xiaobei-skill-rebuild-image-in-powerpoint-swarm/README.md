# PowerPoint 蜂群绘图

给 Codex 一张科研参考图，蜂群模式会把它拆成版面、插图、连线和标注等绘图任务。在 Windows 桌面版 PowerPoint 中，六个角色光标在同一张幻灯片上协同作图，按照对象的位置、层级和连接关系逐步还原整页。

重绘完成后，你可以直接在 PPT 里修改实验分组名称、调整流程箭头，或替换一张显微图。文字、形状和连线都是原生 PowerPoint 对象，复杂插图作为独立图片保留。

## 使用方式

实时绘图使用 Windows 桌面版 PowerPoint、带 Tkinter 的 Python 3.10+ 和 `pywin32`。在 Codex 中安装后，附上参考图并输入以下指令。

```text
用 $xiaobei-skill-rebuild-image-in-powerpoint-swarm 启动蜂群模式，在 PowerPoint 里协同重绘这张参考图。保留原图的分区、文字和箭头关系，交付可继续编辑的 PPT。
```

[查看绘图案例与独立安装包](https://github.com/xiao24bei/xiaobei-skill/releases/tag/rebuild-image-in-powerpoint-swarm-20261125) · [查看 Skill 工作流程](SKILL.md)
