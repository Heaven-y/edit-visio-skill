# 通用 Visio 工具包

创建、编辑、检查、按需导出 Visio 图形，并把参考图片重建为原生可编辑的
`.vsdx`。核心由 Markdown、PowerShell 和本机 Visio COM 构成，不限定调用方：
任何能够读取说明并运行本地脚本的 Agent 或自动化流程都可以使用。
`agents/openai.yaml` 只是可选界面元数据。

## 使用边界

支持 Create、Edit、Rebuild-image、Inspect、Export 五种工作模式。
默认只交付 VSDX，不额外输出 SVG、PDF、PPTX。图片重建面向结构图、流程图、
架构图和概念示意图，不把整张参考图嵌入页面冒充可编辑内容。
真实数据指标图应由真实数据和可追溯的 Python/R 绘图脚本生成。

完整绘制和渲染需要 Windows、PowerShell 7（`pwsh`）与已授权的 Microsoft Visio。
请在 `pwsh` 中执行下方命令。COM 后端会在 Windows PowerShell 5.1 中提前报错，
避免在不支持的宿主中挂起；只读包检查和母版目录扫描不启动 COM。
PowerPoint 仅在用户请求 PPTX 时需要。缺失软件先说明环境要求和安装来源；
不自动下载或部署需要用户授权的 Office 产品。

## 安装与维护

只维护一份工作目录，再按各调用方的发现机制注册或链接。
使用 CC Switch 时，以其管理的 skill 目录为唯一来源，不再分别克隆到多个
Agent 的目录。已经安装时直接更新现有目录。

仓库：[Heaven-y/edit-visio-skill](https://github.com/Heaven-y/edit-visio-skill)。
这是独立维护的个人工具包，不自动追踪参考项目的上游改动。
实际同步以本仓库的提交与推送结果为准，不把本地修改称为远程更新。

## 图标与母版

普通框、矩阵、坐标轴和连线用基本形状；服务器、设备和领域对象先查本机
Stencil。按稳定的 NameU 查找，再渲染样例确认外观与语义。

母版名相近不等于语义或外观匹配。部分母版的子形状是备选样式或隐藏层，
不一定是独立部件；改色前必须检查。
没有合适母版时使用原生曲线和分组部件，不默认导入自制 SVG。
仓库只记录母版名称和路径，不复制或分发 Microsoft Stencil 文件。

入口：

- [图标策略](references/icon-strategy.md)
- [常用母版速查](references/stencil-reference.md)
- [完整索引](references/visio-stencil-index.md)
- [图片重建准则](references/rebuild-guidelines.md)
- [版式配方](references/template-library.md)，不是已打包的 VSDX 模板
- [画布、连接线与最终尺寸](references/canvas-and-connectors.md)

## 创建、编辑与重建

脚手架支持三种明确的写入模式，`PageIndex` 从 1 开始：

- `Create`：目标必须不存在；若提供模板，保留模板内容。
- `Edit`：目标必须存在；保留已有对象和页面尺寸，由回调只修改指定对象。
- `Rebuild`：清空指定页重绘，不自动清空其他页。旧调用默认此模式。

新建无参考图的流程图：

```powershell
& "$skillRoot/scripts/visio_rebuild_scaffold.ps1" `
  -Mode Create -VsdxPath "C:/path/workflow.vsdx" `
  -PageW 12 -PageH 7.5 -DrawingScript "C:/path/draw.ps1"
```

只编辑第二页：

```powershell
& "$skillRoot/scripts/visio_rebuild_scaffold.ps1" `
  -Mode Edit -PageIndex 2 -VsdxPath "C:/path/workflow.vsdx" `
  -DrawingScript "C:/path/edit.ps1"
```

编辑回调先按稳定的 NameU 或 ID 定位对象，不能顺带全局改色、重设共享母版或
删除其他页。脚手架负责隔离保存；具体回调仍须遵守用户指定的编辑范围。

图片重建流程：

1. 从参考图实际读取像素宽高，根据长宽比计算页面尺寸；不写死示例画布。
2. 按参考图整理布局、文字、箭头关系、配色与图标，命名并分组主要模块。
3. 用任务绘图回调重绘暂存文件，不直接清空并保存用户的原文件。
4. 关闭独立 COM 会话，检查暂存 VSDX，再替换目标。
5. 查看渲染预览，与参考图逐区比较后交付。

脚手架接受 `-Phase 1|2|3`。这些是完整重绘到指定阶段的检查点，不是多次运行
自动续画。简单图可以只运行最终阶段，不规定形状数量或固定耗时。

```powershell
& "$skillRoot/scripts/visio_rebuild_scaffold.ps1" `
  -Mode Rebuild -VsdxPath "C:/path/diagram.vsdx" `
  -ReferenceImagePath "C:/path/reference.png" `
  -PageWidthMm 180 `
  -DrawingScript "C:/path/draw.ps1"
```

`draw.ps1` 应支持 `param([switch]$LoadDrawing)`，并定义
`Draw-VisioPage([int]$Phase)`，兼容旧名称 `Draw-ReferenceFigure`。回调可使用脚手架的
`RectTL/TextTL/OvalTL/LineTL`、面板局部坐标工具和 `Connect-VisioShapes`。
通过 `$script:Page`、`$script:Visio` 访问本次会话。
无参考图的新建/重建可指定 `PageW/PageH` 英寸尺寸，或 `RefW/RefH` 坐标画布。
编辑时若不指定坐标画布，辅助函数使用原页的英寸宽高，左上为原点；直接 COM
操作仍以左下为原点。有参考图时，像素宽高自动读取，页面比例跟随原图。

页面支持 `PageWidthMm/PageHeightMm` 毫米尺寸。只给宽度或高度时，另一边按
参考图比例推导；同一方向不能同时给毫米与英寸。旧调用的 16 英寸默认宽度
仅为兼容，不代表适合论文排版。已知最终使用宽度时应直接采用该尺寸。

固定纸张与参考图比例不同时，可用 `-CanvasFit Contain -MarginMm 10` 等比例居中
留白，不裁切或拉伸。`VX/VY` 转换位置，`VL/VPT` 分别将坐标长度转换为英寸/pt，
避免把居中偏移加到圆角、宽度或线宽中。工程比例尺页面需单独处理和验收。

`Connect-VisioShapes` 支持 `FromSide/ToSide` 端口方向、`Routing` 原生直线或
正交路由、线宽、虚线、箭头和 `LabelOffsetXPt/LabelOffsetYPt` 标签偏移。
连接会胶合到节点，移动节点后仍应逐条检查走线、标签和交叉关系。

## 预览、保存与清理

自动检查用 PNG 默认是临时文件，结束时删除。需要 Agent 查看预览时，将
`-PreviewPath` 指向任务临时目录，实际查看后清理；用户明确请求保留的 PNG 则保留。
预览和导出均不得覆盖输入参考图片。
脚手架 `PreviewDpi` 和单独导出的 `PngDpi` 默认 144，可按用途调整，例如打印用
300。脚本临时设置并恢复 Visio 的栅格导出分辨率和尺寸，不遗留全局导出设置。
提高 DPI 只增加像素数，不能补救最终物理尺寸下过小的字或过细的线。

失败重绘不改变原目标，不需要默认堆积 backup。只有 `-KeepBackup` 才保留旧文件，
备份名唯一，不覆盖已有备份。目标被锁定时停止；覆盖许可不等于丢弃未保存界面修改。

使用独立的 `Visio.InvisibleApp`，只关闭自动化拥有的文档和进程。
`-CloseOpenDocument` 不会退出用户的 Visio；若目标有未保存修改，必须明确选择
`-SaveOpenDocument` 或 `-DiscardOpenDocument`。

## 检查与导出

`scripts/visio_quality_gates.ps1` 检查包结构、原生对象、媒体、必要文字和颜色、
页面及预览比例、几何边界、只读重开和真实 Windows 进程退出。
`visio_validate.ps1` 复用同一实现，避免两套检查逐渐不一致。
文字、颜色、比例和边界按 `-PageIndex` 检查，多页修改应逐页验收。
`visio_page_tools.ps1 -InspectPackage` 只读列出全部页，包括背景页及实际 XML 路径。

默认严格检查原生内容。用户明确允许保留既有 Logo 或导入外部资源时，使用
`-AllowMedia`，并如实说明文档含非原生对象，不能宣称完全原生可编辑。
不要为了通过检查而删除用户已有图片。

自动检查不能证明文字不挤压、图标正确、图形没有视觉重叠或图片相似。
缺失的检查报告 SKIPPED，不把参数值或未执行的检查称为 PASS。

`MinFontPt/MinLinePt` 是可选的任务阈值；`FinalWidthMm` 用于检查页面在使用时
等比例缩小后的名义字号、线宽。不强制把某一本期刊的阈值用于所有图。
检查递归覆盖所选页的分组，但不是字形大小、隐藏状态、文字溢出或碰撞检测；
背景页需单独检查。明确需要空白页时才使用 `AllowEmptyPage`。

只有明确请求时才导出 PNG、SVG、PDF、PPTX：

```powershell
& "$skillRoot/scripts/visio_page_tools.ps1" `
  -VsdxPath "C:/path/diagram.vsdx" -ExportFormats svg,pdf -OutputDir "C:/path/exports"
```

PNG、SVG、PPTX 导出 `-PageIndex` 选定页，PDF 导出全部前景页并应用各自背景。
PPTX 包含 Visio 页面的 SVG 渲染，不承诺拆分为 PowerPoint 原生形状。
全部格式从同一保存后的 VSDX 生成，不能分别重画。

## 代码入口

- `SKILL.md`：工作模式、执行约束和验收入口。
- `scripts/visio_runtime.ps1`：图片尺寸、隔离会话、COM 释放与胶合连接器。
- `scripts/visio_canvas.ps1`：毫米/英寸尺寸、等比例坐标与留白适配。
- `scripts/visio_package.ps1`：不启动 Visio 的多页关系、文字与媒体读取。
- `scripts/visio_rebuild_scaffold.ps1`：绘图回调、坐标工具、暂存与替换。
- `scripts/visio_stencil_helpers.ps1`：母版查找、只读打开与放置。
- `scripts/visio_stencil_catalog.ps1`：无 COM 的本机母版索引扫描。
- `scripts/visio_export_formats.ps1`：页面等比例导出。
- `scripts/visio_page_tools.ps1`：包检查、显式备份、关闭指定文档及导出。

维护者可运行无 COM 回归检查，或在本机 Visio 中生成隔离测试文档：

```powershell
pwsh -NoProfile -File tests/visio_regression.ps1
pwsh -NoProfile -File tests/visio_regression.ps1 -WithCom
```

测试默认清理输出，仅保留测试源码。需要查看渲染结果时加 `-KeepArtifacts`，
查看后删除该次输出目录；不使用用户论文图作为回归测试目标。

## 来源与许可

本项目作者、维护者与主版权归属为 **Heaven-y**，采用 [MIT License](LICENSE)。
以下是不同层面的参考来源，不是项目共同作者或自动同步的上游：

- [deermiya/visio-skill](https://github.com/deermiya/visio-skill)：通用图表模式、
  Stencil 发现、布局间距与连接关系的处理思路；不捆绑其自动布局/OCR 实现。
- [pengjunchi0/codex-visio-paper-figure-skill](https://github.com/pengjunchi0/codex-visio-paper-figure-skill)：
  原始绘图/导出结构、图片重建与面板局部坐标。保留/改编内容的 MIT 声明见
  [第三方声明](THIRD_PARTY_NOTICES.md)，不替代本项目的作者和版权。
- [Yuan1z0825/nature-skills](https://github.com/Yuan1z0825/nature-skills)：
  最终物理尺寸、文字与线条层级、逐区渲染验收的方法参考。未复制其代码，
  不强制 Nature 配色、固定字号、Python/R 后端或 SVG/PDF 交付要求。
- [Imbad0202/academic-research-skills-codex](https://github.com/Imbad0202/academic-research-skills-codex)：
  证据、推断和建议分离，以及检查结论不超出检查范围的方法参考。未复制其
  代码或提示词，不引入研究流水线、跨模型调用、网络上传或额外运行依赖。

科学示意图里的箭头、标签和小图不能暗示资料未支持的定量结果或因果结论。
普通流程图、网络图等不因此被强制套入论文工作流。
Microsoft 图标库仍受其自身许可约束，仓库不分发其二进制文件或图标资产。
