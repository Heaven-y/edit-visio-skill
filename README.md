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
  -DrawingScript "C:/path/draw.ps1"
```

`draw.ps1` 应支持 `param([switch]$LoadDrawing)`，并定义
`Draw-VisioPage([int]$Phase)`，兼容旧名称 `Draw-ReferenceFigure`。回调可使用脚手架的
`RectTL/TextTL/OvalTL/LineTL`、面板局部坐标工具和 `Connect-VisioShapes`。
通过 `$script:Page`、`$script:Visio` 访问本次会话。
无参考图的新建/重建可指定 `PageW/PageH` 英寸尺寸，或 `RefW/RefH` 坐标画布。
编辑时若不指定坐标画布，辅助函数使用原页的英寸宽高，左上为原点；直接 COM
操作仍以左下为原点。有参考图时，像素宽高自动读取，页面比例跟随原图。

## 预览、保存与清理

自动检查用 PNG 默认是临时文件，结束时删除。需要 Agent 查看预览时，将
`-PreviewPath` 指向任务临时目录，实际查看后清理；用户明确请求保留的 PNG 则保留。
预览和导出均不得覆盖输入参考图片。

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
- `scripts/visio_package.ps1`：不启动 Visio 的多页关系、文字与媒体读取。
- `scripts/visio_rebuild_scaffold.ps1`：绘图回调、坐标工具、暂存与替换。
- `scripts/visio_stencil_helpers.ps1`：母版查找、只读打开与放置。
- `scripts/visio_stencil_catalog.ps1`：无 COM 的本机母版索引扫描。
- `scripts/visio_export_formats.ps1`：页面等比例导出。
- `scripts/visio_page_tools.ps1`：包检查、显式备份、关闭指定文档及导出。

## 来源与许可

设计参考了 [deermiya/visio-skill](https://github.com/deermiya/visio-skill)
的通用模式、Stencil 索引及图片反推思路，也参考了
[pengjunchi0/codex-visio-paper-figure-skill](https://github.com/pengjunchi0/codex-visio-paper-figure-skill)
的论文图处理思路。运行时和输出规则按本工具包的实际行为维护。

本项目采用 [MIT License](LICENSE)。Microsoft 图标库仍受其自身许可约束。
