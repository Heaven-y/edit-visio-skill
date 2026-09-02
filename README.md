# 通用 Visio 自动化工具包

这是一个面向 Microsoft Visio 的通用自动化工具包，当前保持为初版。它支持创建、编辑、检查、显式导出和图片反推，并将结果维护为 **`.vsdx` 原生可编辑图形**。

核心内容由 Markdown 说明、PowerShell/Python 脚本和可选矢量资产组成。任何能读取这些说明并执行本地脚本的 AI Agent 或自动化流程都可以复用；`agents/openai.yaml` 只是 Codex 的可选界面元数据，不是功能依赖。

## 注意！
图片重建适合结构图、流程图、架构图和多面板示意图。真实数据指标图仍应优先使用 Python/R 生成，再按需把版式元素整理到 Visio。

当前后端使用 Visio COM，因此完整绘图需要 Windows 和 Microsoft Visio。工具包不把调用方限定为某个 Agent；如果要接入 Draw.io 或其他绘图应用，应保留这里的对象清单、输出边界和验收规则，再替换对应后端。


## 工作模式

- `Create`：创建流程图、架构图、时序图、网络拓扑、泳道图或科学示意图。
- `Edit`：修改已有 `.vsdx` 的文字、布局、配色、字体、分组和连接线。
- `Rebuild-image`：把 PNG/JPG/截图/扫描图重建为原生 Visio 形状。
- `Inspect`：检查页面、形状数量、关键文字、媒体条目、文件锁和可编辑性。
- `Export`：只导出用户明确要求的 PNG、SVG、PDF 或 PPTX。

默认只交付 `.vsdx`。预览和其他格式必须显式请求；临时 JSON、预览图和日志应放到构建目录。
写入期间只保留一个滚动备份用于回滚；验收成功后清理备份和临时验证文件，失败时保留备份。

实现参考了 [deermiya/visio-skill](https://github.com/deermiya/visio-skill) 的通用模式划分、时序图、Stencil 和图片反推思路；COM 会话、备份和输出策略按本地 Windows/Visio 环境重新实现。

## 图标与 Stencil 策略

不要为了“有图标”而把低质量剪贴画、Emoji 或整张位图塞进 Visio。按下面的顺序选择：

1. 普通流程、数据和连接关系使用 Visio 原生基本形状。
2. 服务器、数据库、网络设备等标准对象优先使用本机 Visio Stencil 母版。
3. 玉米、DNA、植物器官等领域符号在本机没有合适母版时，使用小型自制 SVG/EMF 矢量资产，保持轮廓、线宽和配色可控。
4. 只有用户明确接受不可编辑细节时，才使用小幅位图；整张参考图始终不能作为最终页面。

图标应满足：轮廓清楚、缩小后仍可辨认、使用有限的填充色和统一线宽，不依赖渐变或照片质感。农业/生物图建议使用“茎、叶、果穗/种子”等少量语义部件表达对象，而不是用多个随意椭圆堆叠。详细选择、导入和许可规则见 [`references/icon-strategy.md`](references/icon-strategy.md)。

`scripts/visio_stencil_catalog.ps1` 可以只读扫描本机 `.vssx/.vss` 文件并列出母版名称，避免凭文件名猜测图标库。

## 适用场景

适合：

- 根据 PNG/JPG/截图重建 Visio 图。
- 将 AI 生成的论文模型图转成可编辑 `.vsdx`。
- 按参考图修改已有 Visio 文件的布局、配色、字体或模块结构。
- 对复杂多面板科学图进行结构化复刻。
- 检查 `.vsdx` 是否误用了整张参考图嵌入。
- 给 Visio 图统一论文风格字体、配色和线条规范。
- 按需从保存后的 `.vsdx` 导出 SVG、PDF、PPTX 或 PNG。
- 对复杂多面板图先做面板四角/边界标定，减少子模块移位、串区和重叠。

不适合：

- 只需要把图片插入 Visio 页面。
- 只需要普通图片编辑、抠图或美化。
- 不要求 Visio 原生可编辑性的纯位图复刻。

## 核心原则

最终交付的 `.vsdx` 应尽量由以下对象构成：

- Visio 原生矩形、圆形、线条、箭头、连接线。
- 可编辑文本。
- 可编辑分组。
- 原生近似绘制的小图表、热图、节点图、立方体、堆叠图。

禁止用整张参考图作为最终页面内容来冒充还原。参考图只能作为临时描摹依据；最终文件中不应留下完整的大尺寸参考 PNG/JPG。

`.vsdx` 是可编辑母版。SVG/PDF/PPTX 是从这个母版导出的交付物，不应该单独重画出彼此不一致的版本。

## 仓库结构

```text
.
├── README.md
├── SKILL.md
├── agents/
│   └── openai.yaml
├── references/
│   ├── icon-strategy.md
│   ├── python-com-backend.md
│   └── rebuild-guidelines.md
├── assets/
│   └── icons/
│       └── maize-ear.svg
└── scripts/
    ├── visio_export_formats.ps1
    ├── visio_page_tools.ps1
    ├── visio_rebuild_scaffold.ps1
    └── visio_stencil_catalog.ps1
```

文件说明：

- `SKILL.md`：通用入口，包含 Visio 模式、工作流、验收标准和安全规则。
- `agents/openai.yaml`：可选的 Codex UI 元数据，不影响脚本在其他 Agent 中使用。
- `references/icon-strategy.md`：内置母版、自制 SVG/EMF 和原生形状的选择规则。
- `assets/icons/maize-ear.svg`：仓库自有的扁平矢量玉米图标，仅在需要时导入或按其结构改绘。
- `references/python-com-backend.md`：PowerShell COM 不稳定时的可选 Python `pywin32` 后端约定。
- `references/rebuild-guidelines.md`：复杂科学图还原准则，包括面板拆解、绘图顺序、样式参数、导出策略和验证 rubric。
- `scripts/visio_export_formats.ps1`：可复用导出函数，支持 PNG、SVG、PDF、PPTX。
- `scripts/visio_page_tools.ps1`：辅助检查脚本，用于备份、导出、检查 `.vsdx` 包结构。
- `scripts/visio_rebuild_scaffold.ps1`：Visio 原生绘图脚手架，用于新建或重建图形，并内置全局坐标和面板局部坐标 helper。
- `scripts/visio_stencil_catalog.ps1`：只读扫描本机 Stencil 并输出母版目录。

## 环境要求

推荐环境：

- Windows。
- Microsoft Visio。
- PowerShell。
- Microsoft PowerPoint，用于 PPTX 导出。
- Git。
- 任意支持本地文件和脚本调用的 AI Agent 或自动化流程（Codex 可通过 `agents/openai.yaml` 获得额外界面集成）。

说明：

- 完整 Visio 自动绘图依赖 Visio COM Automation，因此主要面向 Windows + Microsoft Visio。
- SVG 和 PNG 由 Visio 页面导出。
- PDF 由 Visio 固定格式导出。
- PPTX 默认由 PowerPoint COM 创建单页演示文稿，并插入 Visio 导出的 SVG 页面渲染。
- 导出脚本默认隐藏 Visio 和 PowerPoint；只有显式传入 `-Visible` 才显示窗口。
- 不安装 Visio 时，仍可做 `.vsdx` 包结构检查或有限 XML 修改，但不适合完整一比一重建。

## 安装方式

将本仓库克隆或复制到调用方可读取的 skill/tool 目录。核心功能不依赖 Codex；若使用 Codex，可额外放入其 skills 目录以启用自动发现。

Windows 示例：

```powershell
git clone https://github.com/Heaven-y/edit-visio-skill.git "$env:USERPROFILE\.codex\skills\visio-image-rebuilder"
```

在具体 Agent 中按其 skill 发现机制注册该目录；Codex 用户安装后重启 Codex 或开启新会话即可重新发现。

## 推荐使用方式

示例请求：

```text
使用 visio-image-rebuilder，根据这张参考图片重建 C:\path\model.vsdx，要求最终是 Visio 原生可编辑形状，不要整图嵌入，并导出 SVG、PDF、PPTX。
```

```text
把这个 .vsdx 按参考图更换配色，保持布局不变，最终仍然可编辑，并给我一个 PDF 预览和 PPTX。
```

```text
检查这个 .vsdx 是否只是嵌入了整张 PNG，如果是，请改成原生 Visio 形状重建，再导出 SVG。
```

## 面板标定与防重叠

复杂多面板图使用面板局部坐标和边界断言。推荐流程是：

1. 先标定整张参考图尺寸和 Visio 页面尺寸。
2. 再标定每个主要 panel 的左上角、宽高，必要时记录四角点。
3. panel 内部元素使用 0-1 局部坐标绘制，而不是直接手写全图坐标。
4. 导出预览后检查子模块是否越出父 panel、相邻 panel 是否重叠、箭头和文字是否穿过无关模块。

`visio_rebuild_scaffold.ps1` 中已提供：

- `RectRel`
- `TextRel`
- `OvalRel`
- `LineRel`
- `Assert-RelBox`
- `Assert-RelPoint`

这些 helper 会把局部坐标映射回全局参考坐标，并在局部元素越出 panel 边界时直接报错，避免复杂图后半部分出现整体移位或重叠。

## 多格式导出

导出已有 `.vsdx`：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_page_tools.ps1 `
  -VsdxPath "C:\path\model.vsdx" `
  -ExportFormats svg,pdf,pptx `
  -OutputDir "C:\path\exports" `
  -InspectPackage
```

重建并导出：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_rebuild_scaffold.ps1 `
  -VsdxPath "C:\path\model.vsdx" `
  -PageW 16 `
  -PageH 9 `
  -RefW 1600 `
  -RefH 900 `
  -PreviewPath "C:\path\exports\model.png" `
  -ExportFormats svg,pdf,pptx `
  -OutputDir "C:\path\exports"
```

## 验收标准

一个合格的 Visio 还原结果应满足：

- 主体布局和参考图一致。
- 主要模块、标题、编号、箭头和说明文字齐全。
- 文字可编辑。
- 图形对象可单独选中和修改。
- 没有整张参考图作为最终底图。
- 配色、字体和线条风格统一。
- 领域图标使用合适的 Visio 母版或小型矢量资产，不使用 Emoji、低质量剪贴画或大幅位图替代。
- 复杂多面板图的内部元素不应明显移位、跨 panel 串区或互相重叠。
- 有原文件备份。
- 请求的 PNG/SVG/PDF/PPTX 从同一个保存后的 `.vsdx` 导出，并且文件非空。

## 当前范围

这是一个可直接使用的初版，已覆盖五种工作模式、原生形状重建、面板局部坐标、COM 资源释放、按需导出和包结构检查。后续新增能力应直接更新当前说明和脚本，不再维护容易失真的版本历史表。

## 开源许可证

本项目采用 [MIT License](LICENSE) 开源。
