# 图标与 Visio Stencil 策略

图标的目标是表达语义并保持可编辑，不是用装饰性图片填充页面。优先复用用户
电脑上已经安装的 Visio 母版；仓库不复制 Microsoft 的 `.vss/.vssx` 文件。

## 选择顺序

1. **Visio 原生基本形状**：箭头、节点、容器、矩阵、简单数据流和普通流程图。
2. **本机 Visio Stencil Master**：服务器、数据库、网络设备、云服务、标准人物或
   领域对象。先查 [visio-stencil-index.md](visio-stencil-index.md)，不要猜文件名或
   Master 名称。
3. **原生形状组合 fallback**：母版不存在、外观或语义不匹配时，
   用可编辑曲线、线段和多边形表达必要轮廓，并按部件分组。
4. **外部 SVG/EMF**：只有用户明确要求、确认来源和许可后才导入；不得作为默认
   图标方案，也不得把整张参考图导入页面。

## 本机母版的可靠调用

脚本应使用 `scripts/visio_stencil_helpers.ps1`，不要重复手写 COM 细节：

```powershell
. "$PSScriptRoot\visio_stencil_helpers.ps1"
$stencil = Open-VisioStencil -Visio $visio -Path 'NETSYM_M.VSSX'
try {
    $shape = Drop-VisioStencilMaster `
        -Page $page `
        -Stencil $stencil `
        -MasterName 'Router' `
        -PinX 8.0 -PinY 4.0 `
        -Width 1.2 -Height 0.9
    try { [void]($shape.Text = '') } finally { Release-VisioComObject $shape }
} finally {
    Close-VisioStencil $stencil
}
```

助手的关键约定：

- 通过 `Documents.OpenEx(path, 66)` 只读隐藏打开母版文件；2 是只读，64 是隐藏。
- 优先用 `Masters.ItemU(name)` 按稳定名称查找；失败时再按 `Name` 和 `NameU` 遍历，
  并在报错中列出候选名称。
- 用 `Page.Drop(master, pinX, pinY)` 放置后再设置 `Width`、`Height`、`PinX`、`PinY`。
  `PinX/PinY` 是页面英寸坐标，通常使用图标中心点。
- 放置完成后释放 Shape、关闭并释放 Stencil 文档；不要让 COM 对象进入 PowerShell
  管道，也不要关闭用户已经打开的其他 Visio 文档。
- 中文 `2052` 和英文 `1033` 的显示名称不同。跨语言脚本优先使用索引中的 `NameU`，
  只有在需要面向用户显示时才使用当前 Stencil 的本地化 `Name`。

## 速查入口

常用网络、服务器、终端、农业和云服务母版见
[stencil-reference.md](stencil-reference.md)；所有本机可读母版见
[visio-stencil-index.md](visio-stencil-index.md)。两个文件都只记录名称和路径，不
复制或重新分发 Microsoft Stencil 文件。

有些 Stencil 包含连接线母版；它不是节点图标。需要随对象移动的关系使用
`Connect-VisioShapes` 胶合连接线；`DrawLine` 适合固定的轴线、分隔线等几何图元。

## 领域图标规则

- 网络、办公、工程和其他领域对象都先查本机 Stencil，不能把某一台电脑的
  索引当作所有 Visio 安装都具备的资源清单。
- Stencil 风格与图表其他部分差异过大时，保留母版的语义轮廓，只统一尺寸和周围
  文字；确实无法匹配时采用原生 fallback，不必强行使用。
- 新的领域图标应先渲染样例，检查轮廓、默认比例与样式控制项。子形状可能是
  备选样式或隐藏层，不一定是独立语义部件；具体母版注意事项见速查表。
- 没有合适母版时，fallback 必须保留决定语义的轮廓和少量局部结构，并让每个部件
  可单独选择、缩放和移动。
- 外部资产若获批准，必须记录来源、许可证、导入格式和回退方案；导入后检查
  `.vsdx` 的 `visio/media`，确保没有意外生成大幅位图。

## 发现与验证

先搜索已有索引；仅在文件缺失、名称不匹配或目标电脑不同的时候重新扫描。
将扫描结果写到任务临时目录，不自动修改安装的 skill：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\visio_stencil_catalog.ps1 `
  -RootPath 'C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052' `
  -Format markdown `
  -OutputPath C:\path\task-temp\visio-stencil-index.md
```

该脚本只读取 `.vssx/.vss` 文件，不启动 Visio；`.vssx` 从压缩包中的
`visio/masters/masters.xml` 读取本地化名称、`NameU` 和 ID。旧式二进制 `.vss` 只记
录路径，并提示需要 COM 枚举，不会伪造 Master 名称。
