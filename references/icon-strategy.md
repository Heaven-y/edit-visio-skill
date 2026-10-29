# 图标与 Visio Stencil 策略

图标首先要准确传意、在交付尺寸下清楚美观，并与整图风格一致。原生可编辑性是实现选择，不是高于交付质量的硬约束；用户明确要求全原生时才按该范围执行。仓库不复制Microsoft母版文件。

## 按用途选择

- 文字、矩阵、节点与连线通常适合原生形状。
- 本机母版、原生曲线、许可复用的矢量图或高清位图均可用于领域图标，按语义和效果选择，不强制先尝试母版或反复增加曲线细节。
- 概念或测量示意可以只表现关键器官、动作、单位或时间关系，不必绘制完整生物体；精细不等于写实。
- 沿用会话中已有的媒体和编辑性授权。仅在未获授权、且会实质改变交付范围时澄清，不为每种格式或每个来源重复确认。
- 第三方素材须核对许可并保留必要归属；归属可放在配套说明或图注，非必要不占据图面。若照片可能被误认为本研究证据，应在适当说明中消除歧义。自行绘制的普通概念图不机械添加“非实测”声明。

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

- 选择母版时先核对本机文件、NameU和实际轮廓，不猜测资源存在。
- 任何候选形式都须在目标尺寸下检查语义、比例与风格。母版子形状可能是备选样式或隐藏层，不能盲目按编号改色。
- 原生图标按语义分组；图片保持纵横比，文字和关系连线单独编辑。
- 使用外部资产时记录来源、许可证及实际处理，核对媒体清单；只禁止非预期的栅格化，不把位图本身视为质量失败。

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
