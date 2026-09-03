# Visio Stencil 常用速查

本表只列通用图表中最常用的母版。名称来自本机 Visio `2052` 中文内容目录；
真正调用前应运行 `scripts/visio_stencil_catalog.ps1`，确认目标电脑的文件和名称。
英文 `1033` 目录通常使用同一 `NameU` 对应的英文显示名称。

## 文件位置

```text
C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052\
C:\Program Files\Microsoft Office\root\Office16\Visio Content\1033\
```

## 网络与基础设施

### `NETSYM_M.VSSX`

| 本地化 Name | NameU | 语义 |
| --- | --- | --- |
| `路由器` | `Router` | 路由器 |
| `工作组交换机` | `Workgroup Switch` | 交换机 |
| `小型集线器` | `Mini Hub` | 集线器 |
| `网桥` | `Bridge` | 网桥 |
| `网关` | `Gateway` | 网关 |
| `主机` | `Host` | 主机 |
| `关系数据库` | `Relational Database` | 数据库 |

### `SERVER_M.VSSX`

| 本地化 Name | NameU | 语义 |
| --- | --- | --- |
| `服务器` | `Server` | 通用服务器 |
| `文件服务器` | `File Server` | 文件服务 |
| `Web 服务器` | `Web Server` | Web 服务 |
| `数据库服务器` | `Database Server` | 数据库服务 |
| `应用程序服务器` | `Application Server` | 应用服务 |
| `打印服务器` | `Print Server` | 打印服务 |

### `COMPS_M.VSSX`

| 本地化 Name | NameU | 语义 |
| --- | --- | --- |
| `PC` | `PC` | 台式机 |
| `笔记本电脑` | `Laptop` | 笔记本 |
| `LCD 显示器` | `LCD Monitor` | 显示器 |
| `终端` | `Terminal` | 终端 |
| `平板电脑` | `Tablet` | 平板 |

## 论文和领域符号

### `HOLIDAYS_M.VSSX`

| 本地化 Name | NameU | 语义 |
| --- | --- | --- |
| `玉米` | `Corn` | 玉米植株/果穗 |

当前 Figure 3-1 使用的玉米图标应调用 `Corn`，不应再导入自制 SVG。

### `OFFACC_VISIO2013_M.VSSX`

| 本地化 Name | NameU | 语义 |
| --- | --- | --- |
| `植物` | `Plant` | 通用植物 |
| `小型植物` | `Small plant` | 小型植物 |
| `大植物` | `Large plant` | 大型植物 |

### `OFFACC_M.VSSX`

| 本地化 Name | NameU | 语义 |
| --- | --- | --- |
| `叶子` | `Foliage` | 叶片/枝叶 |
| `开花` | `Flowering` | 开花植物 |

### `WEATHERSEASONS_M.VSSX`

| 本地化 Name | NameU | 语义 |
| --- | --- | --- |
| `云` | `Cloud` | 云 |
| `叶子` | `Foliage` | 叶片 |
| `植物` | `Plant` | 植物 |

### `ANALYTICS_M.VSSX`

| 本地化 Name | NameU | 语义 |
| --- | --- | --- |
| `条形图` | `Bar Chart` | 条形图 |
| `靶心` | `Bull's Eye` | 目标/指标 |
| `计算器` | `Calculator` | 计算或评分 |
| `数据库` | `Database` | 数据库 |
| `下降趋势` | `Down Trend` | 下降趋势 |
| `仪表` | `Gauge` | 仪表指标 |

## 云服务

云 Stencil 的文件和 Master 数量随 Office/Visio 安装内容变化。常见文件包括：

- `AZURECLOUD_M.VSSX`
- `AZURECOMPUTE_M.VSSX`
- `AZURENETWORKING_M.VSSX`
- `AZUREDATABASES_M.VSSX`
- `AZURESTORAGE_M.VSSX`
- `AWSCOMPUTE_M.VSSX`
- `AWSSTORAGE_M.VSSX`
- `AWSNETCONTENTDELIVERY_M.VSSX`
- `KUBERNETESVISIOSTENCIL_M.VSSX`

云服务名称很容易因版本和语言变化，必须从完整索引或实时 COM 枚举中复制，
不要凭产品名称拼接 Master 名称。

## 调用要点

```powershell
. "$PSScriptRoot\..\scripts\visio_stencil_helpers.ps1"
$stencil = Open-VisioStencil -Visio $visio -Path 'NETSYM_M.VSSX'
try {
    $shape = Drop-VisioStencilMaster -Page $page -Stencil $stencil `
        -MasterName '路由器' -PinX 4.0 -PinY 3.0 -Width 1.2 -Height 0.9
    try { $shape.Text = 'Core Router' } finally {
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($shape) | Out-Null
    }
} finally {
    Close-VisioStencil $stencil
}
```

`动态连接线` 是许多 Stencil 中附带的连接线母版，不应作为节点图标使用。连接关系
应使用 Visio 连接线工具或 `ConnectorToolDataObject`。完整 362 个文件及所有可读
Master 见 [visio-stencil-index.md](visio-stencil-index.md)。
