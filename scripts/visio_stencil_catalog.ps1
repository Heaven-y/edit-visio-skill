param(
    [string[]]$RootPath,
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-DefaultStencilRoots {
    $roots = @()
    if ($env:ProgramFiles) {
        $roots += Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\Visio Content'
    }
    if (${env:ProgramFiles(x86)}) {
        $roots += Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\Visio Content'
    }
    if ($env:USERPROFILE) {
        $roots += Join-Path $env:USERPROFILE 'Documents\My Shapes'
    }
    return @($roots | Select-Object -Unique)
}

function Read-VssxMasters([string]$path) {
    $names = New-Object System.Collections.Generic.List[string]
    $archive = $null
    try {
        $archive = [IO.Compression.ZipFile]::OpenRead($path)
        $entry = $archive.Entries |
            Where-Object { $_.FullName -ieq 'visio/masters/masters.xml' } |
            Select-Object -First 1
        if ($null -eq $entry) { return @($names) }

        $reader = New-Object IO.StreamReader($entry.Open())
        try {
            [xml]$xml = $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }

        foreach ($master in @($xml.SelectNodes("//*[local-name()='Master']"))) {
            $name = [string]$master.NameU
            if (-not $name) { $name = [string]$master.Name }
            if ($name -and -not $names.Contains($name)) {
                $names.Add($name)
            }
        }
    } finally {
        if ($archive -ne $null) { $archive.Dispose() }
    }
    return @($names)
}

$roots = @($RootPath)
if ($roots.Count -eq 0) { $roots = Get-DefaultStencilRoots }

$files = New-Object System.Collections.Generic.List[string]
foreach ($root in $roots) {
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        continue
    }
    foreach ($file in @(Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in @('.vssx', '.vss') })) {
        if (-not $files.Contains($file.FullName)) { $files.Add($file.FullName) }
    }
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($filePath in $files) {
    $file = Get-Item -LiteralPath $filePath
    $names = @()
    $note = ''
    if ($file.Extension -ieq '.vssx') {
        try {
            $names = @(Read-VssxMasters $file.FullName)
        } catch {
            $note = "读取 VSSX 失败: $($_.Exception.Message)"
        }
    } else {
        $note = '旧版二进制 VSS，未启动 Visio，无法读取母版名称'
    }
    $sample = (@($names | Select-Object -First 8) -join ' | ')
    $results.Add([pscustomobject]@{
        File = $file.Name
        Path = $file.FullName
        MasterCount = if ($names.Count -gt 0) { $names.Count } else { $null }
        SampleMasters = $sample
        Note = $note
    })
}

if ($OutputPath) {
    $parent = Split-Path -Parent ([IO.Path]::GetFullPath($OutputPath))
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
    $results | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
    Write-Output "已写入目录: $OutputPath"
} else {
    if ($results.Count -eq 0) {
        Write-Output '未找到 .vssx 或 .vss 文件。请使用 -RootPath 指定 Visio Content 或 My Shapes 目录。'
    } else {
        $results | Sort-Object File | Format-Table File, MasterCount, SampleMasters, Note -AutoSize
    }
}
