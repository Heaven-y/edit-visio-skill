param(
    [Parameter(Mandatory = $true)]
    [string]$VsdxPath,

    [string[]]$RequiredText,
    [switch]$SkipCom,
    [switch]$StrictProcess
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'visio_stencil_helpers.ps1')

function Read-VsdxPackage([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $pageEntries = @($zip.Entries | Where-Object { $_.FullName -match '^visio/pages/page\d+\.xml$' })
        $mediaEntries = @($zip.Entries | Where-Object { $_.FullName -like 'visio/media/*' })
        $shapeCount = 0
        $texts = New-Object System.Collections.Generic.List[string]
        foreach ($entry in $pageEntries) {
            $reader = New-Object IO.StreamReader($entry.Open())
            try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $shapeCount += @($xml.SelectNodes("//*[local-name()='Shape']")).Count
            foreach ($node in @($xml.SelectNodes("//*[local-name()='Text']"))) {
                if ($node.InnerText) { $texts.Add([string]$node.InnerText) }
            }
        }
        [pscustomobject]@{
            PageCount = $pageEntries.Count
            ShapeCount = $shapeCount
            Media = $mediaEntries
            Text = @($texts)
        }
    } finally {
        $zip.Dispose()
    }
}

if (-not (Test-Path -LiteralPath $VsdxPath -PathType Leaf)) {
    throw "VSDX not found: $VsdxPath"
}
$fullPath = (Resolve-Path -LiteralPath $VsdxPath).Path
$package = Read-VsdxPackage $fullPath
$rasterMedia = @($package.Media | Where-Object {
        $_.Length -gt 1000000 -or $_.FullName -match '\.(png|jpg|jpeg|gif|bmp)$'
    })

Write-Output "VSDX: $fullPath"
Write-Output "Pages: $($package.PageCount)"
Write-Output "Native shapes: $($package.ShapeCount)"
Write-Output "Media entries: $($package.Media.Count)"
Write-Output "Raster or large media entries: $($rasterMedia.Count)"
if ($package.Media.Count -gt 0) {
    foreach ($entry in $package.Media) { Write-Output ("Media: {0} ({1} bytes)" -f $entry.FullName, $entry.Length) }
}
if ($package.PageCount -lt 1) { throw 'VSDX contains no Visio page XML.' }
if ($package.ShapeCount -lt 1) { throw 'VSDX contains no native shapes.' }
if ($rasterMedia.Count -gt 0) { throw 'VSDX contains raster or large media; inspect before delivery.' }

foreach ($required in @($RequiredText)) {
    if ([string]::IsNullOrWhiteSpace($required)) { continue }
    $found = @($package.Text | Where-Object {
            $_.IndexOf($required, [StringComparison]::OrdinalIgnoreCase) -ge 0
        }).Count -gt 0
    if (-not $found) { throw "Required text not found: $required" }
    Write-Output "Required text found: $required"
}

if (-not $SkipCom) {
    $beforeIds = @((Get-Process -Name VISIO -ErrorAction SilentlyContinue).Id)
    $visio = $null
    $doc = $null
    $page = $null
    try {
        $visio = New-Object -ComObject Visio.Application
        $visio.Visible = $false
        $doc = $visio.Documents.Open($fullPath)
        $page = $doc.Pages.Item(1)
        Write-Output ("COM reopen: pages={0}, shapes={1}" -f $doc.Pages.Count, $page.Shapes.Count)
    } finally {
        if ($doc -ne $null) { try { $doc.Saved = $true } catch {}; try { $doc.Close() } catch {} }
        Release-VisioComObject $page
        Release-VisioComObject $doc
        if ($visio -ne $null) { try { $visio.Quit() } catch {}; Release-VisioComObject $visio }
    }

    $remaining = @()
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        $remaining = @((Get-Process -Name VISIO -ErrorAction SilentlyContinue) |
            Where-Object { $beforeIds -notcontains $_.Id })
        if ($remaining.Count -eq 0) { break }
        Start-Sleep -Milliseconds 100
    }
    if ($remaining.Count -gt 0) {
        $message = "New Visio process(es) remain after validation: " + (($remaining.Id) -join ', ')
        if ($StrictProcess) { throw $message }
        Write-Warning $message
    } else {
        Write-Output 'COM process cleanup: OK'
    }
}

Write-Output 'Validation: OK'
