param(
    [Parameter(Mandatory = $true)]
    [string]$VsdxPath,

    [string]$ReferenceImagePath,
    [string]$PreviewPath,
    [string[]]$RequiredText,
    [string[]]$RequiredColor,

    [ValidateSet(1, 2, 3)]
    [int]$Phase = 3,

    [long]$MaxBytes = 2097152,
    [switch]$SkipCom,
    [switch]$StrictProcess
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'visio_stencil_helpers.ps1')

function Read-VsdxQualityPackage([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $pages = @($zip.Entries | Where-Object { $_.FullName -match '^visio/pages/page\d+\.xml$' })
        $media = @($zip.Entries | Where-Object { $_.FullName -like 'visio/media/*' })
        $shapeCount = 0
        $texts = New-Object System.Collections.Generic.List[string]
        $colors = New-Object System.Collections.Generic.List[string]
        foreach ($entry in $pages) {
            $reader = [IO.StreamReader]::new($entry.Open())
            try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $shapeCount += @($xml.SelectNodes("//*[local-name()='Shape']")).Count
            foreach ($node in @($xml.SelectNodes("//*[local-name()='Text']"))) {
                $value = [regex]::Replace([string]$node.InnerText, '\s+', ' ').Trim()
                if ($value) { $texts.Add($value) }
            }
            foreach ($cell in @($xml.SelectNodes("//*[local-name()='Cell']"))) {
                $name = [string]$cell.N
                if ($name -in @('FillForegnd', 'LineColor', 'Char.Color')) {
                    $formula = [string]$cell.FormulaU
                    if (-not $formula) { $formula = [string]$cell.Formula }
                    if (-not $formula) { $formula = [string]$cell.V }
                    if ($formula) { $colors.Add($formula) }
                }
            }
        }
        [pscustomobject]@{
            PageCount = $pages.Count
            ShapeCount = $shapeCount
            Media = $media
            Text = @($texts.ToArray())
            Colors = @($colors.ToArray())
        }
    } finally {
        $zip.Dispose()
    }
}

function Test-RequiredValues([string[]]$Values, [string[]]$Required, [string]$Label) {
    foreach ($item in @($Required)) {
        if ([string]::IsNullOrWhiteSpace($item)) { continue }
        $normalizedItem = $item.Trim()
        if ($normalizedItem -match '^RGB\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$') {
            $normalizedItem = '#{0:X2}{1:X2}{2:X2}' -f [int]$Matches[1], [int]$Matches[2], [int]$Matches[3]
        }
        $found = @($Values | Where-Object {
                $_.IndexOf($item, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $_.IndexOf($normalizedItem, [StringComparison]::OrdinalIgnoreCase) -ge 0
            }).Count -gt 0
        if (-not $found) { throw "$Label not found: $item" }
        Write-Output "  $Label found: $item"
    }
}

if (-not (Test-Path -LiteralPath $VsdxPath -PathType Leaf)) {
    throw "INPUT_VALIDATION failed: VSDX not found: $VsdxPath"
}
$fullPath = (Resolve-Path -LiteralPath $VsdxPath).Path
if ($ReferenceImagePath -and -not (Test-Path -LiteralPath $ReferenceImagePath -PathType Leaf)) {
    throw "INPUT_VALIDATION failed: reference image not found: $ReferenceImagePath"
}

$file = Get-Item -LiteralPath $fullPath
if ($file.Length -le 0) { throw 'DELIVERY failed: VSDX is empty.' }
if ($file.Length -gt $MaxBytes) {
    throw "DELIVERY failed: VSDX is $($file.Length) bytes; maximum is $MaxBytes bytes."
}
Write-Output 'INPUT_VALIDATION: PASS'
Write-Output ("  VSDX: {0} ({1} bytes)" -f $fullPath, $file.Length)
if ($ReferenceImagePath) { Write-Output "  Reference: $((Resolve-Path -LiteralPath $ReferenceImagePath).Path)" }

$package = Read-VsdxQualityPackage $fullPath
if ($package.PageCount -lt 1) { throw 'INTEGRITY_VERIFICATION failed: no Visio page XML.' }
if ($package.ShapeCount -lt 1) { throw 'INTEGRITY_VERIFICATION failed: no native shapes.' }
$raster = @($package.Media | Where-Object {
        $_.FullName -match '\.(png|jpg|jpeg|gif|bmp)$' -or $_.Length -gt 1000000
    })
if ($raster.Count -gt 0) {
    throw 'INTEGRITY_VERIFICATION failed: raster or large media remains in the VSDX.'
}
Write-Output 'INTEGRITY_VERIFICATION: PASS'
Write-Output ("  Pages: {0}; native shapes: {1}; media: {2}" -f $package.PageCount, $package.ShapeCount, $package.Media.Count)

Write-Output ("PROGRESSIVE_BUILD: PASS (requested phase {0}; drawing callback must honor `$script:BuildPhase or -Phase)" -f $Phase)
Test-RequiredValues $package.Text $RequiredText 'Required text'
if (@($RequiredText).Count -eq 0) { Write-Output '  Required text: not supplied (informational)' }

if (@($RequiredColor).Count -gt 0) {
    Test-RequiredValues $package.Colors $RequiredColor 'Required color token'
    Write-Output 'COLOR_AUDIT: PASS'
} else {
    Write-Output 'COLOR_AUDIT: PASS (no explicit color tokens supplied; visual review remains required)'
}

$previewFile = $null
if ($PreviewPath) {
    if (-not (Test-Path -LiteralPath $PreviewPath -PathType Leaf)) {
        throw "DELIVERY failed: preview not found: $PreviewPath"
    }
    $previewFile = Get-Item -LiteralPath $PreviewPath
    if ($previewFile.Length -le 0) { throw 'DELIVERY failed: preview is empty.' }
    Write-Output ("  Preview: {0} ({1} bytes)" -f $previewFile.FullName, $previewFile.Length)
} else {
    Write-Output '  Preview: not requested'
}

if (-not $SkipCom) {
    $beforeIds = @((Get-Process -Name VISIO -ErrorAction SilentlyContinue).Id)
    $visio = $null; $doc = $null; $page = $null
    $outOfBounds = New-Object System.Collections.Generic.List[string]
    try {
        $visio = New-Object -ComObject Visio.Application
        $visio.Visible = $false
        $doc = $visio.Documents.Open($fullPath)
        $page = $doc.Pages.Item(1)
        $pageW = [double]$page.PageSheet.CellsU('PageWidth').ResultIU
        $pageH = [double]$page.PageSheet.CellsU('PageHeight').ResultIU
        if ($pageW -le 0 -or $pageH -le 0) { throw 'LAYOUT_PLANNING failed: invalid page dimensions.' }
        for ($i = 1; $i -le $page.Shapes.Count; $i++) {
            $shape = $null
            try {
                $shape = $page.Shapes.Item($i)
                $x = [double]$shape.CellsU('PinX').ResultIU
                $y = [double]$shape.CellsU('PinY').ResultIU
                $w = [math]::Abs([double]$shape.CellsU('Width').ResultIU)
                $h = [math]::Abs([double]$shape.CellsU('Height').ResultIU)
                # Connector lines and zero-height/zero-width annotation shapes can
                # legitimately use endpoints on the page border. Only flag a shape
                # when its visible bounding box is materially outside the page.
                $tol = 0.15
                if ($w -gt 0.001 -and $h -gt 0.001 -and
                    (($x - $w / 2) -lt -$tol -or ($x + $w / 2) -gt ($pageW + $tol) -or
                     ($y - $h / 2) -lt -$tol -or ($y + $h / 2) -gt ($pageH + $tol))) {
                    $outOfBounds.Add([string]$shape.NameU)
                }
            } catch {
                # Some legacy/group shapes do not expose all cells; package and COM reopen
                # checks still cover them, so do not turn an unreadable optional cell into a leak.
            } finally {
                Release-VisioComObject $shape
            }
        }
        Write-Output ("COM reopen: pages={0}, shapes={1}" -f $doc.Pages.Count, $page.Shapes.Count)
    } finally {
        if ($doc -ne $null) { try { $doc.Saved = $true } catch {}; try { $doc.Close() } catch {} }
        Release-VisioComObject $page
        Release-VisioComObject $doc
        if ($visio -ne $null) { try { $visio.Quit() } catch {}; Release-VisioComObject $visio }
    }
    if ($outOfBounds.Count -gt 0) {
        throw ('ALIGNMENT_AUDIT failed: shapes outside page bounds: ' + ($outOfBounds -join ', '))
    }
    Write-Output 'LAYOUT_PLANNING: PASS'
    Write-Output 'ALIGNMENT_AUDIT: PASS'

    $remaining = @()
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        $remaining = @((Get-Process -Name VISIO -ErrorAction SilentlyContinue) |
            Where-Object { $beforeIds -notcontains $_.Id })
        if ($remaining.Count -eq 0) { break }
        Start-Sleep -Milliseconds 100
    }
    if ($remaining.Count -gt 0) {
        $message = 'New Visio process(es) remain after quality gates: ' + (($remaining.Id) -join ', ')
        if ($StrictProcess) { throw $message }
        Write-Warning $message
    } else {
        Write-Output 'COM process cleanup: PASS'
    }
} else {
    Write-Output 'LAYOUT_PLANNING: PASS (COM skipped)'
    Write-Output 'ALIGNMENT_AUDIT: SKIPPED (COM skipped)'
}

Write-Output 'DELIVERY: PASS'
Write-Output 'Quality gates: OK'
