param([switch]$WithCom, [switch]$KeepArtifacts)

$ErrorActionPreference = 'Stop'
$skillRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $skillRoot 'scripts/visio_runtime.ps1')
. (Join-Path $skillRoot 'scripts/visio_canvas.ps1')
. (Join-Path $skillRoot 'scripts/visio_export_formats.ps1')
. (Join-Path $skillRoot 'scripts/visio_package.ps1')
$scaffold = Join-Path $skillRoot 'scripts/visio_rebuild_scaffold.ps1'
$gates = Join-Path $skillRoot 'scripts/visio_quality_gates.ps1'
$script:checks = 0

function Assert-True([bool]$Condition, [string]$Label) {
    if (-not $Condition) { throw "FAILED: $Label" }
    $script:checks++
    Write-Output "PASS: $Label"
}
function Assert-Near([double]$Actual, [double]$Expected, [string]$Label, [double]$Tolerance = 0.000001) {
    Assert-True ([math]::Abs($Actual - $Expected) -le $Tolerance) "$Label ($Actual vs $Expected)"
}
function Assert-Throws([scriptblock]$Action, [string]$Pattern, [string]$Label) {
    $message = ''
    try { & $Action | Out-Null } catch { $message = $_.Exception.Message }
    Assert-True ($message -match $Pattern) "$Label [$message]"
}
function Cell-Value($Shape, [string]$Name) {
    $cell = $null
    try { $cell = $Shape.CellsU($Name); return [double]$cell.ResultIU }
    finally { Release-VisioComObject $cell }
}
function Raster-Settings($Settings) {
    [int]$r = 0; [double]$rw = 0; [double]$rh = 0; [int]$ru = 0
    [int]$s = 0; [double]$sw = 0; [double]$sh = 0; [int]$su = 0
    $Settings.GetRasterExportResolution([ref]$r, [ref]$rw, [ref]$rh, [ref]$ru)
    $Settings.GetRasterExportSize([ref]$s, [ref]$sw, [ref]$sh, [ref]$su)
    return (@($r, $rw, $rh, $ru, $s, $sw, $sh, $su) -join '|')
}

Get-ChildItem -LiteralPath (Join-Path $skillRoot 'scripts') -Filter '*.ps1' | ForEach-Object {
    $tokens = $null; $parseErrors = $null
    [void][Management.Automation.Language.Parser]::ParseFile($_.FullName, [ref]$tokens, [ref]$parseErrors)
    Assert-True ($parseErrors.Count -eq 0) "syntax $($_.Name)"
}
$canvas = Resolve-VisioCanvas -RefW 1200 -RefH 600 -PageWidthMm 180
Assert-Near ($canvas.PageH * 25.4) 90 'width-only physical canvas'
Assert-Near ($canvas.Scale * 72) (180 / 25.4 * 72 / 1200) 'reference length to points'
$canvas = Resolve-VisioCanvas -RefW 600 -RefH 1200 -PageHeightMm 180
Assert-Near ($canvas.PageW * 25.4) 90 'height-only portrait canvas'
$canvas = Resolve-VisioCanvas -RefW 1200 -RefH 600 -PageWidthMm 210 -PageHeightMm 297 -CanvasFit Contain -MarginMm 10
Assert-Near ($canvas.ContentW * 25.4) 190 'contain width with margin'
Assert-Near ($canvas.ContentH * 25.4) 95 'contain preserves aspect'
Assert-Near ($canvas.OffsetX * 25.4) 10 'contain horizontal offset'
Assert-Near ($canvas.OffsetY * 25.4) 101 'contain vertical offset'
$canvas = Resolve-VisioCanvas -Edit -ExistingWidth 8 -ExistingHeight 5
Assert-Near $canvas.PageW 8 'edit preserves width'
Assert-Near $canvas.RefH 5 'edit default coordinate grid'
Assert-Throws { Resolve-VisioCanvas -PageW 8 -PageWidthMm 180 -RefW 2 -RefH 1 } 'inches OR' 'conflicting units rejected'
Assert-Throws { Resolve-VisioCanvas -PageW 8 -PageH 5 -RefW 2 -RefH 1 } 'aspect ratio' 'stretch rejected'
Assert-Throws { Resolve-VisioCanvas -PageW 8 -PageH 5 -MarginMm 1 } 'requires' 'implicit crop/padding rejected'
Assert-Throws { Resolve-VisioCanvas -PageW 1 -PageH 1 -CanvasFit Contain -MarginMm 20 } 'no drawable' 'oversized margin rejected'
Assert-Throws { Resolve-VisioCanvas -PageW ([double]::NaN) -PageH 1 } 'finite' 'nonfinite canvas rejected'

$catalogScratch = Join-Path ([IO.Path]::GetTempPath()) ('visio-catalog-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($catalogScratch)
try {
    foreach ($catalogCase in @(
        @{ Name = 'duplicate-names.vssx'; Xml = '<Masters><Master ID="1" Name="Same" NameU="First"/><Master ID="2" Name="Same" NameU="Second"/></Masters>' },
        @{ Name = 'unsafe.vssx'; Xml = '<!DOCTYPE Masters [<!ENTITY item "unsafe">]><Masters>&item;</Masters>' }
    )) {
        $archive = [IO.Compression.ZipFile]::Open((Join-Path $catalogScratch $catalogCase.Name), 'Create')
        try {
            $entry = $archive.CreateEntry('visio/masters/masters.xml')
            $writer = [IO.StreamWriter]::new($entry.Open())
            try { $writer.Write($catalogCase.Xml) } finally { $writer.Dispose() }
        } finally { $archive.Dispose() }
    }
    $catalog = & (Join-Path $skillRoot 'scripts/visio_stencil_catalog.ps1') -RootPath $catalogScratch -Format json | ConvertFrom-Json
    $duplicate = $catalog | Where-Object File -eq 'duplicate-names.vssx'
    $unsafe = $catalog | Where-Object File -eq 'unsafe.vssx'
    Assert-True ($duplicate.MasterCount -eq 2) 'catalog preserves distinct IDs sharing a display name'
    Assert-True ($null -eq $unsafe.MasterCount -and $unsafe.Note -match 'Failed to read') 'unsafe XML reported unavailable, not empty'
} finally { [IO.Directory]::Delete($catalogScratch, $true) }

if (-not $WithCom) { Write-Output "REGRESSION: PASS ($script:checks checks; COM explicitly skipped)"; return }
Assert-VisioComHost
$scratch = Join-Path ([IO.Path]::GetTempPath()) ('visio-regression-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($scratch)
$ownedBefore = @(Get-Process VISIO -ErrorAction SilentlyContinue | ForEach-Object Id)
try {
    $drawing = Join-Path $scratch 'draw.ps1'
    # Runtime-generated fixture, not a user document or a domain-specific example.
    @'
param([switch]$LoadDrawing)
function Draw-VisioPage([int]$Phase) {
    $a = $null; $b = $null; $edge = $null
    try {
        TextRel 0 0 $RefW $RefH 0.05 0.04 0.9 0.10 'Request workflow' 12
        $a = RectRel 0 0 $RefW $RefH 0.08 0.30 0.29 0.24 'Request' $C.BlueSoft $C.Blue 10 $false 0.8 -PassThru
        $b = RectRel 0 0 $RefW $RefH 0.62 0.46 0.29 0.24 'Review' $C.GreenSoft $C.Green 10 $false 0.8 -PassThru
        $a.NameU = 'Request'; $b.NameU = 'Review'
        Style-Shape $a $C.BlueSoft $C.Blue 0.8 1 0
        $labelX = if ($RefH -gt $RefW) { 14 } else { 0 }
        $labelY = if ($RefH -gt $RefW) { 0 } else { 8 }
        $edge = Connect-VisioShapes $a $b -FromSide Right -ToSide Left -Routing Orthogonal -EndArrow 4 `
            -LabelOffsetXPt $labelX -LabelOffsetYPt $labelY -PassThru
        $edge.NameU = 'RequestToReview'
        Set-Text $edge 'next' 8
        $radius = OvalRel 0 0 $RefW $RefH 0.08 0.72 0.1 (0.1 * $RefW / $RefH) '' $C.Teal 'none' -PassThru
        try { $radius.NameU = 'ScaleCircle' } finally { Release-VisioComObject $radius }
        TextRel 0 0 $RefW $RefH 0.30 0.78 0.58 0.12 'Native editable nodes' 9
    } finally {
        foreach ($item in @($edge, $a, $b)) { Release-VisioComObject $item }
    }
}
'@ | Set-Content -LiteralPath $drawing -Encoding utf8
    $landscape = Join-Path $scratch 'landscape.png'
    $portrait = Join-Path $scratch 'portrait.png'
    Add-Type -AssemblyName System.Drawing
    foreach ($imageCase in @(@($landscape, 1200, 600), @($portrait, 600, 1200))) {
        $bitmap = [Drawing.Bitmap]::new($imageCase[1], $imageCase[2])
        try { $bitmap.Save($imageCase[0], [Drawing.Imaging.ImageFormat]::Png) }
        finally { $bitmap.Dispose() }
    }
    $workflow = Join-Path $scratch 'workflow.vsdx'
    $workflowPreview = Join-Path $scratch 'workflow-preview.png'
    & $scaffold -Mode Create -VsdxPath $workflow -DrawingScript $drawing -ReferenceImagePath $landscape `
        -PageWidthMm 180 -PreviewPath $workflowPreview -MinFontPt 7 -MinLinePt 0.5
    $previewSize = Get-ReferenceImageDimensions $workflowPreview
    Assert-Near $previewSize.Width (180 / 25.4 * 144) '144 DPI PNG width' 3
    Assert-Near $previewSize.Height (90 / 25.4 * 144) '144 DPI PNG height' 3
    Assert-True ((Read-VsdxPackage $workflow).Media.Count -eq 0) 'native-only workflow'
    $portraitPath = Join-Path $scratch 'portrait.vsdx'
    & $scaffold -Mode Create -VsdxPath $portraitPath -DrawingScript $drawing -ReferenceImagePath $portrait `
        -PageHeightMm 180 -PreviewPath (Join-Path $scratch 'portrait-preview.png')
    $containedPath = Join-Path $scratch 'contained.vsdx'
    & $scaffold -Mode Create -VsdxPath $containedPath -DrawingScript $drawing -ReferenceImagePath $landscape `
        -PageWidthMm 210 -PageHeightMm 297 -CanvasFit Contain -MarginMm 10 `
        -PreviewPath (Join-Path $scratch 'contained-preview.png') -PreviewDpi 96
    Assert-Throws { & $gates -VsdxPath $workflow -MinFontPt 7 -FinalWidthMm 90 } 'STYLE_SIZE failed' 'font reduction detected at final width'
    Assert-Throws { & $gates -VsdxPath $workflow -MinLinePt 0.5 -FinalWidthMm 90 } 'STYLE_SIZE failed' 'thin lines detected at final width'
    $beforeHash = (Get-FileHash -LiteralPath $workflow).Hash
    Assert-Throws { & $scaffold -Mode Create -VsdxPath $workflow -DrawingScript $drawing } 'refuses to overwrite' 'create refuses existing target'
    $badDrawing = Join-Path $scratch 'fail.ps1'
    'param([switch]$LoadDrawing); function Draw-VisioPage([int]$Phase) { RectTL 1 1 1 1; throw "intentional failure" }' |
        Set-Content -LiteralPath $badDrawing -Encoding utf8
    Assert-Throws { & $scaffold -Mode Rebuild -VsdxPath $workflow -DrawingScript $badDrawing -PageW 8 -PageH 4 } 'intentional failure' 'failed rebuild rejected'
    Assert-True ((Get-FileHash -LiteralPath $workflow).Hash -eq $beforeHash) 'failed rebuild preserves original bytes'
    Assert-Throws { & $scaffold -Mode Rebuild -VsdxPath $workflow -DrawingScript $drawing -ReferenceImagePath $landscape -PageWidthMm 180 -MinFontPt 20 } 'STYLE_SIZE failed' 'failed quality gate blocks replacement'
    Assert-True ((Get-FileHash -LiteralPath $workflow).Hash -eq $beforeHash) 'quality rejection preserves original bytes'
    Assert-Throws { & $scaffold -Mode Rebuild -VsdxPath $workflow -DrawingScript $drawing -ReferenceImagePath $landscape -PreviewPath $landscape } 'overwrite the reference' 'source reference protected'
    Assert-Throws { & $scaffold -Mode Rebuild -VsdxPath $workflow -DrawingScript $drawing -ReferenceImagePath $landscape -RefW 100 -RefH 50 } 'must match the image' 'measured dimensions cannot be overridden'
    Assert-True (@(Get-ChildItem -LiteralPath $scratch -Directory -Filter '.visio-build-*').Count -eq 0) 'failed staging cleaned'

    $app = $null; $docs = $null; $doc = $null; $pages = $null; $page = $null; $shapes = $null
    $a = $null; $b = $null; $edge = $null; $circle = $null; $otherPage = $null; $otherShape = $null; $straight = $null
    $nestedPath = Join-Path $scratch 'nested.vsdx'
    try {
        $app = New-VisioApplication; $docs = $app.Documents
        $doc = $docs.OpenEx($workflow, 64); $pages = $doc.Pages; $page = $pages.Item(1); $shapes = $page.Shapes
        $a = $shapes.ItemU('Request'); $b = $shapes.ItemU('Review'); $edge = $shapes.ItemU('RequestToReview'); $circle = $shapes.ItemU('ScaleCircle')
        Assert-Near (Cell-Value $a 'Rounding') 0 'zero radius resets previous rounding'
        Assert-Near (Cell-Value $circle 'Width') (Cell-Value $circle 'Height') 'uniform canvas keeps circle round'
        Assert-Near (Cell-Value $edge 'ShapeRouteStyle') 1 'orthogonal route selected'
        $oldBeginY = Cell-Value $edge 'BeginY'
        Set-VisioCell $a 'PinY' ((Cell-Value $a 'PinY') + 0.2).ToString([Globalization.CultureInfo]::InvariantCulture)
        Assert-Near (Cell-Value $edge 'BeginY') ($oldBeginY + 0.2) 'glued connector follows moved node'
        Assert-Near (Cell-Value $edge 'BeginX') ((Cell-Value $a 'PinX') + (Cell-Value $a 'Width') / 2) 'right-side attachment'
        $straight = Connect-VisioShapes $b $a -FromSide Bottom -ToSide Bottom -Routing Straight -Dash 2 -EndArrow 4 -PassThru
        Assert-Near (Cell-Value $straight 'ShapeRouteStyle') 2 'straight route selected'
        Assert-Near (Cell-Value $straight 'LinePattern') 2 'dashed connector selected'
        $straight.Delete()
        $otherPage = $pages.Add(); $otherPage.NameU = 'UnchangedPage'
        $otherPage.Background = $true
        $page.BackPage = $otherPage.NameU
        $otherShape = $otherPage.DrawRectangle(0.1, 0.05, 6.8, 0.3); $otherShape.Text = 'Preserve this page'
        Set-VisioCell $otherShape 'LinePattern' '0'
        Set-VisioCell $otherShape 'FillPattern' '0'
        Set-VisioCell $otherShape 'Char.Size' '6 pt'
        Assert-Throws { Connect-VisioShapes $a $otherShape } 'same page' 'cross-page connection rejected'
        [void]$doc.Save()
        $settings = $app.Settings
        $originalSettings = Raster-Settings $settings
        try {
            Assert-Throws { Export-VisioPng $page (Join-Path $scratch 'missing/export.png') -Dpi 200 } '.' 'PNG export failure reported'
            Assert-True ((Raster-Settings $settings) -eq $originalSettings) 'raster settings restored after failure'
        } finally { Release-VisioComObject $settings }
        Export-VisioPageFormats $doc $page $workflow @('png', 'svg', 'pdf') -OutputDir $scratch -OutputBaseName 'export' -PngDpi 300
        $settings = $app.Settings
        try { Assert-True ((Raster-Settings $settings) -eq $originalSettings) 'raster settings restored after success' }
        finally { Release-VisioComObject $settings }
        $exportSize = Get-ReferenceImageDimensions (Join-Path $scratch 'export.png')
        Assert-Near $exportSize.Width (180 / 25.4 * 300) '300 DPI exported width' 3
        $nestedDoc = $null; $nestedPages = $null; $nestedPage = $null
        $small = $null; $normal = $null; $selection = $null; $group = $null
        try {
            $nestedDoc = $docs.Add(''); $nestedPages = $nestedDoc.Pages; $nestedPage = $nestedPages.Item(1)
            $small = $nestedPage.DrawRectangle(1, 1, 3, 2); $small.Text = 'Nested small label'
            Set-VisioCell $small 'Char.Size' '4 pt'
            $normal = $nestedPage.DrawRectangle(1, 3, 3, 4); $normal.Text = 'Normal label'
            Set-VisioCell $normal 'Char.Size' '10 pt'
            $selection = $nestedPage.CreateSelection(1, 0, 0)
            $group = $selection.Group()
            [void]$nestedDoc.SaveAs($nestedPath)
        } finally {
            foreach ($item in @($group, $selection, $normal, $small, $nestedPage, $nestedPages)) { Release-VisioComObject $item }
            if ($nestedDoc) { try { $nestedDoc.Saved = $true; $nestedDoc.Close() } finally { Release-VisioComObject $nestedDoc } }
        }
    } finally {
        try {
            foreach ($item in @($straight, $otherShape, $otherPage, $circle, $edge, $b, $a, $shapes, $page, $pages)) { Release-VisioComObject $item }
            if ($doc) { try { $doc.Saved = $true; $doc.Close() } finally { Release-VisioComObject $doc } }
        } finally { Release-VisioComObject $docs; Stop-VisioApplication $app }
    }
    Assert-Throws { & $gates -VsdxPath $nestedPath -MinFontPt 7 } '4.000 pt < 7' 'nested small text detected'
    $editDrawing = Join-Path $scratch 'edit.ps1'
    @'
param([switch]$LoadDrawing)
function Draw-VisioPage([int]$Phase) {
    $shapes = $script:Page.Shapes; $shape = $null
    try { $shape = $shapes.ItemU('Request'); $shape.Text = 'Changed request' }
    finally { Release-VisioComObject $shape; Release-VisioComObject $shapes }
}
'@ | Set-Content -LiteralPath $editDrawing -Encoding utf8
    & $scaffold -Mode Edit -VsdxPath $workflow -DrawingScript $editDrawing -RequiredText 'Changed request'
    $package = Read-VsdxPackage $workflow
    Assert-True ($package.PageCount -eq 2 -and $package.Pages[1].Text -contains 'Preserve this page') 'edit preserves other page'
    Assert-True $package.Pages[1].Background 'background page identity preserved'
    $backgroundCheck = & $gates -VsdxPath $workflow -PageIndex 2 -RequiredText 'Preserve this page'
    Assert-True ($backgroundCheck -contains 'COM_REOPEN: PASS') 'background selection matches COM'
    $blank = Join-Path $scratch 'blank.vsdx'
    $blankDrawing = Join-Path $scratch 'blank.ps1'
    'param([switch]$LoadDrawing); function Draw-VisioPage([int]$Phase) {}' | Set-Content -LiteralPath $blankDrawing -Encoding utf8
    & $scaffold -Mode Create -VsdxPath $blank -DrawingScript $blankDrawing -PageW 3 -PageH 2 -AllowEmptyPage
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $scratch 'blank.png'))) 'no default PNG deliverable'
    Assert-Throws { & $gates -VsdxPath $blank -SkipCom } 'no shapes' 'blank page rejected by default'
    Assert-True (@(Get-ChildItem -LiteralPath $scratch -Filter '*.backup-*').Count -eq 0) 'no default backups'
    $ownedAfter = @(Get-Process VISIO -ErrorAction SilentlyContinue | ForEach-Object Id)
    foreach ($newId in @($ownedAfter | Where-Object { $_ -notin $ownedBefore })) {
        $process = Get-Process -Id $newId -ErrorAction SilentlyContinue
        if ($process) { try { [void]$process.WaitForExit(5000) } finally { $process.Dispose() } }
    }
    $ownedAfter = @(Get-Process VISIO -ErrorAction SilentlyContinue | ForEach-Object Id)
    Assert-True (@($ownedAfter | Where-Object { $_ -notin $ownedBefore }).Count -eq 0) 'no remaining owned Visio processes'
    Write-Output "REGRESSION: PASS ($script:checks checks)"
} finally {
    if ($KeepArtifacts) { Write-Output "Test artifacts: $scratch" }
    elseif ([IO.Directory]::Exists($scratch)) { [IO.Directory]::Delete($scratch, $true) }
}
