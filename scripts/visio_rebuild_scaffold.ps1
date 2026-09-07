param(
    [Parameter(Mandatory=$true)]
    [string]$VsdxPath,

    [Parameter(Mandatory=$true)]
    [string]$DrawingScript,

    [ValidateSet('Create', 'Rebuild', 'Edit')]
    [string]$Mode = 'Rebuild',
    [ValidateRange(1, 2147483647)]
    [int]$PageIndex = 1,

    [double]$PageW = 0.0,
    [double]$PageH = 0.0,
    [double]$RefW = 0.0,
    [double]$RefH = 0.0,
    [double]$PageWidthMm = 0.0,
    [double]$PageHeightMm = 0.0,
    [ValidateSet('MatchReference', 'Contain')]
    [string]$CanvasFit = 'MatchReference',
    [double]$MarginMm = 0.0,
    [ValidateRange(1, 2400)][int]$PreviewDpi = 144,
    [ValidateRange(0, 1000)][double]$MinFontPt = 0,
    [ValidateRange(0, 1000)][double]$MinLinePt = 0,
    [ValidateRange(0, 100000)][double]$FinalWidthMm = 0,

    [ValidateSet(1, 2, 3)]
    [int]$Phase = 3,

    [string]$ReferenceImagePath,
    [string[]]$RequiredText,
    [string[]]$RequiredColor,

    [string]$PreviewPath,

    [string[]]$ExportFormats,

    [string]$OutputDir,
    [string]$OutputBaseName,

    [string]$TemplatePath = '',
    [string]$FontName = 'Arial',
    [switch]$KeepBackup,
    [switch]$SkipPreview,
    [switch]$SkipQualityGates,
    [switch]$AllowMedia,
    [switch]$AllowEmptyPage,
    [switch]$Visible
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'visio_export_formats.ps1')
. (Join-Path $PSScriptRoot 'visio_stencil_helpers.ps1')
. (Join-Path $PSScriptRoot 'visio_canvas.ps1')

function VX([double]$x) { $script:Canvas.OffsetX + $x * $script:Canvas.Scale }
function VY([double]$y) { $script:PageH - $script:Canvas.OffsetY - $y * $script:Canvas.Scale }
function VL([double]$length) { $length * $script:Canvas.Scale }
function VPT([double]$length) { (VL $length) * 72.0 }
function RGBF([int]$r, [int]$g, [int]$b) { "RGB($r,$g,$b)" }

$C = @{
    Blue = RGBF 31 95 184
    Purple = RGBF 122 88 166
    Green = RGBF 90 132 64
    Teal = RGBF 20 132 150
    Orange = RGBF 231 116 26
    Black = RGBF 17 17 17
    Gray = RGBF 95 95 95
    White = RGBF 255 255 255
    BlueSoft = RGBF 243 248 255
    PurpleSoft = RGBF 251 248 255
    GreenSoft = RGBF 246 251 241
    OrangeSoft = RGBF 255 248 239
}

function Set-Cell($shape, [string]$cell, [string]$formula) {
    Set-VisioCell $shape $cell $formula
}

function Style-Shape($shape, [string]$fill, [string]$line, [double]$linePt = 0.8, [int]$dash = 1, [double]$roundPx = 0) {
    if ($fill -eq 'none') {
        Set-Cell $shape 'FillPattern' '0'
    } else {
        Set-Cell $shape 'FillPattern' '1'
        Set-Cell $shape 'FillForegnd' $fill
    }
    if ($line -eq 'none') {
        Set-Cell $shape 'LinePattern' '0'
    } else {
        Set-Cell $shape 'LinePattern' ([string]$dash)
        Set-Cell $shape 'LineColor' $line
        Set-Cell $shape 'LineWeight' "$linePt pt"
    }
    if ($roundPx -lt 0) {
        Set-Cell $shape 'Rounding' '0.06 in'
    } else {
        Set-Cell $shape 'Rounding' ((VL $roundPx).ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
    }
}

function Set-Text($shape, [string]$text, [double]$size = 10, [string]$color = $C.Black, [bool]$bold = $false, [bool]$italic = $false, [int]$align = 1) {
    [void]($shape.Text = $text)
    $fontFormula = 'FONT("' + ($script:FontName -replace '"', '""') + '")'
    Set-Cell $shape 'Char.Font' $fontFormula
    Set-Cell $shape 'Char.Size' "$size pt"
    Set-Cell $shape 'Char.Color' $color
    $style = 0
    if ($bold) { $style += 1 }
    if ($italic) { $style += 2 }
    Set-Cell $shape 'Char.Style' ([string]$style)
    Set-Cell $shape 'Para.HorzAlign' ([string]$align)
    Set-Cell $shape 'VerticalAlign' '1'
    foreach ($m in 'LeftMargin','RightMargin','TopMargin','BottomMargin') {
        Set-Cell $shape $m '1 pt'
    }
}

function RectTL([double]$x, [double]$y, [double]$w, [double]$h, [string]$text = '', [string]$fill = 'none', [string]$line = $C.Black, [double]$size = 10, [bool]$bold = $false, [double]$linePt = 0.8, [int]$dash = 1, [double]$roundPx = -1, [switch]$PassThru) {
    $s = $script:Page.DrawRectangle((VX $x), (VY ($y + $h)), (VX ($x + $w)), (VY $y))
    $returned = $false
    try {
        Style-Shape $s $fill $line $linePt $dash $roundPx
        if ($text -ne '') { Set-Text $s $text $size $C.Black $bold }
        if ($PassThru) { $returned = $true; return $s }
    } finally { if (-not $returned) { Release-VisioComObject $s } }
}

function TextTL([double]$x, [double]$y, [double]$w, [double]$h, [string]$text, [double]$size = 10, [string]$color = $C.Black, [bool]$bold = $false, [bool]$italic = $false, [int]$align = 1, [switch]$PassThru) {
    $s = RectTL $x $y $w $h '' 'none' 'none' $size $bold 0 1 0 -PassThru
    $returned = $false
    try {
        Set-Text $s $text $size $color $bold $italic $align
        if ($PassThru) { $returned = $true; return $s }
    } finally { if (-not $returned) { Release-VisioComObject $s } }
}

function OvalTL([double]$x, [double]$y, [double]$w, [double]$h, [string]$text = '', [string]$fill = $C.White, [string]$line = $C.Black, [double]$size = 8, [bool]$bold = $false, [double]$linePt = 0.8, [switch]$PassThru) {
    $s = $script:Page.DrawOval((VX $x), (VY ($y + $h)), (VX ($x + $w)), (VY $y))
    $returned = $false
    try {
        Style-Shape $s $fill $line $linePt 1 0
        if ($text -ne '') { Set-Text $s $text $size $C.Black $bold }
        if ($PassThru) { $returned = $true; return $s }
    } finally { if (-not $returned) { Release-VisioComObject $s } }
}

function LineTL([double]$x1, [double]$y1, [double]$x2, [double]$y2, [string]$color = $C.Black, [double]$linePt = 0.8, [bool]$arrowEnd = $false, [bool]$arrowBegin = $false, [int]$dash = 1, [switch]$PassThru) {
    $s = $script:Page.DrawLine((VX $x1), (VY $y1), (VX $x2), (VY $y2))
    $returned = $false
    try {
        Set-VisioLineStyle $s -Color $color -LinePt $linePt -Dash $dash `
            -EndArrow $(if ($arrowEnd) { 4 } else { 0 }) -BeginArrow $(if ($arrowBegin) { 4 } else { 0 })
        if ($PassThru) { $returned = $true; return $s }
    } finally { if (-not $returned) { Release-VisioComObject $s } }
}

function DotTL([double]$cx, [double]$cy, [double]$r, [string]$fill, [string]$line = $C.White, [switch]$PassThru) {
    return OvalTL ($cx - $r) ($cy - $r) (2 * $r) (2 * $r) '' $fill $line 6 $false 0.4 -PassThru:$PassThru
}

function Assert-RelBox([double]$u, [double]$v, [double]$uw, [double]$vh, [string]$label = 'relative box') {
    if ($u -lt 0 -or $v -lt 0 -or $uw -lt 0 -or $vh -lt 0 -or ($u + $uw) -gt 1 -or ($v + $vh) -gt 1) {
        throw "$label is outside calibrated panel bounds. Use 0-1 local coordinates or enlarge the parent panel."
    }
}

function Assert-RelPoint([double]$u, [double]$v, [string]$label = 'relative point') {
    if ($u -lt 0 -or $v -lt 0 -or $u -gt 1 -or $v -gt 1) {
        throw "$label is outside calibrated panel bounds. Use 0-1 local coordinates or enlarge the parent panel."
    }
}

function RX([double]$x0, [double]$w0, [double]$u) { $x0 + $w0 * $u }
function RY([double]$y0, [double]$h0, [double]$v) { $y0 + $h0 * $v }

function RectRel([double]$x0, [double]$y0, [double]$w0, [double]$h0, [double]$u, [double]$v, [double]$uw, [double]$vh, [string]$text = '', [string]$fill = 'none', [string]$line = $C.Black, [double]$size = 10, [bool]$bold = $false, [double]$linePt = 0.8, [int]$dash = 1, [double]$roundPx = -1, [switch]$PassThru) {
    Assert-RelBox $u $v $uw $vh $text
    return RectTL (RX $x0 $w0 $u) (RY $y0 $h0 $v) ($w0 * $uw) ($h0 * $vh) $text $fill $line $size $bold $linePt $dash $roundPx -PassThru:$PassThru
}

function TextRel([double]$x0, [double]$y0, [double]$w0, [double]$h0, [double]$u, [double]$v, [double]$uw, [double]$vh, [string]$text, [double]$size = 10, [string]$color = $C.Black, [bool]$bold = $false, [bool]$italic = $false, [int]$align = 1, [switch]$PassThru) {
    Assert-RelBox $u $v $uw $vh $text
    return TextTL (RX $x0 $w0 $u) (RY $y0 $h0 $v) ($w0 * $uw) ($h0 * $vh) $text $size $color $bold $italic $align -PassThru:$PassThru
}

function OvalRel([double]$x0, [double]$y0, [double]$w0, [double]$h0, [double]$u, [double]$v, [double]$uw, [double]$vh, [string]$text = '', [string]$fill = $C.White, [string]$line = $C.Black, [double]$size = 8, [bool]$bold = $false, [double]$linePt = 0.8, [switch]$PassThru) {
    Assert-RelBox $u $v $uw $vh $text
    return OvalTL (RX $x0 $w0 $u) (RY $y0 $h0 $v) ($w0 * $uw) ($h0 * $vh) $text $fill $line $size $bold $linePt -PassThru:$PassThru
}

function LineRel([double]$x0, [double]$y0, [double]$w0, [double]$h0, [double]$u1, [double]$v1, [double]$u2, [double]$v2, [string]$color = $C.Black, [double]$linePt = 0.8, [bool]$arrowEnd = $false, [bool]$arrowBegin = $false, [int]$dash = 1, [switch]$PassThru) {
    Assert-RelPoint $u1 $v1 'line start'
    Assert-RelPoint $u2 $v2 'line end'
    return LineTL (RX $x0 $w0 $u1) (RY $y0 $h0 $v1) (RX $x0 $w0 $u2) (RY $y0 $h0 $v2) $color $linePt $arrowEnd $arrowBegin $dash -PassThru:$PassThru
}


$VsdxPath = [IO.Path]::GetFullPath($VsdxPath)
if ([IO.Path]::GetExtension($VsdxPath) -ne '.vsdx') { throw 'Target must be a .vsdx file.' }
if (-not (Test-Path -LiteralPath $DrawingScript -PathType Leaf)) { throw "Drawing script not found: $DrawingScript" }
$DrawingScript = (Resolve-Path -LiteralPath $DrawingScript).Path
$formats = @(Normalize-VisioExportFormats $ExportFormats)
$targetExists = Test-Path -LiteralPath $VsdxPath -PathType Leaf
if ($Mode -eq 'Create' -and $targetExists) { throw 'Create mode refuses to overwrite an existing VSDX. Use Edit or Rebuild explicitly.' }
if ($Mode -eq 'Edit' -and -not $targetExists) { throw 'Edit mode requires an existing VSDX.' }
if ($TemplatePath -and $targetExists) { throw 'TemplatePath applies only when creating a document.' }
foreach ($dimension in @($PageW, $PageH, $PageWidthMm, $PageHeightMm, $RefW, $RefH, $MarginMm)) {
    if ($dimension -lt 0 -or [double]::IsNaN($dimension) -or [double]::IsInfinity($dimension)) {
        throw 'Canvas dimensions must be finite and non-negative (0 selects the default).'
    }
}
if ($ReferenceImagePath) {
    $imageSize = Get-ReferenceImageDimensions $ReferenceImagePath
    if (($RefW -gt 0 -and $RefW -ne $imageSize.Width) -or ($RefH -gt 0 -and $RefH -ne $imageSize.Height)) {
        throw 'Reference dimensions must match the image. Rescale drawing coordinates, not the measured canvas.'
    }
    $RefW = $imageSize.Width; $RefH = $imageSize.Height
}
if (($RefW -gt 0) -xor ($RefH -gt 0)) { throw 'Supply both RefW and RefH, or neither.' }
$canvasArgs = @{ PageW = $PageW; PageH = $PageH; RefW = $RefW; RefH = $RefH
    PageWidthMm = $PageWidthMm; PageHeightMm = $PageHeightMm; CanvasFit = $CanvasFit; MarginMm = $MarginMm }
if ($Mode -ne 'Edit') { $script:Canvas = Resolve-VisioCanvas @canvasArgs }
$effectivePreview = $PreviewPath
if (-not $effectivePreview -and $formats -contains 'png') {
    $effectivePreview = Resolve-VisioExportPath $VsdxPath 'png' $OutputDir $OutputBaseName
}
if ($effectivePreview -and [IO.Path]::GetExtension($effectivePreview) -ne '.png') {
    throw 'PreviewPath must use the .png extension.'
}
if ($ReferenceImagePath -and $effectivePreview -and
    [string]::Equals([IO.Path]::GetFullPath($ReferenceImagePath), [IO.Path]::GetFullPath($effectivePreview), [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Preview/export would overwrite the reference image. Supply a different PreviewPath or output base name.'
}
$outputPaths = @{}
foreach ($format in $formats) {
    $outputPaths[$format] = Resolve-VisioExportPath $VsdxPath $format $OutputDir $OutputBaseName -PreviewPath $effectivePreview
}
if ($effectivePreview) { $outputPaths['png'] = [IO.Path]::GetFullPath($effectivePreview) }
foreach ($outPath in $outputPaths.Values) {
    if (Test-Path -LiteralPath $outPath -PathType Container) { throw "Output path is a directory: $outPath" }
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $outPath))
    if (Test-Path -LiteralPath $outPath -PathType Leaf) {
        $outputStream = [IO.File]::Open($outPath, 'Open', 'ReadWrite', 'None')
        $outputStream.Dispose()
    }
}
$targetDir = Split-Path -Parent $VsdxPath
[void][IO.Directory]::CreateDirectory($targetDir)
$buildDir = Join-Path $targetDir ('.visio-build-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($buildDir)
$stagePath = Join-Path $buildDir 'drawing.vsdx'
$stagePreview = if ($effectivePreview -or -not $SkipPreview) { Join-Path $buildDir 'preview.png' } else { $null }
$visio = $null; $doc = $null; $page = $null; $pageSheet = $null; $pages = $null; $documents = $null
try {
    try {
        # Check the target lock before creating an automation application.
        if ($targetExists) {
            $stream = [IO.File]::Open($VsdxPath, 'Open', 'Read', 'None')
            try {
                $stageStream = [IO.File]::Create($stagePath)
                try { $stream.CopyTo($stageStream) } finally { $stageStream.Dispose() }
            } finally { $stream.Dispose() }
            $originalHash = (Get-FileHash -LiteralPath $stagePath -Algorithm SHA256).Hash
        }
        $visio = New-VisioApplication -Visible:$Visible
        $script:Visio = $visio
        $documents = $visio.Documents
        if (Test-Path -LiteralPath $stagePath) { $doc = $documents.OpenEx($stagePath, 64) }
        else {
            $template = if ($TemplatePath) { Resolve-VisioContentPath $TemplatePath } else { '' }
            $doc = $documents.Add($template)
        }
        $pages = $doc.Pages
        if ($PageIndex -gt $pages.Count) { throw "PageIndex $PageIndex exceeds page count $($pages.Count)." }
        $page = $pages.Item($PageIndex)
        $script:Page = $page
        $existingSize = Get-VisioPageSize $page
        if ($Mode -eq 'Edit') {
            $script:Canvas = Resolve-VisioCanvas @canvasArgs -Edit -ExistingWidth $existingSize.Width -ExistingHeight $existingSize.Height
        }
        $script:PageW = $Canvas.PageW; $script:PageH = $Canvas.PageH
        $script:RefW = $Canvas.RefW; $script:RefH = $Canvas.RefH
        $pageSheet = $page.PageSheet
        if ($Mode -ne 'Edit' -or $PageW -ne $existingSize.Width -or $PageH -ne $existingSize.Height) {
            Set-VisioCell $pageSheet 'PageWidth' ($PageW.ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
            Set-VisioCell $pageSheet 'PageHeight' ($PageH.ToString([Globalization.CultureInfo]::InvariantCulture) + ' in')
        }
        if ($Mode -eq 'Rebuild') {
            $shapes = $page.Shapes
            try {
                while ($shapes.Count -gt 0) {
                    $oldShape = $shapes.Item(1)
                    try { $oldShape.Delete() } finally { Release-VisioComObject $oldShape }
                }
            } finally { Release-VisioComObject $shapes }
        }
        $script:BuildPhase = $Phase
        & {
            . $DrawingScript -LoadDrawing
            if (Get-Command Draw-VisioPage -CommandType Function -ErrorAction SilentlyContinue) {
                Draw-VisioPage -Phase $BuildPhase
            } elseif (Get-Command Draw-ReferenceFigure -CommandType Function -ErrorAction SilentlyContinue) {
                Draw-ReferenceFigure -Phase $BuildPhase
            } else { throw 'DrawingScript must define Draw-VisioPage (or legacy Draw-ReferenceFigure).' }
        } | ForEach-Object {
            if ([Runtime.InteropServices.Marshal]::IsComObject($_)) {
                throw 'Drawing callback emitted a COM object. Assign -PassThru results to a variable and release them; do not send COM objects to the output pipeline.'
            }
            Write-Output $_
        }
        [void]$doc.SaveAs($stagePath)
        if ($stagePreview) {
            Export-VisioPageFormats $doc $page $stagePath @('png') -PreviewPath $stagePreview -PngDpi $PreviewDpi
        }
        Write-Output ("Canvas: page={0:F2}x{1:F2} mm; fit={2}; grid unit={3:F4} pt" -f ($PageW * 25.4), ($PageH * 25.4), $CanvasFit, (VPT 1))
        Write-Output ("Drawing callback completed: mode={0}, page={1}, phase={2}, coordinates={3}x{4}" -f $Mode, $PageIndex, $Phase, $RefW, $RefH)
    } finally {
        try {
            Release-VisioComObject $pageSheet
            Release-VisioComObject $page
            Release-VisioComObject $pages
            if ($doc) {
                try { $doc.Saved = $true; $doc.Close() }
                finally { Release-VisioComObject $doc }
            }
        } finally {
            try { Release-VisioComObject $documents }
            finally { Stop-VisioApplication $visio }
        }
    }
    if (-not $SkipQualityGates) {
        $qualityArgs = @{ VsdxPath = $stagePath; PageIndex = $PageIndex; Phase = $Phase; StrictProcess = $true
            AllowMedia = $AllowMedia; AllowEmptyPage = $AllowEmptyPage; CanvasFit = $CanvasFit; MarginMm = $MarginMm
            MinFontPt = $MinFontPt; MinLinePt = $MinLinePt; FinalWidthMm = $FinalWidthMm }
        if ($ReferenceImagePath) { $qualityArgs.ReferenceImagePath = $ReferenceImagePath }
        if ($stagePreview) { $qualityArgs.PreviewPath = $stagePreview }
        if ($RequiredText) { $qualityArgs.RequiredText = $RequiredText }
        if ($RequiredColor) { $qualityArgs.RequiredColor = $RequiredColor }
        & (Join-Path $PSScriptRoot 'visio_quality_gates.ps1') @qualityArgs
    } else { Write-Warning 'Quality gates explicitly skipped; visual review is still required.' }
    $remainingFormats = @($formats | Where-Object { $_ -ne 'png' })
    if ($remainingFormats.Count -gt 0) {
        Export-VisioDocumentFormats -VsdxPath $stagePath -Formats $remainingFormats `
            -OutputDir $buildDir -OutputBaseName 'export' -PageIndex $PageIndex
    }
    if ($targetExists) {
        if ((Get-FileHash -LiteralPath $VsdxPath -Algorithm SHA256).Hash -ne $originalHash) {
            throw 'Target changed during the build; refusing to overwrite concurrent edits.'
        }
        $backup = if ($KeepBackup) {
            Join-Path $targetDir (([IO.Path]::GetFileNameWithoutExtension($VsdxPath)) + '.backup-' + [guid]::NewGuid().ToString('N') + '.vsdx')
        } else { [NullString]::Value }
        [IO.File]::Replace($stagePath, $VsdxPath, $backup)
        if ($KeepBackup) { Write-Output "Retained backup: $backup" }
    } else { [IO.File]::Move($stagePath, $VsdxPath) }
    if ($effectivePreview) {
        $effectivePreview = [IO.Path]::GetFullPath($effectivePreview)
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $effectivePreview))
        Copy-Item -LiteralPath $stagePreview -Destination $effectivePreview -Force
        Write-Output "PNG: $effectivePreview"
    }
    foreach ($format in $remainingFormats) {
        $outPath = $outputPaths[$format]
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $outPath))
        Copy-Item -LiteralPath (Join-Path $buildDir "export.$format") -Destination $outPath -Force
        Write-Output "Export: $outPath"
    }
    Write-Output "Saved: $VsdxPath"
    Write-Output 'VISUAL_REVIEW: REQUIRED (automated checks do not establish reference fidelity)'
} finally {
    # This directory was created by this invocation, under the verified target directory.
    if ([IO.Directory]::Exists($buildDir)) { [IO.Directory]::Delete($buildDir, $true) }
}
