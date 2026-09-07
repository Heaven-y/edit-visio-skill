param(
    [Parameter(Mandatory = $true)]
    [string]$VsdxPath,

    [string]$ReferenceImagePath,
    [string]$PreviewPath,
    [string[]]$RequiredText,
    [string[]]$RequiredColor,

    [ValidateRange(1, 2147483647)]
    [int]$PageIndex = 1,
    [switch]$AllowMedia,

    [ValidateSet(1, 2, 3)]
    [int]$Phase = 3,

    [long]$MaxBytes = 0,
    [switch]$SkipCom,
    # Compatibility switch: owned-process cleanup is always verified now.
    [switch]$StrictProcess
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'visio_stencil_helpers.ps1')
. (Join-Path $PSScriptRoot 'visio_package.ps1')
$RequiredText = @($RequiredText | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$RequiredColor = @($RequiredColor | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

function Test-RequiredValues([string[]]$Values, [string[]]$Required, [string]$Label) {
    foreach ($item in @($Required)) {
        if ([string]::IsNullOrWhiteSpace($item)) { continue }
        $normalizedItem = $item.Trim()
        if ($normalizedItem -match '^RGB\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$') {
            $normalizedItem = '#{0:X2}{1:X2}{2:X2}' -f [int]$Matches[1], [int]$Matches[2], [int]$Matches[3]
        }
        $found = @($Values | Where-Object {
                $candidate = [string]$_
                if ($candidate -match '^RGB\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$') {
                    $candidate = '#{0:X2}{1:X2}{2:X2}' -f [int]$Matches[1], [int]$Matches[2], [int]$Matches[3]
                }
                $_.IndexOf($item, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $candidate.IndexOf($normalizedItem, [StringComparison]::OrdinalIgnoreCase) -ge 0
            }).Count -gt 0
        if (-not $found) { throw "$Label not found: $item" }
        Write-Output "  $Label found: $item"
    }
}

function Assert-AspectRatio([double]$ActualWidth, [double]$ActualHeight,
    [double]$ExpectedWidth, [double]$ExpectedHeight, [string]$Label,
    [double]$Tolerance = 0.015) {
    if ($ActualWidth -le 0 -or $ActualHeight -le 0 -or
        $ExpectedWidth -le 0 -or $ExpectedHeight -le 0) {
        throw "ASPECT_RATIO failed: invalid dimensions for $Label."
    }
    $actualRatio = $ActualWidth / $ActualHeight
    $expectedRatio = $ExpectedWidth / $ExpectedHeight
    $relativeError = [math]::Abs($actualRatio - $expectedRatio) / $expectedRatio
    if ($relativeError -gt $Tolerance) {
        throw ("ASPECT_RATIO failed for {0}: actual {1:F6}, expected {2:F6}, error {3:P2}." -f
            $Label, $actualRatio, $expectedRatio, $relativeError)
    }
    Write-Output ("  Aspect ratio {0}: {1:F6} (expected {2:F6})" -f $Label, $actualRatio, $expectedRatio)
}

if (-not (Test-Path -LiteralPath $VsdxPath -PathType Leaf)) {
    throw "INPUT_VALIDATION failed: VSDX not found: $VsdxPath"
}
$fullPath = (Resolve-Path -LiteralPath $VsdxPath).Path
if ($ReferenceImagePath -and -not (Test-Path -LiteralPath $ReferenceImagePath -PathType Leaf)) {
    throw "INPUT_VALIDATION failed: reference image not found: $ReferenceImagePath"
}
$referenceSize = $null
if ($ReferenceImagePath) {
    $referenceSize = Get-ReferenceImageDimensions $ReferenceImagePath
    if ($referenceSize.Width -le 0 -or $referenceSize.Height -le 0) {
        throw 'INPUT_VALIDATION failed: reference image has invalid dimensions.'
    }
}

$file = Get-Item -LiteralPath $fullPath
if ($file.Length -le 0) { throw 'DELIVERY failed: VSDX is empty.' }
if ($MaxBytes -gt 0 -and $file.Length -gt $MaxBytes) {
    throw "DELIVERY failed: VSDX is $($file.Length) bytes; maximum is $MaxBytes bytes."
}
Write-Output 'INPUT_VALIDATION: PASS'
Write-Output ("  VSDX: {0} ({1} bytes)" -f $fullPath, $file.Length)
if ($ReferenceImagePath) { Write-Output "  Reference: $((Resolve-Path -LiteralPath $ReferenceImagePath).Path)" }

$package = Read-VsdxPackage $fullPath
if ($package.PageCount -lt 1) { throw 'INTEGRITY_VERIFICATION failed: no Visio page XML.' }
if ($PageIndex -gt $package.PageCount) { throw "PageIndex $PageIndex exceeds page count $($package.PageCount)." }
$selectedPage = $package.Pages[$PageIndex - 1]
if ($selectedPage.ShapeCount -lt 1) { throw 'INTEGRITY_VERIFICATION failed: selected page has no shapes.' }
if (-not $AllowMedia -and ($package.Media.Count -gt 0 -or $package.ForeignCount -gt 0)) {
    throw 'INTEGRITY_VERIFICATION failed: non-native media remains in the VSDX.'
}
Write-Output 'INTEGRITY_VERIFICATION: PASS'
Write-Output ("  Pages: {0}; shapes: {1}; media: {2}; foreign shapes (pages/masters): {3}" -f
    $package.PageCount, $package.ShapeCount, $package.Media.Count, $package.ForeignCount)
Write-Output ("  Selected page: {0} ({1}, {2})" -f $PageIndex, $selectedPage.NameU, $selectedPage.PartName)
if ($AllowMedia) { Write-Output 'NATIVE_CONTENT: SKIPPED (media explicitly allowed; full native editability is not established)' }
else { Write-Output 'NATIVE_CONTENT: PASS' }

Write-Output ("PHASE_SCOPE: INFO (requested phase {0}; package alone cannot verify callback completeness)" -f $Phase)
Test-RequiredValues $selectedPage.Text $RequiredText 'Required text'
if (@($RequiredText).Count -eq 0) { Write-Output '  Required text: not supplied (informational)' }

if (@($RequiredColor).Count -gt 0) {
    Test-RequiredValues $selectedPage.Colors $RequiredColor 'Required color token'
    Write-Output 'COLOR_AUDIT: PASS'
} else {
    Write-Output 'COLOR_AUDIT: SKIPPED (no explicit color tokens supplied)'
}

$previewFile = $null
if ($PreviewPath) {
    if (-not (Test-Path -LiteralPath $PreviewPath -PathType Leaf)) {
        throw "DELIVERY failed: preview not found: $PreviewPath"
    }
    $previewFile = Get-Item -LiteralPath $PreviewPath
    if ($previewFile.Length -le 0) { throw 'DELIVERY failed: preview is empty.' }
    $previewSize = Get-ReferenceImageDimensions $PreviewPath
    if ($referenceSize) {
        Assert-AspectRatio $previewSize.Width $previewSize.Height $referenceSize.Width $referenceSize.Height 'preview/reference'
    }
    Write-Output ("  Preview: {0} ({1} bytes)" -f $previewFile.FullName, $previewFile.Length)
} else {
    Write-Output '  Preview: not requested'
}


if (-not $SkipCom) {
    $visio = $null; $documents = $null; $doc = $null; $pages = $null; $page = $null; $shapes = $null
    $outOfBounds = New-Object System.Collections.Generic.List[string]
    try {
        $visio = New-VisioApplication
        $documents = $visio.Documents
        $doc = $documents.OpenEx($fullPath, 66)
        $pages = $doc.Pages
        $page = $pages.Item($PageIndex)
        if ([string]$page.ID -ne $selectedPage.ID) {
            throw 'Package page order does not match Visio. Inspect page IDs before validating this selection.'
        }
        $size = Get-VisioPageSize $page
        $pageW = $size.Width; $pageH = $size.Height
        if ($pageW -le 0 -or $pageH -le 0) { throw 'Invalid page dimensions.' }
        if ($referenceSize) {
            Assert-AspectRatio $pageW $pageH $referenceSize.Width $referenceSize.Height 'page/reference'
        }
        if ($previewFile) {
            Assert-AspectRatio $previewSize.Width $previewSize.Height $pageW $pageH 'preview/page'
        }
        $shapes = $page.Shapes
        for ($i = 1; $i -le $shapes.Count; $i++) {
            $shape = $shapes.Item($i)
            try {
                [double]$left = 0; [double]$bottom = 0; [double]$right = 0; [double]$top = 0
                # Drawing-coordinate geometry bounds include rotation, groups and lines.
                $shape.BoundingBox(8196, [ref]$left, [ref]$bottom, [ref]$right, [ref]$top)
                $tol = 0.02
                if ($left -lt -$tol -or $bottom -lt -$tol -or $right -gt ($pageW + $tol) -or $top -gt ($pageH + $tol)) {
                    $outOfBounds.Add("page $PageIndex / $($shape.NameU)")
                }
            } finally { Release-VisioComObject $shape }
        }
        Write-Output ("COM reopen: page={0}, top-level shapes={1}, size={2:F4}x{3:F4}in" -f $PageIndex, $shapes.Count, $pageW, $pageH)
    } finally {
        try {
            Release-VisioComObject $shapes
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
    if ($outOfBounds.Count -gt 0) {
        throw ('PAGE_BOUNDS failed: ' + ($outOfBounds -join ', '))
    }
    Write-Output 'COM_REOPEN: PASS'
    Write-Output 'PAGE_BOUNDS: PASS (not an overlap or text-fit test)'
} else {
    Write-Output 'COM_REOPEN: SKIPPED'
    Write-Output 'PAGE_BOUNDS: SKIPPED'
}
Write-Output 'AUTOMATED_CHECKS: PASS'
Write-Output 'VISUAL_REVIEW: REQUIRED (inspect text, icons, layout and meaning against the reference)'
