function Resolve-VisioCanvas {
    param(
        [double]$PageW = 0, [double]$PageH = 0,
        [double]$PageWidthMm = 0, [double]$PageHeightMm = 0,
        [double]$RefW = 0, [double]$RefH = 0,
        [ValidateSet('MatchReference', 'Contain')][string]$CanvasFit = 'MatchReference',
        [double]$MarginMm = 0,
        [double]$ExistingWidth = 0, [double]$ExistingHeight = 0,
        [switch]$Edit
    )
    foreach ($value in @($PageW, $PageH, $PageWidthMm, $PageHeightMm, $RefW, $RefH,
        $MarginMm, $ExistingWidth, $ExistingHeight)) {
        if ($value -lt 0 -or [double]::IsNaN($value) -or [double]::IsInfinity($value)) {
            throw 'Canvas dimensions must be finite and non-negative; 0 selects the default.'
        }
    }
    if (($PageW -gt 0 -and $PageWidthMm -gt 0) -or ($PageH -gt 0 -and $PageHeightMm -gt 0)) {
        throw 'Specify each page dimension in inches OR millimeters, not both.'
    }
    if ($PageWidthMm -gt 0) { $PageW = $PageWidthMm / 25.4 }
    if ($PageHeightMm -gt 0) { $PageH = $PageHeightMm / 25.4 }
    if (($RefW -gt 0) -xor ($RefH -gt 0)) { throw 'Supply both RefW and RefH, or neither.' }
    if ($MarginMm -gt 0 -and $CanvasFit -ne 'Contain') { throw 'MarginMm requires CanvasFit Contain.' }
    if ($Edit) {
        if ($PageW -eq 0) { $PageW = $ExistingWidth }
        if ($PageH -eq 0) { $PageH = $ExistingHeight }
    }
    if ($RefW -eq 0) {
        if ($PageW -le 0 -or $PageH -le 0) { throw 'Without a reference, supply both page dimensions or RefW/RefH.' }
        $RefW = $PageW; $RefH = $PageH
    }
    $margin = $MarginMm / 25.4
    # Retain the legacy scale only when neither physical dimension was supplied.
    if ($PageW -eq 0 -and $PageH -eq 0) { $PageW = 16.0 }
    if ($PageW -eq 0) { $PageW = ($PageH - 2 * $margin) * $RefW / $RefH + 2 * $margin }
    if ($PageH -eq 0) { $PageH = ($PageW - 2 * $margin) * $RefH / $RefW + 2 * $margin }
    $innerW = $PageW - 2 * $margin; $innerH = $PageH - 2 * $margin
    if ($innerW -le 0 -or $innerH -le 0) { throw 'Margins leave no drawable page area.' }
    if ($CanvasFit -eq 'MatchReference' -and [math]::Abs(($innerW / $innerH) / ($RefW / $RefH) - 1) -gt 0.000001) {
        throw 'Page aspect ratio differs from the reference. Use CanvasFit Contain for intentional whitespace, or correct the page dimensions.'
    }
    $scale = [math]::Min($innerW / $RefW, $innerH / $RefH)
    [pscustomobject]@{
        PageW = $PageW; PageH = $PageH; RefW = $RefW; RefH = $RefH
        Scale = $scale
        OffsetX = ($PageW - $RefW * $scale) / 2
        OffsetY = ($PageH - $RefH * $scale) / 2
        ContentW = $RefW * $scale; ContentH = $RefH * $scale
        CanvasFit = $CanvasFit; MarginMm = $MarginMm
    }
}
