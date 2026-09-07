param(
    [Parameter(Mandatory = $true)][string]$VsdxPath,
    [string[]]$RequiredText,
    [string[]]$RequiredColor,
    [string]$ReferenceImagePath,
    [string]$PreviewPath,
    [ValidateRange(1, 2147483647)][int]$PageIndex = 1,
    [switch]$AllowMedia,
    [switch]$AllowEmptyPage,
    [ValidateSet('MatchReference', 'Contain')][string]$CanvasFit = 'MatchReference',
    [double]$MarginMm = 0,
    [ValidateRange(0, 1000)][double]$MinFontPt = 0,
    [ValidateRange(0, 1000)][double]$MinLinePt = 0,
    [ValidateRange(0, 100000)][double]$FinalWidthMm = 0,
    [ValidateSet(1, 2, 3)][int]$Phase = 3,
    [long]$MaxBytes = 0,
    [switch]$SkipCom,
    [switch]$StrictProcess
)

# Preserve the public entry point while maintaining one implementation of validation.
& (Join-Path $PSScriptRoot 'visio_quality_gates.ps1') @PSBoundParameters
