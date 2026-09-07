param(
    [Parameter(Mandatory = $true)][string]$VsdxPath,
    [string[]]$RequiredText,
    [string[]]$RequiredColor,
    [string]$ReferenceImagePath,
    [string]$PreviewPath,
    [ValidateRange(1, 2147483647)][int]$PageIndex = 1,
    [switch]$AllowMedia,
    [ValidateSet(1, 2, 3)][int]$Phase = 3,
    [long]$MaxBytes = 0,
    [switch]$SkipCom,
    [switch]$StrictProcess
)

# Preserve the public entry point while maintaining one implementation of validation.
& (Join-Path $PSScriptRoot 'visio_quality_gates.ps1') @PSBoundParameters
