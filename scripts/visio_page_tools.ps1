param(
    [Parameter(Mandatory=$true)]
    [string]$VsdxPath,

    [string]$PreviewPath,

    [string[]]$ExportFormats,

    [string]$OutputDir,
    [string]$OutputBaseName,
    [ValidateRange(1, 2147483647)][int]$PageIndex = 1,

    [switch]$Backup,
    [switch]$ExportPreview,
    [switch]$InspectPackage,
    [switch]$CloseOpenDocument,
    [switch]$SaveOpenDocument,
    [switch]$DiscardOpenDocument,
    [switch]$Visible
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'visio_export_formats.ps1')
. (Join-Path $PSScriptRoot 'visio_package.ps1')

function Close-VisioDocument([string]$path, [switch]$Save, [switch]$Discard) {
    if ($Save -and $Discard) { throw 'Choose save OR discard, not both.' }
    Assert-VisioComHost
    $targetPath = [IO.Path]::GetFullPath($path)
    $visio = $null; $documents = $null
    try {
        try { $visio = Get-ActiveVisioApplication }
        catch { Write-Output 'No active Visio application.'; return }
        $documents = $visio.Documents
        for ($i = $documents.Count; $i -ge 1; $i--) {
            $doc = $documents.Item($i)
            try {
                if (-not [string]::Equals([string]$doc.FullName, $targetPath, [StringComparison]::OrdinalIgnoreCase)) { continue }
                if (-not $doc.Saved -and -not $Save -and -not $Discard) {
                    throw 'Target has unsaved edits. Explicitly choose -SaveOpenDocument or -DiscardOpenDocument.'
                }
                if ($Save) { $doc.Save() }
                if ($Discard) { $doc.Saved = $true }
                $doc.Close()
                Write-Output "Closed target document: $targetPath"
                return
            } finally { Release-VisioComObject $doc }
        }
        Write-Output 'Target was not found in the active Visio application.'
    } finally {
        Release-VisioComObject $documents
        # The active UI belongs to the user, even if no documents remain.
        Release-VisioComObject $visio
    }
}

function Backup-Vsdx([string]$path) {
    $fullPath = (Resolve-Path -LiteralPath $path).Path
    $stem = [IO.Path]::GetFileNameWithoutExtension($fullPath)
    $backupPath = Join-Path (Split-Path -Parent $fullPath) ($stem + '.backup-' + [guid]::NewGuid().ToString('N') + '.vsdx')
    Copy-Item -LiteralPath $fullPath -Destination $backupPath
    Write-Output "Backup: $backupPath"
}

function Inspect-VsdxPackage([string]$path) {
    $package = Read-VsdxPackage $path
    Write-Output "Pages: $($package.PageCount); total shapes: $($package.ShapeCount)"
    foreach ($page in $package.Pages) {
        Write-Output ("Page {0}: {1}; shapes={2}; foreign={3}; part={4}" -f
            $page.Index, $page.NameU, $page.ShapeCount, $page.ForeignCount, $page.PartName)
    }
    foreach ($media in $package.Media) {
        Write-Output ("Media: {0} ({1} bytes)" -f $media.FullName, $media.Length)
    }
    Write-Output "Media entries: $($package.Media.Count)"
}

if (($SaveOpenDocument -or $DiscardOpenDocument) -and -not $CloseOpenDocument) {
    throw 'Save/discard requires -CloseOpenDocument.'
}
if ($CloseOpenDocument) {
    Close-VisioDocument $VsdxPath -Save:$SaveOpenDocument -Discard:$DiscardOpenDocument
}
if ($Backup) { Backup-Vsdx $VsdxPath }
if ($InspectPackage) { Inspect-VsdxPackage $VsdxPath }

$formatsToExport = New-Object System.Collections.Generic.List[string]
if ($ExportPreview -and -not $formatsToExport.Contains('png')) {
    $formatsToExport.Add('png') | Out-Null
}
foreach ($format in @($ExportFormats)) {
    if ($format -and -not $formatsToExport.Contains($format.ToLowerInvariant())) {
        $formatsToExport.Add($format.ToLowerInvariant()) | Out-Null
    }
}

if ($formatsToExport.Count -gt 0) {
    Export-VisioDocumentFormats `
        -VsdxPath $VsdxPath `
        -Formats @($formatsToExport) `
        -OutputDir $OutputDir `
        -OutputBaseName $OutputBaseName `
        -PageIndex $PageIndex `
        -PreviewPath $PreviewPath `
        -Visible:$Visible
}
