. (Join-Path $PSScriptRoot 'visio_runtime.ps1')

$script:VisioExportSupportedFormats = @('png', 'svg', 'pdf', 'pptx')

function Normalize-VisioExportFormats {
    param([string[]]$Formats)

    $normalized = New-Object System.Collections.Generic.List[string]
    foreach ($formatGroup in $Formats) {
        if (-not $formatGroup) { continue }
        foreach ($format in ($formatGroup -split ',')) {
            if (-not $format) { continue }
            $name = $format.Trim().TrimStart('.').ToLowerInvariant()
            if ($script:VisioExportSupportedFormats -notcontains $name) {
                throw "Unsupported export format '$format'. Supported formats: $($script:VisioExportSupportedFormats -join ', ')"
            }
            if (-not $normalized.Contains($name)) {
                $normalized.Add($name)
            }
        }
    }
    return @($normalized)
}

function Resolve-VisioExportPath {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [Parameter(Mandatory=$true)]
        [string]$Format,

        [string]$OutputDir,
        [string]$OutputBaseName,
        [string]$PreviewPath
    )

    if ($Format -eq 'png' -and $PreviewPath) {
        if ([IO.Path]::GetExtension($PreviewPath) -ne '.png') { throw 'PreviewPath must use the .png extension.' }
        return [IO.Path]::GetFullPath($PreviewPath)
    }

    if (-not $OutputDir) {
        $OutputDir = Split-Path -Parent $SourcePath
    }
    if (-not $OutputBaseName) {
        $OutputBaseName = [IO.Path]::GetFileNameWithoutExtension($SourcePath)
    }
    if ([IO.Path]::GetFileName($OutputBaseName) -ne $OutputBaseName) {
        throw 'OutputBaseName must be a filename stem, not a path. Use OutputDir for directories.'
    }
    if ($OutputBaseName.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) {
        throw 'OutputBaseName contains invalid filename characters.'
    }

    return [IO.Path]::GetFullPath((Join-Path $OutputDir "$OutputBaseName.$Format"))
}

function Export-VisioPdf {
    param(
        [Parameter(Mandatory=$true)]
        $Document,

        [Parameter(Mandatory=$true)]
        [string]$OutPath
    )

    $dir = Split-Path -Parent $OutPath
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }

    try {
        # 1 = PDF, 1 = print quality, 0 = all pages. Numeric constants avoid requiring Visio interop assemblies.
        [void]$Document.ExportAsFixedFormat(1, $OutPath, 1, 0)
    } catch {
        throw "PDF export failed: $($_.Exception.Message)"
    }
}

function Export-VisioPng {
    param(
        [Parameter(Mandatory = $true)]$Page,
        [Parameter(Mandatory = $true)][string]$OutPath,
        [ValidateRange(1, 2400)][int]$Dpi = 144
    )
    $size = Get-VisioPageSize $Page
    if ($size.Width * $size.Height * $Dpi * $Dpi -gt 100000000) {
        throw 'PNG would exceed 100 megapixels. Select a lower DPI or export a vector format.'
    }
    $app = $null; $settings = $null; $captured = $false
    [int]$resolution = 0; [double]$resW = 0; [double]$resH = 0; [int]$resUnits = 0
    [int]$rasterSize = 0; [double]$sizeW = 0; [double]$sizeH = 0; [int]$sizeUnits = 0
    try {
        $app = $Page.Application
        $settings = $app.Settings
        $settings.GetRasterExportResolution([ref]$resolution, [ref]$resW, [ref]$resH, [ref]$resUnits)
        $settings.GetRasterExportSize([ref]$rasterSize, [ref]$sizeW, [ref]$sizeH, [ref]$sizeUnits)
        $captured = $true
        # 3 = custom resolution, 0 = pixels/inch; 2 = source page size.
        $settings.SetRasterExportResolution(3, $Dpi, $Dpi, 0)
        $settings.SetRasterExportSize(2, 0, 0, 0)
        [void]$Page.Export($OutPath)
    } finally {
        try {
            if ($captured) {
                try { $settings.SetRasterExportResolution($resolution, $resW, $resH, $resUnits) }
                finally { $settings.SetRasterExportSize($rasterSize, $sizeW, $sizeH, $sizeUnits) }
            }
        } finally {
            Release-VisioComObject $settings
            Release-VisioComObject $app
        }
    }
}

function Export-VisioPptx {
    param(
        [Parameter(Mandatory=$true)]
        $Page,

        [Parameter(Mandatory=$true)]
        [string]$OutPath
    )

    # PowerPoint COM can reuse an interactive application instead of isolating it.
    if (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue) {
        throw 'PPTX export cannot own an isolated PowerPoint session while PowerPoint is already running. Close it yourself or request SVG/PDF.'
    }
    $dir = Split-Path -Parent $OutPath
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }

    $scratch = Join-Path ([IO.Path]::GetTempPath()) ("visio-export-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $scratch | Out-Null
    $svgPath = Join-Path $scratch 'page.svg'

    $powerPoint = $null
    $presentation = $null
    $slide = $null
    $picture = $null
    $presentations = $null; $pageSetup = $null; $slides = $null; $slideShapes = $null
    try {
        [void]$Page.Export($svgPath)
        $size = Get-VisioPageSize $Page
        $pageWidthPt = $size.Width * 72.0
        $pageHeightPt = $size.Height * 72.0

        try {
            $powerPoint = New-Object -ComObject PowerPoint.Application
        } catch {
            throw "PowerPoint COM is required for PPTX export: $($_.Exception.Message)"
        }

        $presentations = $powerPoint.Presentations
        $presentation = $presentations.Add($false)
        $pageSetup = $presentation.PageSetup
        $pageSetup.SlideWidth = $pageWidthPt
        $pageSetup.SlideHeight = $pageHeightPt

        # 12 = ppLayoutBlank, 24 = ppSaveAsOpenXMLPresentation.
        $slides = $presentation.Slides
        $slide = $slides.Add(1, 12)
        $slideShapes = $slide.Shapes
        $picture = $slideShapes.AddPicture($svgPath, $false, $true, 0, 0, $pageWidthPt, $pageHeightPt)
        $presentation.SaveAs($OutPath, 24)
    } catch {
        throw "PPTX export failed: $($_.Exception.Message)"
    } finally {
        if ($presentation -ne $null) {
            try { $presentation.Saved = -1; $presentation.Close() } catch {}
        }
        foreach ($comObject in @($picture, $slideShapes, $slide, $slides, $pageSetup, $presentation, $presentations)) {
            if ($comObject -ne $null) {
                try { [Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) | Out-Null } catch {}
            }
        }
        if ($powerPoint -ne $null) {
            try { $powerPoint.Quit() } catch {}
            try { [Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null } catch {}
        }
        if (Test-Path -LiteralPath $scratch) {
            Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Export-VisioPageFormats {
    param(
        [Parameter(Mandatory=$true)]
        $Document,

        [Parameter(Mandatory=$true)]
        $Page,

        [Parameter(Mandatory=$true)]
        [string]$SourcePath,

        [string[]]$Formats = @('png'),
        [string]$OutputDir,
        [string]$OutputBaseName,
        [string]$PreviewPath,
        [ValidateRange(1, 2400)][int]$PngDpi = 144
    )

    $formatsToExport = Normalize-VisioExportFormats $Formats
    $exportBoundsShape = $null
    try {
        # Visio Page.Export may crop to visible geometry. A temporary page-sized
        # transparent shape keeps the page bounds without covering background pages;
        # it is deleted before the caller can save the document again.
        $size = Get-VisioPageSize $Page
        $pageWidth = $size.Width
        $pageHeight = $size.Height
        if ($pageWidth -gt 0 -and $pageHeight -gt 0 -and
            ($formatsToExport -contains 'png' -or $formatsToExport -contains 'svg' -or $formatsToExport -contains 'pptx')) {
            $exportBoundsShape = $Page.DrawRectangle(0, 0, $pageWidth, $pageHeight)
            try {
                Set-VisioCell $exportBoundsShape 'FillPattern' '1'
                Set-VisioCell $exportBoundsShape 'FillForegnd' 'RGB(255,255,255)'
                Set-VisioCell $exportBoundsShape 'FillForegndTrans' '100%'
                Set-VisioCell $exportBoundsShape 'LinePattern' '0'
                $exportBoundsShape.SendToBack()
            } catch {
                # If an older Visio build lacks SendToBack, fail explicitly rather
                # than silently exporting a shape that could cover the artwork.
                throw "Unable to prepare page-bounds export shape: $($_.Exception.Message)"
            }
        }

        foreach ($format in $formatsToExport) {
            $outPath = Resolve-VisioExportPath -SourcePath $SourcePath -Format $format -OutputDir $OutputDir -OutputBaseName $OutputBaseName -PreviewPath $PreviewPath
            if ([string]::Equals([IO.Path]::GetFullPath($SourcePath), $outPath, [StringComparison]::OrdinalIgnoreCase)) {
                throw 'Export must not overwrite the source document.'
            }
            [void][IO.Directory]::CreateDirectory((Split-Path -Parent $outPath))
            switch ($format) {
                'png' { Export-VisioPng $Page $outPath -Dpi $PngDpi }
                'svg' { [void]$Page.Export($outPath) }
                'pdf' { Export-VisioPdf -Document $Document -OutPath $outPath }
                'pptx' { Export-VisioPptx -Page $Page -OutPath $outPath }
            }

            $bytes = (Get-Item -LiteralPath $outPath).Length
            Write-Output ("{0}: {1} ({2} bytes)" -f $format.ToUpperInvariant(), $outPath, $bytes)
        }
    } finally {
        if ($exportBoundsShape -ne $null) {
            try { [void]$exportBoundsShape.Delete() } catch {}
            try { [Runtime.InteropServices.Marshal]::ReleaseComObject($exportBoundsShape) | Out-Null } catch {}
        }
    }
}

function Export-VisioDocumentFormats {
    param(
        [Parameter(Mandatory=$true)]
        [string]$VsdxPath,

        [string[]]$Formats = @('png'),
        [string]$OutputDir,
        [string]$OutputBaseName,
        [ValidateRange(1, 2147483647)][int]$PageIndex = 1,
        [string]$PreviewPath,
        [ValidateRange(1, 2400)][int]$PngDpi = 144,
        [switch]$Visible
    )

    $visio = $null
    $doc = $null
    $page = $null
    $documents = $null
    $pages = $null
    try {
        $visio = New-VisioApplication -Visible:$Visible
        $documents = $visio.Documents
        $doc = $documents.OpenEx([IO.Path]::GetFullPath($VsdxPath), 66)
        $pages = $doc.Pages
        if ($PageIndex -gt $pages.Count) { throw "PageIndex $PageIndex exceeds page count $($pages.Count)." }
        $page = $pages.Item($PageIndex)
        Export-VisioPageFormats -Document $doc -Page $page -SourcePath $VsdxPath -Formats $Formats -OutputDir $OutputDir -OutputBaseName $OutputBaseName -PreviewPath $PreviewPath -PngDpi $PngDpi
    } finally {
        try {
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
}
