function Read-VisioPackageXml($Archive, [string]$PartName) {
    $entry = $Archive.GetEntry($PartName)
    if (-not $entry) { throw "Missing Visio package part: $PartName" }
    $stream = $entry.Open()
    $reader = $null
    try {
        $settings = [Xml.XmlReaderSettings]::new()
        $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $reader = [Xml.XmlReader]::Create($stream, $settings)
        $xml = [Xml.XmlDocument]::new()
        $xml.XmlResolver = $null
        $xml.Load($reader)
        return ,$xml
    } finally {
        if ($reader) { $reader.Dispose() }
        $stream.Dispose()
    }
}

function Read-VsdxPackage([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead([IO.Path]::GetFullPath($Path))
    try {
        [void](Read-VisioPackageXml $zip '[Content_Types].xml')
        [void](Read-VisioPackageXml $zip 'visio/document.xml')
        $pagesXml = Read-VisioPackageXml $zip 'visio/pages/pages.xml'
        $relsXml = Read-VisioPackageXml $zip 'visio/pages/_rels/pages.xml.rels'
        $rels = @{}
        foreach ($rel in $relsXml.DocumentElement.ChildNodes) {
            if ($rel.LocalName -eq 'Relationship') { $rels[$rel.Id] = $rel }
        }
        $pages = [Collections.Generic.List[object]]::new()
        $totalShapes = 0
        $foreignCount = 0
        $baseUri = [uri]'https://vsdx.invalid/visio/pages/pages.xml'
        $definitions = @($pagesXml.SelectNodes("/*[local-name()='Pages']/*[local-name()='Page']"))
        # XML may serialize background dependencies first; COM lists foreground
        # pages before backgrounds, preserving order within each group.
        $orderedDefinitions = @($definitions | Where-Object { $_.GetAttribute('Background') -notin @('1', 'true') }) +
            @($definitions | Where-Object { $_.GetAttribute('Background') -in @('1', 'true') })
        foreach ($definition in $orderedDefinitions) {
            # Resolve relationships rather than assuming pageN.xml equals tab N.
            $relNode = $definition.SelectSingleNode("*[local-name()='Rel']")
            if (-not $relNode) { throw 'Page definition has no relationship.' }
            $relId = $relNode.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $rel = $rels[$relId]
            if (-not $rel -or $rel.TargetMode -eq 'External') { throw "Invalid page relationship: $relId" }
            $partUri = [uri]::new($baseUri, [string]$rel.Target)
            $partName = [uri]::UnescapeDataString($partUri.AbsolutePath.TrimStart('/'))
            if ($partUri.Authority -ne $baseUri.Authority -or $partUri.Scheme -ne $baseUri.Scheme -or
                $partName -notmatch '^visio/pages/[^/]+\.xml$') {
                throw "Invalid page part target: $($rel.Target)"
            }
            $xml = Read-VisioPackageXml $zip $partName
            $shapeCount = $xml.SelectNodes("//*[local-name()='Shape']").Count
            $pageForeign = $xml.SelectNodes("//*[local-name()='Shape' and @Type='Foreign']").Count
            $texts = [Collections.Generic.List[string]]::new()
            $colors = [Collections.Generic.List[string]]::new()
            foreach ($node in $xml.SelectNodes("//*[local-name()='Text']")) {
                $value = [regex]::Replace($node.InnerText, '\s+', ' ').Trim()
                if ($value) { $texts.Add($value) }
            }
            foreach ($cell in $xml.SelectNodes("//*[local-name()='Cell']")) {
                if ($cell.N -notin @('FillForegnd', 'LineColor', 'Char.Color', 'Color')) { continue }
                foreach ($attribute in @('F', 'Formula', 'V')) {
                    $value = $cell.GetAttribute($attribute)
                    if ($value) { $colors.Add($value) }
                }
            }
            $pages.Add([pscustomobject]@{
                Index = $pages.Count + 1
                ID = [string]$definition.ID
                NameU = [string]$definition.NameU
                Background = $definition.GetAttribute('Background') -in @('1', 'true')
                PartName = $partName
                ShapeCount = $shapeCount
                ForeignCount = $pageForeign
                Text = $texts.ToArray()
                Colors = $colors.ToArray()
            })
            $totalShapes += $shapeCount
            $foreignCount += $pageForeign
        }
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -match '^visio/masters/master\d+\.xml$') {
                $xml = Read-VisioPackageXml $zip $entry.FullName
                $foreignCount += $xml.SelectNodes("//*[local-name()='Shape' and @Type='Foreign']").Count
            }
        }
        $media = @($zip.Entries | Where-Object { $_.FullName -like 'visio/media/*' -and $_.Name } |
            ForEach-Object { [pscustomobject]@{ FullName = $_.FullName; Length = $_.Length } })
        return [pscustomobject]@{
            Pages = $pages.ToArray()
            PageCount = $pages.Count
            ShapeCount = $totalShapes
            ForeignCount = $foreignCount
            Media = $media
        }
    } finally { $zip.Dispose() }
}
