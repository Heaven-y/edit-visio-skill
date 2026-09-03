function Release-VisioComObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$ComObject
    )

    if ($ComObject -ne $null) {
        try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) | Out-Null } catch {}
    }
}

function Get-VisioContentRoots {
    [CmdletBinding()]
    param(
        [string]$PreferredRoot
    )

    $roots = New-Object System.Collections.Generic.List[string]
    foreach ($root in @(
        $PreferredRoot,
        $(if ($env:ProgramFiles) { Join-Path $env:ProgramFiles 'Microsoft Office\root\Office16\Visio Content' }),
        $(if (${env:ProgramFiles(x86)}) { Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\Visio Content' })
    )) {
        if ($root -and -not $roots.Contains($root)) { $roots.Add($root) }
    }
    return @($roots | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
}

function Resolve-VisioContentPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string]$RootPath
    )

    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        return (Resolve-Path -LiteralPath $Path).Path
    }

    $leaf = [IO.Path]::GetFileName($Path)
    foreach ($root in Get-VisioContentRoots -PreferredRoot $RootPath) {
        $match = Get-ChildItem -LiteralPath $root -Recurse -File -Filter $leaf -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($match) { return $match.FullName }
    }
    throw "Visio content not found: $Path"
}

function Resolve-VisioStencilPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string]$RootPath = 'C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052'
    )

    try {
        return Resolve-VisioContentPath -Path $Path -RootPath $RootPath
    } catch {
        throw "Visio stencil not found: $Path"
    }
}

function Open-VisioStencil {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Visio,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string]$RootPath = 'C:\Program Files\Microsoft Office\root\Office16\Visio Content\2052'
    )

    $resolved = Resolve-VisioStencilPath -Path $Path -RootPath $RootPath
    try {
        # 64 = visOpenRO | visOpenDocked. Never open a stencil for editing.
        return $Visio.Documents.OpenEx($resolved, 64)
    } catch {
        throw "Failed to open Visio stencil '$resolved': $($_.Exception.Message)"
    }
}

function Get-VisioStencilMaster {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Stencil,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $masters = $null
    try {
        try {
            # Item(string) is the fast path and supports the localized display name.
            return $Stencil.Masters.Item($Name)
        } catch {}

        $masters = $Stencil.Masters
        $matches = @()
        for ($i = 1; $i -le $masters.Count; $i++) {
            $master = $null
            $keepMaster = $false
            try {
                $master = $masters.Item($i)
                if ([string]::Equals([string]$master.Name, $Name, [StringComparison]::OrdinalIgnoreCase) -or
                    [string]::Equals([string]$master.NameU, $Name, [StringComparison]::OrdinalIgnoreCase)) {
                    $keepMaster = $true
                    return $master
                }
                if ($matches.Count -lt 12) {
                    $matches += if ($master.Name) { [string]$master.Name } else { [string]$master.NameU }
                }
            } catch {
                # A corrupt or legacy master should not hide the remaining candidates.
                continue
            } finally {
                # Do not release a matching master that is being returned to the caller.
                if ($master -ne $null -and -not $keepMaster) {
                    Release-VisioComObject $master
                }
            }
        }
        $hint = if ($matches.Count -gt 0) { $matches -join ', ' } else { '(no readable masters)' }
        throw "Master '$Name' not found in '$($Stencil.Name)'. Candidates: $hint"
    } finally {
        Release-VisioComObject $masters
    }
}

function Find-VisioStencilMaster {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Stencil,

        [Parameter(Mandatory = $true)]
        [string]$Query,

        [int]$MaxResults = 20
    )

    if ([string]::IsNullOrWhiteSpace($Query)) { throw 'Stencil query cannot be empty.' }
    $masters = $null
    $results = New-Object System.Collections.Generic.List[object]
    try {
        $masters = $Stencil.Masters
        for ($i = 1; $i -le $masters.Count; $i++) {
            $master = $null
            try {
                $master = $masters.Item($i)
                $name = [string]$master.Name
                $nameU = [string]$master.NameU
                $nameMatch = $name.IndexOf($Query, [StringComparison]::OrdinalIgnoreCase)
                $nameUMatch = $nameU.IndexOf($Query, [StringComparison]::OrdinalIgnoreCase)
                if ($nameMatch -ge 0 -or $nameUMatch -ge 0) {
                    $score = if ([string]::Equals($nameU, $Query, [StringComparison]::OrdinalIgnoreCase)) { 3 }
                        elseif ([string]::Equals($name, $Query, [StringComparison]::OrdinalIgnoreCase)) { 2 }
                        elseif ($nameUMatch -eq 0 -or $nameMatch -eq 0) { 1 }
                        else { 0 }
                    $results.Add([pscustomobject]@{
                        Name = $name
                        NameU = $nameU
                        ID = [string]$master.ID
                        Score = $score
                    })
                }
            } catch {
                continue
            } finally {
                Release-VisioComObject $master
            }
        }
    } finally {
        Release-VisioComObject $masters
    }
    return @($results.ToArray() |
        Sort-Object -Property @{ Expression = 'Score'; Descending = $true }, NameU |
        Select-Object -First $MaxResults)
}

function Drop-VisioStencilMaster {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Page,

        [Parameter(Mandatory = $true)]
        [object]$Stencil,

        [Parameter(Mandatory = $true)]
        [string]$MasterName,

        [Parameter(Mandatory = $true)]
        [double]$PinX,

        [Parameter(Mandatory = $true)]
        [double]$PinY,

        [double]$Width,
        [double]$Height
    )

    $master = $null
    try {
        $master = Get-VisioStencilMaster -Stencil $Stencil -Name $MasterName
        $shape = $Page.Drop($master, $PinX, $PinY)
        if ($Width -gt 0) { $shape.CellsU('Width').ResultIU = $Width }
        if ($Height -gt 0) { $shape.CellsU('Height').ResultIU = $Height }
        $shape.CellsU('PinX').ResultIU = $PinX
        $shape.CellsU('PinY').ResultIU = $PinY
        return $shape
    } finally {
        Release-VisioComObject $master
    }
}

function Close-VisioStencil {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Stencil
    )

    if ($Stencil -eq $null) { return }
    try { $Stencil.Saved = $true } catch {}
    try { $Stencil.Close() } catch {}
    Release-VisioComObject $Stencil
}
