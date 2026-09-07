function Release-VisioComObject([AllowNull()][object]$ComObject) {
    if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Get-ReferenceImageDimensions([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Reference image not found: $Path"
    }
    Add-Type -AssemblyName System.Drawing
    $image = [System.Drawing.Image]::FromFile((Resolve-Path -LiteralPath $Path).Path)
    try {
        [pscustomobject]@{ Width = [double]$image.Width; Height = [double]$image.Height }
    } finally { $image.Dispose() }
}

function Assert-VisioComHost {
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        throw 'This Visio COM backend requires PowerShell 7 (pwsh). Windows PowerShell 5.1 is not supported; no Visio session was started.'
    }
    if (-not $IsWindows) { throw 'Visio COM automation requires Windows and licensed Microsoft Visio.' }
}

function New-VisioApplication([switch]$Visible) {
    Assert-VisioComHost
    # InvisibleApp creates an isolated automation session, not the user's active UI.
    $app = New-Object -ComObject Visio.InvisibleApp
    try {
        $app.AlertResponse = 7
        if ($Visible) { $app.Visible = $true }
        # Return one COM object without collection enumeration.
        return ,$app
    } catch {
        $app.Quit()
        Release-VisioComObject $app
        throw
    }
}

function Initialize-VisioNativeMethods {
    if (-not ('VisioRuntime.NativeMethods' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace VisioRuntime {
    public static class NativeMethods {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
        [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        public static extern void CLSIDFromProgID(string progId, out Guid clsid);
        [DllImport("oleaut32.dll", PreserveSig = false)]
        public static extern void GetActiveObject(ref Guid clsid, IntPtr reserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object instance);
    }
}
'@
    }
}

function Get-ActiveVisioApplication {
    Assert-VisioComHost
    Initialize-VisioNativeMethods
    $clsid = [guid]::Empty
    [VisioRuntime.NativeMethods]::CLSIDFromProgID('Visio.Application', [ref]$clsid)
    $instance = $null
    [VisioRuntime.NativeMethods]::GetActiveObject([ref]$clsid, [intptr]::Zero, [ref]$instance)
    return ,$instance
}

function Get-VisioProcessId($Application) {
    Initialize-VisioNativeMethods
    # Visio.Application.ProcessID is an internal Visio ID, not the Windows PID.
    [uint32]$ownedId = 0
    [void][VisioRuntime.NativeMethods]::GetWindowThreadProcessId([intptr][long]$Application.WindowHandle32, [ref]$ownedId)
    if ($ownedId -eq 0) { throw 'Could not identify the owned Visio Windows process.' }
    return [int]$ownedId
}

function Stop-VisioApplication([AllowNull()][object]$Application) {
    if ($null -eq $Application) { return }
    try { $ownedId = Get-VisioProcessId $Application }
    finally {
        try { $Application.Quit() }
        finally { Release-VisioComObject $Application }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $process = Get-Process -Id $ownedId -ErrorAction SilentlyContinue
    if ($process) {
        try {
            if (-not $process.WaitForExit(5000)) {
                throw "Owned Visio process $ownedId did not exit; no user process was terminated."
            }
        } finally { $process.Dispose() }
    }
    Write-Output "COM process cleanup: PASS (owned PID $ownedId)"
}

function Set-VisioCell($Shape, [string]$Name, [string]$Formula) {
    $cell = $null
    try { $cell = $Shape.CellsU($Name); $cell.FormulaU = $Formula }
    catch { throw "Cannot set Visio cell '$Name' to '$Formula': $($_.Exception.Message)" }
    finally { Release-VisioComObject $cell }
}

function Get-VisioPageSize($Page) {
    $sheet = $null; $width = $null; $height = $null
    try {
        $sheet = $Page.PageSheet
        $width = $sheet.CellsU('PageWidth')
        $height = $sheet.CellsU('PageHeight')
        return [pscustomobject]@{ Width = [double]$width.ResultIU; Height = [double]$height.ResultIU }
    } finally {
        foreach ($item in @($width, $height, $sheet)) { Release-VisioComObject $item }
    }
}

function Set-VisioLineStyle {
    param(
        [Parameter(Mandatory = $true)]$Shape,
        [string]$Color = 'RGB(17,17,17)',
        [ValidateRange(0, 1000)][double]$LinePt = 0.8,
        [ValidateRange(0, 23)][int]$Dash = 1,
        [ValidateRange(0, 45)][int]$BeginArrow = 0,
        [ValidateRange(0, 45)][int]$EndArrow = 0,
        [ValidateRange(0, 6)][int]$ArrowSize = 2
    )
    Set-VisioCell $Shape 'LinePattern' ([string]$Dash)
    Set-VisioCell $Shape 'LineColor' $Color
    Set-VisioCell $Shape 'LineWeight' ($LinePt.ToString([Globalization.CultureInfo]::InvariantCulture) + ' pt')
    Set-VisioCell $Shape 'BeginArrow' ([string]$BeginArrow)
    Set-VisioCell $Shape 'EndArrow' ([string]$EndArrow)
    Set-VisioCell $Shape 'BeginArrowSize' ([string]$ArrowSize)
    Set-VisioCell $Shape 'EndArrowSize' ([string]$ArrowSize)
}

function Set-VisioEndpointGlue($Cell, $Shape, [string]$Side) {
    if ($Side -eq 'Auto') {
        $pin = $null
        try { $pin = $Shape.CellsU('PinX'); [void]$Cell.GlueTo($pin) }
        finally { Release-VisioComObject $pin }
        return
    }
    $point = switch ($Side) {
        'Left' { @(0.0, 0.5) }
        'Right' { @(1.0, 0.5) }
        'Top' { @(0.5, 1.0) }
        'Bottom' { @(0.5, 0.0) }
    }
    # GlueToPos creates a real connection point in the target's local coordinates.
    [void]$Cell.GlueToPos($Shape, $point[0], $point[1])
}

function Connect-VisioShapes {
    param(
        [Parameter(Mandatory = $true)]$From,
        [Parameter(Mandatory = $true)]$To,
        [switch]$PassThru,
        [ValidateSet('Auto', 'Left', 'Right', 'Top', 'Bottom')][string]$FromSide = 'Auto',
        [ValidateSet('Auto', 'Left', 'Right', 'Top', 'Bottom')][string]$ToSide = 'Auto',
        [ValidateSet('Auto', 'Orthogonal', 'Straight')][string]$Routing = 'Auto',
        [string]$Color = 'RGB(17,17,17)',
        [double]$LinePt = 0.8,
        [int]$Dash = 1,
        [int]$BeginArrow = 0,
        [int]$EndArrow = 0,
        [int]$ArrowSize = 2,
        [ValidateRange(-1000, 1000)][double]$LabelOffsetXPt = 0,
        [ValidateRange(-1000, 1000)][double]$LabelOffsetYPt = 0
    )
    $page = $null; $app = $null; $tool = $null; $connector = $null
    $begin = $null; $end = $null; $toPage = $null
    $returned = $false
    try {
        $page = $From.ContainingPage
        $toPage = $To.ContainingPage
        if (-not [object]::ReferenceEquals($page, $toPage)) {
            throw 'Connector endpoints must belong to the same page.'
        }
        if ($From.OneD -ne 0 -or $To.OneD -ne 0) { throw 'This helper connects 2-D shapes, not line endpoints.' }
        $app = $page.Application
        $tool = $app.ConnectorToolDataObject
        $connector = $page.Drop($tool, 0, 0)
        Set-VisioLineStyle $connector -Color $Color -LinePt $LinePt -Dash $Dash `
            -BeginArrow $BeginArrow -EndArrow $EndArrow -ArrowSize $ArrowSize
        if ($Routing -ne 'Auto') {
            $routeStyle = if ($Routing -eq 'Orthogonal') { '1' } else { '2' }
            Set-VisioCell $connector 'ShapeRouteStyle' $routeStyle
        }
        $begin = $connector.CellsU('BeginX'); $end = $connector.CellsU('EndX')
        Set-VisioEndpointGlue $begin $From $FromSide
        Set-VisioEndpointGlue $end $To $ToSide
        foreach ($axis in @('X', 'Y')) {
            $offset = if ($axis -eq 'X') { $LabelOffsetXPt } else { $LabelOffsetYPt }
            if ($offset -eq 0) { continue }
            $labelCell = $null
            try {
                $labelCell = $connector.CellsU("TxtPin$axis")
                $baseFormula = $labelCell.FormulaU
                # Bypass SETATREF write redirection or the control would refer to itself.
                $labelCell.FormulaForceU = '(' + $baseFormula + ') + ' + $offset.ToString([Globalization.CultureInfo]::InvariantCulture) + ' pt'
            } finally { Release-VisioComObject $labelCell }
        }
        if ($PassThru) { $returned = $true; return ,$connector }
    } catch {
        if ($connector) { $connector.Delete() }
        throw
    } finally {
        foreach ($item in @($begin, $end, $tool, $app, $toPage, $page)) {
            Release-VisioComObject $item
        }
        if (-not $returned) { Release-VisioComObject $connector }
    }
}
