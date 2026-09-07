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

function Connect-VisioShapes {
    param(
        [Parameter(Mandatory = $true)]$From,
        [Parameter(Mandatory = $true)]$To,
        [switch]$PassThru
    )
    $page = $null; $app = $null; $tool = $null; $connector = $null
    $begin = $null; $end = $null; $fromPin = $null; $toPin = $null
    $returned = $false
    try {
        $page = $From.ContainingPage
        $app = $page.Application
        $tool = $app.ConnectorToolDataObject
        $connector = $page.Drop($tool, 0, 0)
        $begin = $connector.CellsU('BeginX'); $end = $connector.CellsU('EndX')
        $fromPin = $From.CellsU('PinX'); $toPin = $To.CellsU('PinX')
        $begin.GlueTo($fromPin)
        $end.GlueTo($toPin)
        if ($PassThru) { $returned = $true; return ,$connector }
    } catch {
        if ($connector) { $connector.Delete() }
        throw
    } finally {
        foreach ($item in @($begin, $end, $fromPin, $toPin, $tool, $app, $page)) {
            Release-VisioComObject $item
        }
        if (-not $returned) { Release-VisioComObject $connector }
    }
}
