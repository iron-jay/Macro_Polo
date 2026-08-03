<#
.SYNOPSIS
    Checks that the add-in classes expose to COM everything Office is going to ask them for.

.DESCRIPTION
    Every defect that took Word and Excel down during the move off VSTO was of one kind: a COM
    contract that compiles cleanly, has no unit-test surface, and fails only when the host calls
    into it - taking the whole process with it, with no managed exception and nothing in the log.

    Two got through:

      * IDTExtensibility2 declared as IDispatch-only instead of dual, so Office's call to
        OnConnection ran off the end of the exposed vtable slots.
      * The shared base class inheriting assembly-level ComVisible(false), which makes the CLR
        refuse QueryInterface for IDispatch on the derived class - and IDispatch is how every
        ribbon callback is dispatched.

    This script builds the CCW in-process and asserts on it directly, so both are caught in a few
    seconds without installing anything or opening Office.

.PARAMETER Configuration
    Which build output to check. Defaults to Release.

.EXAMPLE
    powershell -File build\Test-ComSurface.ps1
#>
[CmdletBinding()]
param(
    [string] $Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

$targets = @(
    @{ Assembly = "Macro_Polo_Word\bin\$Configuration\Macro_Polo_Word.dll";   Class = 'Macro_Polo_Word.WordAddIn' }
    @{ Assembly = "Macro_Polo_Excel\bin\$Configuration\Macro_Polo_Excel.dll"; Class = 'Macro_Polo_Excel.ExcelAddIn' }
)

# The interfaces Office queries for, by IID. Names are resolved from the loaded add-in assembly,
# which carries embedded copies of the Office interop types.
$required = @(
    @{ Name = 'Macro_Polo.Core.IDTExtensibility2';                Why = 'add-in lifetime' }
    @{ Name = 'Microsoft.Office.Core.IRibbonExtensibility';       Why = 'ribbon XML' }
    @{ Name = 'Microsoft.Office.Core.ICustomTaskPaneConsumer';    Why = 'task pane factory' }
)

$marshal = [System.Runtime.InteropServices.Marshal]
$failures = @()

function Test-Vtable {
    param($Pointer, $Label, $MinimumSlots)

    $vtable = $marshal::ReadIntPtr($Pointer)
    if ($vtable -eq [IntPtr]::Zero) {
        return "$Label : null vtable"
    }

    for ($i = 0; $i -lt $MinimumSlots; $i++) {
        $slot = $marshal::ReadIntPtr($vtable, $i * [IntPtr]::Size)
        if ($slot -eq [IntPtr]::Zero) {
            return "$Label : vtable slot $i is null"
        }
    }

    return $null
}

foreach ($target in $targets) {
    $path = Join-Path $repoRoot $target.Assembly
    if (-not (Test-Path $path)) {
        throw "Not built: $path"
    }

    Write-Host "`n=== $($target.Class) ===" -ForegroundColor Cyan

    # LoadFrom rather than Add-Type: the add-in's dependency on Macro_Polo.Core resolves out of the
    # same output directory, which a simple-name lookup would not find.
    $assembly = [System.Reflection.Assembly]::LoadFrom($path)
    $type = $assembly.GetType($target.Class, $true)
    $instance = [Activator]::CreateInstance($type)

    # The check that matters most. A ComVisible(false) type anywhere in the base chain makes the
    # CLR refuse this, and Office cannot dispatch a single ribbon callback without it.
    try {
        $dispatch = $marshal::GetIDispatchForObject($instance)
        $problem = Test-Vtable -Pointer $dispatch -Label 'IDispatch' -MinimumSlots 7
        $marshal::Release($dispatch) | Out-Null

        if ($problem) { $failures += "$($target.Class): $problem" ; Write-Host "  FAIL IDispatch - $problem" -ForegroundColor Red }
        else { Write-Host '  ok   IDispatch (ribbon callbacks)' -ForegroundColor Green }
    }
    catch {
        $failures += "$($target.Class): IDispatch unavailable - $($_.Exception.Message)"
        Write-Host "  FAIL IDispatch - $($_.Exception.Message)" -ForegroundColor Red
    }

    foreach ($entry in $required) {
        $interface = $null
        foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
            $candidate = $assembly.GetType($entry.Name, $false)
            if ($candidate) { $interface = $candidate; break }
        }

        if (-not $interface) {
            $failures += "$($target.Class): interface $($entry.Name) not found"
            Write-Host "  FAIL $($entry.Name) - type not found" -ForegroundColor Red
            continue
        }

        if (-not $interface.IsInstanceOfType($instance)) {
            $failures += "$($target.Class): does not implement $($entry.Name)"
            Write-Host "  FAIL $($entry.Name) - not implemented" -ForegroundColor Red
            continue
        }

        try {
            $pointer = $marshal::GetComInterfaceForObject($instance, $interface)
            # 3 IUnknown + 4 IDispatch + at least one interface method. A dispatch-only declaration
            # of a dual interface shows up here as a vtable that stops after slot 6.
            $problem = Test-Vtable -Pointer $pointer -Label $entry.Name -MinimumSlots 8
            $marshal::Release($pointer) | Out-Null

            if ($problem) { $failures += "$($target.Class): $problem"; Write-Host "  FAIL $($entry.Name) - $problem" -ForegroundColor Red }
            else { Write-Host "  ok   $($entry.Name) ($($entry.Why))" -ForegroundColor Green }
        }
        catch {
            $failures += "$($target.Class): QueryInterface for $($entry.Name) failed - $($_.Exception.Message)"
            Write-Host "  FAIL $($entry.Name) - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# The task pane control. ICTPFactory.CreateCTP instantiates it by ProgID and sites it as an ActiveX
# control, so the CCW has to answer for IOleObject. This is the last COM contract in the add-in that
# can be checked without launching Office.
Write-Host "`n=== Macro_Polo.Core.MacroStatusBanner ===" -ForegroundColor Cyan
$corePath = Join-Path $repoRoot "Macro_Polo.Core\bin\$Configuration\net472\Macro_Polo.Core.dll"
if (-not (Test-Path $corePath)) { throw "Not built: $corePath" }

$core = [System.Reflection.Assembly]::LoadFrom($corePath)
$bannerType = $core.GetType('Macro_Polo.Core.MacroStatusBanner', $true)
$banner = [Activator]::CreateInstance($bannerType)

$oleInterfaces = @(
    @{ Name = 'IOleObject';  Iid = '00000112-0000-0000-C000-000000000046' }
    @{ Name = 'IOleControl'; Iid = 'B196B288-BAB4-101A-B69C-00AA00341D07' }
    @{ Name = 'IViewObject'; Iid = '0000010D-0000-0000-C000-000000000046' }
)

$unknown = $marshal::GetIUnknownForObject($banner)
try {
    foreach ($ole in $oleInterfaces) {
        $iid = [Guid]$ole.Iid
        $ptr = [IntPtr]::Zero
        $hr = $marshal::QueryInterface($unknown, [ref]$iid, [ref]$ptr)
        if ($hr -eq 0 -and $ptr -ne [IntPtr]::Zero) {
            $marshal::Release($ptr) | Out-Null
            Write-Host "  ok   $($ole.Name)" -ForegroundColor Green
        }
        else {
            $failures += "MacroStatusBanner: QueryInterface for $($ole.Name) returned 0x{0:X8}" -f $hr
            Write-Host ("  FAIL $($ole.Name) - hr=0x{0:X8}" -f $hr) -ForegroundColor Red
        }
    }
}
finally {
    $marshal::Release($unknown) | Out-Null
    $banner.Dispose()
}

Write-Host ''
if ($failures) {
    Write-Host "COM surface check FAILED:" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}

Write-Host 'COM surface check passed. Office has everything it queries for.' -ForegroundColor Green
