# Plants the add-in registration inside the Office Click-to-Run registry overlay. Run elevated.
#
# Click-to-Run Office runs behind an App-V registry overlay, and the overlay owns the
# Software\Microsoft\VBA namespace: it is where the editor's own machine-level VBA values live
# (Vbe71DllPath), and on some machines reads of the per-user Addins64 key resolve into it and come
# back empty, so a correct HKCU registration is never seen and the Add-in Manager lists nothing.
# Writing the registration into the overlay puts it in the view those reads actually consult.
#
# The per-user registration under HKCU remains the primary one; this supplements it. Office
# updates can rebuild the overlay, which silently removes these keys - re-run afterwards.

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    # Take the overlay registration out again.
    [switch] $Remove
)

$ErrorActionPreference = 'Stop'

$progId = 'Xlide.VbeAddIn'
$clsid = '{588903F2-4CDE-4607-828A-6870A1F3FDC1}'
$shim = 'F:\GitHub\xlide\xlide_vbide\artifacts\publish\Xlide.Vbe.Shim\release_win-x64\Xlide.Vbe.Shim.dll'

$overlay = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE'
if (-not (Test-Path "$overlay\SOFTWARE\Microsoft\VBA")) {
    throw 'No Click-to-Run overlay with a VBA branch here; this fix only applies to C2R Office.'
}

$addin = "$overlay\SOFTWARE\Microsoft\VBA\VBE\6.0\Addins64\$progId"
$classes = "$overlay\SOFTWARE\Classes"

if ($Remove) {
    foreach ($key in @($addin, "$classes\$progId", "$classes\CLSID\$clsid")) {
        if (Test-Path $key) { Remove-Item $key -Recurse -Force; Write-Host "removed $key" }
    }
    return
}

if (-not (Test-Path $shim)) { throw "No shim at $shim" }

New-Item -Path $addin -Force | Out-Null
Set-ItemProperty -Path $addin -Name 'Description' -Value 'Modern VBA development inside the Visual Basic Editor.'
Set-ItemProperty -Path $addin -Name 'FriendlyName' -Value 'xlide'
Set-ItemProperty -Path $addin -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $addin -Name 'CommandLineSafe' -Value 0 -Type DWord

# Activation resolves in the same view as discovery, so the class goes in beside the add-in key.
New-Item -Path "$classes\$progId\CLSID" -Force | Out-Null
Set-ItemProperty -Path "$classes\$progId" -Name '(default)' -Value 'xlide'
Set-ItemProperty -Path "$classes\$progId\CLSID" -Name '(default)' -Value $clsid

New-Item -Path "$classes\CLSID\$clsid\InprocServer32" -Force | Out-Null
Set-ItemProperty -Path "$classes\CLSID\$clsid" -Name '(default)' -Value 'xlide'
Set-ItemProperty -Path "$classes\CLSID\$clsid\InprocServer32" -Name '(default)' -Value $shim
Set-ItemProperty -Path "$classes\CLSID\$clsid\InprocServer32" -Name 'ThreadingModel' -Value 'Apartment'

Write-Host 'Overlay registration written.'
Write-Host "  $addin"
Write-Host "  $classes\CLSID\$clsid"
Write-Host 'Open Excel from the Start menu and press Alt+F11. If an Office update removes these keys, re-run this script.'
