# Registers the development build of the add-in for the current user. No administrator rights.
#
# Run this from YOUR OWN terminal, not from an agent or sandboxed shell: sandboxed environments
# can virtualize registry writes, leaving a registration that only the sandbox and its children
# can see while Excel launched normally sees nothing. That exact mirage cost this project a full
# diagnostic day. The published shim path is stable across rebuilds, so one real registration
# survives the whole dev loop.

$ErrorActionPreference = 'Stop'

$progId = 'Xlide.VbeAddIn'
$clsid = '{588903F2-4CDE-4607-828A-6870A1F3FDC1}'
$shim = 'F:\GitHub\xlide\xlide_vbide\artifacts\publish\Xlide.Vbe.Shim\release_win-x64\Xlide.Vbe.Shim.dll'

if (-not (Test-Path $shim)) { throw "No shim at $shim - run tools\dev.ps1 -NoRun first." }

$classes = 'HKCU:\Software\Classes'
New-Item -Path "$classes\$progId\CLSID" -Force | Out-Null
Set-Item -Path "$classes\$progId" -Value 'xlide'
Set-Item -Path "$classes\$progId\CLSID" -Value $clsid

New-Item -Path "$classes\CLSID\$clsid\InprocServer32" -Force | Out-Null
Set-Item -Path "$classes\CLSID\$clsid" -Value 'xlide'
Set-Item -Path "$classes\CLSID\$clsid\InprocServer32" -Value $shim
Set-ItemProperty -Path "$classes\CLSID\$clsid\InprocServer32" -Name 'ThreadingModel' -Value 'Apartment'

$addin = "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\$progId"
New-Item -Path $addin -Force | Out-Null
Set-ItemProperty -Path $addin -Name 'Description' -Value 'Modern VBA development inside the Visual Basic Editor.'
Set-ItemProperty -Path $addin -Name 'FriendlyName' -Value 'xlide'
Set-ItemProperty -Path $addin -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $addin -Name 'CommandLineSafe' -Value 0 -Type DWord

Write-Host 'Registered for the current user.'
Write-Host "  $addin"
Write-Host "  $classes\CLSID\$clsid"
Write-Host 'Check in regedit: HKCU\Software\Microsoft\VBA\VBE should now exist. Then open Excel and press Alt+F11.'
