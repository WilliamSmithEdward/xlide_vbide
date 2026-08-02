# Mirrors the per-user registration into the machine hive. Run elevated.
#
# Written for hosts whose per-process registry view masks the per-user VBA branch: the editor
# enumerates machine add-ins as well as user ones, and the machine hive shows through views that
# swallow the user key. Values mirror RegistrationPlan exactly.

#Requires -RunAsAdministrator

$ErrorActionPreference = 'Stop'

$progId = 'Xlide.VbeAddIn'
$clsid = '{588903F2-4CDE-4607-828A-6870A1F3FDC1}'
$shim = 'F:\GitHub\xlide\xlide_vbide\artifacts\publish\Xlide.Vbe.Shim\release_win-x64\Xlide.Vbe.Shim.dll'

if (-not (Test-Path $shim)) { throw "No shim at $shim" }

$classes = "HKLM:\SOFTWARE\Classes"
New-Item -Path "$classes\$progId\CLSID" -Force | Out-Null
Set-ItemProperty -Path "$classes\$progId" -Name '(default)' -Value 'xlide'
Set-ItemProperty -Path "$classes\$progId\CLSID" -Name '(default)' -Value $clsid

New-Item -Path "$classes\CLSID\$clsid\InprocServer32" -Force | Out-Null
Set-ItemProperty -Path "$classes\CLSID\$clsid" -Name '(default)' -Value 'xlide'
Set-ItemProperty -Path "$classes\CLSID\$clsid\InprocServer32" -Name '(default)' -Value $shim
Set-ItemProperty -Path "$classes\CLSID\$clsid\InprocServer32" -Name 'ThreadingModel' -Value 'Apartment'

$addins = "HKLM:\SOFTWARE\Microsoft\VBA\VBE\6.0\Addins64\$progId"
New-Item -Path $addins -Force | Out-Null
Set-ItemProperty -Path $addins -Name 'Description' -Value 'Modern VBA development inside the Visual Basic Editor.'
Set-ItemProperty -Path $addins -Name 'FriendlyName' -Value 'xlide'
Set-ItemProperty -Path $addins -Name 'LoadBehavior' -Value 3 -Type DWord
Set-ItemProperty -Path $addins -Name 'CommandLineSafe' -Value 0 -Type DWord

Write-Host 'Machine-wide registration written.'
Write-Host "  $addins"
Write-Host "  $classes\CLSID\$clsid"
