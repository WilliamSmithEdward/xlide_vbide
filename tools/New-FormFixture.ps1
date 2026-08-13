<#
.SYNOPSIS
    Builds FormFixture.xlsm: a workbook whose UserForm holds every standard control.

.DESCRIPTION
    The designer milestone's fixture (docs/userform-designer.md, M1). One form carrying the
    whole toolbox - every standard control, a Frame with children, a MultiPage with a control
    on its first page - and a code-behind that names real controls, so the workbook COMPILES.

    Built through the designer route, which is the same model the native toolbox calls, so
    "Trust access to the VBA project object model" stays OFF. The form's declaration lives in
    tools\harness\form-plan.mjs, shared verbatim with designer-features.mjs - the suite that
    verifies a read against it - so the fixture and its expectations cannot drift apart.

.EXAMPLE
    tools\New-FormFixture.ps1
#>
[CmdletBinding()]
param(
    [string] $Path,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Path) {
    $Path = Join-Path $repoRoot 'artifacts\fixtures\FormFixture.xlsm'
}

if (-not [System.IO.Path]::IsPathRooted($Path)) {
    $Path = Join-Path $repoRoot $Path
}

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureLaunch -Path $Path

Write-Host '3. Building the form through the designer route.'
& node (Join-Path $PSScriptRoot 'harness\build-form-fixture.mjs') | Write-Host
if ($LASTEXITCODE -ne 0) { throw 'the form fixture could not be built through the api' }

if (-not $Quiet) {
    Write-Host ''
    Write-Host "Fixture written to $Path"
    Write-Host ''
    Write-Host '  EntryForm   every standard control; a Frame holding option buttons; a MultiPage'
    Write-Host '              with a control on Page1; a code-behind that compiles against them'
    Write-Host ''
    Write-Host '  The declaration is tools\harness\form-plan.mjs, shared with the suite.'
    Write-Host ''
}
