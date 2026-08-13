<#
.SYNOPSIS
    The three-phase driver every fixture generator shares: blank workbook, editor session,
    modules written through the debug api.

.DESCRIPTION
    Five generators carried this stretch verbatim - the COM maker for the empty .xlsm, the
    Start-Excel launch that loads the add-in, the JSON plan handed to build-fixture.mjs, and
    the cleanup - which meant a fifth fixture was a copy-paste and a change to how
    build-fixture.mjs is invoked was five edits with nothing to catch a miss (the audit's
    B24). A generator is now its VBA bodies, its summary text, and one call here.

    Dot-source it, then call Invoke-FixtureBuild.
#>

Set-StrictMode -Version Latest

function Invoke-FixtureBuild {
    [CmdletBinding()]
    param(
        # Where the .xlsm lands. Already resolved to a full path by the caller.
        [Parameter(Mandatory)] [string] $Path,

        # Ordered name -> @{ Kind = <vbext kind number>; Code = <module text> }.
        [Parameter(Mandatory)] [System.Collections.IDictionary] $Modules,

        # The Sheet1 document module's text.
        [Parameter(Mandatory)] [string] $SheetCode,

        # The module left showing when the build finishes.
        [Parameter(Mandatory)] [string] $OpenAtEnd
    )

    New-Item -ItemType Directory -Force (Split-Path -Parent $Path) | Out-Null
    if (Test-Path $Path) { Remove-Item $Path -Force }

    $harness = Join-Path $PSScriptRoot 'harness'

    Write-Host '1. Making an empty macro workbook.'
    $maker = New-Object -ComObject Excel.Application
    $maker.Visible = $false
    $maker.DisplayAlerts = $false
    try {
        $blank = $maker.Workbooks.Add()
        $blank.SaveAs($Path, 52)
        $blank.Close($false)
    }
    finally {
        try { $maker.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($maker) | Out-Null
    }

    Write-Host '2. Opening it with the editor, which is what loads the add-in.'
    & (Join-Path $harness 'Start-Excel.ps1') -Workbook $Path -Fresh | Write-Host

    # The module texts go to node as JSON, so that quoting, CRLFs and VBA's own doubled quotes
    # cross once rather than being escaped through two shells.
    $plan = @{
        modules = @(
            foreach ($name in $Modules.Keys) {
                @{ name = $name; kind = $Modules[$name].Kind; code = $Modules[$name].Code }
            }
        )
        sheetCode = $SheetCode
        openAtEnd = $OpenAtEnd
    }

    # Written WITHOUT a byte-order mark. In PowerShell 5.1 `-Encoding utf8` means "UTF-8 with a
    # BOM", and JSON.parse refuses one - naming a character that does not appear to be in the
    # file. This lesson was recorded in four generators before it was recorded once, here.
    $planPath = Join-Path ([System.IO.Path]::GetTempPath()) "xlide-fixture-$PID.json"
    [System.IO.File]::WriteAllText(
        $planPath,
        ($plan | ConvertTo-Json -Depth 5),
        (New-Object System.Text.UTF8Encoding $false))

    Write-Host '3. Writing the components through the debug api.'
    try {
        & node (Join-Path $harness 'build-fixture.mjs') $planPath | Write-Host
        if ($LASTEXITCODE -ne 0) { throw 'the fixture could not be built through the api' }
    }
    finally {
        Remove-Item $planPath -ErrorAction SilentlyContinue
    }
}
