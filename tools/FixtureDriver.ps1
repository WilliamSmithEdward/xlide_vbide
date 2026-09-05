<#
.SYNOPSIS
    The three-phase driver every fixture generator shares: blank workbook, editor session,
    modules written through the xlide api.

.DESCRIPTION
    Five generators carried this stretch verbatim - the COM maker for the empty .xlsm, the
    Start-Excel launch that loads the add-in, the JSON plan handed to build-fixture.mjs, and
    the cleanup - which meant a fifth fixture was a copy-paste and a change to how
    build-fixture.mjs is invoked was five edits with nothing to catch a miss (the audit's
    B24). A generator is now its VBA bodies, its summary text, and one call here.

    Dot-source it, then call Invoke-FixtureBuild.
#>

Set-StrictMode -Version Latest

# The first two phases on their own: the empty macro workbook only Excel can make, opened by
# the launcher that loads the add-in. The form fixture needs exactly these and then a different
# third phase, and copying them is how the driver's own lesson (the audit's B24) gets unlearned.
function Invoke-FixtureLaunch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path
    )

    New-Item -ItemType Directory -Force (Split-Path -Parent $Path) | Out-Null
    if (Test-Path $Path) { Remove-Item $Path -Force }

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
    & (Join-Path $PSScriptRoot 'harness\Start-Excel.ps1') -Workbook $Path -Fresh | Write-Host

    # The builder's own session, named for every harness call that follows: a Word session
    # beside the fresh Excel is a designed state now (2026-08-19), and a bare open() rightly
    # refuses to guess between two. -Fresh closed every Excel, so the one Excel IS the
    # builder's - read its pid from the process list rather than parsing console text.
    $env:XLIDE_PID = (Get-Process EXCEL | Select-Object -First 1).Id
}

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
        [Parameter(Mandatory)] [string] $OpenAtEnd,

        # Forms: an array of @{ Name; Controls = @(@{ Type; Name; Left; Top; Width; Height });
        # Code } - the component through the door, each control through the designer route, the
        # code-behind written like any module's.
        [array] $Forms = @(),

        # The ThisWorkbook document module's text, when the fixture wants one.
        [string] $WorkbookCode
    )

    $harness = Join-Path $PSScriptRoot 'harness'

    Invoke-FixtureLaunch -Path $Path
    $blankStamp = (Get-Item $Path).LastWriteTimeUtc

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

    if ($Forms.Count -gt 0) {
        $plan.forms = @(
            foreach ($form in $Forms) {
                $shaped = @{ name = $form.Name }
                if ($form.Contains('Controls')) {
                    $shaped.controls = @(
                        foreach ($control in $form.Controls) {
                            $one = @{ type = $control.Type; name = $control.Name }
                            foreach ($side in 'Left', 'Top', 'Width', 'Height') {
                                if ($control.Contains($side)) { $one[$side.ToLowerInvariant()] = $control[$side] }
                            }
                            $one
                        }
                    )
                }
                if ($form.Contains('Code')) { $shaped.code = $form.Code }
                $shaped
            }
        )
    }

    if ($WorkbookCode) {
        $plan.workbookCode = $WorkbookCode
    }

    # Written WITHOUT a byte-order mark. In PowerShell 5.1 `-Encoding utf8` means "UTF-8 with a
    # BOM", and JSON.parse refuses one - naming a character that does not appear to be in the
    # file. This lesson was recorded in four generators before it was recorded once, here.
    # Depth 8: the forms shape bottoms out at exactly five levels, and ConvertTo-Json TRUNCATES
    # beyond its depth silently - stringified hashtables where objects were meant - so the limit
    # sits well above the deepest plan rather than exactly at it.
    $planPath = Join-Path ([System.IO.Path]::GetTempPath()) "xlide-fixture-$PID.json"
    [System.IO.File]::WriteAllText(
        $planPath,
        ($plan | ConvertTo-Json -Depth 8),
        (New-Object System.Text.UTF8Encoding $false))

    Write-Host '3. Writing the components through the xlide api.'
    try {
        & node (Join-Path $harness 'build-fixture.mjs') $planPath | Write-Host
        if ($LASTEXITCODE -ne 0) { throw 'the fixture could not be built through the api' }
    }
    finally {
        Remove-Item $planPath -ErrorAction SilentlyContinue
    }

    # THE SAVE IS NOT OVER WHEN THE SAVE COMMAND RETURNS. The api's `save` executes the editor's
    # menu control and answers when the control has been told; Excel writes the file on its own
    # message loop afterwards. A generator's -Quiet close stopped Excel the instant node returned,
    # and on 2026-09-05 that was before the write: the file left behind was the blank workbook
    # from phase 1, the build said "saved; 6 component(s) hold code" (true of the session, not
    # the disk), and the suite that opened it next found no module named Registry. The file's
    # own write time is the only witness to the save, so it is waited for here, where the path
    # is known, before the caller gets to close anything.
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Item $Path).LastWriteTimeUtc -le $blankStamp) {
        if ((Get-Date) -gt $deadline) {
            throw "the editor's save did not reach $Path within 30 seconds; the file is still the blank workbook from phase 1"
        }
        Start-Sleep -Milliseconds 200
    }
}
