<#
.SYNOPSIS
    Builds DebugFixture.xlsm: a workbook that COMPILES, for exercising the debugger.

.DESCRIPTION
    Neither other fixture can be run. The rename fixture deliberately does not compile — two
    modules declare the same public name, which is the whole point of it — and the language
    fixture carries a module of deliberate defects. Pressing Run on either raises a compile
    error, and a compile error is a modal, so a debugger test against them tests the dialog
    guard instead of the debugger.

    This one compiles, and is shaped for the run-and-stop cycle:

      a procedure with several straight-line statements, so stepping has somewhere to go
      locals of more than one type, so the Locals panel has something to be wrong about
      a call into another module, so Step Into and Step Over differ visibly
      a loop, because a breakpoint inside one is hit repeatedly and that is its own case

    Nothing here has side effects on the worksheet: a debugger test that leaves cells changed
    is a test that cannot be run twice.

.EXAMPLE
    tools\New-DebugFixture.ps1
#>
[CmdletBinding()]
param(
    [string] $Path,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Path) {
    $Path = Join-Path $repoRoot 'artifacts\fixtures\DebugFixture.xlsm'
}

if (-not [System.IO.Path]::IsPathRooted($Path)) {
    $Path = Join-Path $repoRoot $Path
}

New-Item -ItemType Directory -Force (Split-Path -Parent $Path) | Out-Null
if (Test-Path $Path) { Remove-Item $Path -Force }

$StandardModule = 1

$modules = [ordered]@{}

$modules['Runner'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' THE PROCEDURE UNDER TEST. Straight-line statements with locals of more than one type, so a
' breakpoint has somewhere to stop and the Locals panel has something to show.
'
' Nothing here touches the worksheet. A debugger test that leaves cells changed cannot be run
' twice, and a fixture that has to be rebuilt between runs is a fixture nobody runs.

Public Sub Walk()
    Dim counter As Long
    Dim label As String
    Dim ratio As Double

    counter = 1
    label = "start"
    ratio = 0.5

    counter = counter + 1
    label = Helper.Describe(counter)
    ratio = ratio * 2

    Debug.Print label, counter, ratio
End Sub

' A LOOP, because a breakpoint inside one is hit repeatedly and that is its own case.
Public Sub Loops()
    Dim n As Long
    Dim total As Long

    For n = 1 To 5
        total = total + n
    Next n

    Debug.Print total
End Sub

' A CALL into another module, so Step Into and Step Over differ visibly.
Public Sub Calls()
    Dim answer As String
    answer = Helper.Describe(7)
    Debug.Print answer
End Sub
'@ }

$modules['Helper'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' The other side of a call. Step Into arrives here; Step Over does not.

Public Function Describe(ByVal value As Long) As String
    Dim prefix As String
    prefix = "value="
    Describe = prefix & CStr(value)
End Function
'@ }

$sheetCode = @'
Option Explicit

' Empty on purpose. The debugger fixture touches no cells.
'@

# ---------------------------------------------------------------- building it

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

$plan = @{
    modules = @(
        foreach ($name in $modules.Keys) {
            @{ name = $name; kind = $modules[$name].Kind; code = $modules[$name].Code }
        }
    )
    sheetCode = $sheetCode
    openAtEnd = 'Runner'
}

$planPath = Join-Path ([System.IO.Path]::GetTempPath()) "xlide-debug-fixture-$PID.json"
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

if (-not $Quiet) {
    Write-Host ''
    Write-Host "Fixture written to $Path"
    Write-Host ''
    Write-Host '  Runner.Walk    straight-line statements and three locals, for stopping and stepping'
    Write-Host '  Runner.Loops   a For loop, because a breakpoint inside one is hit repeatedly'
    Write-Host '  Runner.Calls   a call into Helper, so Step Into and Step Over differ'
    Write-Host '  Helper         the other side of that call'
    Write-Host ''
    Write-Host '  It COMPILES, which neither other fixture does.'
    Write-Host ''
}
