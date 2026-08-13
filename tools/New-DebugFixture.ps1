<#
.SYNOPSIS
    Builds DebugFixture.xlsm: a workbook that COMPILES, for exercising the debugger.

.DESCRIPTION
    Neither other fixture can be run. The rename fixture deliberately does not compile - two
    modules declare the same public name, which is the whole point of it - and the language
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

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureBuild -Path $Path -Modules $modules -SheetCode $sheetCode -OpenAtEnd 'Runner'

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
