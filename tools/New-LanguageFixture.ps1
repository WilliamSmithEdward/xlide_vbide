<#
.SYNOPSIS
    Builds LanguageFixture.xlsm: a workbook for exercising the language features.

.DESCRIPTION
    The rename fixture is shaped for renaming - one method on one class, and a project that
    deliberately does not compile. That makes it the wrong workbook for asking what IntelliSense
    offers: a receiver with one member proves almost nothing, and a project full of deliberate
    errors buries the one error a quick-fix test wants to see.

    This one is shaped for the questions the language features raise:

      a class with several members of every kind, so a dot menu has something to be wrong about
      an enum and a user-defined type, which resolve by different paths
      a module that USES them, with receivers of each kind on their own lines
      Excel's own object model as a receiver, which comes from the type libraries rather than
        from the project
      a procedure with parameters, for signature help
      one module of deliberate, ISOLATED defects, so a quick-fix test has exactly one finding
        to look at rather than a project's worth

    Every module here compiles except Defects, which is named so nobody wonders.

    Built through the debug api rather than through Workbook.VBProject, so "Trust access to the
    VBA project object model" stays OFF. See New-RenameFixture.ps1 for why that matters.

.EXAMPLE
    tools\New-LanguageFixture.ps1
#>
[CmdletBinding()]
param(
    # Where to write it. Defaults beside the other fixtures.
    [string] $Path,

    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot

if (-not $Path) {
    $Path = Join-Path $repoRoot 'artifacts\fixtures\LanguageFixture.xlsm'
}

if (-not [System.IO.Path]::IsPathRooted($Path)) {
    $Path = Join-Path $repoRoot $Path
}

$StandardModule = 1
$ClassModule = 2

$modules = [ordered]@{}

# A class with one of everything a dot menu should list: a Sub, a Function, a read-only
# Property Get, a settable pair, and a private member that must NOT be offered from outside.
$modules['Gadget'] = @{ Kind = $ClassModule; Code = @'
Option Explicit

' A CLASS WITH SEVERAL MEMBERS. A receiver with one member cannot tell a working dot menu
' from one that answers with whatever it found first.

Private mName As String
Private mCount As Long

' Private: this must never appear in another module's dot menu.
Private Sub Reset()
    mCount = 0
End Sub

Public Sub Spin(ByVal turns As Long)
    mCount = mCount + turns
End Sub

Public Function Describe(ByVal prefix As String) As String
    Describe = prefix & mName
End Function

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property

Public Property Get Count() As Long
    Count = mCount
End Property
'@ }

# An enum and a user-defined type: both resolve by their own path, and both are receivers.
$modules['Shapes'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' AN ENUM. `Corner.` should list its members, and so should a variable declared As Corner.
Public Enum Corner
    TopLeft
    TopRight
    BottomLeft
    BottomRight
End Enum

' A USER-DEFINED TYPE. `point.` should list its fields, which is a different resolution
' from a class and from an enum.
Public Type Point
    X As Double
    Y As Double
    Label As String
End Type

Public Function Corners() As Long
    Corners = 4
End Function
'@ }

# The module that USES all of it. Every receiver sits on its own line so a test can name a
# line rather than hunt for an offset.
$modules['Uses'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' RECEIVERS, one per line, each the interesting half of a dot menu.

Public Sub ClassReceiver()
    Dim g As Gadget
    Set g = New Gadget
    g.Spin 1
End Sub

Public Sub TypeReceiver()
    Dim p As Point
    p.X = 1
End Sub

Public Sub EnumReceiver()
    Dim c As Corner
    c = Corner.TopLeft
End Sub

Public Sub HostReceiver()
    ' From the type libraries rather than from this project, which is a different source
    ' entirely and the one most likely to be missing.
    Application.Calculate
End Sub

Public Sub SheetReceiver()
    ActiveSheet.Calculate
End Sub

Public Sub SignatureHere()
    ' Signature help wants a call with parameters, mid-argument.
    Dim g As Gadget
    Set g = New Gadget
    g.Describe "prefix"
End Sub
'@ }

# ONE module of deliberate defects, isolated so a quick-fix test has exactly one thing to look
# at. Everything else in this workbook compiles.
$modules['Defects'] = @{ Kind = $StandardModule; Code = @'
Option Explicit

' DELIBERATE. Every finding a quick-fix test needs, and nothing else in this workbook has any.

Public Sub Undeclared()
    ' Option Explicit is on and this is never declared.
    missingVariable = 1
End Sub

Public Sub TypeMismatch()
    Dim n As Long
    ' A string into a Long.
    n = "not a number"
End Sub
'@ }

$sheetCode = @'
Option Explicit

' Nothing here. The sheet exists because a workbook has one, not because the fixture needs it.
'@

# ---------------------------------------------------------------- building it
#
# Three phases, exactly as New-RenameFixture.ps1 does them and for the same reasons: only Excel
# can make an empty macro workbook, only an ordinarily-started Excel loads the add-in, and only
# the add-in's door can add components without the trust setting.

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureBuild -Path $Path -Modules $modules -SheetCode $sheetCode -OpenAtEnd 'Uses'

if (-not $Quiet) {
    Write-Host ''
    Write-Host "Fixture written to $Path"
    Write-Host ''
    Write-Host '  Gadget    a class with a Sub, a Function, two Property Gets, a Let, and a'
    Write-Host '            PRIVATE Sub that must never appear in another module''s dot menu'
    Write-Host '  Shapes    an Enum and a user-defined Type, which resolve by their own paths'
    Write-Host '  Uses      one receiver per line: class, type, enum, Application, ActiveSheet,'
    Write-Host '            and a call with parameters for signature help'
    Write-Host '  Defects   the only module with findings, so a quick-fix test sees exactly one'
    Write-Host ''
}
