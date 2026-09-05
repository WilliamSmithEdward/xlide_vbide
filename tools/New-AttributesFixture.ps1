<#
.SYNOPSIS
    Builds the workbook the attribute annotations are tested against.

.DESCRIPTION
    A VBA module carries attributes the code pane never shows - VB_PredeclaredId, VB_Description,
    VB_UserMemId, the Excel hotkey, a variable's description - and the editor offers no way to set
    them. The Rubberduck convention names each one in a comment ('@PredeclaredId, '@Description
    and the rest), and xlide writes the attribute to match. Every annotation the reader knows has
    a module here that carries it, and every refusal has a module that earns it, so the live suite
    (tools\harness\attributes.mjs) can ask one question per rule:

      Registry     a class with '@PredeclaredId and '@ModuleDescription, a described member, and
                   a described module-level variable
      Bag          a class with a '@DefaultMember and an '@Enumerator, the collection-wrapper shape
      Macros       a standard module with an '@ExcelHotkey, a described procedure, and a module
                   description
      Uses         calls Registry.Lookup as a value, which the analyzer accepts only once the class
                   is predeclared - the proof the analyzer hears about an applied attribute
      Misplaced    a '@PredeclaredId on a standard module and a '@Description above a variable,
                   both reported rather than written
      Sheet1       a document module with a '@ModuleDescription, which cannot be applied because a
                   document module cannot be imported

    Built through the xlide api, so "Trust access to the VBA project object model" does NOT have
    to be on. A Debug build must be registered and loading. Every module compiles.

    The build's own save writes the annotations into the attributes (that is what a save does, by
    default), so the file as built carries them. The suite takes them away at its start, with the
    setting off, and again at its end, so every run begins from the same place.

.EXAMPLE
    tools\New-AttributesFixture.ps1
    Builds it, saves it, and leaves Excel open on it.

.EXAMPLE
    tools\New-AttributesFixture.ps1 -Quiet
    Builds it and closes Excel afterwards.
#>
[CmdletBinding()]
param(
    # Where to save it. Defaults beside the repo's other build output.
    [string] $Path,

    # Build and close, rather than leaving it open to work in.
    [switch] $Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $Path) {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $fixtures = Join-Path $repoRoot 'artifacts\fixtures'
    if (-not (Test-Path $fixtures)) { New-Item -ItemType Directory -Force $fixtures | Out-Null }
    $Path = Join-Path $fixtures 'AttributesFixture.xlsm'
}

# Component types, as VBComponents.Add takes them.
$StandardModule = 1
$ClassModule = 2

$modules = [ordered]@{}

$modules['Registry'] = @{ Kind = $ClassModule; Code = @"
'@ModuleDescription("Where things are looked up.")
'@PredeclaredId
Option Explicit

' THE PREDECLARED CLASS. With VB_PredeclaredId = True its own name is an object, so Uses can
' write Registry.Lookup without New; without it that line is an undeclared name.

'@VariableDescription("How many lookups so far.")
Public Count As Long

'@Description("Finds a thing by its name.")
Public Function Lookup(ByVal name As String) As String
    Count = Count + 1
    Lookup = "found " & name
End Function

Public Sub Reset()
    Count = 0
End Sub
"@ }

$modules['Bag'] = @{ Kind = $ClassModule; Code = @"
Option Explicit

' THE COLLECTION WRAPPER: a default member so bag(1) means bag.Item(1), and an enumerator so
' For Each walks it. Both are VB_UserMemId values the code pane cannot show.

Private items As Collection

Private Sub Class_Initialize()
    Set items = New Collection
End Sub

Public Sub Add(ByVal thing As Variant)
    items.Add thing
End Sub

'@DefaultMember
Public Function Item(ByVal index As Long) As Variant
    Item = items(index)
End Function

'@Enumerator
Public Function NewEnum() As IUnknown
    Set NewEnum = items.[_NewEnum]
End Function
"@ }

$modules['Macros'] = @{ Kind = $StandardModule; Code = @"
'@ModuleDescription("Macros for the sheet.")
Option Explicit

' A STANDARD MODULE: a description on the module, a description on a procedure, and an Excel
' hotkey, which is Ctrl+Shift+D because the letter is upper case.

'@ExcelHotkey("D")
Public Sub DoIt()
    Debug.Print "did it"
End Sub

' THE DESCRIPTION IS NOT ASCII ON PURPOSE. An en dash and a curly apostrophe are 0x96 and 0x92
' in the page the editor exports in, and a control character each in Latin-1: the range a
' Latin-1 read gets wrong, so the suite proves the write and the package read agree with the code.
'@Description("Prints a greeting $([char]0x2013) the day$([char]0x2019)s first.")
Public Sub Hello()
    Debug.Print "hello"
End Sub
"@ }

$modules['Uses'] = @{ Kind = $StandardModule; Code = @"
Option Explicit

' THE PROOF THE ANALYZER HEARS ABOUT AN APPLIED ATTRIBUTE. Registry.Lookup names the class as a
' value, which only a predeclared class allows: before the annotation is applied the analyzer
' reports the line, and after it the report goes.

Public Sub Try()
    Debug.Print Registry.Lookup("x")
End Sub
"@ }

$modules['Misplaced'] = @{ Kind = $StandardModule; Code = @"
'@PredeclaredId
Option Explicit

' TWO ANNOTATIONS THAT CANNOT MEAN WHAT THEY SAY: a predeclared standard module, and a member
' description above a variable. Both are reported on their lines and neither is written.

'@Description("This is a variable, not a procedure.")
Public Total As Long

Public Sub Nothing_To_See()
    Total = 1
End Sub
"@ }

$sheetCode = @"
'@ModuleDescription("The first sheet.")
Option Explicit

' A DOCUMENT MODULE cannot be imported, so its annotation is reported as inapplicable.

Public Sub Refresh()
    Debug.Print "sheet"
End Sub
"@

# ---------------------------------------------------------------- building it

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureBuild -Path $Path -Modules $modules -SheetCode $sheetCode -OpenAtEnd 'Registry'

Write-Host ''
Write-Host "Fixture written to $Path"
Write-Host ''
Write-Host '  Registry    predeclared class, module and member and variable descriptions'
Write-Host '  Bag         default member and enumerator'
Write-Host '  Macros      hotkey and descriptions on a standard module'
Write-Host '  Uses        Registry.Lookup as a value: reported until the class is predeclared'
Write-Host '  Misplaced   annotations that cannot mean what they say'
Write-Host '  Sheet1      a document module, which cannot take attributes'
Write-Host ''
Write-Host '  tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\AttributesFixture.xlsm'
Write-Host '  node tools\harness\attributes.mjs'
Write-Host ''

if ($Quiet) {
    # ONLY THE BUILDER'S OWN SESSION, which Invoke-FixtureLaunch named in XLIDE_PID. This used
    # to stop every Excel on the machine, the owner's open workbooks included (2026-09-05).
    if ($env:XLIDE_PID) { Stop-Process -Id ([int] $env:XLIDE_PID) -Force -ErrorAction SilentlyContinue }
}
