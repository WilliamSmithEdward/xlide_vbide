<#
.SYNOPSIS
    Builds the workbook the explorer's FOLDER LAYOUT is tested against (#23).

.DESCRIPTION
    A folder is a comment at the top of a module - '@Folder("Parent.Child"), the Rubberduck
    convention - and the tree draws what the comments say. Every rule the reader follows has a
    module here that exercises it, so the live suite (tools\harness\folders.mjs) can ask the tree
    one question per rule and get an answer only that rule can give:

      Ledger, Posting      the documented form, two modules of two kinds in one nested folder
      Invoice              a sibling folder under the same parent
      Helpers              the documented form, in a folder the twin also uses by name
      Tools                the loose spellings: lower-case name, quotes without parentheses
      Bare                 the barest spelling, no quotes and no parentheses, three levels deep
      Loose                no annotation at all, so it sits at the workbook's root
      Late                 an annotation BELOW the first procedure, which is a comment, not a folder
      Twice                two annotations, of which the first is the folder
      Procedures           the parity module: the status bar's "current procedure" is held to the
                           editor's own ProcOfLine on every line of it
      Sheet1               a DOCUMENT module in a folder
      ReminderForm         a USERFORM in a folder, three levels deep beside Bare

    Meant to be opened ALONGSIDE FolderTwinFixture.xlsm (tools\New-FolderTwinFixture.ps1), whose
    module names and folder names collide with these on purpose: a Helpers in another folder, a
    Ledger in no folder, and a "Shared" folder of its own that must not merge with this one.

        tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\FolderFixture.xlsm,artifacts\fixtures\FolderTwinFixture.xlsm
        node tools\harness\folders.mjs

    Built through the xlide api, so "Trust access to the VBA project object model" does NOT have
    to be on. A Debug build must be registered and loading, which is what makes the door exist.
    Every module compiles: the suite runs beside the debugger suites and nothing here may trip
    them.

.EXAMPLE
    tools\New-FolderFixture.ps1
    Builds it, saves it, and leaves Excel open on it.

.EXAMPLE
    tools\New-FolderFixture.ps1 -Quiet
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
    $Path = Join-Path $fixtures 'FolderFixture.xlsm'
}

# Component types, as VBComponents.Add takes them.
$StandardModule = 1
$ClassModule = 2

$modules = [ordered]@{}

$modules['Ledger'] = @{ Kind = $StandardModule; Code = @"
'@Folder("Accounts.Ledger")
Option Explicit

' THE DOCUMENTED FORM, in a folder two levels deep. The twin holds a Ledger too, with no
' annotation at all, so the same name sits in a folder here and at the root there.

Public Sub Post(ByVal amount As Currency)
    Debug.Print "ledger " & amount
End Sub
"@ }

$modules['Posting'] = @{ Kind = $ClassModule; Code = @"
'@Folder("Accounts.Ledger")
Option Explicit

' A CLASS in the same folder as Ledger: a folder mixes kinds, and the flat tree's order within
' it (module before class) is the order the folder keeps.

Private mAmount As Currency

Public Property Get Amount() As Currency
    Amount = mAmount
End Property

Public Property Let Amount(ByVal value As Currency)
    mAmount = value
End Property
"@ }

$modules['Invoice'] = @{ Kind = $ClassModule; Code = @"
'@Folder("Accounts.Billing")
Option Explicit

' A SIBLING folder under Accounts, so Accounts has two children to sort.

Public Function Total() As Currency
    Total = 0
End Function
"@ }

$modules['Helpers'] = @{ Kind = $StandardModule; Code = @"
'@Folder("Shared")
Option Explicit

' THE COLLIDING MODULE. FolderTwinFixture.xlsm holds a Helpers too, in Accounts.Ledger. A tree
' that filed modules by name alone would put one of them in the wrong folder.

Public Sub Recalculate(ByVal label As String)
    Debug.Print "main " & label
End Sub
"@ }

$modules['Tools'] = @{ Kind = $StandardModule; Code = @"
Option Explicit

' THE LOOSE SPELLING: the name in lower case, the path in quotes but without parentheses, and
' the annotation on the second line rather than the first. The same folder as Helpers, so the
' folder is drawn once and spelled the way Helpers spelled it.
'@folder "shared"

Public Sub Tidy()
    Debug.Print "tidy"
End Sub
"@ }

$modules['Bare'] = @{ Kind = $StandardModule; Code = @"
'@Folder Accounts.Billing.Reminders
Option Explicit

' THE BAREST SPELLING, three levels deep. The form below sits beside it.

Public Sub Remind()
    Debug.Print "remind"
End Sub
"@ }

$modules['Loose'] = @{ Kind = $StandardModule; Code = @"
Option Explicit

' NO ANNOTATION. This module sits at the workbook's root, below the folders.

Public Sub Wander()
    Debug.Print "loose"
End Sub
"@ }

$modules['Late'] = @{ Kind = $StandardModule; Code = @"
Option Explicit

' The annotation is BELOW the first procedure, where it is a comment like any other. Root.

Public Sub First()
    Debug.Print "first"
End Sub

'@Folder("Nowhere")
Public Sub Second()
    Debug.Print "second"
End Sub
"@ }

$modules['Twice'] = @{ Kind = $StandardModule; Code = @"
'@Folder("Accounts")
'@Folder("Ignored")
Option Explicit

' TWO ANNOTATIONS: the first one is the folder, and the second is ignored rather than argued
' with. Directly under Accounts, beside its two subfolders.

Public Sub Once()
    Debug.Print "once"
End Sub
"@ }

$modules['Procedures'] = @{ Kind = $StandardModule; Code = @"
'@Folder("Shared")
Option Explicit

' THE PARITY MODULE. The suite puts the caret on EVERY line of this and compares the page's
' "current procedure" with what the editor's own ProcOfLine says, so this holds every shape
' that decides which procedure a line is in: a declarations section with comments and a
' Declare, a comment introducing a Sub, blank lines between procedures, a property pair sharing
' a name, a Static function, and comments trailing the last End.

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private mTotal As Long

' Adds one.
Public Sub Add()
    mTotal = mTotal + 1
End Sub


' The total so far.
Public Property Get Total() As Long
    Total = mTotal
End Property

Public Property Let Total(ByVal value As Long)
    mTotal = value
End Property

Private Static Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function

' Trailing, after the last procedure.
"@ }

$sheetCode = @"
'@Folder("Accounts")
Option Explicit

' A DOCUMENT module in a folder. Sheet1 has a folder here and none in the twin.

Public Sub Refresh()
    Debug.Print "sheet"
End Sub
"@

$forms = @(
    @{
        Name = 'ReminderForm'
        Controls = @(
            @{ Type = 'CommandButton'; Name = 'SendButton'; Left = 24; Top = 24; Width = 96; Height = 24 }
        )
        Code = @"
'@Folder("Accounts.Billing.Reminders")
Option Explicit

' A USERFORM in a folder, three levels deep beside Bare: its designer row sits under it there.

Private Sub SendButton_Click()
    Debug.Print "sent"
End Sub
"@
    }
)

# ---------------------------------------------------------------- building it

. (Join-Path $PSScriptRoot 'FixtureDriver.ps1')
Invoke-FixtureBuild -Path $Path -Modules $modules -SheetCode $sheetCode -Forms $forms -OpenAtEnd 'Ledger'

Write-Host ''
Write-Host "Fixture written to $Path"
Write-Host ''
Write-Host '  Accounts                      Twice, Sheet1'
Write-Host '    Billing                     Invoice'
Write-Host '      Reminders                 Bare, ReminderForm'
Write-Host '    Ledger                      Ledger, Posting'
Write-Host '  Shared                        Helpers, Procedures, Tools'
Write-Host '  (root)                        ThisWorkbook, Late, Loose'
Write-Host ''
Write-Host 'Open it ALONGSIDE FolderTwinFixture.xlsm for the cross-workbook cases:'
Write-Host '  tools\harness\Start-Excel.ps1 -Fresh -Workbook artifacts\fixtures\FolderFixture.xlsm,artifacts\fixtures\FolderTwinFixture.xlsm'
Write-Host '  node tools\harness\folders.mjs'
Write-Host ''

if ($Quiet) {
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
