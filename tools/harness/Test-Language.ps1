<#
.SYNOPSIS
    Non-ASCII text through the paths only a running host has: COM, the door, and the page.

.DESCRIPTION
    engine\test\language.mjs covers the engine over its pipe, headless, and it is a gate step.
    This covers what that one cannot reach, and what no unit test can:

      COM        a module's text is read and written through the VBA object model as a BSTR.
      the door   the debug api carries module text as an HTTP body.
      the page   the text crosses to WebView2 as a JSON message and becomes a monaco model.
      the tree   component and procedure names are rendered by the explorer.

    THE HOST'S OWN LIMIT IS THE HEADLINE. VBA stores module text in the system ANSI code page,
    not in Unicode, so a character outside that page does not survive being written at all: it
    comes back as a question mark. On a Western European system, accented Latin survives and CJK
    does not. That is Excel, not xlide, and it is worth measuring rather than assuming, because
    the answer depends on the machine this runs on.

    Run it after Start-Excel.ps1. It writes into a scratch module of the open workbook and puts
    the module back when it is done.

.EXAMPLE
    tools\harness\Test-Language.ps1
#>
[CmdletBinding()]
param(
    # The module to write into. Its original text is restored at the end.
    [string] $Module = 'HelpersExtra'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$probe = Join-Path $PSScriptRoot 'language-live-probe.mjs'

if (-not (Test-Path $probe)) {
    throw "The probe is missing: $probe"
}

node $probe --module $Module
if ($LASTEXITCODE -ne 0) {
    throw 'the live language probe failed'
}
