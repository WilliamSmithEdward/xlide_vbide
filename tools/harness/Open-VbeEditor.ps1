<#
.SYNOPSIS
    Opens the VBA editor in a running Excel, without needing the VBA project object model trusted.

.DESCRIPTION
    `$excel.VBE.MainWindow.Visible = $true` is the obvious way and it needs "Trust access to the
    VBA project object model" turned on, because `Application.VBE` is one of the two properties
    that setting gates. With it off that property is NULL — not an exception, a null — so the
    assignment fails with "You cannot call a method on a null-valued expression", or worse, a
    try/catch swallows it and the script carries on believing the editor is open.

    `CommandBars.ExecuteMso('VisualBasic')` is Excel executing its own Developer > Visual Basic
    button. It is not gated, and it opens the same editor.

    That matters because opening the editor is what LOADS the add-in, so every harness script
    starts here — and none of them should require a machine to lower a security setting that
    exists to stop macros rewriting each other (2026-08-07).

.EXAMPLE
    & (Join-Path $PSScriptRoot 'Open-VbeEditor.ps1') -Excel $excel
#>
[CmdletBinding()]
param(
    # A live Excel.Application, however it was attached.
    [Parameter(Mandatory = $true)] $Excel
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Excel.CommandBars.ExecuteMso('VisualBasic')
