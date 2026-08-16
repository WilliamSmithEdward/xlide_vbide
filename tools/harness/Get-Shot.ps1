<#
.SYNOPSIS
    A picture of the running editor - or of one widget in it - as a PNG.

.DESCRIPTION
    The debug api renders windows as BMP, because a bitmap is what PrintWindow gives and
    encoding anything else inside an ahead-of-time shim is work the shim should not carry.
    Most things that want to LOOK at the result want a PNG, so the conversion happens here.

    The `-Selector` form is the useful one during interface work: the whole frame is a large
    picture in which a 54-pixel drop zone cannot be seen, and a surface built by reading
    numbers rather than looking at it is a surface built with one eye shut.

.EXAMPLE
    tools\harness\Get-Shot.ps1
    The whole editor frame, to shot.png.

.EXAMPLE
    tools\harness\Get-Shot.ps1 -Selector '.drop-compass' -Out compass.png
    Just the drag compass, with a little padding around it.

.EXAMPLE
    tools\harness\Get-Shot.ps1 -Selector '#dock-bottom' -Pad 0

.EXAMPLE
    tools\harness\Get-Shot.ps1 -Window form -Caption 'Quarter Entry' -Out live.png
    The form as the RUNTIME paints it, which is the only way to check a designer's work
    against the thing the developer will actually see.
#>
[CmdletBinding()]
param(
    # A CSS selector to crop to. Omitted, the whole window is returned.
    [string] $Selector,

    # Pixels of breathing room around the element.
    [int] $Pad = 8,

    # Which surface: the editor frame, the Object Browser palette, or a RUNNING form.
    [ValidateSet('frame', 'palette', 'form')]
    [string] $Window = 'frame',

    # With -Window form, which running form: any caption it holds. The first one standing
    # otherwise. A running form is the only picture of a designer's work the object model
    # cannot answer for, because MSForms draws its controls without windows of their own.
    [string] $Caption,

    # Where to write the PNG.
    [string] $Out = 'shot.png'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'XlideApi.psm1') -Force
$found = Get-XlideApi
$api = $found.Base

$query = "window=$Window"
if ($Selector) {
    $query += "&selector=$([uri]::EscapeDataString($Selector))&pad=$Pad"
}
if ($Caption) {
    $query += "&caption=$([uri]::EscapeDataString($Caption))"
}

$bmp = [IO.Path]::GetTempFileName() + '.bmp'
try {
    Invoke-WebRequest "$api/capture?$query" -OutFile $bmp -TimeoutSec 30

    # An error comes back as JSON, not an image, and a "png" that is really an error message
    # is a confusing thing to open.
    $head = [IO.File]::ReadAllBytes($bmp)[0..1]
    if ($head[0] -ne 0x42 -or $head[1] -ne 0x4D) {
        throw (Get-Content $bmp -Raw)
    }

    Add-Type -AssemblyName System.Drawing
    $image = [Drawing.Image]::FromFile($bmp)
    try {
        $full = if ([IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path (Get-Location) $Out }
        $image.Save($full, [Drawing.Imaging.ImageFormat]::Png)
        Write-Output ("{0}  {1}x{2}" -f $full, $image.Width, $image.Height)
    } finally {
        $image.Dispose()
    }
} finally {
    Remove-Item $bmp -Force -ErrorAction SilentlyContinue
}
