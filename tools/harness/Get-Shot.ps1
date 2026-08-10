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
#>
[CmdletBinding()]
param(
    # A CSS selector to crop to. Omitted, the whole window is returned.
    [string] $Selector,

    # Pixels of breathing room around the element.
    [int] $Pad = 8,

    # Which surface: the editor frame, or the Object Browser palette.
    [ValidateSet('frame', 'palette')]
    [string] $Window = 'frame',

    # Where to write the PNG.
    [string] $Out = 'shot.png'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $excel) { throw 'No Excel is running.' }

$discovery = Join-Path $env:LOCALAPPDATA "xlide_vbide\debug-api-$($excel.Id).json"
if (-not (Test-Path $discovery)) { throw "No debug api for Excel $($excel.Id); is this a Debug build?" }

$d = Get-Content $discovery -Raw | ConvertFrom-Json
$api = "http://127.0.0.1:$($d.port)/$($d.token)"

$query = "window=$Window"
if ($Selector) {
    $query += "&selector=$([uri]::EscapeDataString($Selector))&pad=$Pad"
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
