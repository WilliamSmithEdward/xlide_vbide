<#
.SYNOPSIS
    Builds the installer's icon from the product logo.

.DESCRIPTION
    Windows picks a different size out of an icon depending on where it draws it -- 16 pixels in a
    file list, 32 on the desktop, 256 in the preview pane -- and downscales itself if the size it
    wants is not there. Its downscaling is worse than doing it once, properly, up front, so every
    size ships.

    The result is committed. This runs when the logo changes, not as part of a build, which is why
    the icon is in source control and this script is not wired into anything.

.EXAMPLE
    tools\Generate-Icon.ps1
#>
[CmdletBinding()]
param(
    [string] $Source,
    [string] $Destination
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $Source) { $Source = Join-Path $repoRoot 'assets\images\extension_logo.png' }
if (-not $Destination) { $Destination = Join-Path $repoRoot 'assets\xlide.ico' }

if (-not (Test-Path $Source)) { throw "No logo at $Source." }

Add-Type -AssemblyName System.Drawing

# 256 is stored last on purpose: Windows reads the directory in order and some shells stop at the
# first entry large enough, so the small, hand-tuned sizes want to come first.
$sizes = @(16, 24, 32, 48, 64, 128, 256)

$logo = [System.Drawing.Image]::FromFile($Source)
try {
    $images = @()
    foreach ($size in $sizes) {
        $bitmap = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
            $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $graphics.DrawImage($logo, (New-Object System.Drawing.Rectangle(0, 0, $size, $size)))
        } finally {
            $graphics.Dispose()
        }

        $buffer = New-Object System.IO.MemoryStream
        $bitmap.Save($buffer, [System.Drawing.Imaging.ImageFormat]::Png)
        $bitmap.Dispose()
        $images += , @{ Size = $size; Bytes = $buffer.ToArray() }
        $buffer.Dispose()
    }
} finally {
    $logo.Dispose()
}

# An .ico is a six-byte header, then a sixteen-byte directory entry per image, then the images. Each
# entry here holds a PNG rather than a bitmap, which every Windows since Vista reads and which keeps
# the 256-pixel entry from costing 256 KB.
$output = New-Object System.IO.MemoryStream
$writer = New-Object System.IO.BinaryWriter($output)
try {
    $writer.Write([UInt16] 0)                 # reserved
    $writer.Write([UInt16] 1)                 # 1 = icon
    $writer.Write([UInt16] $images.Count)

    $offset = 6 + (16 * $images.Count)
    foreach ($image in $images) {
        # 256 is written as 0: the field is one byte and 256 does not fit.
        $dimension = if ($image.Size -ge 256) { 0 } else { $image.Size }
        $writer.Write([Byte] $dimension)      # width
        $writer.Write([Byte] $dimension)      # height
        $writer.Write([Byte] 0)               # palette size, 0 for truecolour
        $writer.Write([Byte] 0)               # reserved
        $writer.Write([UInt16] 1)             # colour planes
        $writer.Write([UInt16] 32)            # bits per pixel
        $writer.Write([UInt32] $image.Bytes.Length)
        $writer.Write([UInt32] $offset)
        $offset += $image.Bytes.Length
    }

    foreach ($image in $images) { $writer.Write($image.Bytes) }
    $writer.Flush()

    [System.IO.File]::WriteAllBytes($Destination, $output.ToArray())
} finally {
    $writer.Dispose()
    $output.Dispose()
}

$sizeList = ($images | ForEach-Object { $_.Size }) -join ', '
Write-Host ("Wrote {0} ({1:N1} KB, sizes {2})" -f $Destination, ((Get-Item $Destination).Length / 1KB), $sizeList)
