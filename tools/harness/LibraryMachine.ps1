<#
.SYNOPSIS
    The machine a Windows executable or library was built for, read from its own header.

.DESCRIPTION
    Dot-source it. Get-LibraryMachine answers 0x8664 for 64-bit and 0x014C for 32-bit, and
    Get-LibraryWidth answers 64 or 32. Nothing is loaded or run: four kilobytes are read.
#>

function Get-LibraryMachine([string] $Path) {
    # Resolved here: .NET opens a relative path against the process's folder, not PowerShell's.
    $Path = (Resolve-Path -LiteralPath $Path).Path
    $stream = [IO.File]::OpenRead($Path)
    try {
        $header = New-Object byte[] 4096
        $read = $stream.Read($header, 0, $header.Length)
    } finally {
        $stream.Dispose()
    }

    $at = [BitConverter]::ToInt32($header, 0x3C)
    if ($at -le 0 -or ($at + 6) -gt $read -or $header[$at] -ne 0x50 -or $header[$at + 1] -ne 0x45) {
        throw "$Path is not a Windows executable or library."
    }

    return [BitConverter]::ToUInt16($header, $at + 4)
}

function Get-LibraryWidth([string] $Path) {
    $machine = Get-LibraryMachine $Path
    switch ($machine) {
        0x8664 { return 64 }
        0x014C { return 32 }
        default { throw ("{0} is built for machine 0x{1:X4}, which is neither width of Office." -f $Path, $machine) }
    }
}
