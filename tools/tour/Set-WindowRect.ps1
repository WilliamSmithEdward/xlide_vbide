<#
.SYNOPSIS
    Puts a window at an exact place and size, in physical pixels.

.DESCRIPTION
    For the tour: PrintWindow captures at the window's own size, so the stills' dimensions ARE
    whatever the frame was left at. Two things make "exact" true. The process opts into
    per-monitor DPI awareness first, because an unaware caller works in scaled coordinates and
    the window lands at 1.5x whatever was asked on a scaled monitor. And the window is restored
    first, because a maximized window ignores SetWindowPos sizes entirely.

.EXAMPLE
    tools\tour\Set-WindowRect.ps1 -Handle 67588 -Width 1280 -Height 800
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][long] $Handle,
    [Parameter(Mandatory)][int] $Width,
    [Parameter(Mandatory)][int] $Height,
    [int] $X = 40,
    [int] $Y = 40
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -Namespace XlideTour -Name Mover -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr handle, int command);
[DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr handle, IntPtr after, int x, int y, int width, int height, uint flags);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr handle);
[DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr handle, out RECT rect);
public struct RECT { public int Left, Top, Right, Bottom; }
'@

# PER_MONITOR_AWARE_V2. Refused when the process already committed to a DPI mode; that is fine,
# it means the default applied and the coordinates below are already physical on this machine.
[void][XlideTour.Mover]::SetProcessDpiAwarenessContext([IntPtr]::new(-4))

$window = [IntPtr]::new($Handle)
[void][XlideTour.Mover]::ShowWindow($window, 9)   # SW_RESTORE: a maximized window ignores sizes
[void][XlideTour.Mover]::SetWindowPos($window, [IntPtr]::Zero, $X, $Y, $Width, $Height, 0x4)  # SWP_NOZORDER
[void][XlideTour.Mover]::SetForegroundWindow($window)

$rect = New-Object XlideTour.Mover+RECT
[void][XlideTour.Mover]::GetWindowRect($window, [ref] $rect)
Write-Output ("{0},{1} {2}x{3}" -f $rect.Left, $rect.Top, ($rect.Right - $rect.Left), ($rect.Bottom - $rect.Top))
