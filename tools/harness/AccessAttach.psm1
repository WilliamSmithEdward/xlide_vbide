<#
.SYNOPSIS
    Reaches a running Access's own Application object, by process id.

.DESCRIPTION
    Shared by Start-Access.ps1 and New-AccessFixture.ps1, which both need it and needed it at
    different moments: the launcher to press the editor's ribbon button, the generator to close a
    session still holding the file it is about to replace and to save each module by name.

    ACCESS ANSWERS ON ITS FRAME, and answers with the Application ITSELF. Excel's worksheet pane
    and Word's document pane answer OBJID_NATIVEOM with a Window whose `.Application` is the host;
    Access answers on the top-level `OMain` window and what comes back reads "Microsoft Access"
    from `.Name`. There is no document window to walk to either - Access's MDI children are forms
    and reports, so a database with none open would have nothing to ask.

    Not the running object table, for the reason every launcher here repeats: a host publishes
    itself there ten to forty seconds after it is visibly ready, and the window answers at once
    and names the instance rather than guessing between two.
#>

if (-not ([System.Management.Automation.PSTypeName]'XlideHarness.AttachAccess').Type) {
    Add-Type -Namespace XlideHarness -Name AttachAccess -MemberDefinition @'
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object o);
delegate bool EnumProc(IntPtr h, IntPtr l);

const uint NativeObjectModel = 0xFFFFFFF0u;

public static object ApplicationOf(int processId)
{
    IntPtr frame = IntPtr.Zero;
    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId) { return true; }
        var name = new System.Text.StringBuilder(128);
        GetClassNameW(h, name, 128);
        if (name.ToString() == "OMain") { frame = h; return false; }
        return true;
    }, IntPtr.Zero);
    if (frame == IntPtr.Zero) { return null; }
    var dispatch = new Guid("00020400-0000-0000-C000-000000000046");
    object application;
    return AccessibleObjectFromWindow(frame, NativeObjectModel, ref dispatch, out application) == 0 ? application : null;
}
'@
}

<#
.SYNOPSIS
    The Access Application in a process, or null when its frame cannot be reached.
#>
function Get-AccessApplication {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [int] $ProcessId,
        [int] $TimeoutSeconds = 0
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $application = [XlideHarness.AttachAccess]::ApplicationOf($ProcessId)
        if ($null -ne $application) { return $application }
        if ((Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 50 }
    } while ((Get-Date) -lt $deadline)

    return $null
}

<#
.SYNOPSIS
    The full path of the database a running Access holds, or null when it cannot be read.

.DESCRIPTION
    Unreadable is NOT empty: a process whose database cannot be read is reported as null so that
    callers treat it as a stranger rather than as an empty session worth closing.
#>
function Get-AccessDatabasePath {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [int] $ProcessId)

    try {
        $application = Get-AccessApplication -ProcessId $ProcessId
        if ($null -eq $application) { return $null }
        return $application.CurrentProject.FullName
    }
    catch {
        return $null
    }
}

Export-ModuleMember -Function Get-AccessApplication, Get-AccessDatabasePath
