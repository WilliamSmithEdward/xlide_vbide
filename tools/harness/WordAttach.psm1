<#
.SYNOPSIS
    Reaches a running Word's document Window, and the documents it holds, by process id.

.DESCRIPTION
    Shared by Start-Word.ps1 and New-WordFixture.ps1, the way AccessAttach.psm1 is shared by the
    Access pair: the launcher to press the editor's ribbon button and to sort the Words it may
    close from the ones it may not, the generator to close a session still holding the file it is
    about to replace.

    WORD ANSWERS ON ITS DOCUMENT PANE. Asking the `_WwG` child of the `OpusApp` frame for its
    native object model (OBJID_NATIVEOM) answers a Window whose `.Application` is Word - the same
    shape as Excel's worksheet pane, a few levels deeper in the tree. Not the running object
    table, for the reason every launcher here repeats: a host publishes itself there ten to forty
    seconds after it is visibly ready, and the window answers at once and names the instance
    rather than guessing between two.

    The Add-Type is guarded because the generator and the launcher run in ONE PowerShell process:
    the generator imports this module, then runs the launcher, which imports it again, and a type
    added twice under one name is an error rather than a no-op.
#>

if (-not ([System.Management.Automation.PSTypeName]'XlideHarness.AttachWord').Type) {
    Add-Type -Namespace XlideHarness -Name AttachWord -MemberDefinition @'
[DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc cb, IntPtr l);
[DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
[DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, System.Text.StringBuilder s, int m);
[DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object o);

delegate bool EnumProc(IntPtr h, IntPtr l);

// OBJID_NATIVEOM. Asking Word's document pane for its native object model yields the document
// Window, and its Application, without going through the running object table.
const uint NativeObjectModel = 0xFFFFFFF0u;

public static object DocumentWindowOf(int processId)
{
    IntPtr pane = IntPtr.Zero;

    EnumWindows((h, l) =>
    {
        int owner;
        GetWindowThreadProcessId(h, out owner);
        if (owner != processId) { return true; }

        EnumChildWindows(h, (child, l2) =>
        {
            var name = new System.Text.StringBuilder(128);
            GetClassNameW(child, name, 128);
            if (name.ToString() == "_WwG") { pane = child; return false; }
            return true;
        }, IntPtr.Zero);

        return pane == IntPtr.Zero;
    }, IntPtr.Zero);

    if (pane == IntPtr.Zero) { return null; }

    var dispatch = new Guid("00020400-0000-0000-C000-000000000046");
    object window;
    return AccessibleObjectFromWindow(pane, NativeObjectModel, ref dispatch, out window) == 0 ? window : null;
}
'@
}

<#
.SYNOPSIS
    The document Window of a Word process, or null when its pane cannot be reached in time.

.DESCRIPTION
    The caller owns the wrapper: release it with Marshal.ReleaseComObject before the process it
    names is stopped. A wrapper the collector finalises after its Word has been killed makes DCOM
    start a fresh hidden Word to answer the release (the Excel launcher, 2026-09-08).
#>
function Get-WordDocumentWindow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [int] $ProcessId,
        [int] $TimeoutSeconds = 0
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $window = [XlideHarness.AttachWord]::DocumentWindowOf($ProcessId)
        if ($null -ne $window) { return $window }
        if ((Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 50 }
    } while ((Get-Date) -lt $deadline)

    return $null
}

<#
.SYNOPSIS
    The full paths of the documents a running Word holds, or null when they cannot be read.

.DESCRIPTION
    Unreadable is NOT empty: a process whose documents cannot be read is reported as null so that
    callers treat it as a stranger rather than as an empty session worth closing. Every wrapper
    the read takes - the window, its Application, the Documents collection, each document - is
    given back before it returns, while the process can still answer the release.
#>
function Get-WordDocumentPaths {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)] [int] $ProcessId)

    $window = $null
    $application = $null
    $documents = $null
    try {
        $window = [XlideHarness.AttachWord]::DocumentWindowOf($ProcessId)
        if ($null -eq $window) { return $null }

        $application = $window.Application
        $documents = $application.Documents
        $held = @()
        foreach ($document in $documents) {
            $held += $document.FullName
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
        }
        return , $held
    }
    catch {
        return $null
    }
    finally {
        foreach ($wrapper in @($documents, $application, $window)) {
            if ($null -ne $wrapper) {
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wrapper) | Out-Null } catch { }
            }
        }
    }
}

Export-ModuleMember -Function Get-WordDocumentWindow, Get-WordDocumentPaths
