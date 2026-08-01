using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Xlide.Vbe.Shim.Interop;

/// <summary>
/// Ties child processes to this one, so they die when it does.
///
/// A child started here is not reaped by anything if the host is terminated rather than closed:
/// nothing of ours runs, and the child is left with no client and no reason to exit. Users end
/// tasks, hosts crash, and debuggers stop processes, so that is not a rare path.
///
/// The operating system will do it for us. A job object configured to kill on close terminates
/// every assigned process the moment the last handle to the job goes away, which happens when this
/// process ends however it ends. It needs no cooperation from the child and no code of ours to run.
/// </summary>
internal sealed partial class KillOnCloseJob : IDisposable
{
    private nint _handle;

    private KillOnCloseJob(nint handle) => _handle = handle;

    public static KillOnCloseJob? Create()
    {
        var handle = CreateJobObject(0, null);
        if (handle == 0)
        {
            return null;
        }

        var information = new ExtendedLimitInformation
        {
            BasicLimitInformation = new BasicLimitInformation { LimitFlags = LimitKillOnJobClose },
        };

        var size = Marshal.SizeOf<ExtendedLimitInformation>();
        var buffer = Marshal.AllocHGlobal(size);

        try
        {
            Marshal.StructureToPtr(information, buffer, fDeleteOld: false);

            if (!SetInformationJobObject(handle, ExtendedLimitInformationClass, buffer, (uint)size))
            {
                CloseHandle(handle);
                return null;
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }

        return new KillOnCloseJob(handle);
    }

    public bool Assign(Process process)
    {
        ArgumentNullException.ThrowIfNull(process);
        return _handle != 0 && AssignProcessToJobObject(_handle, process.Handle);
    }

    public void Dispose()
    {
        var handle = _handle;
        _handle = 0;

        if (handle != 0)
        {
            // Closing the last handle is what terminates the assigned processes.
            CloseHandle(handle);
        }
    }

    private const int ExtendedLimitInformationClass = 9;
    private const uint LimitKillOnJobClose = 0x00002000;

    [StructLayout(LayoutKind.Sequential)]
    private struct BasicLimitInformation
    {
        public long PerProcessUserTimeLimit;
        public long PerJobUserTimeLimit;
        public uint LimitFlags;
        public nuint MinimumWorkingSetSize;
        public nuint MaximumWorkingSetSize;
        public uint ActiveProcessLimit;
        public nuint Affinity;
        public uint PriorityClass;
        public uint SchedulingClass;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct IoCounters
    {
        public ulong ReadOperationCount;
        public ulong WriteOperationCount;
        public ulong OtherOperationCount;
        public ulong ReadTransferCount;
        public ulong WriteTransferCount;
        public ulong OtherTransferCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct ExtendedLimitInformation
    {
        public BasicLimitInformation BasicLimitInformation;
        public IoCounters IoInfo;
        public nuint ProcessMemoryLimit;
        public nuint JobMemoryLimit;
        public nuint PeakProcessMemoryUsed;
        public nuint PeakJobMemoryUsed;
    }

    [LibraryImport("kernel32.dll", EntryPoint = "CreateJobObjectW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    private static partial nint CreateJobObject(nint attributes, string? name);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool SetInformationJobObject(nint job, int infoClass, nint info, uint length);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool AssignProcessToJobObject(nint job, nint process);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool CloseHandle(nint handle);
}
