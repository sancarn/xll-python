using Microsoft.Win32.SafeHandles;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PythonExcel
{
    // Closing the last handle releases the IDE's entire process tree, including
    // language servers and the debug adapter, even if Excel exits unexpectedly.
    internal sealed class ProcessJob : SafeHandleZeroOrMinusOneIsInvalid
    {
        [StructLayout(LayoutKind.Sequential)]
        struct Limits {
            public long ProcessTime, JobTime;
            public uint Flags;
            public UIntPtr MinWorkingSet, MaxWorkingSet;
            public uint ActiveProcesses;
            public UIntPtr Affinity;
            public uint Priority, SchedulingClass;
            public ulong ReadOps, WriteOps, OtherOps, ReadBytes, WriteBytes, OtherBytes;
            public UIntPtr ProcessMemory, JobMemory, PeakProcessMemory, PeakJobMemory;
        }
        [DllImport("kernel32", SetLastError=true)]
        static extern IntPtr CreateJobObject(IntPtr attributes, string name);
        [DllImport("kernel32", SetLastError=true)]
        static extern bool SetInformationJobObject(ProcessJob job, int infoClass, ref Limits limits, int size);
        [DllImport("kernel32", SetLastError=true)]
        static extern bool AssignProcessToJobObject(ProcessJob job, IntPtr process);
        [DllImport("kernel32")]
        static extern bool CloseHandle(IntPtr value);

        internal ProcessJob() : base(true)
        {
            SetHandle(CreateJobObject(IntPtr.Zero, null));
            if (IsInvalid) throw new Win32Exception();
            var limits = new Limits { Flags = 0x2000 }; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            if (!SetInformationJobObject(this, 9, ref limits, Marshal.SizeOf(limits))) throw new Win32Exception();
        }
        internal void Add(Process process)
        {
            if (!AssignProcessToJobObject(this, process.Handle)) throw new Win32Exception();
        }
        protected override bool ReleaseHandle() => CloseHandle(handle);
    }
}
