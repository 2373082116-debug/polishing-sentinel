Set-StrictMode -Version 2.0

function Initialize-ExperimentExcelOwnership {
    if ('SentinelOwnedExcelJob' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32.SafeHandles;

public sealed class SentinelOwnedExcelJob : IDisposable {
    [StructLayout(LayoutKind.Sequential)] struct BasicLimits {
        public long ProcessTime, JobTime;
        public uint Flags;
        public UIntPtr MinimumWorkingSet, MaximumWorkingSet;
        public uint ActiveProcesses;
        public UIntPtr Affinity;
        public uint Priority, Scheduling;
    }
    [StructLayout(LayoutKind.Sequential)] struct IoCounters {
        public ulong ReadOps, WriteOps, OtherOps, ReadBytes, WriteBytes, OtherBytes;
    }
    [StructLayout(LayoutKind.Sequential)] struct ExtendedLimits {
        public BasicLimits Basic;
        public IoCounters Io;
        public UIntPtr ProcessMemory, JobMemory, PeakProcessMemory, PeakJobMemory;
    }
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern IntPtr CreateJobObject(IntPtr security, string name);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool SetInformationJobObject(SafeFileHandle job, int kind, ref ExtendedLimits limits, uint size);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool AssignProcessToJobObject(SafeFileHandle job, IntPtr process);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool TerminateJobObject(SafeFileHandle job, uint exitCode);
    [DllImport("user32.dll", SetLastError = true)]
    static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    readonly SafeFileHandle job;
    readonly object sync = new object();
    Timer watchdog;
    bool disposed, timedOut;
    public bool TimedOut { get { lock (sync) { return timedOut; } } }
    public int TerminationError { get; private set; }
    public static int WindowProcessId(IntPtr window) {
        uint id;
        if (GetWindowThreadProcessId(window, out id) == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
        return checked((int)id);
    }
    public SentinelOwnedExcelJob(IntPtr processHandle, int timeoutSeconds) {
        if (timeoutSeconds < 1 || timeoutSeconds > 86400) throw new ArgumentOutOfRangeException("timeoutSeconds");
        job = new SafeFileHandle(CreateJobObject(IntPtr.Zero, null), true);
        if (job.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
        try {
            ExtendedLimits limits = new ExtendedLimits();
            // Last handle closes on normal disposal or controller termination.
            limits.Basic.Flags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            if (!SetInformationJobObject(job, 9, ref limits, (uint)Marshal.SizeOf(typeof(ExtendedLimits))))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            if (!AssignProcessToJobObject(job, processHandle))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            watchdog = new Timer(OnTimeout, null, timeoutSeconds * 1000, Timeout.Infinite);
        }
        catch { job.Dispose(); throw; }
    }
    void OnTimeout(object ignored) {
        lock (sync) {
            if (disposed) return;
            timedOut = true;
            try { if (!TerminateJobObject(job, 1460)) TerminationError = Marshal.GetLastWin32Error(); }
            catch { TerminationError = -1; }
        }
    }
    public void Dispose() {
        lock (sync) {
            if (disposed) return;
            disposed = true;
            if (watchdog != null) watchdog.Dispose();
            job.Dispose();
        }
    }
}
'@
}

function New-ExperimentExcelSession {
    [CmdletBinding()]
    param([int] $TimeoutSeconds = 0)
    if (-not $PSBoundParameters.ContainsKey('TimeoutSeconds')) {
        $policy = Get-Variable -Name ExperimentExcelTimeoutSeconds -Scope Script -ErrorAction SilentlyContinue
        $TimeoutSeconds = if ($policy -and $null -ne $policy.Value) { [int]$policy.Value } else { 600 }
    }
    if ($TimeoutSeconds -lt 1 -or $TimeoutSeconds -gt 86400) { throw 'Excel 超时必须在 1 至 86400 秒之间。' }
    Initialize-ExperimentExcelOwnership
    $existingIds = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $excel = $null; $process = $null; $guard = $null; $owned = $false
    $health = $null; $operationToken = 0
    try {
        $excel = New-Object -ComObject Excel.Application
        $processId = [SentinelOwnedExcelJob]::WindowProcessId([IntPtr]$excel.Hwnd)
        if ($existingIds -contains $processId -or $excel.UserControl -or $excel.Workbooks.Count -ne 0) {
            throw 'Excel 未创建独立的自动化实例；为保护已有窗口，已取消操作。'
        }
        $process = Get-Process -Id $processId -ErrorAction Stop
        if ($process.ProcessName -ne 'EXCEL') { throw 'Excel 进程身份核验失败。' }
        $owned = $true
        $guard = [SentinelOwnedExcelJob]::new($process.Handle, $TimeoutSeconds)
        $healthVariable = Get-Variable -Name SentinelCurrentHealth -Scope Script -ErrorAction SilentlyContinue
        if ($healthVariable -and $null -ne $healthVariable.Value) {
            $health = $healthVariable.Value
            $operationToken = $health.BeginOperation($TimeoutSeconds)
        }
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        return [pscustomobject]@{ Application = $excel; Process = $process; Guard = $guard; TimeoutSeconds = $TimeoutSeconds; Health = $health; OperationToken = $operationToken }
    }
    catch {
        $timedOut = $null -ne $guard -and $guard.TimedOut
        try { if ($owned -and $null -ne $excel -and -not $timedOut) { $excel.Quit() } }
        catch { Write-Warning "新建 Excel 实例清理失败：$($_.Exception.Message)" }
        finally {
            Release-ComObject $excel
            if ($null -ne $guard) { $guard.Dispose() }
            if ($null -ne $health) { $health.EndOperation($operationToken) }
            if ($null -ne $process) {
                try {
                    if ($owned -and -not $process.HasExited) { $process.Kill(); [void]$process.WaitForExit(5000) }
                }
                finally { $process.Dispose() }
            }
        }
        if ($timedOut) { throw [TimeoutException]::new("Excel 会话超过 $TimeoutSeconds 秒，已回收专用实例；未完成的写入不提交。") }
        throw
    }
}

function Close-ExperimentExcelSession {
    [CmdletBinding()]
    param([AllowNull()] $Session, [AllowNull()] $Workbook = $null)
    if ($null -eq $Session) { return }
    $cleanupErrors = @()
    $timedOut = $Session.Guard.PSObject.Properties['TimedOut'] -and $Session.Guard.TimedOut
    try {
        try { if ($null -ne $Workbook -and -not $timedOut) { $Workbook.Close($false) } }
        catch { $cleanupErrors += "关闭工作簿失败：$($_.Exception.Message)" }
        $timedOut = $Session.Guard.PSObject.Properties['TimedOut'] -and $Session.Guard.TimedOut
        try { if ($null -ne $Session.Application -and -not $timedOut) { $Session.Application.Quit() } }
        catch { $cleanupErrors += "退出 Excel 失败：$($_.Exception.Message)" }
    }
    finally {
        try {
            Release-ComObject $Workbook
            Release-ComObject $Session.Application
        }
        finally {
            try {
                if ($null -ne $Session.Guard) { $Session.Guard.Dispose() }
                $timedOut = $Session.Guard.PSObject.Properties['TimedOut'] -and $Session.Guard.TimedOut
                if ($Session.PSObject.Properties['Health'] -and $null -ne $Session.Health) { $Session.Health.EndOperation($Session.OperationToken) }
            }
            finally {
                if ($null -ne $Session.Process) {
                    try {
                        if (-not $Session.Process.WaitForExit(5000)) { throw '本次 Excel 实例尚未退出，需要检查。' }
                    }
                    finally { $Session.Process.Dispose() }
                }
            }
        }
    }
    if ($timedOut) { throw [TimeoutException]::new("Excel 会话超过 $($Session.TimeoutSeconds) 秒，已回收专用实例；未完成的写入不提交。") }
    foreach ($message in $cleanupErrors) { Write-Warning $message }
}
