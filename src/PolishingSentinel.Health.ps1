Set-StrictMode -Version 2.0
if (-not ('SentinelHealthMonitor' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Collections.Generic;
using System.Globalization;
public sealed class SentinelHealthMonitor : IDisposable {
    public sealed class Beat {
        public int OwnerProcessId;
        public long OwnerStartTimeUtcTicks, HeartbeatUtcTicks, StageSequence, StageStartedUtcTicks, OperationToken, OperationDeadlineUtcTicks;
    }
    static readonly Dictionary<string, SentinelHealthMonitor> monitors = new Dictionary<string, SentinelHealthMonitor>(StringComparer.OrdinalIgnoreCase);
    readonly object gate = new object();
    readonly string directory, path;
    readonly int pid;
    readonly long start;
    readonly Dictionary<long, long> operations = new Dictionary<long, long>();
    Timer timer;
    bool disposed;
    long sequence, stage, token;
    public static SentinelHealthMonitor Get(string root, int pid, long start) {
        root = Path.GetFullPath(root);
        lock(monitors) {
            SentinelHealthMonitor m;
            if (!monitors.TryGetValue(root, out m) || m.disposed) {
                m = new SentinelHealthMonitor(root, pid, start); monitors[root] = m;
            }
            return m;
        }
    }
    public static void Stop(string root) {
        lock(monitors) {
            SentinelHealthMonitor m; root = Path.GetFullPath(root);
            if (monitors.TryGetValue(root, out m)) { m.Dispose(); monitors.Remove(root); }
        }
    }
    SentinelHealthMonitor(string root, int id, long ticks) {
        directory = Path.Combine(root, "state"); pid = id; start = ticks;
        path = Path.Combine(directory, "heartbeat-" + pid.ToString(CultureInfo.InvariantCulture) + "-" + start.ToString(CultureInfo.InvariantCulture) + ".json");
        stage = DateTime.UtcNow.Ticks; Publish();
        timer = new Timer(delegate { lock(gate) { if (Available()) Publish(); } }, null, 1000, 1000);
    }
    bool Available() {
        if (disposed) return false;
        if (!Directory.Exists(directory)) { Dispose(); return false; }
        return true;
    }
    public Beat Advance() {
        lock(gate) { if (Available()) { sequence++; stage = DateTime.UtcNow.Ticks; Publish(); } return Snapshot(); }
    }
    public long BeginOperation(int seconds) {
        lock(gate) {
            if (!Available()) return 0;
            if (seconds < 1 || seconds > 86400) throw new ArgumentOutOfRangeException("seconds");
            long next = ++token; operations.Add(next, DateTime.UtcNow.AddSeconds(seconds).Ticks); Publish(); return next;
        }
    }
    public void EndOperation(long id) { lock(gate) { if (Available()) { operations.Remove(id); Publish(); } } }
    public Beat Snapshot() {
        lock(gate) {
            Beat b = new Beat { OwnerProcessId = pid, OwnerStartTimeUtcTicks = start, HeartbeatUtcTicks = DateTime.UtcNow.Ticks, StageSequence = sequence, StageStartedUtcTicks = stage };
            foreach (var op in operations) {
                if (b.OperationDeadlineUtcTicks == 0 || op.Value < b.OperationDeadlineUtcTicks) { b.OperationToken = op.Key; b.OperationDeadlineUtcTicks = op.Value; }
            }
            return b;
        }
    }
    void Publish() {
        try {
            Beat b = Snapshot();
            string json = String.Format(CultureInfo.InvariantCulture,
                "{{\"OwnerProcessId\":{0},\"OwnerStartTimeUtcTicks\":{1},\"HeartbeatUtcTicks\":{2},\"StageSequence\":{3},\"StageStartedUtcTicks\":{4},\"OperationToken\":{5},\"OperationDeadlineUtcTicks\":{6}}}",
                b.OwnerProcessId, b.OwnerStartTimeUtcTicks, b.HeartbeatUtcTicks, b.StageSequence, b.StageStartedUtcTicks, b.OperationToken, b.OperationDeadlineUtcTicks);
            WriteAtomic(path, json);
        } catch { /* Heartbeat I/O failure must not crash the host. */ }
    }
    public static void WriteAtomic(string target, string json) {
        string temp = target + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try {
            File.WriteAllText(temp, json, new UTF8Encoding(true));
            if (File.Exists(target)) File.Replace(temp, target, null);
            else {
                try { File.Move(temp, target); }
                catch (IOException) { if (!File.Exists(target)) throw; File.Replace(temp, target, null); }
            }
        } finally { if (File.Exists(temp)) File.Delete(temp); }
    }
    public static string ReadAtomic(string path) {
        for (int attempt = 0; ; attempt++) {
            try {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (var reader = new StreamReader(stream, Encoding.UTF8)) return reader.ReadToEnd();
            } catch (IOException) { if (attempt >= 4) throw; Thread.Sleep(10); }
        }
    }
    public void Dispose() {
        lock(gate) {
            if (disposed) return;
            disposed = true; if (timer != null) timer.Dispose(); operations.Clear();
            try { File.Delete(path); } catch { }
        }
    }
}
'@
}
function Stop-PolishingSentinelHealth {
    param([Parameter(Mandatory = $true)] [string] $ScriptRoot)
    [SentinelHealthMonitor]::Stop($ScriptRoot)
}
function Add-PolishingSentinelHealthState {
    param($Status, [string] $ScriptRoot)
    $terminal = $Status.State -in @('Completed', 'PartialSuccess', 'Error', 'DryRun', 'TimedOut', 'Unknown', 'NeedsConfirmation', 'RetryPending')
    $health = 'Stale'; $text = '缺少可核验的后台心跳'; $elapsed = 0; $age = $null
    $now = [datetime]::UtcNow.Ticks
    if ($Status.PSObject.Properties['StageStartedUtcTicks']) {
        $elapsed = [math]::Max(0, [math]::Floor(($now - [long]$Status.StageStartedUtcTicks) / 10000000.0))
    }
    if ($Status.PSObject.Properties['OwnerProcessId'] -and $Status.PSObject.Properties['OwnerStartTimeUtcTicks']) {
        $owner = Get-Process -Id $Status.OwnerProcessId -ErrorAction SilentlyContinue
        $alive = $false
        try { $alive = $null -ne $owner -and $owner.StartTime.ToUniversalTime().Ticks -eq [long]$Status.OwnerStartTimeUtcTicks }
        catch { $alive = $false }
        finally { if ($null -ne $owner) { $owner.Dispose() } }
        if (-not $alive) { $health = 'Stopped'; $text = '后台进程已退出' }
        else {
            $path = Join-Path $ScriptRoot ('state/heartbeat-{0}-{1}.json' -f $Status.OwnerProcessId, $Status.OwnerStartTimeUtcTicks)
            try {
                $beat = [SentinelHealthMonitor]::ReadAtomic($path) | ConvertFrom-Json
                if ($beat.OwnerProcessId -ne $Status.OwnerProcessId -or $beat.OwnerStartTimeUtcTicks -ne $Status.OwnerStartTimeUtcTicks) { throw 'Heartbeat owner mismatch' }
                $age = ($now - [long]$beat.HeartbeatUtcTicks) / 10000000.0
                if ($age -gt 5 -or $age -lt -5) { $text = '后台心跳已过期，请检查监控进程' }
                else {
                    $health = 'Healthy'; $text = '后台心跳正常（不代表任务已有进展）'
                    if (-not $terminal -and [long]$beat.OperationDeadlineUtcTicks -gt 0 -and $now -gt [long]$beat.OperationDeadlineUtcTicks) {
                        $health = 'TimedOut'; $text = 'Excel 操作已超过期限，正在等待回收结果'
                    }
                    elseif (-not $terminal -and $Status.State -ne 'WaitingForUsb' -and $elapsed -gt 300 -and [long]$beat.OperationDeadlineUtcTicks -eq 0) {
                        $health = 'SuspectedStall'; $text = '本阶段超过 5 分钟未更新，可能耗时较长或已卡住'
                    }
                }
            } catch { $health = 'Stale'; $text = '后台心跳缺失或不可读取' }
        }
    }
    if (-not $terminal) {
        if ($health -eq 'Stopped') { $Status.State = 'Exited' }
        elseif ($health -in @('Stale', 'TimedOut')) { $Status.State = $health }
    }
    $Status | Add-Member NoteProperty HealthState $health -Force
    $Status | Add-Member NoteProperty HealthText $text -Force
    $Status | Add-Member NoteProperty StageElapsedSeconds $elapsed -Force
    $Status | Add-Member NoteProperty HeartbeatAgeSeconds $age -Force
    return $Status
}
