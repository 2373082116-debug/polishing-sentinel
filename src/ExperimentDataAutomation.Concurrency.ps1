Set-StrictMode -Version 2.0

function Enter-ExperimentWorkbookLock {
    param(
        [Parameter(Mandatory = $true)] [string] $WorkbookPath,
        [ValidateSet('Workbook', 'Monitor')] [string] $Purpose = 'Workbook',
        [int] $WaitMilliseconds = -1
    )
    $path = [IO.Path]::GetFullPath($WorkbookPath).Replace('/', '\').ToUpperInvariant()
    $hash = [Security.Cryptography.SHA256]::Create()
    try { $key = [BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($path))).Replace('-', '') }
    finally { $hash.Dispose() }
    $mutex = New-Object Threading.Mutex($false, ('Global\PolishingSentinel.{0}.{1}' -f $Purpose, $key))
    try {
        try { $acquired = $mutex.WaitOne($WaitMilliseconds) }
        catch [Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { $mutex.Dispose(); return $null }
        return $mutex
    }
    catch { $mutex.Dispose(); throw }
}

function Exit-ExperimentWorkbookLock {
    param([Parameter(Mandatory = $true)] [Threading.Mutex] $Lease)
    try { $Lease.ReleaseMutex() } finally { $Lease.Dispose() }
}

function Invoke-WithExperimentWorkbookLock {
    param(
        [Parameter(Mandatory = $true)] [string] $WorkbookPath,
        [Parameter(Mandatory = $true)] [scriptblock] $Action
    )
    $lease = Enter-ExperimentWorkbookLock -WorkbookPath $WorkbookPath
    try { & $Action } finally { Exit-ExperimentWorkbookLock -Lease $lease }
}
