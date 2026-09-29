$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptRoot 'PolishingSentinel.Status.ps1')

$processes = @(Get-PolishingSentinelMonitorProcesses -ScriptRoot $scriptRoot)

if ($processes.Count -eq 0) {
    Write-Host '未发现正在运行的实验数据自动整理监控进程。'
    exit 0
}

foreach ($process in $processes) {
    Stop-Process -Id $process.ProcessId -Force
    Write-Host "已停止监控进程：$($process.ProcessId)"
}
