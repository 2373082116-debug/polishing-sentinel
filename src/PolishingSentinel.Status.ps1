Set-StrictMode -Version 2.0

. (Join-Path $PSScriptRoot 'PolishingSentinel.Health.ps1')

function Get-PolishingSentinelStatusPath {
    param([Parameter(Mandatory = $true)] [string] $ScriptRoot)

    return (Join-Path (Join-Path $ScriptRoot 'state') 'runtime-status.json')
}

function Update-PolishingSentinelStatus {
    param(
        [Parameter(Mandatory = $true)] [string] $ScriptRoot,
        [Parameter(Mandatory = $true)] [string] $State,
        [Parameter(Mandatory = $true)] [string] $Message,
        [hashtable] $Data = @{}
    )

    $statusPath = Get-PolishingSentinelStatusPath -ScriptRoot $ScriptRoot
    $statusDirectory = Split-Path -Parent $statusPath
    if (-not (Test-Path -LiteralPath $statusDirectory)) {
        New-Item -ItemType Directory -Path $statusDirectory -Force | Out-Null
    }

    $process = Get-Process -Id $PID
    try { $startTicks = $process.StartTime.ToUniversalTime().Ticks } finally { $process.Dispose() }
    $script:SentinelCurrentHealth = [SentinelHealthMonitor]::Get($ScriptRoot, $PID, $startTicks)
    $beat = $script:SentinelCurrentHealth.Advance()
    $status = [pscustomobject]@{
        AppName   = '抛光哨站'
        State     = $State
        Message   = $Message
        UpdatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Data      = $Data
        OwnerProcessId = $PID
        OwnerStartTimeUtcTicks = $startTicks
        StageSequence = $beat.StageSequence
        StageStartedUtcTicks = $beat.StageStartedUtcTicks
    }

    [SentinelHealthMonitor]::WriteAtomic($statusPath, ($status | ConvertTo-Json -Depth 8))
    return $status
}

function Read-PolishingSentinelStatus {
    param([Parameter(Mandatory = $true)] [string] $ScriptRoot)

    $statusPath = Get-PolishingSentinelStatusPath -ScriptRoot $ScriptRoot
    if (-not (Test-Path -LiteralPath $statusPath)) {
        return [pscustomobject]@{
            AppName   = '抛光哨站'
            State     = 'Unknown'
            Message   = '尚未写入运行状态'
            UpdatedAt = ''
            Data      = [pscustomobject]@{}
        }
    }

    $status = [SentinelHealthMonitor]::ReadAtomic($statusPath) | ConvertFrom-Json
    return (Add-PolishingSentinelHealthState -Status $status -ScriptRoot $ScriptRoot)
}

function Test-PolishingSentinelPathSet {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Settings)

    $items = @(
        @{ Key = 'MasterWorkbookPath'; Label = '主实验数据表'; Path = $Settings.MasterWorkbookPath; Type = 'File' },
        @{ Key = 'MatlabScriptPath'; Label = 'MATLAB 程序'; Path = $Settings.MatlabScriptPath; Type = 'File' },
        @{ Key = 'MatlabExePath'; Label = 'MATLAB 可执行文件'; Path = $Settings.MatlabExePath; Type = 'File' },
        @{ Key = 'LogDirectory'; Label = '日志目录'; Path = $Settings.LogDirectory; Type = 'Directory' },
        @{ Key = 'BackupDirectory'; Label = '备份目录'; Path = $Settings.BackupDirectory; Type = 'Directory' }
    )

    foreach ($item in $items) {
        if ($item.Key -like 'Matlab*' -and (-not $Settings.PSObject.Properties['SaEngine'] -or $Settings.SaEngine -eq 'BuiltIn')) { continue }
        $path = [string] $item.Path
        [pscustomobject]@{
            Key    = $item.Key
            Label  = $item.Label
            Path   = $path
            Type   = $item.Type
            Exists = ((-not [string]::IsNullOrWhiteSpace($path)) -and (Test-Path -LiteralPath $path))
        }
    }
}

function Get-PolishingSentinelMonitorProcesses {
    param(
        [Parameter(Mandatory = $true)] [string] $ScriptRoot,
        [object[]] $ProcessList = $null
    )

    $startScript = Join-Path $ScriptRoot 'Start-ExperimentDataAutomation.ps1'
    $resolvedStartScript = Resolve-Path -LiteralPath $startScript -ErrorAction SilentlyContinue
    if ($resolvedStartScript) {
        $startScriptPath = [string] $resolvedStartScript.ProviderPath
    }
    else {
        $startScriptPath = [System.IO.Path]::GetFullPath($startScript)
    }

    $normalizedStartScript = $startScriptPath.Replace('/', '\')
    $startScriptPattern = '(^|[\s"''])' + [regex]::Escape($normalizedStartScript) + '($|[\s"''])'
    $processes = $ProcessList
    if ($null -eq $processes) {
        $processes = @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe' OR Name = 'pwsh.exe'" -ErrorAction Stop)
    }

    @($processes |
        Where-Object {
            $commandLine = ([string] $_.CommandLine).Replace('/', '\')
            $_.ProcessId -ne $PID -and
            [regex]::IsMatch($commandLine, $startScriptPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        } |
        Select-Object ProcessId, CommandLine)
}

function Get-PolishingSentinelTaskStatus {
    param([Parameter(Mandatory = $true)] [string] $TaskName)

    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not $task) {
        return [pscustomobject]@{
            TaskName = $TaskName
            Exists   = $false
            State    = ''
        }
    }

    return [pscustomobject]@{
        TaskName = $TaskName
        Exists   = $true
        State    = [string] $task.State
    }
}
