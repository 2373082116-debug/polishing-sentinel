param(
    [switch] $RunOnce,
    [switch] $DryRun,
    [int] $PollSeconds = 0,
    [int] $RescanSeconds = 0,
    [string] $SettingsPath = ''
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptRoot 'ExperimentDataAutomation.Core.ps1')
. (Join-Path $scriptRoot 'ExperimentDataAutomation.Backfill.ps1')
. (Join-Path $scriptRoot 'PolishingSentinel.Config.ps1')
. (Join-Path $scriptRoot 'PolishingSentinel.Status.ps1')
$monitorLease = $null

try {
    if (-not [string]::IsNullOrWhiteSpace($SettingsPath) -and (Test-Path -LiteralPath $SettingsPath)) {
        $settingsText = Get-Content -LiteralPath $SettingsPath -Raw -Encoding UTF8
        $loadedSettings = $settingsText | ConvertFrom-Json
        $settings = Merge-PolishingSentinelSettings -Defaults (New-PolishingSentinelDefaultSettings -ScriptRoot $scriptRoot) -Loaded $loadedSettings
    }
    else {
        $settings = Read-PolishingSentinelSettings -ScriptRoot $scriptRoot
    }

    $pollOverride = $null
    if ($PollSeconds -gt 0) {
        $pollOverride = $PollSeconds
    }

    $config = ConvertTo-ExperimentAutomationConfigFromSettings -ScriptRoot $scriptRoot -Settings $settings -DryRun:$DryRun -PollSecondsOverride $pollOverride
    if ($RescanSeconds -le 0) {
        $RescanSeconds = [int] $settings.RescanSeconds
    }
    Initialize-ExperimentAutomationEnvironment -Config $config

    if (-not $RunOnce) {
        $monitorLease = Enter-ExperimentWorkbookLock -WorkbookPath $config.MasterWorkbookPath -Purpose Monitor -WaitMilliseconds 0
        if ($null -eq $monitorLease) {
            Write-ExperimentLog -Config $config -Message '后台监控已在运行，本次重复启动已忽略。'
            exit 0
        }
    }
    Write-ExperimentLog -Config $config -Message '实验数据自动整理程序已启动。'
}
catch {
    if ($null -ne $monitorLease) { Exit-ExperimentWorkbookLock $monitorLease; $monitorLease = $null }
    $message = "自动整理启动出错：$($_.Exception.Message)"
    $fallbackLogDirectory = Join-Path $scriptRoot 'logs'
    $fallbackStateDirectory = Join-Path $scriptRoot 'state'

    try {
        foreach ($path in @($fallbackLogDirectory, $fallbackStateDirectory)) {
            if (-not (Test-Path -LiteralPath $path)) {
                New-Item -ItemType Directory -Path $path -Force | Out-Null
            }
        }
    }
    catch {
    }

    try {
        Update-PolishingSentinelStatus -ScriptRoot $scriptRoot -State 'Error' -Message $message -Data @{ ProcessId = $PID; Stage = 'Startup' } | Out-Null
    }
    catch {
    }

    try {
        $fallbackLogPath = Join-Path $fallbackLogDirectory 'automation.log'
        $line = '{0:yyyy-MM-dd HH:mm:ss} [ERROR] {1}' -f (Get-Date), $message
        Add-Content -LiteralPath $fallbackLogPath -Value $line -Encoding UTF8
    }
    catch {
    }

    throw
}

if ($RunOnce) {
    try {
        $config.RetryFailed = $true
        Invoke-ExperimentDataAutomationOnce -Config $config | Format-List
        exit 0
    }
    catch {
        $message = "自动整理出错：$($_.Exception.Message)"
        Write-ExperimentLog -Config $config -Message $message -Level 'ERROR'
        throw
    }
}

$lastDrivePresent = $false
$lastScanAt = Get-Date '2000-01-01'

try {
while ($true) {
    try {
        $volume = Get-UsbDriveByLabel -VolumeLabel $config.VolumeLabel -ExpectedDriveLetter $config.ExpectedDriveLetter
        if ($volume) {
            $now = Get-Date
            if (-not $lastDrivePresent) {
                $lastDrivePresent = $true
                $lastScanAt = Get-Date '2000-01-01'
                Write-ExperimentLog -Config $config -Message "检测到 U 盘插入：$($volume.DriveLetter):"
            }

            if (($now - $lastScanAt).TotalSeconds -ge $RescanSeconds) {
                $lastScanAt = $now
                Invoke-ExperimentDataAutomationOnce -Config $config -SilentWhenNoDrive | Out-Null
            }
        }
        else {
            if ($lastDrivePresent) {
                Write-ExperimentLog -Config $config -Message 'U 盘已移除，等待下次插入。'
            }
            $lastDrivePresent = $false
            $idleLease = Enter-ExperimentWorkbookLock -WorkbookPath $config.MasterWorkbookPath -WaitMilliseconds 0
            if ($null -ne $idleLease) {
                try { Update-PolishingSentinelStatus -ScriptRoot $scriptRoot -State 'WaitingForUsb' -Message '等待 U 盘插入' -Data @{ ExpectedDriveLetter = $config.ExpectedDriveLetter; VolumeLabel = $config.VolumeLabel } | Out-Null }
                finally { Exit-ExperimentWorkbookLock $idleLease }
            }
        }
    }
    catch {
        $message = "自动整理出错：$($_.Exception.Message)"
        Write-ExperimentLog -Config $config -Message $message -Level 'ERROR'
        Show-ExperimentNotification -Message $message
    }

    Start-Sleep -Seconds $config.PollSeconds
}
}
finally { if ($null -ne $monitorLease) { Exit-ExperimentWorkbookLock $monitorLease } }
