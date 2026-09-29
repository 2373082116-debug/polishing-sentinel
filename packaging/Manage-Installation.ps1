param([ValidateSet('Check','Install','Remove')][string]$Mode,[Parameter(Mandatory=$true)][string]$AppRoot)
$ErrorActionPreference='Stop'
try {
    $AppRoot=[IO.Path]::GetFullPath($AppRoot).TrimEnd('\')
    $data=Join-Path $env:LOCALAPPDATA 'PolishingSentinel'
    if ($Mode -eq 'Check') {
        $running=@(Get-CimInstance Win32_Process | Where-Object {
            $_.ProcessId -ne $PID -and $_.Name -in @('pwsh.exe','powershell.exe','PolishingSentinel.exe') -and
            $_.CommandLine -and $_.CommandLine.IndexOf($AppRoot+'\',[StringComparison]::OrdinalIgnoreCase) -ge 0
        })
        if ($running.Count) {exit 10}
        exit 0
    }
    foreach ($name in @('config','logs','state','temp','backups')) {
        $link=Join-Path $AppRoot $name; $target=Join-Path $data $name
        if ($Mode -eq 'Install') {
            [void][IO.Directory]::CreateDirectory($target)
            if (Test-Path -LiteralPath $link) {
                $item=Get-Item -LiteralPath $link -Force
                if ($item.LinkType -ne 'Junction' -or [IO.Path]::GetFullPath([string]$item.Target) -ne $target) {throw 'Unexpected data directory; installation stopped.'}
            } else {New-Item -ItemType Junction -Path $link -Target $target | Out-Null}
        } elseif (Test-Path -LiteralPath $link) {
            $item=Get-Item -LiteralPath $link -Force
            if ($item.LinkType -ne 'Junction' -or [IO.Path]::GetFullPath([string]$item.Target) -ne $target) {throw 'Unexpected data directory; removal stopped.'}
            [IO.Directory]::Delete($link)
        }
    }
    if ($Mode -eq 'Install') {
        . (Join-Path $AppRoot 'PolishingSentinel.Config.ps1')
        $path=Get-PolishingSentinelSettingsPath $data
        if (-not (Test-Path -LiteralPath $path)) {
            $settings=New-PolishingSentinelDefaultSettings $AppRoot
            $settings.MasterWorkbookPath=''; $settings.MatlabScriptPath=''; $settings.MatlabExePath=''
            $settings.EnableStartupTask=$false
            $settings.TaskName='PolishingSentinel.Installed'
            foreach ($pair in @(@('LogDirectory','logs'),@('StateDirectory','state'),@('TempDirectory','temp'),@('BackupDirectory','backups'))) {$settings.($pair[0])=Join-Path $data $pair[1]}
            Save-PolishingSentinelSettings $data $settings | Out-Null
        }
    } else {
        foreach ($task in @(Get-ScheduledTask -ErrorAction Stop)) {
            foreach ($action in $task.Actions) {
                if ($action.Arguments -and $action.Arguments.IndexOf((Join-Path $AppRoot 'Start-ExperimentDataAutomation.ps1'),[StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false
                    break
                }
            }
        }
    }
    exit 0
} catch {
    [IO.File]::WriteAllText((Join-Path $env:TEMP 'PolishingSentinel-InstallError.txt'),($_ | Out-String))
    exit 1
}
