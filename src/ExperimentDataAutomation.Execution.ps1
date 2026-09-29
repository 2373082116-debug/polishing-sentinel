Set-StrictMode -Version 2.0

function Invoke-ExperimentTimedProcess {
    param(
        [Parameter(Mandatory = $true)] [string] $FilePath,
        [Parameter(Mandatory = $true)] [string] $Arguments,
        [ValidateRange(1, 86400)] [int] $TimeoutSeconds = 1200,
        [Parameter(Mandatory = $true)] [string] $LogPrefix
    )
    $stdout = $LogPrefix + '.stdout.log'
    $stderr = $LogPrefix + '.stderr.log'
    $process = $null
    try {
        $process = Start-Process -FilePath $FilePath -ArgumentList $Arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr -ErrorAction Stop
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $timeoutError = [TimeoutException]::new("MATLAB 超过 $TimeoutSeconds 秒未完成。日志：$LogPrefix")
            $timeoutError.Data['ProcessId'] = $process.Id
            throw $timeoutError
        }
        $process.WaitForExit()
        return [pscustomobject]@{ ExitCode = $process.ExitCode; StdoutPath = $stdout; StderrPath = $stderr }
    }
    finally {
        if ($null -ne $process) {
            try {
                if (-not $process.HasExited) {
                    # Target only this invocation and its descendants, never other MATLAB sessions.
                    $killer = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32/taskkill.exe') -ArgumentList "/PID $($process.Id) /T /F" -WindowStyle Hidden -PassThru
                    try {
                        if (-not $killer.WaitForExit(10000)) { throw '终止 MATLAB 进程树超时，需要人工检查。' }
                        if (-not $process.WaitForExit(5000)) { throw '未能终止本次 MATLAB，需要人工检查。' }
                    }
                    finally { $killer.Dispose() }
                }
            }
            finally { $process.Dispose() }
        }
    }
}

function Get-ExperimentRetryPath {
    param($Config, $SourceVersion)
    $identity = [IO.Path]::GetFullPath($Config.MasterWorkbookPath).ToUpperInvariant() + '|' + $SourceVersion.Identity + '|' + $SourceVersion.Manifest.Fingerprint
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $key = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($identity))).Replace('-', '') }
    finally { $sha.Dispose() }
    return (Join-Path (Join-Path $Config.StateDirectory 'retries') ($key + '.json'))
}

function Get-ExperimentRetryState {
    param($Config, $SourceVersion)
    $path = Get-ExperimentRetryPath $Config $SourceVersion
    if ([IO.File]::Exists($path)) {
        $state = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($state.Identity -ne $SourceVersion.Identity -or $state.Fingerprint -ne $SourceVersion.Manifest.Fingerprint -or $state.Attempts -lt 0) { throw "重试记录损坏：$path" }
        return $state
    }
    return [pscustomobject]@{
        Identity = $SourceVersion.Identity; Fingerprint = $SourceVersion.Manifest.Fingerprint
        Attempts = 0; State = 'New'; NextAttemptUtcTicks = 0; LastError = ''
    }
}

function Save-ExperimentRetryState {
    param($Config, $SourceVersion, $State)
    $path = Get-ExperimentRetryPath $Config $SourceVersion
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
    Save-ExperimentTransactionJournal -Transaction $State -Path $path
}

function Get-ExperimentWorkProgressData {
    param($Result,[int]$Total,[string]$Folder,[int]$Attempt=0)
    @{
        TotalFolders=$Total
        FinishedFolders=($Result.CompletedFolders+$Result.FailedFolders+$Result.DeferredFolders)
        CompletedFolders=$Result.CompletedFolders; FailedFolders=$Result.FailedFolders; DeferredFolders=$Result.DeferredFolders
        Folder=$Folder; Attempt=$Attempt; ProcessId=$PID
    }
}

function Invoke-ExperimentWorkItems {
    param($Config, [object[]] $Items = @())
    $result = [pscustomobject]@{
        Imported = 0; CompletedFolders = 0; FailedFolders = 0; DeferredFolders = 0
        Backfill = $null; Errors = @(); SharedFailure = $false; StateWarnings = 0; TimedOutFolders = 0
        ConfirmationFolders = 0; NewConfirmations = 0; SaWarnings = 0
    }
    foreach ($item in $Items) {
        if ($Config.DryRun) {
            Write-ExperimentLog -Config $Config -Message "DRY-RUN: 待处理 $($item.Folder.FullName)"
            continue
        }
        if ($result.SharedFailure) { $result.DeferredFolders++; continue }
        $state = $null
        $stage = 'RetryState'
        $attemptStarted = $false
        try {
        if (Get-Command Update-PolishingSentinelStatus -ErrorAction SilentlyContinue) {
            Update-PolishingSentinelStatus -ScriptRoot $Config.ScriptRoot -State 'ValidatingSa' -Message ("正在检查实验数据：{0}" -f $item.Metadata.FolderName) -Data (Get-ExperimentWorkProgressData $result $Items.Count $item.Folder.FullName) | Out-Null
        }
        if ((Get-ExperimentSaEngine $Config) -eq 'BuiltIn') {
            $stage = 'ValidatingSa'
            $inputPlan = Get-ExperimentSaInputPlan $item.Folder.FullName $Config.SaOutputFileName
            Assert-ExperimentSaConfirmation $Config $item.Folder.FullName $inputPlan
        }
        $stage = 'RetryState'
        $state = Get-ExperimentRetryState -Config $Config -SourceVersion $item.SourceVersion
        if ($Config.RetryFailed) { $state.Attempts = 0; $state.NextAttemptUtcTicks = 0 }
        if ($state.Attempts -ge $Config.MaxExperimentAttempts -or [long]$state.NextAttemptUtcTicks -gt [DateTime]::UtcNow.Ticks) {
            $result.DeferredFolders++
            $reason = if ($state.Attempts -ge $Config.MaxExperimentAttempts) { '已达到自动重试上限，请检查日志后点击立即扫描' } else { '等待重试间隔' }
            Write-ExperimentLog -Config $Config -Message "$reason：$($item.Metadata.FolderName)；$($state.LastError)"
            continue
        }
        $state.Attempts++; $state.State = 'Running'
        $state.NextAttemptUtcTicks = [DateTime]::UtcNow.AddSeconds($Config.RetryDelaySeconds).Ticks
        Save-ExperimentRetryState -Config $Config -SourceVersion $item.SourceVersion -State $state
        $attemptStarted = $true
        $stage = if ((Get-ExperimentSaEngine $Config) -eq 'Matlab') { 'RunningMatlab' } else { 'ExtractingSa' }
            if (Get-Command Update-PolishingSentinelStatus -ErrorAction SilentlyContinue) {
                Update-PolishingSentinelStatus -ScriptRoot $Config.ScriptRoot -State $stage -Message ("正在整理 Sa（第 {0} 次尝试）：{1}" -f $state.Attempts, $item.Metadata.FolderName) -Data (Get-ExperimentWorkProgressData $result $Items.Count $item.Folder.FullName $state.Attempts) | Out-Null
            }
            $extraction = Invoke-ExperimentSaExtraction -Config $Config -TargetFolder $item.Folder.FullName
            $output = $extraction.OutputPath
            $result.SaWarnings += $extraction.WarningCount
            $records = @(New-AutoImportRecordsFromFolder -Metadata $item.Metadata -SaOutputPath $output)
            if ($records.Count -eq 0) { throw 'Sa 汇总未返回样品记录，已取消写入。' }
            foreach ($record in $records) { $record | Add-Member -NotePropertyName SourceVersion -NotePropertyValue $item.SourceVersion -Force }
            $stage = 'WritingExcel'
            if (Get-Command Update-PolishingSentinelStatus -ErrorAction SilentlyContinue) {
                Update-PolishingSentinelStatus -ScriptRoot $Config.ScriptRoot -State $stage -Message ("正在提交实验：{0}" -f $item.Metadata.FolderName) -Data (Get-ExperimentWorkProgressData $result $Items.Count $item.Folder.FullName $state.Attempts) | Out-Null
            }
            $commit = Invoke-ExperimentImportCommit -Config $Config -Records $records -SourceUpdates @($item.SourceVersion)
            $result.Imported += $commit.Imported
            $result.CompletedFolders++
            if ($null -ne $commit.Backfill) {
                if ($null -eq $result.Backfill) { $result.Backfill = $commit.Backfill }
                else {
                    foreach ($name in @('BackfilledBlocks', 'PendingRows', 'PreviewNewBlocks', 'CreatedNewBlocks', 'NewBlockFailures')) {
                        if ($result.Backfill.PSObject.Properties[$name] -and $commit.Backfill.PSObject.Properties[$name]) { $result.Backfill.$name += $commit.Backfill.$name }
                    }
                    $result.Backfill.BackupPath = $commit.Backfill.BackupPath
                }
            }
            $state.State = 'Completed'; $state.LastError = ''; $state.Attempts = 0; $state.NextAttemptUtcTicks = 0
            Write-ExperimentLog -Config $Config -Message "实验已独立提交 $($records.Count) 行 Sa：$($item.Metadata.FolderName)"
        }
        catch {
            if ($_.Exception.Data['SaNeedsConfirmation']) {
                $result.DeferredFolders++; $result.ConfirmationFolders++
                if ($_.Exception.Data['Notify']) { $result.NewConfirmations++ }
                if ($attemptStarted) {
                    $state.Attempts--; $state.State = 'AwaitingConfirmation'; $state.NextAttemptUtcTicks = 0
                    Save-ExperimentRetryState -Config $Config -SourceVersion $item.SourceVersion -State $state
                }
                Write-ExperimentLog -Config $Config -Level WARN -Message $_.Exception.Message
                continue
            }
            if (-not $attemptStarted) {
                $result.FailedFolders++
                $result.Errors += [pscustomobject]@{ Folder = $item.Folder.FullName; Stage = $stage; Message = $_.Exception.Message; Attempt = 0 }
                Write-ExperimentLog -Config $Config -Level ERROR -Message "处理准备失败（$stage），已隔离实验：$($item.Metadata.FolderName)；$($_.Exception.Message)"
                continue
            }
            $state.State = 'Failed'; $state.LastError = $_.Exception.Message
            if ($_.Exception -is [TimeoutException]) { $result.TimedOutFolders++ }
            $state.NextAttemptUtcTicks = [DateTime]::UtcNow.AddSeconds($Config.RetryDelaySeconds).Ticks
            $result.FailedFolders++
            $result.Errors += [pscustomobject]@{ Folder = $item.Folder.FullName; Stage = $stage; Message = $state.LastError; Attempt = $state.Attempts }
            Write-ExperimentLog -Config $Config -Level ERROR -Message "实验处理失败（$stage，第 $($state.Attempts) 次）：$($item.Metadata.FolderName)；$($state.LastError)"
            # A shared workbook failure must not launch more expensive computations.
            if ($stage -eq 'WritingExcel') { $result.SharedFailure = $true }
        }
        try { Save-ExperimentRetryState -Config $Config -SourceVersion $item.SourceVersion -State $state }
        catch {
            # The workbook commit remains authoritative even if its retry-state update fails.
            $result.StateWarnings++
            $result.Errors += [pscustomobject]@{ Folder = $item.Folder.FullName; Stage = 'RetryState'; Message = $_.Exception.Message; Attempt = $state.Attempts }
            Write-ExperimentLog -Config $Config -Level ERROR -Message "未能更新重试状态，请检查状态目录；主表已提交的数据不受影响：$($item.Metadata.FolderName)；$($_.Exception.Message)"
        }
    }
    return $result
}
