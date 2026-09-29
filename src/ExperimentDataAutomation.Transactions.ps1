Set-StrictMode -Version 2.0

function Get-ExperimentFileSha256 {
    param([Parameter(Mandatory = $true)] [string] $LiteralPath)
    $stream = $null; $algorithm = $null
    try {
        # Keep integrity checks independent of PowerShell module auto-loading.
        $stream = [IO.File]::Open($LiteralPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $algorithm = [Security.Cryptography.SHA256]::Create()
        return [BitConverter]::ToString($algorithm.ComputeHash($stream)).Replace('-', '')
    }
    finally {
        if ($null -ne $algorithm) { $algorithm.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-ExperimentWorkbookSnapshot {
    param($Workbook)
    $sheets = @()
    foreach ($sheet in $Workbook.Worksheets) {
        $used = $sheet.UsedRange
        $values = $used.Formula
        $rows = [int] $used.Rows.Count; $columns = [int] $used.Columns.Count
        $cells = @{}
        for ($r = 1; $r -le $rows; $r++) {
            for ($c = 1; $c -le $columns; $c++) {
                $value = Get-ExcelArrayValue -Values $values -Row $r -Column $c -RowCount $rows -ColumnCount $columns
                if ($null -eq $value -or [string] $value -eq '') { continue }
                # Tag values so JSON readers cannot coerce date-like cell text into DateTime.
                $cells[('{0}:{1}' -f ($r + [int] $used.Row - 1), ($c + [int] $used.Column - 1))] = 'cell:' + [string] $value
            }
        }
        $sheets += [pscustomobject]@{ Name = $sheet.Name; Cells = $cells }
    }
    [pscustomobject]@{ Sheets = $sheets }
}

function Save-ExperimentWorkbookCheckpoint {
    param($Workbook)
    $snapshot = Get-ExperimentWorkbookSnapshot $Workbook
    $Workbook.Save()
    $active = Get-Variable -Name ExperimentActiveTransaction -Scope Script -ErrorAction SilentlyContinue
    if ($active -and $null -ne $active.Value -and $Workbook.FullName -eq $active.Value.Stage) {
        $active.Value.Expected = $snapshot
    }
}

function Assert-ExperimentWorkbookSnapshot {
    param([string] $WorkbookPath, $Expected)
    if ($null -eq $Expected) { throw '未取得写入核验快照，不能提交主表' }
    $session = $null; $excel = $null; $book = $null
    try {
        $session = New-ExperimentExcelSession
        $excel = $session.Application
        $book = $excel.Workbooks.Open($WorkbookPath, 0, $true)
        $actual = Get-ExperimentWorkbookSnapshot $book
        if ($actual.Sheets.Count -ne $Expected.Sheets.Count) { throw '保存后工作表数量不一致' }
        foreach ($sheet in $Expected.Sheets) {
            $match = @($actual.Sheets | Where-Object { $_.Name -eq $sheet.Name })
            if ($match.Count -ne 1) { throw "保存后缺失工作表：$($sheet.Name)" }
            $expectedCells = @{}
            if ($sheet.Cells -is [Collections.IDictionary]) { $expectedCells = $sheet.Cells }
            else { foreach ($property in $sheet.Cells.PSObject.Properties) { $expectedCells[$property.Name] = $property.Value } }
            if ($match[0].Cells.Count -ne $expectedCells.Count) { throw "保存后非空单元格数量不一致：$($sheet.Name)" }
            foreach ($address in $expectedCells.Keys) {
                if (-not $match[0].Cells.ContainsKey($address) -or [string] $match[0].Cells[$address] -cne [string] $expectedCells[$address]) {
                    throw "保存后数据或公式不一致：$($sheet.Name) $address"
                }
            }
        }
    }
    finally {
        Close-ExperimentExcelSession -Session $session -Workbook $book
    }
}

function Save-ExperimentTransactionJournal {
    param($Transaction, [string] $Path)
    $temp = $Path + '.tmp'
    [IO.File]::WriteAllText($temp, ($Transaction | ConvertTo-Json -Depth 14), [Text.UTF8Encoding]::new($true))
    if ([IO.File]::Exists($Path)) { [IO.File]::Replace($temp, $Path, [System.Management.Automation.Language.NullString]::Value) }
    else { [IO.File]::Move($temp, $Path) }
}

function Repair-ExperimentInterruptedTransactions {
    param($Config)
    $directory = Join-Path $Config.ScriptRoot 'state/transactions'
    if (-not (Test-Path -LiteralPath $directory)) { return }
    foreach ($file in @(Get-ChildItem -LiteralPath $directory -Filter '*.json')) {
        $tx = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($tx.Original -ne [IO.Path]::GetFullPath($Config.MasterWorkbookPath) -or $tx.State -notin @('Prepared', 'Staging', 'NeedsReview')) { continue }
        $currentHash = Get-ExperimentFileSha256 -LiteralPath $tx.Original
        if ($tx.State -eq 'NeedsReview') { throw "存在需要确认的写入事务：$($file.FullName)；备份：$($tx.BackupPath)" }
        if ($currentHash -eq $tx.BeforeHash) { $tx.State = 'AbortedRecovered' }
        elseif ($tx.State -eq 'Prepared' -and $currentHash -eq $tx.StagedHash) {
            try { Assert-ExperimentWorkbookSnapshot -WorkbookPath $tx.Original -Expected $tx.Expected }
            catch {
                $tx.State = 'NeedsReview'
                Save-ExperimentTransactionJournal $tx $file.FullName
                throw "中断后的主表核验未通过，需要确认；备份：$($tx.BackupPath)。$($_.Exception.Message)"
            }
            $tx.State = 'CommittedRecovered'
        }
        else {
            $tx.State = 'NeedsReview'
            Save-ExperimentTransactionJournal $tx $file.FullName
            throw "中断后主表已有其他变化，需要确认；备份：$($tx.BackupPath)"
        }
        Save-ExperimentTransactionJournal $tx $file.FullName
    }
}

function Invoke-ExperimentWorkbookTransaction {
    param(
        [Parameter(Mandatory = $true)] $Config,
        [Parameter(Mandatory = $true)] [scriptblock] $Action,
        [scriptblock] $Verify = { param($candidate, $expected) Assert-ExperimentWorkbookSnapshot -WorkbookPath $candidate -Expected $expected }
    )
    $original = [IO.Path]::GetFullPath($Config.MasterWorkbookPath)
    $lease = Enter-ExperimentWorkbookLock -WorkbookPath $original
    $previous = Get-Variable -Name ExperimentActiveTransaction -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    $tx = $null; $replaced = $false
    try {
        Repair-ExperimentInterruptedTransactions $Config
        $probe = [IO.File]::Open($original, 'Open', 'ReadWrite', 'None')
        $probe.Dispose()
        $id = [guid]::NewGuid().ToString('N')
        $stage = Join-Path ([IO.Path]::GetDirectoryName($original)) ('.sentinel-' + $id + '.xlsx')
        $backupDirectory = $Config.BackupDirectory
        [void] [IO.Directory]::CreateDirectory($backupDirectory)
        $backup = Join-Path $backupDirectory (([IO.Path]::GetFileNameWithoutExtension($original)) + '.' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.' + $id + '.xlsx')
        $journalDirectory = Join-Path $Config.ScriptRoot 'state/transactions'
        [void] [IO.Directory]::CreateDirectory($journalDirectory)
        $journal = Join-Path $journalDirectory ($id + '.json')
        $before = Get-ExperimentFileSha256 -LiteralPath $original
        [IO.File]::Copy($original, $backup, $false)
        [IO.File]::Copy($backup, $stage, $false)
        if ((Get-ExperimentFileSha256 -LiteralPath $backup) -ne $before) { throw '备份期间主表发生变化，已取消写入' }
        $tx = [pscustomobject]@{ State = 'Staging'; CreatedAt = (Get-Date).ToString('o'); Original = $original; Stage = $stage; BackupPath = $backup; BeforeHash = $before; StagedHash = ''; Expected = $null }
        Save-ExperimentTransactionJournal $tx $journal
        $script:ExperimentActiveTransaction = $tx
        $working = $Config.PSObject.Copy()
        $working.MasterWorkbookPath = $stage
        $result = & $Action $working
        & $Verify $stage $tx.Expected
        $tx.StagedHash = Get-ExperimentFileSha256 -LiteralPath $stage
        $tx.State = 'Prepared'
        Save-ExperimentTransactionJournal $tx $journal
        if ((Get-ExperimentFileSha256 -LiteralPath $original) -ne $before) { throw '主表在处理期间发生变化，已取消替换' }
        [IO.File]::Replace($stage, $original, [System.Management.Automation.Language.NullString]::Value)
        $replaced = $true
        & $Verify $original $tx.Expected
        $tx.State = 'Committed'
        Save-ExperimentTransactionJournal $tx $journal
        return [pscustomobject]@{ Result = $result; BackupPath = $backup; JournalPath = $journal }
    }
    catch {
        $failure = $_
        if ($tx) {
            if ($replaced) {
                if ((Get-ExperimentFileSha256 -LiteralPath $original) -eq $tx.StagedHash) {
                    $restore = $tx.Stage
                    [IO.File]::Copy($tx.BackupPath, $restore, $true)
                    [IO.File]::Replace($restore, $original, [System.Management.Automation.Language.NullString]::Value)
                    $tx.State = 'RolledBack'
                }
                else { $tx.State = 'NeedsReview' }
            }
            else { $tx.State = 'Aborted' }
            Save-ExperimentTransactionJournal $tx $journal
        }
        throw $failure
    }
    finally {
        $script:ExperimentActiveTransaction = $previous
        Exit-ExperimentWorkbookLock $lease
    }
}

function Test-ExperimentTransactionActive {
    param([string] $WorkbookPath)
    $active = Get-Variable -Name ExperimentActiveTransaction -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    return ($null -ne $active -and $active.Stage -eq [IO.Path]::GetFullPath($WorkbookPath))
}

function Invoke-ExperimentImportCommit {
    param($Config, [AllowEmptyCollection()] [object[]] $Records = @(), [AllowEmptyCollection()] [object[]] $SourceUpdates = @())
    $transaction = Invoke-ExperimentWorkbookTransaction -Config $Config -Action {
        param($working)
        $imported = 0; $backfill = $null
        if ($Records.Count) { $imported = Append-AutoImportRecords -WorkbookPath $working.MasterWorkbookPath -SheetName $working.AutoSheetName -Records $Records }
        if ($working.EnableBackfill) {
            $backfillConfig = New-ExperimentBackfillConfig -ScriptRoot $working.ScriptRoot -MasterWorkbookPath $working.MasterWorkbookPath -AutoSheetName $working.AutoSheetName -PendingSheetName $working.PendingSheetName -BackupDirectory $working.BackupDirectory -SlurryMappings $working.SlurryMappings -ProcessDefaults $working.ProcessDefaults -NewBlockMode $working.NewBlockMode
            $backfill = Invoke-ExperimentBackfillPhase1 -Config $backfillConfig
        }
        Write-ExperimentSourceVersions -WorkbookPath $working.MasterWorkbookPath -Updates $SourceUpdates
        foreach ($update in $SourceUpdates) {
            $latest = Get-ExperimentSourceManifest -FolderPath $update.FolderPath -OutputFileName $working.SaOutputFileName
            if ($latest.Fingerprint -ne $update.Manifest.Fingerprint) { throw "处理期间源文件发生变化，已取消本轮写入：$($update.FolderPath)" }
        }
        [pscustomobject]@{ Imported = $imported; Backfill = $backfill }
    }
    if ($transaction.Result.Backfill) { $transaction.Result.Backfill.BackupPath = $transaction.BackupPath }
    return $transaction.Result
}
