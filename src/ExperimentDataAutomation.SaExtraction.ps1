Set-StrictMode -Version 2.0

function Get-ExperimentSaEngine {
    param($Config)
    if ($Config.PSObject.Properties['SaEngine']) { return [string]$Config.SaEngine }
    return 'BuiltIn'
}

function Get-ExperimentSaInputPlan {
    param([string] $FolderPath, [string] $OutputFileName)
    $manifest = Get-ExperimentSourceManifest $FolderPath $OutputFileName
    $slots = @{}; $maximum = 0
    foreach ($file in $manifest.Files) {
        if ($file.Name -notmatch '^(\d+)-([123])\.xlsx$') { throw "无法确定球编号和测量次数：$($file.Name)；应为 球编号-次数.xlsx，次数为 1、2、3。" }
        $sample = 0
        if (-not [int]::TryParse($Matches[1], [ref]$sample) -or $sample -lt 1 -or $sample -gt 10000) { throw "球编号超出支持范围 1 至 10000：$($file.Name)" }
        $measurement = [int]$Matches[2]
        $key = "$sample-$measurement"
        if (-not $slots.ContainsKey($key)) { $slots[$key] = @() }
        $slots[$key] += $file.Name
        $maximum = [math]::Max($maximum, $sample)
    }
    $conflicts = @($slots.Keys | Sort-Object | Where-Object { $slots[$_].Count -gt 1 } | ForEach-Object { "编号 $_ ：$($slots[$_] -join '、')" })
    [pscustomobject]@{ Manifest = $manifest; Slots = $slots; Maximum = $maximum; Conflicts = $conflicts }
}

function Get-ExperimentSaConfirmationPath {
    param($Config, [string] $FolderPath)
    $hash = [Security.Cryptography.SHA256]::Create()
    try { $key = [BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes([IO.Path]::GetFullPath($FolderPath).ToUpperInvariant()))).Replace('-', '') }
    finally { $hash.Dispose() }
    Join-Path (Join-Path $Config.StateDirectory 'sa-confirmations') ($key + '.json')
}

function Get-ExperimentSaConfirmations {
    param($Config)
    $directory = Join-Path $Config.StateDirectory 'sa-confirmations'
    if (-not (Test-Path -LiteralPath $directory)) { return }
    foreach ($file in Get-ChildItem -LiteralPath $directory -Filter '*.json' -File) {
        $request = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($request.State -eq 'AwaitingConfirmation') {
            $request | Add-Member NoteProperty RequestPath $file.FullName -Force
            $request
        }
    }
}

function Resolve-ExperimentSaConfirmation {
    param($Config, [string] $RequestPath, [ValidateSet('ContinueWithBlanks')] [string] $Decision,
        [Parameter(Mandatory = $true)] [string] $ExpectedFingerprint)
    $lease = Enter-ExperimentWorkbookLock -WorkbookPath $Config.MasterWorkbookPath -WaitMilliseconds 0
    if ($null -eq $lease) { throw '正在扫描或写入，请稍后重新确认。' }
    try {
        $directory = [IO.Path]::GetFullPath((Join-Path $Config.StateDirectory 'sa-confirmations')).TrimEnd('\')
        $path = [IO.Path]::GetFullPath($RequestPath)
        if ([IO.Path]::GetDirectoryName($path) -ne $directory) { throw '确认记录路径不属于当前软件。' }
        $request = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($request.State -ne 'AwaitingConfirmation' -or $request.Fingerprint -ne $ExpectedFingerprint) { throw '确认请求已经变化，请刷新列表后重新查看。' }
        $current = Get-ExperimentSourceManifest $request.FolderPath $Config.SaOutputFileName
        if ($current.Fingerprint -ne $request.Fingerprint) { throw '源文件已改变，请重新扫描后确认。' }
        $request.State = $Decision
        Save-ExperimentTransactionJournal -Transaction $request -Path $path
    } finally { Exit-ExperimentWorkbookLock $lease }
}

function Assert-ExperimentSaConfirmation {
    param($Config, [string] $FolderPath, $Plan)
    $path = Get-ExperimentSaConfirmationPath $Config $FolderPath
    $old = $null
    if (Test-Path -LiteralPath $path) { $old = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json }
    if ($Plan.Conflicts.Count -eq 0) {
        if ($old -and $old.State -ne 'Cleared') {
            $old.State = 'Cleared'
            Save-ExperimentTransactionJournal -Transaction $old -Path $path
        }
        return
    }
    if ($old -and $old.Fingerprint -eq $Plan.Manifest.Fingerprint -and $old.State -eq 'ContinueWithBlanks') { return }
    $notify = -not ($old -and $old.Fingerprint -eq $Plan.Manifest.Fingerprint -and $old.State -eq 'AwaitingConfirmation')
    if ($notify) {
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
        $request = [pscustomobject]@{ FolderPath = $FolderPath; Fingerprint = $Plan.Manifest.Fingerprint; Conflicts = @($Plan.Conflicts); State = 'AwaitingConfirmation'; CreatedAt = [datetime]::UtcNow.ToString('o') }
        Save-ExperimentTransactionJournal -Transaction $request -Path $path
    }
    $errorObject = [InvalidOperationException]::new("测量编号重复，实验已暂停。请在控制板的待确认与日志页处理编号冲突：$FolderPath")
    $errorObject.Data['SaNeedsConfirmation'] = $true
    $errorObject.Data['Notify'] = $notify
    throw $errorObject
}

function Invoke-ExperimentBuiltInSa {
    param($Config, [string] $TargetFolder)
    if ([IO.Path]::GetFileName($Config.SaOutputFileName) -ne $Config.SaOutputFileName -or $Config.SaOutputFileName -notlike '*.xlsx') { throw 'Sa 输出必须是当前目录中的 xlsx 文件名。' }
    $outputPath = Join-Path $TargetFolder $Config.SaOutputFileName
    if ($Config.DryRun) { return [pscustomobject]@{ OutputPath = $outputPath; WarningCount = 0; Warnings = @() } }
    $plan = Get-ExperimentSaInputPlan $TargetFolder $Config.SaOutputFileName
    Assert-ExperimentSaConfirmation $Config $TargetFolder $plan
    $session = $null; $book = $null
    $warnings = [Collections.Generic.List[string]]::new()
    $data = New-Object 'object[,]' $plan.Maximum, 5
    $stage = Join-Path $Config.TempDirectory ('sa-' + [guid]::NewGuid().ToString('N') + '.xlsx')
    $publish = Join-Path $TargetFolder ('.sa-' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [void][IO.Directory]::CreateDirectory($Config.TempDirectory)
        $session = New-ExperimentExcelSession
        $excel = $session.Application
        $excel.AutomationSecurity = 3
        for ($sample = 1; $sample -le $plan.Maximum; $sample++) {
            $notes = [Collections.Generic.List[string]]::new(); $valid = 0
            for ($measurement = 1; $measurement -le 3; $measurement++) {
                $key = "$sample-$measurement"
                if (-not $plan.Slots.ContainsKey($key)) { $notes.Add("$key 缺测"); continue }
                if ($plan.Slots[$key].Count -gt 1) { $notes.Add("$key 重复，已确认留空"); continue }
                $source = Join-Path $TargetFolder $plan.Slots[$key][0]
                $sheet = $null; $cell = $null
                try {
                    $book = $excel.Workbooks.Open($source, 0, $true)
                    $sheet = $book.Worksheets.Item(1)
                    $cell = $sheet.Range('D5')
                    $value = $cell.Value2
                    $number = $null
                    if (-not $excel.WorksheetFunction.IsError($cell) -and $value -isnot [bool]) {
                        $parsed = 0.0
                        if ($value -is [double] -or $value -is [int]) { $number = [double]$value }
                        elseif ($null -ne $value -and [double]::TryParse(([string]$value).Trim(), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) { $number = $parsed }
                    }
                    if ($null -ne $number -and -not [double]::IsNaN($number) -and -not [double]::IsInfinity($number)) {
                        $data[($sample - 1), ($measurement - 1)] = $number; $valid++
                    } else { $notes.Add("$key D5 无效，留空") }
                } finally {
                    Release-ComObject $cell; Release-ComObject $sheet
                    if ($book) { try { $book.Close($false) } finally { Release-ComObject $book; $book = $null } }
                }
            }
            $data[($sample - 1), 3] = "$valid/3"
            $data[($sample - 1), 4] = $notes -join '；'
            foreach ($note in $notes) { $warnings.Add($note) }
        }
        $book = $excel.Workbooks.Add()
        $sheet = $book.Worksheets.Item(1)
        $sheet.Name = 'Sa'
        $range = $sheet.Range("A1:E$($plan.Maximum)")
        try { $sheet.Columns.Item(4).NumberFormat = '@'; $range.Value2 = $data; $book.SaveAs($stage, 51) }
        finally { Release-ComObject $range; Release-ComObject $sheet }
        Close-ExperimentExcelSession $session $book
        $session = $null; $book = $null
        $verified = @(Read-SaRowsFromWorkbook $stage)
        if ($verified.Count -ne $plan.Maximum) { throw 'Sa 输出行数核验失败，旧汇总表未替换。' }
        foreach ($row in $verified) {
            for ($c = 0; $c -lt 3; $c++) {
                if (-not (Test-ExperimentSaNumberEqual $row.('Sa' + ($c + 1)) $data[($row.SampleIndex - 1), $c])) { throw 'Sa 输出数值核验失败，旧汇总表未替换。' }
            }
        }
        if ((Get-ExperimentSourceManifest $TargetFolder $Config.SaOutputFileName).Fingerprint -ne $plan.Manifest.Fingerprint) { throw '提取期间源文件改变，已取消发布 Sa 汇总。' }
        [IO.File]::Copy($stage, $publish, $false)
        if ([IO.File]::Exists($outputPath)) { [IO.File]::Replace($publish, $outputPath, [NullString]::Value) }
        else { [IO.File]::Move($publish, $outputPath) }
        foreach ($warning in $warnings) { Write-ExperimentLog -Config $Config -Level WARN -Message "$TargetFolder ：$warning" }
        Write-ExperimentLog -Config $Config -Message "内置 Sa 汇总完成：$outputPath；缺测或无效位置 $($warnings.Count) 个。"
        return [pscustomobject]@{ OutputPath = $outputPath; WarningCount = $warnings.Count; Warnings = @($warnings.ToArray()) }
    } finally {
        try { if ($session) { Close-ExperimentExcelSession $session $book } }
        finally { foreach ($path in @($stage, $publish)) { if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) } } }
    }
}

function Invoke-ExperimentSaExtraction {
    param($Config, [string] $TargetFolder)
    switch (Get-ExperimentSaEngine $Config) {
        'BuiltIn' { return Invoke-ExperimentBuiltInSa $Config $TargetFolder }
        'Matlab' { return [pscustomobject]@{ OutputPath = (Invoke-BaiguangMatlab $Config $TargetFolder); WarningCount = 0; Warnings = @() } }
        default { throw '未知 Sa 处理方式。' }
    }
}
