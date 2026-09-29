Set-StrictMode -Version 2.0

function Get-ExperimentSourceIdentity {
    param([string] $FolderPath, [object] $SlurryMappings = $null)
    $meta = Parse-StructuredExperimentFolderName -FolderPath $FolderPath -SlurryMappings $SlurryMappings
    if ($meta.IsStructured -and $meta.ExperimentNo -match '^\d+$' -and $meta.TargetSheetName) {
        return ('{0}|{1}|{2}' -f $meta.SlurryType, $meta.ExperimentNo, $meta.ExperimentDate).ToLowerInvariant()
    }
    return $FolderPath.Substring([IO.Path]::GetPathRoot($FolderPath).Length).Replace('/', '\').ToLowerInvariant()
}

function Get-ExperimentSourceManifest {
    param([string] $FolderPath, [string] $OutputFileName = '所有的Sa数据.xlsx')
    $files = @()
    foreach ($file in @(Get-ChildItem -LiteralPath $FolderPath -File -Filter '*.xlsx' -ErrorAction Stop | Where-Object { $_.Name -notlike '~$*' -and $_.Name -ne $OutputFileName } | Sort-Object Name)) {
        $stream = [IO.File]::Open($file.FullName, 'Open', 'Read', 'Read')
        $hash = [Security.Cryptography.SHA256]::Create()
        try { $digest = [BitConverter]::ToString($hash.ComputeHash($stream)).Replace('-', ''); $length = $stream.Length }
        finally { $hash.Dispose(); $stream.Dispose() }
        $files += [pscustomobject]@{ Name = $file.Name; Length = $length; Hash = $digest }
    }
    if ($files.Count -eq 0) { throw "原始文件已全部移除，需要确认；主表不会被清空：$FolderPath" }
    $json = ConvertTo-Json -InputObject @($files) -Compress
    $hash = [Security.Cryptography.SHA256]::Create()
    try { $digest = [BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($json))).Replace('-', '') }
    finally { $hash.Dispose() }
    [pscustomobject]@{ Fingerprint = $digest; Files = $files; FilesJson = $json }
}

function Get-ExperimentTrackedFolders {
    param([object[]] $Folders, $Versions, [string[]] $SourceRoots, [string] $DriveLetter, $SlurryMappings)
    $byPath = @{}; $identities = @{}
    foreach ($folder in $Folders) {
        $byPath[$folder.FullName] = $folder
        $identities[(Get-ExperimentSourceIdentity $folder.FullName $SlurryMappings)] = $true
    }
    foreach ($identity in $Versions.Keys) {
        if ($identities.ContainsKey($identity)) { continue }
        $old = $Versions[$identity].FolderPath
        $path = $DriveLetter + ':\' + $old.Substring([IO.Path]::GetPathRoot($old).Length)
        foreach ($root in $SourceRoots) {
            if ($path.StartsWith($root.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $root)) {
                $byPath[$path] = [IO.DirectoryInfo]::new($path)
                break
            }
        }
    }
    return @($byPath.Values | Sort-Object FullName)
}

function Get-ExperimentSourceChangeKind {
    param($Previous, $Current)
    if ($null -eq $Previous) { return 'New' }
    if ($Previous.Fingerprint -eq $Current.Fingerprint) { return 'Unchanged' }
    $oldFiles = @($Previous.FilesJson | ConvertFrom-Json)
    foreach ($old in $oldFiles) {
        $matches = @($Current.Files | Where-Object { $_.Name -eq $old.Name -and $_.Hash -eq $old.Hash })
        if ($matches.Count -ne 1) { return 'Modified' }
    }
    if ($Current.Files.Count -gt $oldFiles.Count) { return 'Added' }
    return 'Modified'
}

function Get-ExperimentVersionColumns {
    param($Worksheet, [switch] $Create)
    $columns = @{}
    $last = [int] $Worksheet.UsedRange.Columns.Count
    foreach ($name in @('来源身份', '导入版本', '源文件指纹', '版本状态', '源文件变更')) {
        for ($c = 1; $c -le $last; $c++) {
            if ([string] $Worksheet.Cells.Item(1, $c).Value2 -eq $name) { $columns[$name] = $c; break }
        }
        if (-not $columns.ContainsKey($name) -and $Create) { $last++; $Worksheet.Cells.Item(1, $last).Value2 = $name; $columns[$name] = $last }
    }
    return $columns
}

function Read-ExperimentSourceVersionsFromWorkbook {
    param($Workbook)
    $versions = @{}
    $sheet = @($Workbook.Worksheets | Where-Object { $_.Name -eq '源文件版本' } | Select-Object -First 1)
    if ($sheet.Count -eq 0) { return $versions }
    for ($r = 2; $r -le [int] $sheet[0].UsedRange.Rows.Count; $r++) {
        $key = [string] $sheet[0].Cells.Item($r, 1).Value2
        if (-not $key) { continue }
        $versions[$key] = [pscustomobject]@{
            Revision = [int] $sheet[0].Cells.Item($r, 2).Value2
            FolderPath = [string] $sheet[0].Cells.Item($r, 3).Value2
            Fingerprint = [string] $sheet[0].Cells.Item($r, 4).Value2
            FilesJson = [string] $sheet[0].Cells.Item($r, 5).Value2
        }
    }
    return $versions
}

function Write-ExperimentSourceVersions {
    param([string] $WorkbookPath, [object[]] $Updates)
    if ($Updates.Count -eq 0) { return }
    $session = $null; $excel = $null; $book = $null
    try {
        $session = New-ExperimentExcelSession
        $excel = $session.Application
        $book = $excel.Workbooks.Open($WorkbookPath)
        $sheets = @($book.Worksheets | Where-Object { $_.Name -eq '源文件版本' })
        if ($sheets.Count) { $sheet = $sheets[0] }
        else {
            $sheet = $book.Worksheets.Add(); $sheet.Name = '源文件版本'
            $headers = @('来源身份', '导入版本', '文件夹', '源文件指纹', '原始文件清单', '记录时间', '处理方式')
            for ($c = 0; $c -lt $headers.Count; $c++) { $sheet.Cells.Item(1, $c + 1).Value2 = $headers[$c] }
        }
        $r = [int] $sheet.UsedRange.Rows.Count + 1
        foreach ($item in $Updates) {
            if ($item.Manifest.FilesJson.Length -gt 32000) { throw '原始文件清单过长，无法完整保存版本记录' }
            $values = @($item.Identity, [string] $item.Revision, $item.FolderPath, $item.Manifest.Fingerprint, $item.Manifest.FilesJson, (Get-Date).ToString('o'), $item.ChangeKind)
            for ($c = 0; $c -lt $values.Count; $c++) { $sheet.Cells.Item($r, $c + 1).Value2 = [string] $values[$c] }
            $r++
        }
        Save-ExperimentWorkbookCheckpoint -Workbook $book
    }
    finally {
        Close-ExperimentExcelSession -Session $session -Workbook $book
    }
}

function Set-ExperimentImportRevision {
    param($Worksheet, $Record, [int] $RowNumber)
    $revision = Get-ExperimentObjectProperty $Record 'SourceVersion' $null
    if ($null -eq $revision) { return }
    $cols = Get-ExperimentVersionColumns $Worksheet -Create
    # Archive older rows without deleting the evidence used for conflict review.
    for ($r = 2; $r -lt $RowNumber; $r++) {
        $path = [string] $Worksheet.Cells.Item($r, 3).Value2
        if (-not $path) { continue }
        $identity = [string] $Worksheet.Cells.Item($r, $cols['来源身份']).Value2
        if (-not $identity) { $identity = Get-ExperimentSourceIdentity -FolderPath $path -SlurryMappings $revision.SlurryMappings }
        if ($identity -ne $revision.Identity) { continue }
        $oldVersion = [string] $Worksheet.Cells.Item($r, $cols['导入版本']).Value2
        if ($oldVersion -and [int] $oldVersion -ge $revision.Revision) { continue }
        $Worksheet.Cells.Item($r, $cols['来源身份']).Value2 = $identity
        if (-not $oldVersion) { $Worksheet.Cells.Item($r, $cols['导入版本']).Value2 = '0' }
        $Worksheet.Cells.Item($r, $cols['版本状态']).Value2 = 'Archived'
    }
    $Worksheet.Cells.Item($RowNumber, $cols['来源身份']).Value2 = $revision.Identity
    $Worksheet.Cells.Item($RowNumber, $cols['导入版本']).Value2 = [string] $revision.Revision
    $Worksheet.Cells.Item($RowNumber, $cols['源文件指纹']).Value2 = $revision.Manifest.Fingerprint
    $Worksheet.Cells.Item($RowNumber, $cols['版本状态']).Value2 = 'Current'
    $Worksheet.Cells.Item($RowNumber, $cols['源文件变更']).Value2 = $revision.ChangeKind
}

function Test-ExperimentSafeSaExtension {
    param($Worksheet, $Group, $Snapshot)
    $cols = Get-ExperimentVersionColumns $Worksheet
    if (-not $cols.ContainsKey('源文件变更')) { return $false }
    $first = $Group.AuditRows[0]
    if ([string] $Worksheet.Cells.Item($first, $cols['源文件变更']).Value2 -ne 'Added') { return $false }
    $identity = [string] $Worksheet.Cells.Item($first, $cols['来源身份']).Value2
    $version = [int] $Worksheet.Cells.Item($first, $cols['导入版本']).Value2
    $previous = @()
    for ($r = 2; $r -le [int] $Worksheet.UsedRange.Rows.Count; $r++) {
        if ([string] $Worksheet.Cells.Item($r, $cols['来源身份']).Value2 -ne $identity -or [int] $Worksheet.Cells.Item($r, $cols['导入版本']).Value2 -ne ($version - 1)) { continue }
        $previous += [pscustomobject]@{
            SampleIndex = [int] $Worksheet.Cells.Item($r, 12).Value2
            Sa1 = ConvertTo-NullableDouble $Worksheet.Cells.Item($r, 13).Value2
            Sa2 = ConvertTo-NullableDouble $Worksheet.Cells.Item($r, 14).Value2
            Sa3 = ConvertTo-NullableDouble $Worksheet.Cells.Item($r, 15).Value2
            SaMean = ConvertTo-NullableDouble $Worksheet.Cells.Item($r, 16).Value2
        }
    }
    if (-not (Test-ExperimentSaSnapshotComplete -Snapshot $Snapshot -ExpectedRows $previous)) { return $false }
    foreach ($old in $previous) {
        $new = @($Group.SaRows | Where-Object { $_.SampleIndex -eq $old.SampleIndex })
        if ($new.Count -ne 1) { return $false }
        foreach ($field in @('Sa1', 'Sa2', 'Sa3')) {
            if ($null -ne $old.$field -and -not (Test-ExperimentSaNumberEqual $old.$field $new[0].$field)) { return $false }
        }
    }
    return $true
}
