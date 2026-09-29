Set-StrictMode -Version 2.0

. (Join-Path $PSScriptRoot 'ExperimentDataAutomation.Concurrency.ps1')
. (Join-Path $PSScriptRoot 'ExperimentDataAutomation.Completion.ps1')
. (Join-Path $PSScriptRoot 'ExperimentDataAutomation.SourceVersions.ps1')
. (Join-Path $PSScriptRoot 'ExperimentDataAutomation.Transactions.ps1')
. (Join-Path $PSScriptRoot 'ExperimentDataAutomation.Execution.ps1')
. (Join-Path $PSScriptRoot 'ExperimentDataAutomation.ExcelSession.ps1')
. (Join-Path $PSScriptRoot 'ExperimentDataAutomation.SaExtraction.ps1')

function New-ExperimentAutomationConfig {
    param(
        [Parameter(Mandatory = $true)] [string] $ScriptRoot,
        [string] $VolumeLabel = '虎大子李',
        [string] $ExpectedDriveLetter = 'F',
        [string] $MasterWorkbookPath = 'D:\共享文件\实验数据\实验数据.xlsx',
        [string] $MatlabScriptPath = 'D:\共享文件\matlab\program\baiguang.m',
        [string] $MatlabExePath = 'D:\Program Files\MATLAB\R2025b\bin\matlab.exe',
        [string[]] $SourceRelativeFolders = @(
            '实验数据\抛光实验\水基抛光（光镜+白光）',
            '实验数据\抛光实验\油基抛光'
        ),
        [int] $PollSeconds = 5,
        [switch] $DryRun
    )

    [pscustomobject]@{
        VolumeLabel           = $VolumeLabel
        ExpectedDriveLetter   = $ExpectedDriveLetter
        MasterWorkbookPath    = $MasterWorkbookPath
        MatlabScriptPath      = $MatlabScriptPath
        MatlabExePath         = $MatlabExePath
        SaEngine              = 'BuiltIn'
        SourceRelativeFolders = $SourceRelativeFolders
        PollSeconds           = $PollSeconds
        MatlabTimeoutSeconds  = 1200
        ExcelTimeoutSeconds   = 600
        MaxExperimentAttempts = 3
        RetryDelaySeconds     = 60
        RetryFailed           = $false
        DryRun                = [bool] $DryRun
        ScriptRoot            = $ScriptRoot
        LogDirectory          = Join-Path $ScriptRoot 'logs'
        StateDirectory        = Join-Path $ScriptRoot 'state'
        TempDirectory         = Join-Path $ScriptRoot 'temp'
        AutoSheetName         = '自动导入Sa数据'
        PendingSheetName      = '待确认补齐'
        SaOutputFileName      = '所有的Sa数据.xlsx'
        EnableBackfill        = $true
        NewBlockMode          = 'Preview'
        BackupDirectory       = Join-Path $ScriptRoot 'backups'
        SlurryMappings        = [pscustomobject]@{
            '金刚石' = '油基抛光'
            '硅溶胶' = '探索实验'
        }
        ProcessDefaults       = [pscustomobject]@{
            Water = [pscustomobject]@{
                Plate    = 'PEEK'
                Track    = '4'
                Spacing  = '9'
                Fullness = '92%'
                Pressure = '单球0.1N'
                BallMatch = '0.8mm氧化锆球'
            }
            Oil   = [pscustomobject]@{
                Plate    = '铸铁抛光盘'
                Track    = '3'
                Spacing  = '9'
                Fullness = '92%'
                Pressure = '单球0.1N'
                BallMatch = '0.8mm钢球'
            }
        }
    }
}

function Initialize-ExperimentAutomationEnvironment {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Config)

    foreach ($path in @($Config.LogDirectory, $Config.StateDirectory, $Config.TempDirectory)) {
        if (-not (Test-Path -LiteralPath $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

function Write-ExperimentLog {
    param(
        [Parameter(Mandatory = $true)] [pscustomobject] $Config,
        [Parameter(Mandatory = $true)] [string] $Message,
        [string] $Level = 'INFO'
    )

    Initialize-ExperimentAutomationEnvironment -Config $Config
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level, $Message
    $logPath = Join-Path $Config.LogDirectory 'automation.log'
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
    Write-Host $line
}

function Show-ExperimentNotification {
    param(
        [string] $Title = '实验数据自动整理',
        [Parameter(Mandatory = $true)] [string] $Message
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

        $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
        $notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
        $notifyIcon.Visible = $true
        $notifyIcon.BalloonTipTitle = $Title
        $notifyIcon.BalloonTipText = $Message
        $notifyIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        if (-not (Get-Variable -Name ExperimentNotificationResources -Scope Script -ErrorAction SilentlyContinue)) {
            $script:ExperimentNotificationResources = New-Object System.Collections.ArrayList
        }

        $timer = New-Object System.Timers.Timer
        $timer.Interval = 9000
        $timer.AutoReset = $false
        $resource = [pscustomobject]@{
            NotifyIcon = $notifyIcon
            Timer      = $timer
            Handler    = $null
        }
        $handler = {
            try {
                $resource.NotifyIcon.Visible = $false
                $resource.NotifyIcon.Dispose()
                $resource.Timer.Dispose()
            }
            finally {
                [void] $script:ExperimentNotificationResources.Remove($resource)
            }
        }
        $resource.Handler = $handler
        [void] $timer.add_Elapsed($handler)
        [void] $script:ExperimentNotificationResources.Add($resource)

        $notifyIcon.ShowBalloonTip(8000)
        $timer.Start()
    }
    catch {
        Write-Host "${Title}: $Message"
    }
}

function Release-ComObject {
    param([object] $ComObject)

    if ($null -ne $ComObject) {
        try {
            if ([System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
                [void] [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
            }
        }
        catch {
        }
    }
}

function Normalize-ExperimentText {
    param([AllowNull()] [string] $Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $normalized = $Text.ToLowerInvariant()
    $normalized = $normalized -replace '\s+', ''
    $normalized = $normalized -replace '[\\/:*?"<>|（）()\[\]【】{}《》,，.。;；:：_\-—–+＋=＝%％''"“”‘’]', ''
    return $normalized
}

function ConvertTo-IsoExperimentDate {
    param(
        [Parameter(Mandatory = $true)] [string] $Year,
        [Parameter(Mandatory = $true)] [string] $Month,
        [Parameter(Mandatory = $true)] [string] $Day
    )

    $y = [int] $Year
    if ($Year.Length -eq 2) {
        $y = 2000 + $y
    }

    try {
        $dt = Get-Date -Year $y -Month ([int] $Month) -Day ([int] $Day) -Hour 0 -Minute 0 -Second 0
        return $dt.ToString('yyyy-MM-dd')
    }
    catch {
        return $null
    }
}

function Get-DateMatchFromText {
    param([Parameter(Mandatory = $true)] [string] $Text)

    $patterns = @(
        '(?<!\d)(20\d{2})[.\-/年](\d{1,2})[.\-/月](\d{1,2})',
        '(?<!\d)(\d{2})[.\-/年](\d{1,2})[.\-/月](\d{1,2})'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($Text, $pattern)
        if ($match.Success) {
            return ConvertTo-IsoExperimentDate -Year $match.Groups[1].Value -Month $match.Groups[2].Value -Day $match.Groups[3].Value
        }
    }

    return $null
}

function Get-DateTokens {
    param([AllowNull()] [string] $IsoDate)

    if ([string]::IsNullOrWhiteSpace($IsoDate)) {
        return @()
    }

    $dt = [datetime]::ParseExact($IsoDate, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    @(
        $dt.ToString('yyyy-M-d'),
        $dt.ToString('yyyy.MM.dd'),
        ('{0}.{1}.{2}' -f $dt.Year, $dt.Month, $dt.Day),
        ('{0}年{1}月{2}日' -f $dt.Year, $dt.Month, $dt.Day),
        $IsoDate
    )
}

function Get-FirstRegexValue {
    param(
        [Parameter(Mandatory = $true)] [string] $Text,
        [Parameter(Mandatory = $true)] [string] $Pattern,
        [int] $Group = 1
    )

    $match = [regex]::Match($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($match.Success) {
        return $match.Groups[$Group].Value
    }

    return $null
}

function Get-ExperimentMetadataFromPath {
    param([Parameter(Mandatory = $true)] [string] $FolderPath)

    $cleanPath = $FolderPath.TrimEnd('\')
    $leaf = Split-Path -Path $cleanPath -Leaf
    $combined = "$cleanPath $leaf"

    $slurryType = ''
    if ($combined -like '*油基抛光*') {
        $slurryType = '油基抛光'
    }
    elseif ($combined -like '*水基抛光*') {
        $slurryType = '水基抛光'
    }

    $date = Get-DateMatchFromText -Text $combined
    $experimentNo = Get-FirstRegexValue -Text $leaf -Pattern '^\s*(\d+)\s*-'
    if ([string]::IsNullOrWhiteSpace($experimentNo)) {
        $experimentNo = Get-FirstRegexValue -Text $combined -Pattern '实验\s*(\d+)'
    }
    $concentrationNumber = Get-FirstRegexValue -Text $leaf -Pattern '(\d+(?:\.\d+)?)\s*[%％]'
    $concentration = $null
    if ($concentrationNumber) {
        $concentration = "$concentrationNumber%"
    }

    $particleMatch = [regex]::Match($leaf, '(\d+(?:\.\d+)?)\s*(nm|μm|um|mm)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $particleSize = $null
    if ($particleMatch.Success) {
        $particleSize = ($particleMatch.Groups[1].Value + $particleMatch.Groups[2].Value).Replace('UM', 'um')
    }

    $material = ''
    foreach ($candidate in @('金刚石', '硅溶胶', '氧化铝', '二氧化硅', '碳化硅', 'CeO2', 'SiO2')) {
        if ($combined -like "*$candidate*") {
            $material = $candidate
            break
        }
    }

    $duration = Get-FirstRegexValue -Text $leaf -Pattern '(\d+(?:\.\d+)?\s*h)' -Group 1
    if ($duration) {
        $duration = $duration -replace '\s+', ''
    }

    $keySource = "$slurryType|$date|$experimentNo|$leaf"
    $key = Normalize-ExperimentText -Text $keySource

    [pscustomobject]@{
        FolderPath      = $cleanPath
        FolderName      = $leaf
        SlurryType      = $slurryType
        ExperimentDate  = $date
        ExperimentNo    = $experimentNo
        Concentration   = $concentration
        ParticleSize    = $particleSize
        Material        = $material
        Duration        = $duration
        ConditionText   = $leaf
        ExperimentKey   = $key
    }
}

function Get-UsbDriveByLabel {
    param(
        [Parameter(Mandatory = $true)] [string] $VolumeLabel,
        [string] $ExpectedDriveLetter = 'F'
    )

    $volumes = @(Get-Volume -ErrorAction SilentlyContinue | Where-Object { $_.FileSystemLabel -eq $VolumeLabel -and $_.DriveLetter })
    if ($volumes.Count -gt 0) {
        return $volumes[0]
    }

    if (-not [string]::IsNullOrWhiteSpace($ExpectedDriveLetter)) {
        $driveRoot = "$($ExpectedDriveLetter.TrimEnd(':')):\"
        if (Test-Path -LiteralPath $driveRoot) {
            $volume = Get-Volume -DriveLetter $ExpectedDriveLetter.TrimEnd(':') -ErrorAction SilentlyContinue
            if ($volume -and $volume.FileSystemLabel -eq $VolumeLabel) {
                return $volume
            }
        }
    }

    return $null
}

function Get-SourceRootPaths {
    param(
        [Parameter(Mandatory = $true)] [string] $DriveLetter,
        [Parameter(Mandatory = $true)] [string[]] $RelativeFolders
    )

    $root = "$($DriveLetter.TrimEnd(':')):\"
    foreach ($relative in $RelativeFolders) {
        Join-Path $root $relative
    }
}

function Get-CandidateExperimentFolders {
    param([Parameter(Mandatory = $true)] [string[]] $SourceRoots)

    $foldersByPath = @{}
    foreach ($root in $SourceRoots) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }

        Get-ChildItem -LiteralPath $root -Recurse -File -Filter '*.xlsx' -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -notlike '~$*' -and
                $_.Name -ne '所有的Sa数据.xlsx'
            } |
            ForEach-Object {
                if ($_.Directory.Name -match '^\d') {
                    $foldersByPath[$_.Directory.FullName] = $_.Directory
                }
            }
    }

    $foldersByPath.Values | Sort-Object FullName
}

function ConvertTo-NullableDouble {
    param([AllowNull()] [object] $Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) {
        return [double] $Value
    }

    $text = ([string] $Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $parsed = 0.0
    $style = [System.Globalization.NumberStyles]::Any
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    if ([double]::TryParse($text, $style, $culture, [ref] $parsed)) {
        return $parsed
    }

    $culture = [System.Globalization.CultureInfo]::CurrentCulture
    if ([double]::TryParse($text, $style, $culture, [ref] $parsed)) {
        return $parsed
    }

    return $null
}

function Convert-SaGridToRows {
    param([Parameter(Mandatory = $true)] [object[]] $Rows, [switch] $PreserveEmptyRows)

    $sampleIndex = 0
    foreach ($row in $Rows) {
        $sampleIndex++
        $values = @($row)
        if ($values.Count -eq 1 -and $values[0] -is [System.Array]) {
            $values = @($values[0])
        }

        $sa1 = if ($values.Count -ge 1) { ConvertTo-NullableDouble $values[0] } else { $null }
        $sa2 = if ($values.Count -ge 2) { ConvertTo-NullableDouble $values[1] } else { $null }
        $sa3 = if ($values.Count -ge 3) { ConvertTo-NullableDouble $values[2] } else { $null }
        $numbers = @(@($sa1, $sa2, $sa3) | Where-Object { $null -ne $_ })

        if ($numbers.Count -eq 0 -and -not $PreserveEmptyRows) {
            continue
        }

        $mean = if ($numbers.Count) { [Math]::Round([double](($numbers | Measure-Object -Average).Average), 6) } else { $null }
        [pscustomobject]@{
            SampleIndex = $sampleIndex
            Sa1         = $sa1
            Sa2         = $sa2
            Sa3         = $sa3
            SaMean      = $mean
        }
    }
}

function Get-ExcelArrayValue {
    param(
        [AllowNull()] [object] $Values,
        [int] $Row,
        [int] $Column,
        [int] $RowCount,
        [int] $ColumnCount
    )

    if ($null -eq $Values) {
        return $null
    }

    if ($RowCount -eq 1 -and $ColumnCount -eq 1) {
        return $Values
    }

    try {
        return $Values[$Row, $Column]
    }
    catch {
        try {
            if ($RowCount -eq 1) {
                return $Values[$Column]
            }
            if ($ColumnCount -eq 1) {
                return $Values[$Row]
            }
        }
        catch {
        }
    }

    return $null
}

function Read-SaRowsFromWorkbook {
    param([Parameter(Mandatory = $true)] [string] $SaWorkbookPath)

    if (-not (Test-Path -LiteralPath $SaWorkbookPath)) {
        throw "未找到 Sa 输出表：$SaWorkbookPath"
    }

    $session = $null; $excel = $null
    $workbook = $null
    try {
        $session = New-ExperimentExcelSession
        $excel = $session.Application
        $workbook = $excel.Workbooks.Open($SaWorkbookPath, $null, $true)
        $worksheet = $workbook.Worksheets.Item(1)
        $used = $worksheet.UsedRange
        $rowCount = [int] $used.Row + [int] $used.Rows.Count - 1
        $columnCount = 3
        $readColumns = 3
        $values = $worksheet.Range("A1:C$rowCount").Value2
        $preserveEmpty = [string]$worksheet.Cells.Item(1, 4).Value2 -match '^[0-3]/3$'
        $rows = @()

        for ($r = 1; $r -le $rowCount; $r++) {
            $row = @()
            for ($c = 1; $c -le $readColumns; $c++) {
                $row += Get-ExcelArrayValue -Values $values -Row $r -Column $c -RowCount $rowCount -ColumnCount $columnCount
            }
            $rows += ,$row
        }

        return @(Convert-SaGridToRows -Rows $rows -PreserveEmptyRows:$preserveEmpty)
    }
    finally {
        Close-ExperimentExcelSession -Session $session -Workbook $workbook
    }
}

function Get-WorkbookExperimentContext {
    param(
        [Parameter(Mandatory = $true)] [string] $WorkbookPath,
        [string] $AutoSheetName = '自动导入Sa数据',
        [string[]] $RecordedSheetNames = @('探索实验', '油基抛光')
    )

    if (-not (Test-Path -LiteralPath $WorkbookPath)) {
        throw "未找到主实验数据表：$WorkbookPath"
    }

    $session = $null; $excel = $null
    $workbook = $null
    $blocks = @()
    $allParts = New-Object System.Collections.Generic.List[string]
    $importedKeys = @{}
    $recordedSheetLookup = @{}
    foreach ($name in $RecordedSheetNames) {
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            $recordedSheetLookup[$name] = $true
        }
    }

    try {
        $session = New-ExperimentExcelSession
        $excel = $session.Application
        $workbook = $excel.Workbooks.Open($WorkbookPath, $null, $true)

        foreach ($worksheet in $workbook.Worksheets) {
            $used = $worksheet.UsedRange
            $rowCount = [int] $used.Rows.Count
            $columnCount = [int] $used.Columns.Count
            $values = $used.Value2

            if ($worksheet.Name -eq $AutoSheetName) {
                $versionColumns = Get-ExperimentVersionColumns $worksheet
                for ($r = 2; $r -le $rowCount; $r++) {
                    if ($versionColumns.ContainsKey('版本状态') -and [string] $worksheet.Cells.Item($r, $versionColumns['版本状态']).Value2 -eq 'Archived') { continue }
                    $key = [string] (Get-ExcelArrayValue -Values $values -Row $r -Column 2 -RowCount $rowCount -ColumnCount $columnCount)
                    $folder = [string] (Get-ExcelArrayValue -Values $values -Row $r -Column 3 -RowCount $rowCount -ColumnCount $columnCount)
                    $sampleIndex = [string] (Get-ExcelArrayValue -Values $values -Row $r -Column 12 -RowCount $rowCount -ColumnCount $columnCount)
                    $sampleNumber = 0
                    $hasSampleRecord = [int]::TryParse($sampleIndex, [ref] $sampleNumber) -and $sampleNumber -gt 0

                    if ($hasSampleRecord -and -not [string]::IsNullOrWhiteSpace($key)) {
                        $importedKeys[$key] = $true
                    }
                    if ($hasSampleRecord -and -not [string]::IsNullOrWhiteSpace($folder)) {
                        $importedKeys[(Normalize-ExperimentText $folder)] = $true
                    }
                }
            }

            $scanRecordedBlocks = $recordedSheetLookup.ContainsKey($worksheet.Name)
            if ($scanRecordedBlocks) {
                foreach ($saBlock in @(Get-ExperimentBackfillBlocksFromWorkbook -Workbook $workbook -SheetNames @($worksheet.Name))) {
                    $blocks += [pscustomobject]@{
                        SheetName = $saBlock.SheetName
                        Row = $saBlock.StartRow
                        Column = $saBlock.StartColumn
                        Title = $saBlock.Title
                        Text = $saBlock.SlurryText
                        SaSnapshot = Get-ExperimentSaSnapshot -Worksheet $worksheet -Block $saBlock
                    }
                }
            }
            for ($r = 1; $r -le $rowCount; $r++) {
                for ($c = 1; $c -le $columnCount; $c++) {
                    $value = Get-ExcelArrayValue -Values $values -Row $r -Column $c -RowCount $rowCount -ColumnCount $columnCount
                    if ($null -eq $value) {
                        continue
                    }

                    $text = ([string] $value).Trim()
                    if ([string]::IsNullOrWhiteSpace($text)) {
                        continue
                    }

                    $allParts.Add($text)

                    if ($scanRecordedBlocks -and $text -match '抛光' -and [string]::IsNullOrWhiteSpace((Get-ExperimentNumberFromTitle $text)) -and ($text -match '实验|20\d{2}|\d{2}[.\-/年]\d{1,2}[.\-/月]\d{1,2}')) {
                        $blockParts = New-Object System.Collections.Generic.List[string]
                        $maxRow = [Math]::Min($rowCount, $r + 11)
                        $maxColumn = [Math]::Min($columnCount, $c + 8)
                        for ($rr = $r; $rr -le $maxRow; $rr++) {
                            for ($cc = $c; $cc -le $maxColumn; $cc++) {
                                $blockValue = Get-ExcelArrayValue -Values $values -Row $rr -Column $cc -RowCount $rowCount -ColumnCount $columnCount
                                if ($null -ne $blockValue -and -not [string]::IsNullOrWhiteSpace([string] $blockValue)) {
                                    $blockParts.Add(([string] $blockValue).Trim())
                                }
                            }
                        }

                        $blocks += [pscustomobject]@{
                            SheetName = $worksheet.Name
                            Row       = $r + [int] $used.Row - 1
                            Column    = $c + [int] $used.Column - 1
                            Title     = $text
                            SaSnapshot = $null
                            Text      = ($blockParts -join ' ')
                        }
                    }
                }
            }
        }

        [pscustomobject]@{
            AllText        = ($allParts -join ' ')
            NormalizedText = Normalize-ExperimentText (($allParts -join ' '))
            Blocks         = $blocks
            ImportedKeys   = $importedKeys
            SourceVersions = Read-ExperimentSourceVersionsFromWorkbook $workbook
        }
    }
    finally {
        Close-ExperimentExcelSession -Session $session -Workbook $workbook
    }
}

function Test-NormalizedContainsAny {
    param(
        [Parameter(Mandatory = $true)] [string] $NormalizedHaystack,
        [Parameter(Mandatory = $true)] [string[]] $Needles
    )

    foreach ($needle in $Needles) {
        $normalizedNeedle = Normalize-ExperimentText $needle
        if ($normalizedNeedle.Length -gt 0 -and $NormalizedHaystack.Contains($normalizedNeedle)) {
            return $true
        }
    }

    return $false
}

function Test-ExperimentAlreadyRecorded {
    param(
        [Parameter(Mandatory = $true)] [pscustomobject] $Metadata,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]] $WorkbookBlocks,
        [AllowNull()] [object] $SlurryMappings = $null,
        [AllowEmptyCollection()] [object[]] $ExpectedSaRows = @()
    )

    $hasLeadingExperimentNumber = -not [string]::IsNullOrWhiteSpace($Metadata.ExperimentNo) -and [string] $Metadata.FolderName -match '^\s*\d+\s*-'
    if ($hasLeadingExperimentNumber) {
        $parsed = Parse-StructuredExperimentFolderName -FolderPath $Metadata.FolderPath -SlurryMappings $SlurryMappings
        if ($parsed.Issues.Count -gt 0 -or $ExpectedSaRows.Count -eq 0) { return $false }
        $matches = @($WorkbookBlocks | Where-Object {
            $_.SheetName -eq $parsed.TargetSheetName -and
            (Get-ExperimentNumberFromTitle -Title (Get-ExperimentObjectProperty $_ 'Title' '')) -eq $parsed.ExperimentNo
        })
        if ($matches.Count -ne 1) { return $false }
        $titleDate = Get-DateMatchFromText -Text $matches[0].Title
        if ($titleDate -ne $parsed.ExperimentDate) { return $false }
        $snapshot = Get-ExperimentObjectProperty $matches[0] 'SaSnapshot' $null
        if ($null -eq $snapshot) { return $false }
        return (Test-ExperimentSaSnapshotComplete -Snapshot $snapshot -ExpectedRows $ExpectedSaRows)
    }

    $folderNameNormalized = Normalize-ExperimentText $Metadata.FolderName

    foreach ($block in $WorkbookBlocks) {
        $blockNormalized = Normalize-ExperimentText ([string] $block.Text)

        if ($folderNameNormalized.Length -gt 12 -and $blockNormalized.Contains($folderNameNormalized)) {
            return $true
        }

        $score = 0
        $hasDateOrNumber = $false

        if (-not [string]::IsNullOrWhiteSpace($Metadata.ExperimentDate)) {
            if (Test-NormalizedContainsAny -NormalizedHaystack $blockNormalized -Needles (Get-DateTokens $Metadata.ExperimentDate)) {
                $score += 3
                $hasDateOrNumber = $true
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($Metadata.ExperimentNo)) {
            if (Test-NormalizedContainsAny -NormalizedHaystack $blockNormalized -Needles @("实验$($Metadata.ExperimentNo)")) {
                $score += 3
                $hasDateOrNumber = $true
            }
        }

        foreach ($field in @('Concentration', 'ParticleSize', 'Material', 'SlurryType', 'Duration')) {
            $value = [string] $Metadata.$field
            if ([string]::IsNullOrWhiteSpace($value)) {
                continue
            }

            if (Test-NormalizedContainsAny -NormalizedHaystack $blockNormalized -Needles @($value)) {
                switch ($field) {
                    'Concentration' { $score += 2 }
                    'ParticleSize' { $score += 2 }
                    default { $score += 1 }
                }
            }
        }

        if ($hasDateOrNumber -and $score -ge 5) {
            return $true
        }
    }

    return $false
}

function Ensure-AutoImportWorksheet {
    param(
        [Parameter(Mandatory = $true)] [object] $Workbook,
        [Parameter(Mandatory = $true)] [string] $SheetName
    )

    $worksheet = $null
    foreach ($sheet in $Workbook.Worksheets) {
        if ($sheet.Name -eq $SheetName) {
            $worksheet = $sheet
            break
        }
    }

    if (-not $worksheet) {
        $missing = [System.Type]::Missing
        $worksheet = $Workbook.Worksheets.Add($missing, $Workbook.Worksheets.Item($Workbook.Worksheets.Count))
        $worksheet.Name = $SheetName
    }

    $headers = @(
        '导入时间',
        '实验Key',
        '来源文件夹',
        '抛光类型',
        '实验日期',
        '实验编号',
        '文件夹名',
        '抛光液浓度',
        '磨粒粒径',
        '磨粒/抛光液',
        '时长',
        '样品序号',
        'Sa1(nm)',
        'Sa2(nm)',
        'Sa3(nm)',
        'Sa均值(nm)',
        'MATLAB输出表',
        '备注'
    )

    for ($i = 0; $i -lt $headers.Count; $i++) {
        if ([string]::IsNullOrWhiteSpace([string] $worksheet.Cells.Item(1, $i + 1).Value2)) {
            $worksheet.Cells.Item(1, $i + 1).Value2 = $headers[$i]
        }
    }

    $worksheet.Rows.Item(1).Font.Bold = $true
    return $worksheet
}

function ConvertTo-ExcelCellText {
    param([AllowNull()] [object] $Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [System.IFormattable]) {
        return $Value.ToString($null, [System.Globalization.CultureInfo]::InvariantCulture)
    }

    return [string] $Value
}

function Append-AutoImportRecords {
    param(
        [Parameter(Mandatory = $true)] [string] $WorkbookPath,
        [string] $SheetName = '自动导入Sa数据',
        [Parameter(Mandatory = $true)] [object[]] $Records
    )
    Invoke-WithExperimentWorkbookLock -WorkbookPath $WorkbookPath -Action {
        if (Test-ExperimentTransactionActive $WorkbookPath) {
            Append-AutoImportRecordsCore -WorkbookPath $WorkbookPath -SheetName $SheetName -Records $Records
        }
        else {
            $root = Split-Path -Parent ([IO.Path]::GetFullPath($WorkbookPath))
            $settings = [pscustomobject]@{ ScriptRoot = $root; MasterWorkbookPath = $WorkbookPath; BackupDirectory = (Join-Path $root 'backups') }
            $transaction = Invoke-ExperimentWorkbookTransaction -Config $settings -Action {
                param($working)
                Append-AutoImportRecordsCore -WorkbookPath $working.MasterWorkbookPath -SheetName $SheetName -Records $Records
            }
            $transaction.Result
        }
    }
}

function Append-AutoImportRecordsCore {
    param(
        [Parameter(Mandatory = $true)] [string] $WorkbookPath,
        [Parameter(Mandatory = $true)] [string] $SheetName,
        [Parameter(Mandatory = $true)] [object[]] $Records
    )

    if ($Records.Count -eq 0) {
        return 0
    }

    $session = $null; $excel = $null
    $workbook = $null
    try {
        $session = New-ExperimentExcelSession
        $excel = $session.Application
        $workbook = $excel.Workbooks.Open($WorkbookPath)

        if ($workbook.ReadOnly) {
            throw "主实验数据表现在是只读状态，请关闭已打开的 Excel 文件后重试：$WorkbookPath"
        }

        $worksheet = Ensure-AutoImportWorksheet -Workbook $workbook -SheetName $SheetName
        $used = $worksheet.UsedRange
        $nextRow = [int] $used.Rows.Count + 1
        if ($nextRow -lt 2) {
            $nextRow = 2
        }

        foreach ($record in $Records) {
            $values = @(
                (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),
                $record.Metadata.ExperimentKey,
                $record.Metadata.FolderPath,
                $record.Metadata.SlurryType,
                $record.Metadata.ExperimentDate,
                $record.Metadata.ExperimentNo,
                $record.Metadata.FolderName,
                $record.Metadata.Concentration,
                $record.Metadata.ParticleSize,
                $record.Metadata.Material,
                $record.Metadata.Duration,
                $record.Sa.SampleIndex,
                $record.Sa.Sa1,
                $record.Sa.Sa2,
                $record.Sa.Sa3,
                $record.Sa.SaMean,
                $record.SaOutputPath,
                $record.Note
            )

            for ($column = 1; $column -le $values.Count; $column++) {
                $worksheet.Cells.Item($nextRow, $column).Value2 = ConvertTo-ExcelCellText $values[$column - 1]
            }
            Set-ExperimentImportRevision -Worksheet $worksheet -Record $record -RowNumber $nextRow
            $nextRow++
        }

        $worksheet.Columns.AutoFit() | Out-Null
        Save-ExperimentWorkbookCheckpoint -Workbook $workbook
        return $Records.Count
    }
    finally {
        Close-ExperimentExcelSession -Session $session -Workbook $workbook
    }
}

function ConvertTo-MatlabStringLiteral {
    param([Parameter(Mandatory = $true)] [string] $Path)
    return ($Path -replace "'", "''")
}

function New-TemporaryBaiguangScript {
    param(
        [Parameter(Mandatory = $true)] [string] $OriginalScriptPath,
        [Parameter(Mandatory = $true)] [string] $TargetFolder,
        [Parameter(Mandatory = $true)] [string] $TempDirectory
    )

    if (-not (Test-Path -LiteralPath $OriginalScriptPath)) {
        throw "未找到 MATLAB 程序：$OriginalScriptPath"
    }
    if (-not (Test-Path -LiteralPath $TempDirectory)) {
        New-Item -ItemType Directory -Path $TempDirectory -Force | Out-Null
    }

    $folderForMatlab = $TargetFolder.TrimEnd('\') + '\'
    $matlabFolder = ConvertTo-MatlabStringLiteral -Path $folderForMatlab
    $replacement = "folder_path = '$matlabFolder'; % auto replaced by ExperimentDataAutomation"

    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $text = [System.IO.File]::ReadAllText($OriginalScriptPath, $utf8)
    $regex = New-Object System.Text.RegularExpressions.Regex("(?m)^\s*folder_path\s*=\s*'[^']*';[^\r\n]*")
    if (-not $regex.IsMatch($text)) {
        throw "无法在 MATLAB 程序中找到 folder_path = '...'; 这一行：$OriginalScriptPath"
    }

    $evaluator = [System.Text.RegularExpressions.MatchEvaluator] { param($match) $replacement }
    $patched = $regex.Replace($text, $evaluator, 1)

    $excludeLine = "file_list = file_list(~strcmp({file_list.name}, '所有的Sa数据.xlsx'));"
    if (-not $patched.Contains($excludeLine)) {
        $fileListRegex = New-Object System.Text.RegularExpressions.Regex("(?m)^(\s*file_list\s*=\s*dir\(fullfile\(folder_path\s*,\s*'\*\.xlsx'\)\);[^\r\n]*)(\r?\n)")
        if ($fileListRegex.IsMatch($patched)) {
            $fileListEvaluator = [System.Text.RegularExpressions.MatchEvaluator] {
                param($match)
                return $match.Groups[1].Value + $match.Groups[2].Value + $excludeLine + $match.Groups[2].Value
            }
            $patched = $fileListRegex.Replace($patched, $fileListEvaluator, 1)
        }
        else {
            throw "无法在 MATLAB 程序中找到 file_list = dir(fullfile(folder_path,'*.xlsx')); 这一行：$OriginalScriptPath"
        }
    }

    $lockFilter = "file_list = file_list(~startsWith({file_list.name}, '~`$'));"
    if (-not $patched.Contains($lockFilter)) { $patched = $patched.Replace($excludeLine, ($excludeLine + [Environment]::NewLine + $lockFilter)) }
    $safeName = 'baiguang_auto_{0}.m' -f (Get-Date -Format 'yyyyMMdd_HHmmss_ffff')
    $tempPath = Join-Path $TempDirectory $safeName
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($tempPath, $patched, $utf8Bom)
    return $tempPath
}

function Invoke-BaiguangMatlab {
    param(
        [Parameter(Mandatory = $true)] [pscustomobject] $Config,
        [Parameter(Mandatory = $true)] [string] $TargetFolder
    )

    $outputPath = Join-Path $TargetFolder $Config.SaOutputFileName
    if ($Config.DryRun) {
        Write-ExperimentLog -Config $Config -Message "DRY-RUN: would run MATLAB for $TargetFolder"
        return $outputPath
    }

    if (-not (Test-Path -LiteralPath $Config.MatlabExePath)) {
        $matlabCommand = Get-Command matlab -ErrorAction SilentlyContinue
        if ($matlabCommand) {
            $Config.MatlabExePath = $matlabCommand.Source
        }
        else {
            throw "未找到 MATLAB 可执行文件：$($Config.MatlabExePath)"
        }
    }

    $tempScript = New-TemporaryBaiguangScript -OriginalScriptPath $Config.MatlabScriptPath -TargetFolder $TargetFolder -TempDirectory $Config.TempDirectory
    $scriptForMatlab = ConvertTo-MatlabStringLiteral -Path $tempScript
    $batchCommand = "run('$scriptForMatlab')"
    Write-ExperimentLog -Config $Config -Message "启动 MATLAB：$TargetFolder"

    if ([IO.Path]::GetFileName($Config.SaOutputFileName) -ne $Config.SaOutputFileName) { throw 'Sa 输出文件名不能包含目录路径' }
    $previousOutput = $null
    if ([IO.File]::Exists($outputPath)) {
        $previousOutput = Join-Path $Config.TempDirectory (([guid]::NewGuid().ToString('N')) + '.sa-backup')
        [IO.File]::Copy($outputPath, $previousOutput, $false)
    }
    try {
    # A successful process must produce a new output, not reuse a stale workbook.
    if ([IO.File]::Exists($outputPath)) { [IO.File]::Delete($outputPath) }
    $matlabLogPath = Join-Path $Config.LogDirectory ('matlab_{0}_{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'), [guid]::NewGuid().ToString('N'))
    if ($batchCommand.Contains('"')) { throw 'MATLAB 临时脚本路径不能包含双引号。' }
    $processResult = Invoke-ExperimentTimedProcess -FilePath $Config.MatlabExePath -Arguments ('-wait -batch "' + $batchCommand + '"') -TimeoutSeconds $Config.MatlabTimeoutSeconds -LogPrefix $matlabLogPath
    $exitCode = $processResult.ExitCode

    if ($exitCode -ne 0) {
        throw "MATLAB 运行失败，退出码 $exitCode。日志：$matlabLogPath"
    }

    if (-not (Test-Path -LiteralPath $outputPath)) {
        throw "MATLAB 已结束但未生成输出表：$outputPath。日志：$matlabLogPath"
    }

    Write-ExperimentLog -Config $Config -Message "MATLAB 完成，输出：$outputPath"
    return $outputPath
    }
    catch {
        if ($previousOutput -and [IO.File]::Exists($previousOutput)) { [IO.File]::Copy($previousOutput, $outputPath, $true) }
        elseif ([IO.File]::Exists($outputPath)) { [IO.File]::Delete($outputPath) }
        throw
    }
}

function New-AutoImportRecordsFromFolder {
    param(
        [Parameter(Mandatory = $true)] [pscustomobject] $Metadata,
        [Parameter(Mandatory = $true)] [string] $SaOutputPath
    )

    $saRows = @(Read-SaRowsFromWorkbook -SaWorkbookPath $SaOutputPath)
    foreach ($sa in $saRows) {
        [pscustomobject]@{
            Metadata     = $Metadata
            Sa           = $sa
            SaOutputPath = $SaOutputPath
            Note         = ('自动导入；有效测量 {0}/3' -f @(@($sa.Sa1, $sa.Sa2, $sa.Sa3) | Where-Object { $null -ne $_ }).Count)
        }
    }
}

function Invoke-ExperimentDataAutomationOnce {
    param(
        [Parameter(Mandatory = $true)] [pscustomobject] $Config,
        [switch] $SilentWhenNoDrive
    )
    Invoke-WithExperimentWorkbookLock -WorkbookPath $Config.MasterWorkbookPath -Action {
        $previousExcelPolicy = Get-Variable -Name ExperimentExcelTimeoutSeconds -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        $script:ExperimentExcelTimeoutSeconds = if ($Config.PSObject.Properties['ExcelTimeoutSeconds']) { [int]$Config.ExcelTimeoutSeconds } else { 600 }
        try {
            Invoke-ExperimentDataAutomationOnceCore -Config $Config -SilentWhenNoDrive:$SilentWhenNoDrive
        }
        catch {
            if (Get-Command Update-PolishingSentinelStatus -ErrorAction SilentlyContinue) {
                $errorState = if ($_.Exception -is [TimeoutException]) { 'TimedOut' } else { 'Error' }
                Update-PolishingSentinelStatus -ScriptRoot $Config.ScriptRoot -State $errorState -Message ("自动整理出错：{0}" -f $_.Exception.Message) -Data @{ ProcessId = $PID } | Out-Null
            }
            throw
        }
        finally { $script:ExperimentExcelTimeoutSeconds = $previousExcelPolicy }
    }
}

function Invoke-ExperimentDataAutomationOnceCore {
    param(
        [Parameter(Mandatory = $true)] [pscustomobject] $Config,
        [switch] $SilentWhenNoDrive
    )

    Initialize-ExperimentAutomationEnvironment -Config $Config

    $statusCommand = Get-Command Update-PolishingSentinelStatus -ErrorAction SilentlyContinue
    if (-not $Config.DryRun) {
        if ($statusCommand) {
            Update-PolishingSentinelStatus -ScriptRoot $Config.ScriptRoot -State 'Recovering' -Message '正在检查并恢复中断的事务' -Data @{ ProcessId = $PID } | Out-Null
        }
        Repair-ExperimentInterruptedTransactions $Config
    }
    $statusState = $null
    if ($statusCommand) {
        Update-PolishingSentinelStatus -ScriptRoot $Config.ScriptRoot -State 'CheckingUsb' -Message '正在检测 U 盘' -Data @{ VolumeLabel = $Config.VolumeLabel } | Out-Null
        $statusState = 'CheckingUsb'
    }

    $volume = Get-UsbDriveByLabel -VolumeLabel $Config.VolumeLabel -ExpectedDriveLetter $Config.ExpectedDriveLetter
    if (-not $volume) {
        Write-ExperimentLog -Config $Config -Message "未检测到卷标为 $($Config.VolumeLabel) 的 U 盘。"
        if (-not $SilentWhenNoDrive) {
            Show-ExperimentNotification -Message "未检测到卷标为 $($Config.VolumeLabel) 的 U 盘。"
        }
        if ($statusCommand) {
            Update-PolishingSentinelStatus -ScriptRoot $Config.ScriptRoot -State 'WaitingForUsb' -Message '等待 U 盘插入' -Data @{ ExpectedDriveLetter = $Config.ExpectedDriveLetter; VolumeLabel = $Config.VolumeLabel } | Out-Null
            $statusState = 'WaitingForUsb'
        }
        $result = [pscustomobject]@{ DriveFound = $false; CandidateFolders = 0; NewFolders = 0; ImportedRows = 0 }
        if ($statusCommand) {
            $result | Add-Member -NotePropertyName StatusState -NotePropertyValue $statusState
        }
        return $result
    }

    $driveLetter = [string] $volume.DriveLetter
    Write-ExperimentLog -Config $Config -Message "检测到 U 盘 $driveLetter`:，卷标 $($Config.VolumeLabel)。"
    if ($statusCommand) {
        Update-PolishingSentinelStatus -ScriptRoot $Config.ScriptRoot -State 'Scanning' -Message ("正在扫描 U 盘 {0}:" -f $driveLetter) -Data @{ DriveLetter = $driveLetter } | Out-Null
        $statusState = 'Scanning'
    }

    $sourceRoots = @(Get-SourceRootPaths -DriveLetter $driveLetter -RelativeFolders $Config.SourceRelativeFolders)
    foreach ($root in $sourceRoots) {
        Write-ExperimentLog -Config $Config -Message "扫描目录：$root"
    }

    $folders = @(Get-CandidateExperimentFolders -SourceRoots $sourceRoots)
    $recordedSheetNames = @('探索实验', '油基抛光')
    if ($Config.PSObject.Properties['TargetSheets']) {
        $recordedSheetNames = @($Config.TargetSheets)
    }
    elseif ($Config.PSObject.Properties['SlurryMappings'] -and $null -ne $Config.SlurryMappings) {
        $recordedSheetNames = @($Config.SlurryMappings.PSObject.Properties | ForEach-Object { [string] $_.Value } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    }
    $context = Get-WorkbookExperimentContext -WorkbookPath $Config.MasterWorkbookPath -AutoSheetName $Config.AutoSheetName -RecordedSheetNames $recordedSheetNames
    $newFolders = @()
    $sourceUpdates = @()
    $sourceVersions = Get-ExperimentObjectProperty $context 'SourceVersions' @{}
    $folders = @(Get-ExperimentTrackedFolders -Folders $folders -Versions $sourceVersions -SourceRoots $sourceRoots -DriveLetter $driveLetter -SlurryMappings $Config.SlurryMappings)
    $identities = @{}
    $discoveryErrors = @()
    foreach ($folder in $folders) {
        $identity = Get-ExperimentSourceIdentity -FolderPath $folder.FullName -SlurryMappings $Config.SlurryMappings
        if (-not $identities.ContainsKey($identity)) { $identities[$identity] = 0 }
        $identities[$identity]++
    }
    foreach ($folder in $folders) {
        try {
        $metadata = Get-ExperimentMetadataFromPath -FolderPath $folder.FullName
        $identity = Get-ExperimentSourceIdentity -FolderPath $folder.FullName -SlurryMappings $Config.SlurryMappings
        if ($identities[$identity] -gt 1) { throw "源目录中出现重复实验身份，需要确认：$($folder.FullName)" }
        $manifest = Get-ExperimentSourceManifest -FolderPath $folder.FullName -OutputFileName $Config.SaOutputFileName
        $previousVersion = $sourceVersions[$identity]
        $changeKind = Get-ExperimentSourceChangeKind -Previous $previousVersion -Current $manifest
        $sourceVersion = [pscustomobject]@{
            Identity = $identity; FolderPath = $folder.FullName; Manifest = $manifest
            Revision = if ($previousVersion) { $previousVersion.Revision + 1 } else { 1 }
            ChangeKind = $changeKind; SlurryMappings = $Config.SlurryMappings
        }
        $sourceChanged = $null -ne $previousVersion -and $changeKind -ne 'Unchanged'
        $keyAlreadyImported = $context.ImportedKeys.ContainsKey($metadata.ExperimentKey) -or $context.ImportedKeys.ContainsKey((Normalize-ExperimentText $metadata.FolderPath))
        if ($keyAlreadyImported -and -not $sourceChanged) {
            if ($null -eq $previousVersion) {
                $sourceVersion.Revision = 0; $sourceVersion.ChangeKind = 'BaselineUnverified'
                $sourceUpdates += $sourceVersion
                Write-ExperimentLog -Config $Config -Message "建立当前源文件基线，历史 Sa 未重新核验源文件：$($metadata.FolderName)"
            }
            Write-ExperimentLog -Config $Config -Message "Sa 已导入，不重复计算；后续核验主表回写：$($metadata.FolderName)"
            continue
        }

        $expectedSaRows = @()
        $existingOutput = Join-Path $folder.FullName $Config.SaOutputFileName
        if ($metadata.FolderName -match '^\s*\d+\s*-' -and (Test-Path -LiteralPath $existingOutput)) {
            $expectedSaRows = @(Read-SaRowsFromWorkbook -SaWorkbookPath $existingOutput)
        }
        if (-not $sourceChanged -and (Test-ExperimentAlreadyRecorded -Metadata $metadata -WorkbookBlocks $context.Blocks -SlurryMappings $Config.SlurryMappings -ExpectedSaRows $expectedSaRows)) {
            if ($null -eq $previousVersion) {
                $sourceVersion.Revision = 0; $sourceVersion.ChangeKind = 'BaselineUnverified'
                $sourceUpdates += $sourceVersion
            }
            $matchReason = if ($metadata.FolderName -match '^\s*\d+\s*-') { '实验号、工作表和 Sa 完整性核验通过' } else { '无前置实验号，沿用历史命名匹配规则' }
            Write-ExperimentLog -Config $Config -Message "跳过主表中已记录的实验（$matchReason）：$($metadata.FolderName)"
            continue
        }

        $newFolders += [pscustomobject]@{
            Folder   = $folder
            Metadata = $metadata
            SourceVersion = $sourceVersion
        }
        $sourceUpdates += $sourceVersion
        }
        catch {
            $discoveryErrors += [pscustomobject]@{ Folder = $folder.FullName; Stage = 'Scanning'; Message = $_.Exception.Message; Attempt = 0 }
            Write-ExperimentLog -Config $Config -Level ERROR -Message ("实验目录检查失败，已隔离：{0}；{1}" -f $folder.FullName, $_.Exception.Message)
        }
    }

    if ($newFolders.Count -eq 0) {
        $backfillResult = $null
        $enableBackfill = $true
        if ($Config.PSObject.Properties['EnableBackfill']) {
            $enableBackfill = [bool] $Config.EnableBackfill
        }
        $newBlockMode = 'Preview'
        if ($Config.PSObject.Properties['NewBlockMode'] -and -not [string]::IsNullOrWhiteSpace([string] $Config.NewBlockMode)) {
            $newBlockMode = [string] $Config.NewBlockMode
        }

        if (-not $Config.DryRun -and $enableBackfill -and ($discoveryErrors.Count -eq 0 -or $discoveryErrors.Count -lt $folders.Count) -and (Get-Command Invoke-ExperimentBackfillPhase1 -ErrorAction SilentlyContinue)) {
            $commit = Invoke-ExperimentImportCommit -Config $Config -SourceUpdates $sourceUpdates
            $backfillResult = $commit.Backfill
            $phase2LogText = ''
            if ($backfillResult.PSObject.Properties['PreviewNewBlocks']) {
                $phase2LogText = '，新建块预览 {0} 个，自动新建 {1} 个，新建失败 {2} 个' -f $backfillResult.PreviewNewBlocks, $backfillResult.CreatedNewBlocks, $backfillResult.NewBlockFailures
            }
            Write-ExperimentLog -Config $Config -Message ("回写完成：已有块 {0} 个{1}，待确认 {2} 条。备份：{3}" -f $backfillResult.BackfilledBlocks, $phase2LogText, $backfillResult.PendingRows, $backfillResult.BackupPath)
        }
        elseif (-not $Config.DryRun -and $sourceUpdates.Count -gt 0) {
            [void] (Invoke-ExperimentImportCommit -Config $Config -SourceUpdates $sourceUpdates)
        }

        if ($backfillResult) {
            $phase2Text = ''
            if ($backfillResult.PSObject.Properties['PreviewNewBlocks']) {
                $phase2Text = '，新建块预览 {0} 个，自动新建 {1} 个，新建失败 {2} 个' -f $backfillResult.PreviewNewBlocks, $backfillResult.CreatedNewBlocks, $backfillResult.NewBlockFailures
            }
            $message = "没有发现需要重新导入的新实验文件夹；已继续检查自动导入表，回写已有块 $($backfillResult.BackfilledBlocks) 个$phase2Text，待确认 $($backfillResult.PendingRows) 条。"
        }
        else {
            $message = "没有发现主表缺失的新实验文件夹。"
        }
        $baselineCount = @($sourceUpdates | Where-Object { $_.ChangeKind -eq 'BaselineUnverified' }).Count
        if ($baselineCount) { $message += " 已为 $baselineCount 个历史实验建立源文件基线；基线不代表历史 Sa 已重新计算核验。" }
        $statusState = 'Completed'
        if ($discoveryErrors.Count -gt 0) { $statusState = 'Error'; $message += " 有 $($discoveryErrors.Count) 个实验目录检查失败，详见日志；原有数据保留。" }
        if ($Config.DryRun -and $discoveryErrors.Count -eq 0) { $statusState = 'DryRun' }
        Write-ExperimentLog -Config $Config -Message $message
        if ($statusCommand) {
            Update-PolishingSentinelStatus -ScriptRoot $Config.ScriptRoot -State $statusState -Message $message -Data @{
                CandidateFolders = $folders.Count
                NewFolders       = 0
                ImportedRows     = 0
                FailedFolders    = $discoveryErrors.Count
                Errors           = $discoveryErrors
                BackfilledBlocks = if ($backfillResult) { $backfillResult.BackfilledBlocks } else { 0 }
                PendingRows      = if ($backfillResult) { $backfillResult.PendingRows } else { 0 }
                NewBlockMode     = if ($backfillResult) { $backfillResult.NewBlockMode } else { '' }
                PreviewNewBlocks = if ($backfillResult -and $backfillResult.PSObject.Properties['PreviewNewBlocks']) { $backfillResult.PreviewNewBlocks } else { 0 }
                CreatedNewBlocks = if ($backfillResult -and $backfillResult.PSObject.Properties['CreatedNewBlocks']) { $backfillResult.CreatedNewBlocks } else { 0 }
                NewBlockFailures = if ($backfillResult -and $backfillResult.PSObject.Properties['NewBlockFailures']) { $backfillResult.NewBlockFailures } else { 0 }
            } | Out-Null
        }
        $result = [pscustomobject]@{
            DriveFound       = $true
            CandidateFolders = $folders.Count
            NewFolders       = 0
            ImportedRows     = 0
            CompletedFolders = 0
            FailedFolders    = $discoveryErrors.Count
            DeferredFolders  = 0
            Errors           = $discoveryErrors
            StatusState      = $statusState
            BackfilledBlocks = if ($backfillResult) { $backfillResult.BackfilledBlocks } else { 0 }
            PendingRows      = if ($backfillResult) { $backfillResult.PendingRows } else { 0 }
            NewBlockMode     = if ($backfillResult) { $backfillResult.NewBlockMode } else { '' }
            PreviewNewBlocks = if ($backfillResult -and $backfillResult.PSObject.Properties['PreviewNewBlocks']) { $backfillResult.PreviewNewBlocks } else { 0 }
            CreatedNewBlocks = if ($backfillResult -and $backfillResult.PSObject.Properties['CreatedNewBlocks']) { $backfillResult.CreatedNewBlocks } else { 0 }
            NewBlockFailures = if ($backfillResult -and $backfillResult.PSObject.Properties['NewBlockFailures']) { $backfillResult.NewBlockFailures } else { 0 }
            BackfillBackup   = if ($backfillResult) { $backfillResult.BackupPath } else { '' }
        }
        return $result
    }

    $baselines = @($sourceUpdates | Where-Object { $_.ChangeKind -eq 'BaselineUnverified' })
    if (-not $Config.DryRun -and $baselines.Count -gt 0) {
        [void](Invoke-ExperimentImportCommit -Config $Config -SourceUpdates $baselines)
    }
    $work = Invoke-ExperimentWorkItems -Config $Config -Items $newFolders
    $work.FailedFolders += $discoveryErrors.Count
    $work.Errors = @($discoveryErrors) + @($work.Errors)
    $imported = $work.Imported
    $backfillResult = $work.Backfill

    if ($backfillResult) {
        $phase2Text = ''
        if ($backfillResult.PSObject.Properties['PreviewNewBlocks']) {
            $phase2Text = '，新建块预览 {0} 个，自动新建 {1} 个，新建失败 {2} 个' -f $backfillResult.PreviewNewBlocks, $backfillResult.CreatedNewBlocks, $backfillResult.NewBlockFailures
        }
        $message = "已处理 $($work.CompletedFolders) 个实验文件夹，导入 $imported 行 Sa 数据；回写已有块 $($backfillResult.BackfilledBlocks) 个$phase2Text，待确认 $($backfillResult.PendingRows) 条。"
    }
    else {
        $message = "已处理 $($work.CompletedFolders) 个实验文件夹，导入 $imported 行 Sa 数据。"
    }
    $statusState = 'Completed'
    if ($work.FailedFolders -gt 0 -or $work.DeferredFolders -gt 0) {
        $statusState = if ($work.CompletedFolders -gt 0) { 'PartialSuccess' } elseif ($work.FailedFolders -gt 0) { 'Error' } else { 'RetryPending' }
        $message += " 失败 $($work.FailedFolders) 个，等待重试或检查 $($work.DeferredFolders) 个；成功结果已保留。"
        if ($work.SharedFailure) { $message += ' 主表提交受阻，本轮已停止后续计算。' }
    }
    if ($work.StateWarnings -gt 0) {
        $statusState = if ($work.CompletedFolders -gt 0) { 'PartialSuccess' } else { 'Error' }
        $message += " 重试状态保存异常 $($work.StateWarnings) 项，请检查日志。"
    }
    if ($work.TimedOutFolders -gt 0) {
        $message += " 其中 $($work.TimedOutFolders) 个处理超时，未完成的写入未提交。"
        if ($work.CompletedFolders -eq 0 -and $work.TimedOutFolders -eq $work.FailedFolders) { $statusState = 'TimedOut' }
    }
    if ($work.ConfirmationFolders -gt 0) {
        $message += " 编号冲突待确认 $($work.ConfirmationFolders) 个；请在待确认与日志页处理编号冲突。"
        if ($work.CompletedFolders -eq 0 -and $work.FailedFolders -eq 0) { $statusState = 'NeedsConfirmation' }
    }
    if ($work.SaWarnings -gt 0) {
        $message += " 缺测、无效或已确认冲突位置共 $($work.SaWarnings) 个，已保留空位；详情见 Sa 汇总及日志。"
    }
    if ($Config.DryRun) { $statusState = 'DryRun'; $message = "预检查发现 $($newFolders.Count) 个待处理实验；未运行 MATLAB、未写入主表。" }
    if ($statusCommand) {
        Update-PolishingSentinelStatus -ScriptRoot $Config.ScriptRoot -State $statusState -Message $message -Data @{
            CandidateFolders = $folders.Count
            NewFolders       = $newFolders.Count
            ImportedRows     = $imported
            CompletedFolders = $work.CompletedFolders
            TotalFolders     = ($newFolders.Count + $discoveryErrors.Count)
            FinishedFolders  = ($work.CompletedFolders + $work.FailedFolders + $work.DeferredFolders)
            FailedFolders    = $work.FailedFolders
            DeferredFolders  = $work.DeferredFolders
            StateWarnings    = $work.StateWarnings
            TimedOutFolders  = $work.TimedOutFolders
            ConfirmationFolders = $work.ConfirmationFolders
            SaWarnings       = $work.SaWarnings
            Errors           = $work.Errors
            BackfilledBlocks = if ($backfillResult) { $backfillResult.BackfilledBlocks } else { 0 }
            PendingRows      = if ($backfillResult) { $backfillResult.PendingRows } else { 0 }
            NewBlockMode     = if ($backfillResult) { $backfillResult.NewBlockMode } else { '' }
            PreviewNewBlocks = if ($backfillResult -and $backfillResult.PSObject.Properties['PreviewNewBlocks']) { $backfillResult.PreviewNewBlocks } else { 0 }
            CreatedNewBlocks = if ($backfillResult -and $backfillResult.PSObject.Properties['CreatedNewBlocks']) { $backfillResult.CreatedNewBlocks } else { 0 }
            NewBlockFailures = if ($backfillResult -and $backfillResult.PSObject.Properties['NewBlockFailures']) { $backfillResult.NewBlockFailures } else { 0 }
        } | Out-Null
    }
    Write-ExperimentLog -Config $Config -Message $message
    if ($work.CompletedFolders -gt 0 -or $work.FailedFolders -gt 0 -or $work.NewConfirmations -gt 0) { Show-ExperimentNotification -Message $message }

    $result = [pscustomobject]@{
        DriveFound       = $true
        CandidateFolders = $folders.Count
        NewFolders       = $newFolders.Count
        ImportedRows     = $imported
        CompletedFolders = $work.CompletedFolders
        FailedFolders    = $work.FailedFolders
        DeferredFolders  = $work.DeferredFolders
        StateWarnings    = $work.StateWarnings
        TimedOutFolders  = $work.TimedOutFolders
        Errors           = $work.Errors
        StatusState      = $statusState
        BackfilledBlocks = if ($backfillResult) { $backfillResult.BackfilledBlocks } else { 0 }
        PendingRows      = if ($backfillResult) { $backfillResult.PendingRows } else { 0 }
        NewBlockMode     = if ($backfillResult) { $backfillResult.NewBlockMode } else { '' }
        PreviewNewBlocks = if ($backfillResult -and $backfillResult.PSObject.Properties['PreviewNewBlocks']) { $backfillResult.PreviewNewBlocks } else { 0 }
        CreatedNewBlocks = if ($backfillResult -and $backfillResult.PSObject.Properties['CreatedNewBlocks']) { $backfillResult.CreatedNewBlocks } else { 0 }
        NewBlockFailures = if ($backfillResult -and $backfillResult.PSObject.Properties['NewBlockFailures']) { $backfillResult.NewBlockFailures } else { 0 }
        BackfillBackup   = if ($backfillResult) { $backfillResult.BackupPath } else { '' }
    }
    return $result
}
