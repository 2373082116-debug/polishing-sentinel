Set-StrictMode -Version 2.0

function New-ExperimentBackfillConfig {
    param(
        [Parameter(Mandatory = $true)] [string] $ScriptRoot,
        [Parameter(Mandatory = $true)] [string] $MasterWorkbookPath,
        [string] $AutoSheetName = '自动导入Sa数据',
        [string] $PendingSheetName = '待确认补齐',
        [string] $BackupDirectory = '',
        [AllowNull()] [object] $SlurryMappings = $null,
        [AllowNull()] [object] $ProcessDefaults = $null,
        [string] $NewBlockMode = 'Preview'
    )

    if ([string]::IsNullOrWhiteSpace($BackupDirectory)) {
        $BackupDirectory = Join-Path $ScriptRoot 'backups'
    }
    if ($null -eq $SlurryMappings) {
        $SlurryMappings = [pscustomobject]@{
            '金刚石' = '油基抛光'
            '硅溶胶' = '探索实验'
        }
    }
    if ($null -eq $ProcessDefaults) {
        $ProcessDefaults = [pscustomobject]@{
            Water = [pscustomobject]@{ Plate = 'PEEK'; Track = '4'; Spacing = '9'; Fullness = '92%'; Pressure = '单球0.1N'; BallMatch = '0.8mm氧化锆球' }
            Oil   = [pscustomobject]@{ Plate = '铸铁抛光盘'; Track = '3'; Spacing = '9'; Fullness = '92%'; Pressure = '单球0.1N'; BallMatch = '0.8mm钢球' }
        }
    }
    $targetSheets = @('探索实验', '油基抛光')
    if ($SlurryMappings -is [System.Collections.IDictionary]) {
        $targetSheets = @($SlurryMappings.Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string] $_) } | Select-Object -Unique)
    }
    else {
        $targetSheets = @($SlurryMappings.PSObject.Properties | ForEach-Object { $_.Value } | Where-Object { -not [string]::IsNullOrWhiteSpace([string] $_) } | Select-Object -Unique)
    }
    if ($targetSheets.Count -eq 0) {
        $targetSheets = @('探索实验', '油基抛光')
    }
    if (@('Preview', 'Auto', 'Disabled') -notcontains $NewBlockMode) {
        $NewBlockMode = 'Preview'
    }

    [pscustomobject]@{
        ScriptRoot         = $ScriptRoot
        MasterWorkbookPath = $MasterWorkbookPath
        AutoSheetName      = $AutoSheetName
        PendingSheetName   = $PendingSheetName
        TargetSheets       = $targetSheets
        BackupDirectory    = $BackupDirectory
        SlurryMappings     = $SlurryMappings
        ProcessDefaults    = $ProcessDefaults
        NewBlockMode       = $NewBlockMode
    }
}

function Get-ExperimentObjectProperty {
    param(
        [AllowNull()] [object] $Object,
        [Parameter(Mandatory = $true)] [string] $Name,
        [AllowNull()] [object] $Default = $null
    )

    if ($null -eq $Object) {
        return $Default
    }
    if ($Object -is [System.Collections.IDictionary] -and $Object.Contains($Name)) {
        return $Object[$Name]
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $Default
    }

    return $property.Value
}

function Get-ExperimentProcessDefaults {
    param(
        [Parameter(Mandatory = $true)] [string] $SlurryType,
        [AllowNull()] [object] $SlurryMappings = $null,
        [AllowNull()] [object] $ProcessDefaults = $null
    )

    $defaultMappings = [pscustomobject]@{
        '金刚石' = '油基抛光'
        '硅溶胶' = '探索实验'
    }
    $targetSheetName = [string] (Get-ExperimentObjectProperty -Object $SlurryMappings -Name $SlurryType -Default '')
    if ([string]::IsNullOrWhiteSpace($targetSheetName)) {
        $targetSheetName = [string] (Get-ExperimentObjectProperty -Object $defaultMappings -Name $SlurryType -Default '')
    }

    if ([string]::IsNullOrWhiteSpace($targetSheetName)) {
        return [pscustomobject]@{
            TargetSheetName = ''
            Plate           = ''
            Track           = ''
            Spacing         = ''
            Fullness        = ''
            Pressure        = ''
            BallMatch       = ''
        }
    }

    $groupName = ''
    if ($SlurryType -eq '硅溶胶' -or $targetSheetName -eq '探索实验') {
        $groupName = 'Water'
    }
    elseif ($SlurryType -eq '金刚石' -or $targetSheetName -eq '油基抛光') {
        $groupName = 'Oil'
    }

    $defaultProcessDefaults = [pscustomobject]@{
        Water = [pscustomobject]@{ Plate = 'PEEK'; Track = '4'; Spacing = '9'; Fullness = '92%'; Pressure = '单球0.1N'; BallMatch = '0.8mm氧化锆球' }
        Oil   = [pscustomobject]@{ Plate = '铸铁抛光盘'; Track = '3'; Spacing = '9'; Fullness = '92%'; Pressure = '单球0.1N'; BallMatch = '0.8mm钢球' }
    }
    $defaultGroup = Get-ExperimentObjectProperty -Object $defaultProcessDefaults -Name $groupName -Default $null
    $configuredGroup = Get-ExperimentObjectProperty -Object $ProcessDefaults -Name $groupName -Default $null

    [pscustomobject]@{
        TargetSheetName = $targetSheetName
        Plate           = [string] (Get-ExperimentObjectProperty -Object $configuredGroup -Name 'Plate' -Default (Get-ExperimentObjectProperty -Object $defaultGroup -Name 'Plate' -Default ''))
        Track           = [string] (Get-ExperimentObjectProperty -Object $configuredGroup -Name 'Track' -Default (Get-ExperimentObjectProperty -Object $defaultGroup -Name 'Track' -Default ''))
        Spacing         = [string] (Get-ExperimentObjectProperty -Object $configuredGroup -Name 'Spacing' -Default (Get-ExperimentObjectProperty -Object $defaultGroup -Name 'Spacing' -Default ''))
        Fullness        = [string] (Get-ExperimentObjectProperty -Object $configuredGroup -Name 'Fullness' -Default (Get-ExperimentObjectProperty -Object $defaultGroup -Name 'Fullness' -Default ''))
        Pressure        = [string] (Get-ExperimentObjectProperty -Object $configuredGroup -Name 'Pressure' -Default (Get-ExperimentObjectProperty -Object $defaultGroup -Name 'Pressure' -Default ''))
        BallMatch       = [string] (Get-ExperimentObjectProperty -Object $configuredGroup -Name 'BallMatch' -Default (Get-ExperimentObjectProperty -Object $defaultGroup -Name 'BallMatch' -Default ''))
    }
}

function Get-ExperimentProcessFieldDefinitions {
    return @(
        [pscustomobject]@{ Key = 'Plate'; Label = '盘体'; Required = $true },
        [pscustomobject]@{ Key = 'Track'; Label = '滚道'; Required = $true },
        [pscustomobject]@{ Key = 'Spacing'; Label = '间距'; Required = $true },
        [pscustomobject]@{ Key = 'Fullness'; Label = '充盈度'; Required = $true },
        [pscustomobject]@{ Key = 'Pressure'; Label = '压力'; Required = $true },
        [pscustomobject]@{ Key = 'BallMatch'; Label = '配球'; Required = $true }
    )
}

function Find-ExperimentBlockLabelCell {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [int] $StartRow,
        [Parameter(Mandatory = $true)] [int] $StartColumn,
        [Parameter(Mandatory = $true)] [string] $Label
    )

    for ($r = $StartRow; $r -le ($StartRow + 5); $r++) {
        for ($c = $StartColumn; $c -le ($StartColumn + 9); $c++) {
            $text = ([string] $Worksheet.Cells.Item($r, $c).Text).Trim()
            if ($text -eq $Label) {
                return [pscustomobject]@{ Row = $r; Column = $c }
            }
        }
    }

    return $null
}

function Get-ExperimentBlockFieldCells {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [pscustomobject] $Block
    )

    $cells = @{}
    foreach ($field in Get-ExperimentProcessFieldDefinitions) {
        $labelCell = Find-ExperimentBlockLabelCell -Worksheet $Worksheet -StartRow $Block.StartRow -StartColumn $Block.StartColumn -Label $field.Label
        if ($null -ne $labelCell) {
            $cells[$field.Key] = [pscustomobject]@{ Row = $labelCell.Row; Column = ($labelCell.Column + 1) }
        }
    }

    return $cells
}

function ConvertTo-ExperimentBackfillCellText {
    param([AllowNull()] [object] $Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [System.IFormattable]) {
        return $Value.ToString($null, [System.Globalization.CultureInfo]::InvariantCulture)
    }

    return [string] $Value
}

function Parse-StructuredExperimentFolderName {
    param(
        [Parameter(Mandatory = $true)] [string] $FolderPath,
        [AllowNull()] [object] $SlurryMappings = $null,
        [AllowNull()] [object] $ProcessDefaults = $null
    )

    $trimmedPath = $FolderPath.TrimEnd('\')
    $folderName = Split-Path -Path $trimmedPath -Leaf
    $parts = @($folderName -split '-', 9)
    $issues = New-Object System.Collections.Generic.List[string]

    if ($parts.Count -lt 7) {
        $issues.Add('文件夹字段数量不足')
        return [pscustomobject]@{
            IsStructured    = $false
            FolderPath      = $trimmedPath
            FolderName      = $folderName
            ExperimentNo    = ''
            ExperimentDate  = ''
            Concentration   = ''
            ParticleSize    = ''
            SlurryType      = ''
            TargetSheetName = ''
            BallSource      = ''
            Duration        = ''
            PadType         = ''
            Note            = ''
            Defaults        = $null
            Issues          = @($issues)
        }
    }

    $experimentNo = $parts[0].Trim()
    if ($experimentNo -notmatch '^\d+$') {
        $issues.Add('实验号不是纯数字')
    }
    else {
        $experimentNo = $experimentNo -replace '^0+(?=\d)', ''
    }

    $experimentDate = Get-DateMatchFromText -Text $parts[1].Trim()
    if ([string]::IsNullOrWhiteSpace($experimentDate)) {
        $issues.Add('日期无法解析')
    }

    $combinedCondition = $parts[2].Trim()
    $combinedMatch = [regex]::Match($combinedCondition, '^(\d+(\.\d+)?%)(\d+(\.\d+)?(nm|um|μm|mm))$')
    $hasJoinedConcentrationParticle = $combinedMatch.Success

    if (-not $hasJoinedConcentrationParticle -and $parts.Count -lt 8) {
        $issues.Add('文件夹字段数量不足')
    }

    $concentration = $combinedCondition
    $particleSizeIndex = 3
    $slurryTypeIndex = 4
    $ballSourceIndex = 5
    $durationIndex = 6
    $padTypeIndex = 7
    $noteIndex = 8
    if ($hasJoinedConcentrationParticle) {
        $concentration = $combinedMatch.Groups[1].Value
        $particleSizeIndex = 2
        $slurryTypeIndex = 3
        $ballSourceIndex = 4
        $durationIndex = 5
        $padTypeIndex = 6
        $noteIndex = 7
    }

    if ($concentration -notmatch '^\d+(\.\d+)?%$') {
        $issues.Add('浓度字段格式异常')
    }

    $particleSize = ''
    if ($hasJoinedConcentrationParticle) {
        $particleSize = $combinedMatch.Groups[3].Value
    }
    elseif ($parts.Count -gt $particleSizeIndex) {
        $particleSize = $parts[$particleSizeIndex].Trim()
    }
    if ($particleSize -notmatch '^\d+(\.\d+)?(nm|um|μm|mm)$') {
        $issues.Add('粒径字段格式异常')
    }

    $slurryType = ''
    if ($parts.Count -gt $slurryTypeIndex) {
        $slurryType = $parts[$slurryTypeIndex].Trim()
    }
    $manufacturer = ''
    $supplierMatch = [regex]::Match($slurryType, '^(?:硅溶胶\s*[（(]\s*)?(FUJIMI|新国瑞升|国瑞升)\s*[）)]?$', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($supplierMatch.Success) {
        $manufacturer = $supplierMatch.Groups[1].Value
        $slurryType = '硅溶胶'
    }
    $defaults = Get-ExperimentProcessDefaults -SlurryType $slurryType -SlurryMappings $SlurryMappings -ProcessDefaults $ProcessDefaults
    if ([string]::IsNullOrWhiteSpace($defaults.TargetSheetName)) {
        $issues.Add('抛光液类型无法映射目标表')
    }

    $ballSource = ''
    if ($parts.Count -gt $ballSourceIndex) {
        $ballSource = $parts[$ballSourceIndex].Trim()
    }
    if ([string]::IsNullOrWhiteSpace($ballSource)) {
        $issues.Add('球胚字段为空')
    }

    $duration = ''
    if ($parts.Count -gt $durationIndex) {
        $duration = $parts[$durationIndex].Trim()
    }
    if ($duration -notmatch '^\d+(\.\d+)?h$') {
        $issues.Add('抛光时间字段格式异常')
    }

    $padType = ''
    if ($parts.Count -gt $padTypeIndex) {
        $padType = $parts[$padTypeIndex].Trim()
    }
    if ([string]::IsNullOrWhiteSpace($padType)) {
        $issues.Add('抛光垫字段为空')
    }

    $note = ''
    if ($parts.Count -gt $noteIndex) {
        $noteParts = @($parts[$noteIndex..($parts.Count - 1)])
        $note = ($noteParts -join '-').Trim()
    }
    if ($manufacturer -and $note.IndexOf($manufacturer, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        $note = if ([string]::IsNullOrWhiteSpace($note) -or $note -eq '无') { $manufacturer } else { $note + '；' + $manufacturer }
    }

    [pscustomobject]@{
        IsStructured    = $true
        FolderPath      = $trimmedPath
        FolderName      = $folderName
        ExperimentNo    = $experimentNo
        ExperimentDate  = $experimentDate
        Concentration   = $concentration
        ParticleSize    = $particleSize
        SlurryType      = $slurryType
        TargetSheetName = $defaults.TargetSheetName
        BallSource      = $ballSource
        Duration        = $duration
        PadType         = $padType
        Note            = $note
        Defaults        = $defaults
        Issues          = @($issues)
    }
}

function Get-ExperimentNumberFromTitle {
    param([AllowNull()] [string] $Title)

    if ([string]::IsNullOrWhiteSpace($Title)) {
        return ''
    }

    $match = [regex]::Match($Title, '^\s*(?:[A-Za-z]\s+)?实验\s*0*(\d+)(?!\d)')
    if ($match.Success) {
        return $match.Groups[1].Value
    }

    return ''
}

function Get-ExperimentBackfillBlocksFromWorkbook {
    param(
        [Parameter(Mandatory = $true)] [object] $Workbook,
        [Parameter(Mandatory = $true)] [string[]] $SheetNames
    )

    $blocks = @()
    foreach ($sheetName in $SheetNames) {
        $worksheet = $Workbook.Worksheets.Item($sheetName)
        $used = $worksheet.UsedRange
        $firstRow = [int] $used.Row
        $rowCount = [int] $used.Rows.Count
        $lastRow = $firstRow + $rowCount - 1
        $candidateColumns = @(2, 12, 22)

        for ($row = $firstRow; $row -le $lastRow; $row++) {
            foreach ($column in $candidateColumns) {
                $title = ([string] $worksheet.Cells.Item($row, $column).Text).Trim()
                if ([string]::IsNullOrWhiteSpace($title)) {
                    continue
                }
                if ([string]::IsNullOrWhiteSpace((Get-ExperimentNumberFromTitle -Title $title)) -or $title -notmatch '抛光' -or $title -match '记录表') {
                    continue
                }

                $saHeaderRow = 0
                $saStartColumn = 0
                $maxHeaderRow = [Math]::Min($row + 8, $lastRow)
                for ($scanRow = $row; $scanRow -le $maxHeaderRow -and $saHeaderRow -eq 0; $scanRow++) {
                    for ($scanColumn = $column; $scanColumn -le ($column + 9); $scanColumn++) {
                        $header = ([string] $worksheet.Cells.Item($scanRow, $scanColumn).Text).Trim()
                        if ($header -match '^Sa1') {
                            $saHeaderRow = $scanRow
                            $saStartColumn = $scanColumn
                            break
                        }
                    }
                }

                if ($saHeaderRow -eq 0) {
                    $blocks += [pscustomobject]@{
                        SheetName = $sheetName; StartRow = $row; StartColumn = $column
                        Title = $title; ExperimentNo = Get-ExperimentNumberFromTitle $title
                        ExperimentDate = Get-DateMatchFromText $title; SlurryText = $title
                        SaHeaderRow = 0; SampleStartRow = 0; SaStartColumn = 0
                        BatchAverageRow = $row + 11; BatchAverageColumn = $column + 8
                    }
                    continue
                }

                $batchAverageRow = 0
                $batchAverageColumn = $saStartColumn + 3
                $maxBatchRow = $lastRow
                for ($scanRow = $saHeaderRow + 1; $scanRow -le $maxBatchRow -and $batchAverageRow -eq 0; $scanRow++) {
                    $nextTitle = ([string] $worksheet.Cells.Item($scanRow, $column).Text).Trim()
                    if ($nextTitle -match '实验\s*\d+.*抛光') { break }
                    for ($scanColumn = $column; $scanColumn -le ($column + 9); $scanColumn++) {
                        $label = ([string] $worksheet.Cells.Item($scanRow, $scanColumn).Text).Trim()
                        if ($label -match '单批均值') {
                            $batchAverageRow = $scanRow
                            break
                        }
                    }
                }
                if ($batchAverageRow -eq 0) {
                    $batchAverageRow = $saHeaderRow + 6
                }

                $blockText = New-Object System.Collections.Generic.List[string]
                $maxTextRow = [Math]::Min($row + 5, $lastRow)
                for ($textRow = $row; $textRow -le $maxTextRow; $textRow++) {
                    for ($textColumn = $column; $textColumn -le ($column + 9); $textColumn++) {
                        $textValue = ([string] $worksheet.Cells.Item($textRow, $textColumn).Text).Trim()
                        if (-not [string]::IsNullOrWhiteSpace($textValue)) {
                            $blockText.Add($textValue)
                        }
                    }
                }

                $blocks += [pscustomobject]@{
                    SheetName          = $sheetName
                    StartRow           = $row
                    StartColumn        = $column
                    Title              = $title
                    ExperimentNo       = Get-ExperimentNumberFromTitle -Title $title
                    ExperimentDate     = Get-DateMatchFromText -Text $title
                    SlurryText         = ($blockText -join ' ')
                    SaHeaderRow        = $saHeaderRow
                    SampleStartRow     = $saHeaderRow + 1
                    SaStartColumn      = $saStartColumn
                    BatchAverageRow    = $batchAverageRow
                    BatchAverageColumn = $batchAverageColumn
                }
            }
        }
    }

    return $blocks
}

function Get-ExperimentBackfillBlocks {
    param(
        [Parameter(Mandatory = $true)] [string] $WorkbookPath,
        [Parameter(Mandatory = $true)] [string[]] $SheetNames
    )

    $session = $null; $excel = $null
    $workbook = $null
    try {
        $session = New-ExperimentExcelSession
        $excel = $session.Application
        $workbook = $excel.Workbooks.Open($WorkbookPath, $null, $true)
        return @(Get-ExperimentBackfillBlocksFromWorkbook -Workbook $workbook -SheetNames $SheetNames)
    }
    finally {
        Close-ExperimentExcelSession -Session $session -Workbook $workbook
    }
}

function Get-LatestTemplateBlock {
    param(
        [Parameter(Mandatory = $true)] [object[]] $Blocks,
        [Parameter(Mandatory = $true)] [string] $TargetSheetName
    )

    $candidates = @($Blocks | Where-Object {
        $_.SheetName -eq $TargetSheetName -and
        $_.ExperimentNo -match '^\d+$' -and
        [int] $_.SaHeaderRow -gt 0 -and
        [int] $_.SampleStartRow -gt 0 -and
        [int] $_.BatchAverageRow -gt 0
    } | Sort-Object { [int] $_.ExperimentNo } -Descending)

    if ($candidates.Count -eq 0) {
        return $null
    }

    return $candidates[0]
}

function Get-CommonExperimentBlockColumns {
    param([Parameter(Mandatory = $true)] [object[]] $Blocks)

    $columns = @($Blocks |
        Where-Object { $_.StartColumn -match '^\d+$' } |
        Group-Object StartColumn |
        Sort-Object { [int] $_.Name } |
        ForEach-Object { [int] $_.Name })
    if ($columns.Count -eq 0) {
        return @(2, 12, 22)
    }

    $standardColumns = @(2, 12, 22)
    $onlyStandardColumns = $true
    foreach ($column in $columns) {
        if ($standardColumns -notcontains [int] $column) {
            $onlyStandardColumns = $false
            break
        }
    }
    if ($onlyStandardColumns) {
        return $standardColumns
    }

    return $columns
}

function Get-ExperimentBlockRowStep {
    param([Parameter(Mandatory = $true)] [object[]] $Blocks)

    $rows = @($Blocks |
        Where-Object { $_.StartRow -match '^\d+$' } |
        Select-Object -ExpandProperty StartRow -Unique |
        Sort-Object { [int] $_ })
    if ($rows.Count -lt 2) {
        return 0
    }

    $steps = @()
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $steps += ([int] $rows[$i] - [int] $rows[$i - 1])
    }
    $stepGroups = @($steps | Group-Object | Sort-Object Count -Descending)
    if ($stepGroups.Count -gt 1 -and $stepGroups[0].Count -eq $stepGroups[1].Count) {
        return 0
    }

    return [int] $stepGroups[0].Name
}

function Get-NextExperimentBlockPosition {
    param(
        [Parameter(Mandatory = $true)] [object[]] $Blocks,
        [Parameter(Mandatory = $true)] [string] $TargetSheetName,
        [Parameter(Mandatory = $true)] [pscustomobject] $TemplateBlock
    )

    $sheetBlocks = @($Blocks | Where-Object { $_.SheetName -eq $TargetSheetName })
    $columns = @(Get-CommonExperimentBlockColumns -Blocks $sheetBlocks)
    $sortedColumns = @($columns | Sort-Object)
    $currentColumnIndex = [Array]::IndexOf($sortedColumns, [int] $TemplateBlock.StartColumn)

    if ($currentColumnIndex -ge 0 -and $currentColumnIndex -lt ($sortedColumns.Count - 1)) {
        return [pscustomobject]@{
            IsUsable       = $true
            StartRow       = [int] $TemplateBlock.StartRow
            StartColumn    = [int] $sortedColumns[$currentColumnIndex + 1]
            Reason         = '同一行组下一个列位'
            ColumnSequence = $sortedColumns
            RowStep        = 0
        }
    }

    $rowStep = Get-ExperimentBlockRowStep -Blocks $sheetBlocks
    if ($rowStep -le 0) {
        return [pscustomobject]@{
            IsUsable       = $false
            StartRow       = 0
            StartColumn    = 0
            Reason         = '无法推断下一行组行距'
            ColumnSequence = $sortedColumns
            RowStep        = $rowStep
        }
    }

    # Historical row spacing may predate the mandatory remark row.
    $nextRow = [int] $TemplateBlock.StartRow + $rowStep
    foreach ($block in @($sheetBlocks | Where-Object { $_.StartRow -eq $TemplateBlock.StartRow })) {
        $occupied = Get-ExperimentBlockRange -Block $block
        $nextRow = [Math]::Max($nextRow, $occupied.StartRow + $occupied.Rows + 1)
    }
    return [pscustomobject]@{
        IsUsable       = $true
        StartRow       = $nextRow
        StartColumn    = [int] $sortedColumns[0]
        Reason         = '下一行组第一个列位'
        ColumnSequence = $sortedColumns
        RowStep        = $rowStep
    }
}

function ConvertTo-ExperimentDisplayDate {
    param([Parameter(Mandatory = $true)] [string] $IsoDate)

    $dt = [datetime]::ParseExact($IsoDate, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    return ('{0}.{1}.{2}' -f $dt.Year, $dt.Month, $dt.Day)
}

function Get-ExperimentPolishingType {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Metadata)

    if ($Metadata.SlurryType -eq '金刚石') {
        return '油基抛光'
    }
    if ($Metadata.SlurryType -eq '硅溶胶') {
        return '水基抛光'
    }
    if ($Metadata.TargetSheetName -like '*油*') {
        return '油基抛光'
    }

    return '水基抛光'
}

function New-ExperimentBlockTitle {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Metadata)

    $polishingType = Get-ExperimentPolishingType -Metadata $Metadata
    return ('实验{0} {1}（{2}）{3}' -f $Metadata.ExperimentNo, $polishingType, (ConvertTo-ExperimentDisplayDate -IsoDate $Metadata.ExperimentDate), $Metadata.Duration)
}

function Get-ExperimentBlockRange {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Block)

    $rowCount = [Math]::Max(1, ([int] $Block.BatchAverageRow - [int] $Block.StartRow + 2))
    return [pscustomobject]@{
        StartRow    = [int] $Block.StartRow
        StartColumn = [int] $Block.StartColumn
        Rows        = $rowCount
        Columns     = 10
    }
}

function Update-ExperimentRemarkLabels {
    param([Parameter(Mandatory = $true)] $Workbook, [AllowEmptyCollection()] [object[]] $Blocks)
    foreach ($block in $Blocks) {
        if ([int] $block.SaHeaderRow -le 0) { continue }
        $sheet = $Workbook.Worksheets.Item($block.SheetName)
        $row = [int] $block.BatchAverageRow + 1
        $label = ([string] $sheet.Cells.Item($row, $block.StartColumn).Text).Trim()
        if ($label -eq '小结') {
            Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $sheet -Row $row -Column $block.StartColumn -Value '备注' -FieldName '备注标签')
        }
    }
}

function Test-ExperimentBlockRangeOverlap {
    param(
        [Parameter(Mandatory = $true)] [int] $StartRow,
        [Parameter(Mandatory = $true)] [int] $StartColumn,
        [Parameter(Mandatory = $true)] [int] $Rows,
        [Parameter(Mandatory = $true)] [int] $Columns,
        [AllowEmptyCollection()]
        [Parameter(Mandatory = $true)] [object[]] $ExistingBlocks
    )

    $endRow = $StartRow + $Rows - 1
    $endColumn = $StartColumn + $Columns - 1
    foreach ($block in $ExistingBlocks) {
        $range = Get-ExperimentBlockRange -Block $block
        $blockEndRow = $range.StartRow + $range.Rows - 1
        $blockEndColumn = $range.StartColumn + $range.Columns - 1
        $separate = ($endRow -lt $range.StartRow) -or ($StartRow -gt $blockEndRow) -or ($endColumn -lt $range.StartColumn) -or ($StartColumn -gt $blockEndColumn)
        if (-not $separate) {
            return $true
        }
    }
    return $false
}

function Test-NewBlockTargetRange {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [int] $TargetStartRow,
        [Parameter(Mandatory = $true)] [int] $TargetStartColumn,
        [Parameter(Mandatory = $true)] [pscustomobject] $TemplateBlock,
        [AllowEmptyCollection()]
        [Parameter(Mandatory = $true)] [object[]] $ExistingBlocks
    )

    $templateRange = Get-ExperimentBlockRange -Block $TemplateBlock
    $rows = $templateRange.Rows
    $columns = $templateRange.Columns
    $address = ('R{0}C{1}:R{2}C{3}' -f $TargetStartRow, $TargetStartColumn, ($TargetStartRow + $rows - 1), ($TargetStartColumn + $columns - 1))

    $endRow = [long] $TargetStartRow + $rows - 1
    $endColumn = [long] $TargetStartColumn + $columns - 1
    if ($TargetStartRow -lt 1 -or $TargetStartColumn -lt 1 -or $endRow -gt 1048576 -or $endColumn -gt 16384) {
        return [pscustomobject]@{ IsSafe = $false; Reason = '目标区域超出 Excel 行列范围'; Address = $address; Rows = $rows; Columns = $columns }
    }

    if (Test-ExperimentBlockRangeOverlap -StartRow $TargetStartRow -StartColumn $TargetStartColumn -Rows $rows -Columns $columns -ExistingBlocks $ExistingBlocks) {
        return [pscustomobject]@{ IsSafe = $false; Reason = '目标区域与已有实验块重叠'; Address = $address; Rows = $rows; Columns = $columns }
    }

    for ($r = $TargetStartRow; $r -le ($TargetStartRow + $rows - 1); $r++) {
        for ($c = $TargetStartColumn; $c -le ($TargetStartColumn + $columns - 1); $c++) {
            $cell = $Worksheet.Cells.Item($r, $c)
            if ($cell.MergeCells) {
                $merge = $cell.MergeArea
                if ($merge.Row -lt $TargetStartRow -or $merge.Column -lt $TargetStartColumn -or
                    ([long] $merge.Row + $merge.Rows.Count - 1) -gt $endRow -or
                    ([long] $merge.Column + $merge.Columns.Count - 1) -gt $endColumn) {
                    return [pscustomobject]@{ IsSafe = $false; Reason = '目标区域包含跨边界合并单元格，需要确认'; Address = $address; Rows = $rows; Columns = $columns }
                }
            }
            if ($cell.HasFormula) {
                return [pscustomobject]@{ IsSafe = $false; Reason = '目标区域已有公式，需要确认'; Address = $address; Rows = $rows; Columns = $columns }
            }
            $text = ([string] $cell.Value2).Trim()
            if ($null -ne $cell.Value2 -and $text.Length -eq 0) {
                return [pscustomobject]@{ IsSafe = $false; Reason = '目标区域包含空白文本，需要确认'; Address = $address; Rows = $rows; Columns = $columns }
            }
            if ([string]::IsNullOrWhiteSpace($text)) {
                continue
            }
            if ($text -match '实验\s*\d+') {
                return [pscustomobject]@{ IsSafe = $false; Reason = ('目标区域已有实验标题：{0}' -f $text); Address = $address; Rows = $rows; Columns = $columns }
            }
            return [pscustomobject]@{ IsSafe = $false; Reason = ('目标区域存在非空内容：{0}' -f $text); Address = $address; Rows = $rows; Columns = $columns }
        }
    }

    return [pscustomobject]@{ IsSafe = $true; Reason = '目标区域为空'; Address = $address; Rows = $rows; Columns = $columns }
}

function New-ExperimentBlockFieldSummary {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Metadata)

    return ('实验号={0}; 日期={1}; 浓度={2}; 粒径={3}; 抛光液={4}; 球胚={5}; 时间={6}; 抛光垫={7}; 备注={8}' -f
        $Metadata.ExperimentNo,
        $Metadata.ExperimentDate,
        $Metadata.Concentration,
        $Metadata.ParticleSize,
        $Metadata.SlurryType,
        $Metadata.BallSource,
        $Metadata.Duration,
        $Metadata.PadType,
        $Metadata.Note)
}

function Resolve-NewExperimentBlockPlan {
    param(
        [Parameter(Mandatory = $true)] [pscustomobject] $Metadata,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]] $Blocks,
        [AllowNull()] [object] $Workbook,
        [Parameter(Mandatory = $true)] [pscustomobject] $Config
    )

    $mode = [string] (Get-ExperimentObjectProperty -Object $Config -Name 'NewBlockMode' -Default 'Preview')
    if ([string]::IsNullOrWhiteSpace($mode)) {
        $mode = 'Preview'
    }

    if ($mode -eq 'Disabled') {
        return [pscustomobject]@{ Action = 'QueuePending'; Reason = '阶段二已关闭'; Mode = $mode; Metadata = $Metadata }
    }

    if (-not $Metadata.IsStructured) {
        return [pscustomobject]@{ Action = 'QueuePending'; Reason = '文件夹命名不符合固定字段约定'; Mode = $mode; Metadata = $Metadata }
    }

    if ($Metadata.Issues.Count -gt 0) {
        return [pscustomobject]@{ Action = 'QueuePending'; Reason = ($Metadata.Issues -join '；'); Mode = $mode; Metadata = $Metadata }
    }

    $existing = @($Blocks | Where-Object { $_.SheetName -eq $Metadata.TargetSheetName -and $_.ExperimentNo -eq $Metadata.ExperimentNo })
    if ($existing.Count -gt 0) {
        return [pscustomobject]@{ Action = 'QueuePending'; Reason = '目标表已存在同实验号'; Mode = $mode; Metadata = $Metadata }
    }

    $template = Get-LatestTemplateBlock -Blocks $Blocks -TargetSheetName $Metadata.TargetSheetName
    if ($null -eq $template) {
        return [pscustomobject]@{ Action = 'QueuePending'; Reason = '无法找到有效模板块'; Mode = $mode; Metadata = $Metadata }
    }

    $position = Get-NextExperimentBlockPosition -Blocks $Blocks -TargetSheetName $Metadata.TargetSheetName -TemplateBlock $template
    if (-not $position.IsUsable) {
        return [pscustomobject]@{ Action = 'QueuePending'; Reason = $position.Reason; Mode = $mode; Metadata = $Metadata; TemplateBlock = $template }
    }

    $range = Get-ExperimentBlockRange -Block $template
    $targetAddress = ('R{0}C{1}:R{2}C{3}' -f $position.StartRow, $position.StartColumn, ($position.StartRow + $range.Rows - 1), ($position.StartColumn + $range.Columns - 1))
    $safety = [pscustomobject]@{ IsSafe = $true; Reason = '未执行目标区域检查'; Address = $targetAddress; Rows = $range.Rows; Columns = $range.Columns }

    if ($null -ne $Workbook) {
        $worksheet = $Workbook.Worksheets.Item($Metadata.TargetSheetName)
        $targetSheetBlocks = @($Blocks | Where-Object { $_.SheetName -eq $Metadata.TargetSheetName })
        $safety = Test-NewBlockTargetRange -Worksheet $worksheet -TargetStartRow $position.StartRow -TargetStartColumn $position.StartColumn -TemplateBlock $template -ExistingBlocks $targetSheetBlocks
        if (-not $safety.IsSafe) {
            return [pscustomobject]@{
                Action             = 'QueuePending'
                Reason             = $safety.Reason
                Mode               = $mode
                TargetSheetName    = $Metadata.TargetSheetName
                TemplateBlock      = $template
                TargetStartRow     = [int] $position.StartRow
                TargetStartColumn  = [int] $position.StartColumn
                TargetRangeAddress = $safety.Address
                Title              = New-ExperimentBlockTitle -Metadata $Metadata
                FieldSummary       = New-ExperimentBlockFieldSummary -Metadata $Metadata
                Safety             = $safety
                Metadata           = $Metadata
            }
        }
    }

    return [pscustomobject]@{
        Action             = if ($mode -eq 'Auto') { 'CreateNewBlock' } else { 'PreviewNewBlock' }
        Reason             = if ($mode -eq 'Auto') { '自动模式，等待创建新实验块' } else { '预览模式' }
        Mode               = $mode
        TargetSheetName    = $Metadata.TargetSheetName
        TemplateBlock      = $template
        TargetStartRow     = [int] $position.StartRow
        TargetStartColumn  = [int] $position.StartColumn
        TargetRangeAddress = $targetAddress
        Title              = New-ExperimentBlockTitle -Metadata $Metadata
        FieldSummary       = New-ExperimentBlockFieldSummary -Metadata $Metadata
        Safety             = $safety
        Metadata           = $Metadata
    }
}

function Resolve-ExperimentBackfillDecision {
    param(
        [Parameter(Mandatory = $true)] [pscustomobject] $Metadata,
        [Parameter(Mandatory = $true)] [object[]] $Blocks
    )

    if (-not $Metadata.IsStructured) {
        return [pscustomobject]@{
            Action          = 'QueuePending'
            Reason          = '文件夹命名不符合固定字段约定'
            TargetSheetName = ''
            TargetBlock     = $null
            Confidence      = 0
        }
    }

    if ($Metadata.Issues.Count -gt 0) {
        return [pscustomobject]@{
            Action          = 'QueuePending'
            Reason          = ($Metadata.Issues -join '；')
            TargetSheetName = $Metadata.TargetSheetName
            TargetBlock     = $null
            Confidence      = 0
        }
    }

    $matches = @($Blocks | Where-Object {
        $_.SheetName -eq $Metadata.TargetSheetName -and
        $_.ExperimentNo -eq $Metadata.ExperimentNo
    })

    if ($matches.Count -eq 1) {
        $issue = ''
        if ((Get-ExperimentObjectProperty $matches[0] 'SaHeaderRow' 1) -le 0) {
            $issue = '已找到实验标题，但 Sa 表头缺失，无法安全补齐'
        }
        $blockDate = Get-ExperimentObjectProperty $matches[0] 'ExperimentDate' ''
        if ($blockDate -and $blockDate -ne $Metadata.ExperimentDate) {
            $issue = '实验号相同但日期不同，需要确认目标实验块'
        }
        if ($issue) {
            return [pscustomobject]@{ Action = 'QueuePending'; Reason = $issue; TargetSheetName = $Metadata.TargetSheetName; TargetBlock = $matches[0]; Confidence = 0 }
        }
        return [pscustomobject]@{
            Action          = 'BackfillExisting'
            Reason          = '实验号唯一匹配已有实验块'
            TargetSheetName = $matches[0].SheetName
            TargetBlock     = $matches[0]
            Confidence      = 100
        }
    }

    if ($matches.Count -gt 1) {
        return [pscustomobject]@{
            Action          = 'QueuePending'
            Reason          = '目标实验号匹配到多个块'
            TargetSheetName = $Metadata.TargetSheetName
            TargetBlock     = $null
            Confidence      = 0
        }
    }

    [pscustomobject]@{
        Action          = 'QueuePending'
        Reason          = '阶段一不自动创建新实验块'
        TargetSheetName = $Metadata.TargetSheetName
        TargetBlock     = $null
        Confidence      = 0
    }
}

function New-ExperimentBackfillPlan {
    param(
        [Parameter(Mandatory = $true)] [pscustomobject] $Metadata,
        [Parameter(Mandatory = $true)] [pscustomobject] $Decision,
        [Parameter(Mandatory = $true)] [pscustomobject] $Group,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]] $Blocks,
        [AllowNull()] [object] $Workbook,
        [Parameter(Mandatory = $true)] [pscustomobject] $Config
    )

    if ($Decision.Action -eq 'QueuePending' -and $Decision.Reason -eq '阶段一不自动创建新实验块') {
        $plan = Resolve-NewExperimentBlockPlan -Metadata $Metadata -Blocks $Blocks -Workbook $Workbook -Config $Config
        $plan | Add-Member -NotePropertyName Metadata -NotePropertyValue $Metadata -Force
        $plan | Add-Member -NotePropertyName Group -NotePropertyValue $Group -Force
        return $plan
    }

    return [pscustomobject]@{
        Action          = $Decision.Action
        Reason          = $Decision.Reason
        TargetSheetName = $Decision.TargetSheetName
        TargetBlock     = $Decision.TargetBlock
        Confidence      = $Decision.Confidence
        Metadata        = $Metadata
        Group           = $Group
    }
}

function Ensure-BackfillStatusColumns {
    param([Parameter(Mandatory = $true)] [object] $Worksheet)

    $headers = @('回写状态', '目标工作表', '目标块起始单元格', '回写时间', '回写备注')
    $used = $Worksheet.UsedRange
    $columnCount = [int] $used.Columns.Count
    if ($columnCount -lt 1) {
        $columnCount = 1
    }

    $columns = @{}
    for ($column = 1; $column -le [Math]::Max($columnCount, 18); $column++) {
        $header = ([string] $Worksheet.Cells.Item(1, $column).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($header) -and -not $columns.ContainsKey($header)) {
            $columns[$header] = $column
        }
    }

    $nextColumn = [Math]::Max($columnCount, 18) + 1
    foreach ($header in $headers) {
        if (-not $columns.ContainsKey($header)) {
            $Worksheet.Cells.Item(1, $nextColumn).Value2 = $header
            $columns[$header] = $nextColumn
            $nextColumn++
        }
    }

    $Worksheet.Rows.Item(1).Font.Bold = $true
    return $columns
}

function Ensure-PendingBackfillWorksheet {
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
        '记录时间',
        '原因',
        '来源文件夹',
        '目标工作表',
        '实验号',
        '日期',
        '浓度',
        '粒径',
        '抛光液类型',
        '球胚',
        '抛光时间',
        '抛光垫',
        '备注',
        '固定参数',
        'Sa数据',
        '目标块',
        '置信度'
    )

    for ($i = 0; $i -lt $headers.Count; $i++) {
        if ([string]::IsNullOrWhiteSpace([string] $worksheet.Cells.Item(1, $i + 1).Value2)) {
            $worksheet.Cells.Item(1, $i + 1).Value2 = $headers[$i]
        }
    }
    $worksheet.Rows.Item(1).Font.Bold = $true
    return $worksheet
}

function Convert-SaRowsToReviewText {
    param([Parameter(Mandatory = $true)] [object[]] $SaRows)

    $parts = @()
    foreach ($row in $SaRows) {
        $parts += ('球{0}: Sa1={1}, Sa2={2}, Sa3={3}, 均值={4}' -f
            $row.SampleIndex,
            (ConvertTo-ExperimentBackfillCellText $row.Sa1),
            (ConvertTo-ExperimentBackfillCellText $row.Sa2),
            (ConvertTo-ExperimentBackfillCellText $row.Sa3),
            (ConvertTo-ExperimentBackfillCellText $row.SaMean))
    }
    return ($parts -join '; ')
}

function Append-PendingBackfillRecord {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [pscustomobject] $Metadata,
        [Parameter(Mandatory = $true)] [pscustomobject] $Decision,
        [Parameter(Mandatory = $true)] [object[]] $SaRows
    )

    $used = $Worksheet.UsedRange
    $nextRow = [int] $used.Rows.Count + 1
    if ($nextRow -lt 2) {
        $nextRow = 2
    }

    $defaultsText = ''
    if ($null -ne $Metadata.Defaults) {
        $defaultsText = '盘体={0}; 滚道={1}; 间距={2}; 充盈度={3}; 压力={4}' -f
            $Metadata.Defaults.Plate,
            $Metadata.Defaults.Track,
            $Metadata.Defaults.Spacing,
            $Metadata.Defaults.Fullness,
            $Metadata.Defaults.Pressure
    }

    $targetBlock = ''
    if ($null -ne $Decision.TargetBlock) {
        $targetBlock = 'R{0}C{1}' -f $Decision.TargetBlock.StartRow, $Decision.TargetBlock.StartColumn
    }

    $values = @(
        (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),
        $Decision.Reason,
        $Metadata.FolderPath,
        $Decision.TargetSheetName,
        $Metadata.ExperimentNo,
        $Metadata.ExperimentDate,
        $Metadata.Concentration,
        $Metadata.ParticleSize,
        $Metadata.SlurryType,
        $Metadata.BallSource,
        $Metadata.Duration,
        $Metadata.PadType,
        $Metadata.Note,
        $defaultsText,
        (Convert-SaRowsToReviewText -SaRows $SaRows),
        $targetBlock,
        $Decision.Confidence
    )

    for ($column = 1; $column -le $values.Count; $column++) {
        $Worksheet.Cells.Item($nextRow, $column).Value2 = ConvertTo-ExperimentBackfillCellText $values[$column - 1]
    }

    return 1
}

function Ensure-NewBlockPreviewColumns {
    param([Parameter(Mandatory = $true)] [object] $Worksheet)

    $headers = @('建议动作', '模板工作表', '模板块坐标', '建议新块坐标', '拟写标题', '拟写字段摘要', '目标区域检查', '确认状态', '处理时间')
    $used = $Worksheet.UsedRange
    $columns = @{}
    for ($column = 1; $column -le [int] $used.Columns.Count; $column++) {
        $header = ([string] $Worksheet.Cells.Item(1, $column).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($header) -and -not $columns.ContainsKey($header)) {
            $columns[$header] = $column
        }
    }

    $nextColumn = [int] $used.Columns.Count + 1
    foreach ($header in $headers) {
        if (-not $columns.ContainsKey($header)) {
            $Worksheet.Cells.Item(1, $nextColumn).Value2 = $header
            $columns[$header] = $nextColumn
            $nextColumn++
        }
    }

    $Worksheet.Rows.Item(1).Font.Bold = $true
    return $columns
}

function Append-NewBlockPreviewRecord {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [pscustomobject] $Plan,
        [Parameter(Mandatory = $true)] [object[]] $SaRows
    )

    $columns = Ensure-NewBlockPreviewColumns -Worksheet $Worksheet
    $decision = [pscustomobject]@{
        Action          = 'QueuePending'
        Reason          = $Plan.Reason
        TargetSheetName = $Plan.TargetSheetName
        TargetBlock     = $Plan.TemplateBlock
        Confidence      = 0
    }

    [void] (Append-PendingBackfillRecord -Worksheet $Worksheet -Metadata $Plan.Metadata -Decision $decision -SaRows $SaRows)
    $row = [int] $Worksheet.UsedRange.Rows.Count

    $templateSheet = ''
    $templateCell = ''
    if ($null -ne $Plan.TemplateBlock) {
        $templateSheet = [string] $Plan.TemplateBlock.SheetName
        $templateCell = 'R{0}C{1}' -f $Plan.TemplateBlock.StartRow, $Plan.TemplateBlock.StartColumn
    }

    $safetyText = '预览模式未写入'
    if ($Plan.PSObject.Properties['Safety'] -and $null -ne $Plan.Safety -and -not [string]::IsNullOrWhiteSpace([string] $Plan.Safety.Reason)) {
        $safetyText = [string] $Plan.Safety.Reason
    }

    $Worksheet.Cells.Item($row, $columns['建议动作']).Value2 = '可自动建块'
    $Worksheet.Cells.Item($row, $columns['模板工作表']).Value2 = $templateSheet
    $Worksheet.Cells.Item($row, $columns['模板块坐标']).Value2 = $templateCell
    $Worksheet.Cells.Item($row, $columns['建议新块坐标']).Value2 = $Plan.TargetRangeAddress
    $Worksheet.Cells.Item($row, $columns['拟写标题']).Value2 = $Plan.Title
    $Worksheet.Cells.Item($row, $columns['拟写字段摘要']).Value2 = $Plan.FieldSummary
    $Worksheet.Cells.Item($row, $columns['目标区域检查']).Value2 = $safetyText
    $Worksheet.Cells.Item($row, $columns['确认状态']).Value2 = '未确认'
    $Worksheet.Cells.Item($row, $columns['处理时间']).Value2 = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    return 1
}

function Convert-BlockToTargetBlock {
    param(
        [Parameter(Mandatory = $true)] [pscustomobject] $TemplateBlock,
        [Parameter(Mandatory = $true)] [pscustomobject] $Plan
    )

    $rowOffset = [int] $Plan.TargetStartRow - [int] $TemplateBlock.StartRow
    $columnOffset = [int] $Plan.TargetStartColumn - [int] $TemplateBlock.StartColumn
    $capacity = [int] (Get-ExperimentObjectProperty -Object $Plan -Name 'SampleCapacity' -Default ([int] $TemplateBlock.BatchAverageRow - [int] $TemplateBlock.SampleStartRow))
    return [pscustomobject]@{
        SheetName           = $Plan.TargetSheetName
        StartRow            = [int] $Plan.TargetStartRow
        StartColumn         = [int] $Plan.TargetStartColumn
        Title               = $Plan.Title
        ExperimentNo        = $Plan.Metadata.ExperimentNo
        SaHeaderRow         = [int] $TemplateBlock.SaHeaderRow + $rowOffset
        SampleStartRow      = [int] $TemplateBlock.SampleStartRow + $rowOffset
        SaStartColumn       = [int] $TemplateBlock.SaStartColumn + $columnOffset
        BatchAverageRow     = [int] $TemplateBlock.SampleStartRow + $rowOffset + $capacity
        BatchAverageColumn  = [int] $TemplateBlock.BatchAverageColumn + $columnOffset
    }
}

function Clear-NewExperimentBlockVariableData {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [pscustomobject] $Plan,
        [Parameter(Mandatory = $true)] [pscustomobject] $TemplateBlock
    )

    $targetBlock = Convert-BlockToTargetBlock -TemplateBlock $TemplateBlock -Plan $Plan
    $startRow = [int] $Plan.TargetStartRow
    $startColumn = [int] $Plan.TargetStartColumn

    $metadataValueOffsets = @(
        @(1, 1),
        @(2, 1),
        @(2, 5),
        @(3, 1),
        @(3, 4),
        @(3, 6),
        @(3, 8),
        @(4, 1),
        @(4, 6)
    )
    foreach ($offset in $metadataValueOffsets) {
        $Worksheet.Cells.Item($startRow + [int] $offset[0], $startColumn + [int] $offset[1]).Value2 = $null
    }

    for ($sample = 1; $sample -le ($targetBlock.BatchAverageRow - $targetBlock.SampleStartRow); $sample++) {
        $targetRow = $targetBlock.SampleStartRow + $sample - 1
        for ($column = $startColumn + 1; $column -le $targetBlock.BatchAverageColumn; $column++) {
            $Worksheet.Cells.Item($targetRow, $column).Value2 = $null
        }
    }

    for ($column = $startColumn + 1; $column -le $targetBlock.BatchAverageColumn; $column++) {
        $Worksheet.Cells.Item($targetBlock.BatchAverageRow, $column).Value2 = $null
    }
}

function Write-NewExperimentBlockMetadata {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [pscustomobject] $Plan
    )

    $metadata = $Plan.Metadata
    $defaults = $metadata.Defaults
    $row = [int] $Plan.TargetStartRow
    $column = [int] $Plan.TargetStartColumn
    $padText = [string] $metadata.PadType
    if (-not [string]::IsNullOrWhiteSpace($padText) -and $padText -notmatch '抛光垫') {
        $padText = $padText + '抛光垫'
    }

    $slurryText = '{0}{1}{2}' -f $metadata.Concentration, $metadata.ParticleSize, $metadata.SlurryType
    if (-not [string]::IsNullOrWhiteSpace($metadata.Note)) {
        $slurryText = '{0}（{1}）' -f $slurryText, $metadata.Note
    }

    $templateBlock = Get-ExperimentObjectProperty -Object $Plan -Name 'TemplateBlock' -Default $null
    if ($null -ne $templateBlock) {
        $targetBlock = Convert-BlockToTargetBlock -TemplateBlock $templateBlock -Plan $Plan
    }
    else {
        $targetSheetName = [string] (Get-ExperimentObjectProperty -Object $Plan -Name 'TargetSheetName' -Default $metadata.TargetSheetName)
        $targetBlock = [pscustomobject]@{
            SheetName           = $targetSheetName
            StartRow            = $row
            StartColumn         = $column
            Title               = $Plan.Title
            ExperimentNo        = $metadata.ExperimentNo
            SaHeaderRow         = ($row + 5)
            SampleStartRow      = ($row + 6)
            SaStartColumn       = ($column + 1)
            BatchAverageRow     = ($row + 11)
            BatchAverageColumn  = ($column + 9)
        }
    }
    $fieldCells = Get-ExperimentBlockFieldCells -Worksheet $Worksheet -Block $targetBlock
    $fieldFallbacks = @{
        BallMatch = [pscustomobject]@{ Row = ($row + 2); Column = ($column + 6) }
        Plate     = [pscustomobject]@{ Row = ($row + 3); Column = ($column + 1) }
        Track     = [pscustomobject]@{ Row = ($row + 3); Column = ($column + 4) }
        Spacing   = [pscustomobject]@{ Row = ($row + 3); Column = ($column + 6) }
        Fullness  = [pscustomobject]@{ Row = ($row + 3); Column = ($column + 8) }
        Pressure  = [pscustomobject]@{ Row = ($row + 4); Column = ($column + 6) }
    }
    $fieldValues = @{
        BallMatch = $defaults.BallMatch
        Plate     = $defaults.Plate
        Track     = $defaults.Track
        Spacing   = $defaults.Spacing
        Fullness  = $defaults.Fullness
        Pressure  = $defaults.Pressure
    }

    Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $Worksheet -Row $row -Column $column -Value $Plan.Title -FieldName '标题')
    Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $Worksheet -Row ($row + 1) -Column ($column + 1) -Value $slurryText -FieldName '抛光液')
    Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $Worksheet -Row ($row + 2) -Column ($column + 1) -Value $metadata.BallSource -FieldName '球源')
    Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $Worksheet -Row ($row + 4) -Column ($column + 1) -Value $padText -FieldName '抛光垫')

    foreach ($field in Get-ExperimentProcessFieldDefinitions) {
        $targetCell = $null
        if ($fieldCells.ContainsKey($field.Key)) {
            $targetCell = $fieldCells[$field.Key]
        }
        else {
            $targetCell = $fieldFallbacks[$field.Key]
            if ($field.Key -eq 'BallMatch') {
                Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $Worksheet -Row ($row + 2) -Column ($column + 5) -Value '配球' -FieldName '配球标签')
            }
        }

        Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $Worksheet -Row $targetCell.Row -Column $targetCell.Column -Value $fieldValues[$field.Key] -FieldName $field.Label)
    }

    $remarkRow = [int] $targetBlock.BatchAverageRow + 1
    Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $Worksheet -Row $remarkRow -Column $targetBlock.StartColumn -Value '备注' -FieldName '备注标签')
    Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $Worksheet -Row $remarkRow -Column ($targetBlock.StartColumn + 1) -Value $metadata.Note -FieldName '备注内容')
}

function New-ExperimentBlockFromTemplate {
    param(
        [Parameter(Mandatory = $true)] [object] $Workbook,
        [Parameter(Mandatory = $true)] [pscustomobject] $Plan,
        [Parameter(Mandatory = $true)] [object[]] $SaRows,
        [Parameter(Mandatory = $true)] [object[]] $Blocks
    )

    $worksheet = $Workbook.Worksheets.Item($Plan.TargetSheetName)
    $capacity = [int] $Plan.TemplateBlock.BatchAverageRow - [int] $Plan.TemplateBlock.SampleStartRow
    $indices = @{}
    foreach ($sample in $SaRows) {
        $index = 0
        if (-not [int]::TryParse([string] $sample.SampleIndex, [ref] $index) -or $index -lt 1 -or $index -gt 10000 -or $indices.ContainsKey($index)) {
            return [pscustomobject]@{ Created = $false; Reason = '样品编号无效或重复，需要确认'; TargetBlock = $null }
        }
        $indices[$index] = $true
        $capacity = [Math]::Max($capacity, $index)
    }
    $Plan | Add-Member -NotePropertyName SampleCapacity -NotePropertyValue $capacity -Force
    $safetyTemplate = $Plan.TemplateBlock.PSObject.Copy()
    $safetyTemplate.BatchAverageRow = [int] $safetyTemplate.SampleStartRow + $capacity

    $targetSheetBlocks = @($Blocks | Where-Object { $_.SheetName -eq $Plan.TargetSheetName })
    $safety = Test-NewBlockTargetRange -Worksheet $worksheet -TargetStartRow $Plan.TargetStartRow -TargetStartColumn $Plan.TargetStartColumn -TemplateBlock $safetyTemplate -ExistingBlocks $targetSheetBlocks
    if (-not $safety.IsSafe) {
        return [pscustomobject]@{ Created = $false; Reason = $safety.Reason; TargetBlock = $null }
    }

    $sourceRangeInfo = Get-ExperimentBlockRange -Block $Plan.TemplateBlock
    $source = $worksheet.Range(
        $worksheet.Cells.Item($Plan.TemplateBlock.StartRow, $Plan.TemplateBlock.StartColumn),
        $worksheet.Cells.Item($Plan.TemplateBlock.StartRow + $sourceRangeInfo.Rows - 1, $Plan.TemplateBlock.StartColumn + $sourceRangeInfo.Columns - 1)
    )
    $target = $worksheet.Cells.Item($Plan.TargetStartRow, $Plan.TargetStartColumn)
    [void] $source.Copy($target)
    $targetBlock = Convert-BlockToTargetBlock -TemplateBlock $Plan.TemplateBlock -Plan $Plan
    # Copy sample formatting separately so growing a block never shifts adjacent experiments.
    if ($capacity -gt ([int] $Plan.TemplateBlock.BatchAverageRow - [int] $Plan.TemplateBlock.SampleStartRow)) {
        # Historical templates vertically merge the roundness cells across five samples.
        # Unmerge only the checked new block, never the source or adjacent blocks.
        $sampleArea = $worksheet.Range($worksheet.Cells.Item($targetBlock.SampleStartRow, $Plan.TargetStartColumn), $worksheet.Cells.Item($targetBlock.BatchAverageRow + 1, $Plan.TargetStartColumn + 9))
        [void] $sampleArea.UnMerge()
        $sampleTemplate = $worksheet.Range($worksheet.Cells.Item($targetBlock.SampleStartRow, $Plan.TargetStartColumn), $worksheet.Cells.Item($targetBlock.SampleStartRow, $Plan.TargetStartColumn + 9))
        for ($r = $targetBlock.SampleStartRow + 1; $r -lt $targetBlock.BatchAverageRow; $r++) {
            [void] $sampleTemplate.Copy($worksheet.Cells.Item($r, $Plan.TargetStartColumn))
        }
        $footer = $worksheet.Range($worksheet.Cells.Item($Plan.TemplateBlock.BatchAverageRow, $Plan.TemplateBlock.StartColumn), $worksheet.Cells.Item($Plan.TemplateBlock.BatchAverageRow + 1, $Plan.TemplateBlock.StartColumn + 9))
        [void] $footer.Copy($worksheet.Cells.Item($targetBlock.BatchAverageRow, $Plan.TargetStartColumn))
    }
    for ($r = $targetBlock.SampleStartRow; $r -lt $targetBlock.BatchAverageRow; $r++) {
        Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $worksheet -Row $r -Column $Plan.TargetStartColumn -Value ('球' + ($r - $targetBlock.SampleStartRow + 1)) -FieldName '样品编号')
    }
    $worksheet.Application.CutCopyMode = $false

    Clear-NewExperimentBlockVariableData -Worksheet $worksheet -Plan $Plan -TemplateBlock $Plan.TemplateBlock
    Write-NewExperimentBlockMetadata -Worksheet $worksheet -Plan $Plan
    $targetBlock = Convert-BlockToTargetBlock -TemplateBlock $Plan.TemplateBlock -Plan $Plan
    [void] (Write-SaRowsToExperimentBlock -Worksheet $worksheet -Block $targetBlock -SaRows $SaRows)

    return [pscustomobject]@{ Created = $true; Reason = '已创建新实验块'; TargetBlock = $targetBlock }
}

function Test-NewBlockCreated {
    param(
        [Parameter(Mandatory = $true)] [string] $WorkbookPath,
        [Parameter(Mandatory = $true)] [string] $TargetSheetName,
        [Parameter(Mandatory = $true)] [string] $ExperimentNo
    )

    $blocks = @(Get-ExperimentBackfillBlocks -WorkbookPath $WorkbookPath -SheetNames @($TargetSheetName))
    return ($null -ne ($blocks | Where-Object { $_.SheetName -eq $TargetSheetName -and $_.ExperimentNo -eq $ExperimentNo } | Select-Object -First 1))
}

function Compare-ExcelCellTextWithNumber {
    param(
        [AllowNull()] [string] $CellText,
        [AllowNull()] [object] $Expected
    )

    if ([string]::IsNullOrWhiteSpace($CellText)) {
        return $true
    }

    $expectedText = ConvertTo-ExperimentBackfillCellText $Expected
    if ($CellText.Trim() -eq $expectedText) {
        return $true
    }

    $cellNumber = 0.0
    $expectedNumber = 0.0
    $cellParsed = [double]::TryParse($CellText.Trim(), [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref] $cellNumber)
    $expectedParsed = [double]::TryParse($expectedText, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref] $expectedNumber)

    if ($cellParsed -and $expectedParsed) {
        return ([Math]::Abs($cellNumber - $expectedNumber) -lt 0.000001)
    }

    return $false
}

function Compare-ExcelCellTextForVerification {
    param(
        [AllowNull()] [string] $ActualText,
        [AllowNull()] [object] $Expected
    )

    $actual = ([string] $ActualText).Trim()
    $expectedText = (ConvertTo-ExperimentBackfillCellText $Expected).Trim()
    if ($actual -eq $expectedText) { return $true }

    $actualNumber = 0.0
    $expectedNumber = 0.0
    $actualOk = [double]::TryParse($actual, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref] $actualNumber)
    $expectedOk = [double]::TryParse($expectedText, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref] $expectedNumber)
    return ($actualOk -and $expectedOk -and [Math]::Abs($actualNumber - $expectedNumber) -lt 0.000001)
}

function Set-ExcelCellTextVerified {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [int] $Row,
        [Parameter(Mandatory = $true)] [int] $Column,
        [AllowNull()] [object] $Value,
        [Parameter(Mandatory = $true)] [string] $FieldName
    )

    $address = 'R{0}C{1}' -f $Row, $Column
    $cell = $Worksheet.Cells.Item($Row, $Column)
    $text = [string] (ConvertTo-ExperimentBackfillCellText $Value)
    $cell.Value = $text
    $actualText = ([string] $cell.Text).Trim()
    if (Compare-ExcelCellTextForVerification -ActualText $actualText -Expected $text) {
        return [pscustomobject]@{ Success = $true; Message = ''; Address = $address }
    }

    $actualValue2 = ([string] $cell.Value2).Trim()
    if (Compare-ExcelCellTextForVerification -ActualText $actualValue2 -Expected $text) {
        return [pscustomobject]@{ Success = $true; Message = ''; Address = $address }
    }

    return [pscustomobject]@{
        Success = $false
        Message = ('{0} 写入后读回不一致：{1} 期望 {2}，实际 Text={3} Value2={4}' -f $FieldName, $address, $text, $actualText, $actualValue2)
        Address = $address
    }
}

function Assert-ExcelWriteSucceeded {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Result)

    if (-not $Result.Success) {
        throw $Result.Message
    }
}

function Test-SaCellConflict {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [pscustomobject] $Block,
        [Parameter(Mandatory = $true)] [object[]] $SaRows
    )

    # Check the entire write set before numeric conflicts can take the additive-extension path.
    foreach ($row in $SaRows) {
        $targetRow = $Block.SampleStartRow + [int]$row.SampleIndex - 1
        if ($targetRow -lt $Block.SampleStartRow -or $targetRow -ge $Block.BatchAverageRow) {
            return [pscustomobject]@{ HasConflict = $true; CanExtend = $false; Message = ('球{0} 超出实验块可写入范围' -f $row.SampleIndex) }
        }
        for ($offset = 0; $offset -lt 4; $offset++) {
            if ($Worksheet.Cells.Item($targetRow, $Block.SaStartColumn + $offset).HasFormula) {
                return [pscustomobject]@{ HasConflict = $true; CanExtend = $false; Message = ('球{0} 的 Sa 区域已有公式，需要确认' -f $row.SampleIndex) }
            }
        }
    }
    if ($Worksheet.Cells.Item($Block.BatchAverageRow, $Block.BatchAverageColumn).HasFormula) {
        return [pscustomobject]@{ HasConflict = $true; CanExtend = $false; Message = '单批均值已有公式，需要确认' }
    }
    $indices = @{}
    foreach ($row in $SaRows) {
        $hasValues = @(@($row.Sa1, $row.Sa2, $row.Sa3) | Where-Object { $null -ne $_ }).Count -gt 0
        if ($indices.ContainsKey([int] $row.SampleIndex) -or $hasValues -ne ($null -ne $row.SaMean)) {
            return [pscustomobject]@{ HasConflict = $true; Message = '来源 Sa 数据存在重复样品编号或缺少测量值、均值，需要确认' }
        }
        $indices[[int] $row.SampleIndex] = $true
        $targetRow = $Block.SampleStartRow + [int] $row.SampleIndex - 1
        if ($targetRow -lt $Block.SampleStartRow -or $targetRow -ge $Block.BatchAverageRow) {
            return [pscustomobject]@{ HasConflict = $true; Message = ('球{0} 超出实验块可写入范围' -f $row.SampleIndex) }
        }
        $values = @($row.Sa1, $row.Sa2, $row.Sa3, $row.SaMean)
        for ($offset = 0; $offset -lt 4; $offset++) {
            $cell = $Worksheet.Cells.Item($targetRow, $Block.SaStartColumn + $offset)
            if (-not (Compare-ExcelCellTextWithNumber -CellText ([string] $cell.Value2) -Expected $values[$offset])) {
                return [pscustomobject]@{
                    HasConflict = $true
                    Message     = ('球{0} 的 Sa 列 {1} 已有不同值：{2}' -f $row.SampleIndex, ($offset + 1), $cell.Value2)
                }
            }
        }
    }

    $snapshot = Get-ExperimentSaSnapshot -Worksheet $Worksheet -Block $Block
    foreach ($existing in $snapshot.SaRows) {
        if (-not $indices.ContainsKey([int] $existing.SampleIndex)) {
            return [pscustomobject]@{ HasConflict = $true; Message = '实验块包含来源记录以外的样品数据，需要确认' }
        }
    }
    $measuredRows = @($SaRows | Where-Object { $null -ne $_.SaMean })
    $expectedMean = if ($measuredRows.Count) { [Math]::Round(($measuredRows | Measure-Object -Property SaMean -Average).Average, 6) } else { $null }
    if (-not (Compare-ExcelCellTextWithNumber -CellText ([string] $snapshot.BatchMean) -Expected $expectedMean)) {
        return [pscustomobject]@{ HasConflict = $true; Message = '单批均值已有不同值，需要确认' }
    }
    return [pscustomobject]@{ HasConflict = $false; Message = '' }
}

function Write-SaRowsToExperimentBlock {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [pscustomobject] $Block,
        [Parameter(Mandatory = $true)] [object[]] $SaRows
    )

    $means = @()
    foreach ($row in $SaRows) {
        $targetRow = $Block.SampleStartRow + [int] $row.SampleIndex - 1
        $values = @($row.Sa1, $row.Sa2, $row.Sa3, $row.SaMean)
        for ($offset = 0; $offset -lt 4; $offset++) {
            Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $Worksheet -Row $targetRow -Column ($Block.SaStartColumn + $offset) -Value $values[$offset] -FieldName ('球{0} Sa{1}' -f $row.SampleIndex, ($offset + 1)))
        }
        if ($null -ne $row.SaMean) {
            $means += [double] $row.SaMean
        }
    }

    if ($means.Count -gt 0) {
        $average = [Math]::Round((($means | Measure-Object -Average).Average), 6)
        Assert-ExcelWriteSucceeded (Set-ExcelCellTextVerified -Worksheet $Worksheet -Row $Block.BatchAverageRow -Column $Block.BatchAverageColumn -Value $average -FieldName '单批均值')
    }

    return $SaRows.Count
}

function Backup-ExperimentWorkbook {
    param(
        [Parameter(Mandatory = $true)] [string] $WorkbookPath,
        [Parameter(Mandatory = $true)] [string] $BackupDirectory
    )

    if (-not (Test-Path -LiteralPath $BackupDirectory)) {
        New-Item -ItemType Directory -Path $BackupDirectory -Force | Out-Null
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($WorkbookPath)
    $extension = [System.IO.Path]::GetExtension($WorkbookPath)
    $backupPath = Join-Path $BackupDirectory ('{0}.{1}{2}' -f $baseName, (Get-Date -Format 'yyyyMMdd_HHmmss'), $extension)
    Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -Force
    return $backupPath
}

function Test-AutoImportRowLooksStructured {
    param(
        [AllowNull()] [string] $FolderPath,
        [AllowNull()] [string] $FolderName
    )

    $name = $FolderName
    if ([string]::IsNullOrWhiteSpace($name) -and -not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $name = Split-Path -Path ($FolderPath.TrimEnd('\')) -Leaf
    }

    return ($name -match '^\d+-')
}

function Test-AutoImportBackfillRemarkAllowsReplay {
    param(
        [AllowNull()] [string] $Remark,
        [switch] $IncludePhase2Preview
    )

    $text = ([string] $Remark).Trim()
    if ($text -eq '当前解析仍需确认：抛光液类型无法映射目标表') {
        # Mapping rules can change after review; reparse without bypassing write safety.
        return $true
    }
    if ($text -like '当前解析仍需确认：*') {
        return $false
    }
    if ($text -eq '阶段一不自动创建新实验块') {
        return $true
    }
    if ($IncludePhase2Preview -and $text -in @('目标区域存在非空内容：小结', '目标区域存在非空内容：备注')) {
        # Re-plan with corrected bounds; all target safety checks still apply.
        return $true
    }
    if ($IncludePhase2Preview -and $text -match '^Sa 样品数量 \d+ 超过模板容量 \d+$') { return $true }
    if ($IncludePhase2Preview -and $text -eq '阶段二预览：可自动建块') {
        return $true
    }

    $legacyParserErrorPatterns = @(
        '文件夹字段数量不足',
        '实验号不是纯数字',
        '日期无法解析',
        '浓度字段格式异常',
        '粒径字段格式异常',
        '抛光液类型无法映射目标表',
        '球胚字段为空',
        '抛光时间字段格式异常',
        '抛光垫字段为空'
    )
    foreach ($pattern in $legacyParserErrorPatterns) {
        if ($text -like ('*{0}*' -f $pattern)) {
            return $true
        }
    }

    return $false
}

function Test-BackfillDecisionIsCurrentParserIssue {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Decision)

    if ($Decision.Action -ne 'QueuePending') {
        return $false
    }

    $reason = ([string] $Decision.Reason).Trim()
    if ($reason -eq '文件夹命名不符合固定字段约定') {
        return $true
    }

    $parserIssuePatterns = @(
        '文件夹字段数量不足',
        '实验号不是纯数字',
        '日期无法解析',
        '浓度字段格式异常',
        '粒径字段格式异常',
        '抛光液类型无法映射目标表',
        '球胚字段为空',
        '抛光时间字段格式异常',
        '抛光垫字段为空'
    )
    foreach ($pattern in $parserIssuePatterns) {
        if ($reason -like ('*{0}*' -f $pattern)) {
            return $true
        }
    }

    return $false
}

function Get-AutoImportRowsForBackfill {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [hashtable] $StatusColumns,
        [switch] $IncludePhase2Preview,
        [switch] $IncludeCompleted
    )

    $used = $Worksheet.UsedRange
    $rows = @()
    $versionColumns = Get-ExperimentVersionColumns $Worksheet
    for ($row = 2; $row -le [int] $used.Rows.Count; $row++) {
        if ($versionColumns.ContainsKey('版本状态') -and [string] $Worksheet.Cells.Item($row, $versionColumns['版本状态']).Value2 -eq 'Archived') { continue }
        $status = ''
        if ($StatusColumns.ContainsKey('回写状态')) {
            $status = ([string] $Worksheet.Cells.Item($row, $StatusColumns['回写状态']).Text).Trim()
        }
        $remark = ''
        if ($StatusColumns.ContainsKey('回写备注')) {
            $remark = ([string] $Worksheet.Cells.Item($row, $StatusColumns['回写备注']).Text).Trim()
        }
        if ($status -eq '已回写' -and -not $IncludeCompleted) {
            continue
        }
        $canReplayPending = Test-AutoImportBackfillRemarkAllowsReplay -Remark $remark -IncludePhase2Preview:$IncludePhase2Preview
        if ($status -eq '待确认' -and -not $canReplayPending) {
            continue
        }

        $folderPath = ([string] $Worksheet.Cells.Item($row, 3).Text).Trim()
        $folderName = ([string] $Worksheet.Cells.Item($row, 7).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($folderPath) -or -not (Test-AutoImportRowLooksStructured -FolderPath $folderPath -FolderName $folderName)) {
            continue
        }

        $sampleIndexText = ([string] $Worksheet.Cells.Item($row, 12).Text).Trim()
        $sampleIndex = 0
        if (-not [int]::TryParse($sampleIndexText, [ref] $sampleIndex)) {
            continue
        }

        $rows += [pscustomobject]@{
            RowNumber   = $row
            FolderPath  = $folderPath
            FolderName  = $folderName
            SampleIndex = $sampleIndex
            Sa1         = ConvertTo-NullableDouble ([string] $Worksheet.Cells.Item($row, 13).Text)
            Sa2         = ConvertTo-NullableDouble ([string] $Worksheet.Cells.Item($row, 14).Text)
            Sa3         = ConvertTo-NullableDouble ([string] $Worksheet.Cells.Item($row, 15).Text)
            SaMean      = ConvertTo-NullableDouble ([string] $Worksheet.Cells.Item($row, 16).Text)
        }
    }

    return $rows
}

function Group-AutoImportRowsByFolder {
    param([AllowEmptyCollection()] [object[]] $Rows = @())

    $groups = @()
    foreach ($group in ($Rows | Group-Object FolderPath)) {
        $saRows = @($group.Group | Sort-Object SampleIndex | ForEach-Object {
            [pscustomobject]@{
                SampleIndex = $_.SampleIndex
                Sa1         = $_.Sa1
                Sa2         = $_.Sa2
                Sa3         = $_.Sa3
                SaMean      = $_.SaMean
                AuditRow    = $_.RowNumber
            }
        })

        $auditRows = @()
        foreach ($item in $group.Group) {
            $auditRows += [int] $item.RowNumber
        }

        $groups += [pscustomobject]@{
            FolderPath = $group.Name
            SaRows     = $saRows
            AuditRows  = $auditRows
        }
    }

    return $groups
}

function Get-AutoImportGroupExperimentNumber {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Group)

    $folderName = Split-Path -Path ([string] $Group.FolderPath).TrimEnd('\') -Leaf
    $match = [regex]::Match($folderName, '^\s*(\d+)\s*-')
    if ($match.Success) {
        return [int] $match.Groups[1].Value
    }

    return [int]::MaxValue
}

function Sort-AutoImportGroupsForBackfill {
    param(
        [Parameter(ValueFromPipeline = $true)]
        [AllowNull()]
        [pscustomobject] $Group
    )

    begin {
        $groups = @()
    }
    process {
        if ($null -ne $Group) {
            $groups += $Group
        }
    }
    end {
        return @($groups | Sort-Object @{ Expression = { Get-AutoImportGroupExperimentNumber -Group $_ } }, FolderPath)
    }
}

function Set-AutoImportBackfillStatus {
    param(
        [Parameter(Mandatory = $true)] [object] $Worksheet,
        [Parameter(Mandatory = $true)] [hashtable] $StatusColumns,
        [Parameter(Mandatory = $true)] [int[]] $Rows,
        [Parameter(Mandatory = $true)] [string] $Status,
        [string] $TargetSheetName = '',
        [string] $TargetBlockCell = '',
        [string] $Remark = ''
    )

    foreach ($row in $Rows) {
        $Worksheet.Cells.Item($row, $StatusColumns['回写状态']).Value2 = $Status
        $Worksheet.Cells.Item($row, $StatusColumns['目标工作表']).Value2 = $TargetSheetName
        $Worksheet.Cells.Item($row, $StatusColumns['目标块起始单元格']).Value2 = $TargetBlockCell
        $Worksheet.Cells.Item($row, $StatusColumns['回写时间']).Value2 = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $Worksheet.Cells.Item($row, $StatusColumns['回写备注']).Value2 = $Remark
    }
}

function Invoke-ExperimentBackfillPhase1 {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Config)
    Invoke-WithExperimentWorkbookLock -WorkbookPath $Config.MasterWorkbookPath -Action {
        if (Test-ExperimentTransactionActive $Config.MasterWorkbookPath) {
            Invoke-ExperimentBackfillPhase1Core -Config $Config
        }
        else {
            $transaction = Invoke-ExperimentWorkbookTransaction -Config $Config -Action { param($working) Invoke-ExperimentBackfillPhase1Core -Config $working }
            $transaction.Result.BackupPath = $transaction.BackupPath
            $transaction.Result
        }
    }
}

function Invoke-ExperimentBackfillPhase1Core {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Config)

    $backupPath = $script:ExperimentActiveTransaction.BackupPath
    $session = $null; $excel = $null
    $workbook = $null
    $backfilledBlocks = 0
    $pendingRows = 0
    $previewNewBlocks = 0
    $createdNewBlocks = 0
    $newBlockFailures = 0
    $verifiedGroups = @()

    try {
        $session = New-ExperimentExcelSession
        $excel = $session.Application
        $workbook = $excel.Workbooks.Open($Config.MasterWorkbookPath)
        if ($workbook.ReadOnly) {
            throw "主实验数据表现在是只读状态，请关闭已打开的 Excel 文件后重试：$($Config.MasterWorkbookPath)"
        }

        $auto = Ensure-AutoImportWorksheet -Workbook $workbook -SheetName $Config.AutoSheetName
        $statusColumns = Ensure-BackfillStatusColumns -Worksheet $auto
        $pending = Ensure-PendingBackfillWorksheet -Workbook $workbook -SheetName $Config.PendingSheetName

        $includePhase2Preview = ([string] $Config.NewBlockMode -eq 'Auto')
        $auditRows = @(Get-AutoImportRowsForBackfill -Worksheet $auto -StatusColumns $statusColumns -IncludePhase2Preview:$includePhase2Preview -IncludeCompleted)
        $groups = @(Group-AutoImportRowsByFolder -Rows $auditRows | Sort-AutoImportGroupsForBackfill)
        $blocks = @(Get-ExperimentBackfillBlocksFromWorkbook -Workbook $workbook -SheetNames $Config.TargetSheets)

        Update-ExperimentRemarkLabels -Workbook $workbook -Blocks $blocks

        foreach ($group in $groups) {
            $metadata = Parse-StructuredExperimentFolderName -FolderPath $group.FolderPath -SlurryMappings $Config.SlurryMappings -ProcessDefaults $Config.ProcessDefaults
            $decision = Resolve-ExperimentBackfillDecision -Metadata $metadata -Blocks $blocks

            if ($decision.Action -eq 'BackfillExisting') {
                $worksheet = $workbook.Worksheets.Item($decision.TargetBlock.SheetName)
                $targetBlockCell = 'R{0}C{1}' -f $decision.TargetBlock.StartRow, $decision.TargetBlock.StartColumn
                $snapshot = Get-ExperimentSaSnapshot -Worksheet $worksheet -Block $decision.TargetBlock
                if (Test-ExperimentSaSnapshotComplete -Snapshot $snapshot -ExpectedRows $group.SaRows) {
                    $verifiedGroups += $group
                    Set-AutoImportBackfillStatus -Worksheet $auto -StatusColumns $statusColumns -Rows $group.AuditRows -Status '已回写' -TargetSheetName $decision.TargetBlock.SheetName -TargetBlockCell $targetBlockCell -Remark '已核验实验块 Sa 完整'
                    continue
                }
                $conflict = Test-SaCellConflict -Worksheet $worksheet -Block $decision.TargetBlock -SaRows $group.SaRows
                if ($conflict.HasConflict -and (Get-ExperimentObjectProperty $conflict 'CanExtend' $true) -and (Test-ExperimentSafeSaExtension -Worksheet $auto -Group $group -Snapshot $snapshot)) {
                    $inRange = @($group.SaRows | Where-Object { $_.SampleIndex -lt 1 -or ($decision.TargetBlock.SampleStartRow + $_.SampleIndex - 1) -ge $decision.TargetBlock.BatchAverageRow }).Count -eq 0
                    if ($inRange) { $conflict = [pscustomobject]@{ HasConflict = $false; Message = '' } }
                }
                if ($conflict.HasConflict) {
                    $conflictDecision = [pscustomobject]@{
                        Action          = 'QueuePending'
                        Reason          = $conflict.Message
                        TargetBlock     = $decision.TargetBlock
                        TargetSheetName = $decision.TargetBlock.SheetName
                        Confidence      = 0
                    }
                    $pendingRows += Append-PendingBackfillRecord -Worksheet $pending -Metadata $metadata -Decision $conflictDecision -SaRows $group.SaRows
                    Set-AutoImportBackfillStatus -Worksheet $auto -StatusColumns $statusColumns -Rows $group.AuditRows -Status '待确认' -TargetSheetName $decision.TargetBlock.SheetName -TargetBlockCell $targetBlockCell -Remark $conflict.Message
                    continue
                }

                [void] (Write-SaRowsToExperimentBlock -Worksheet $worksheet -Block $decision.TargetBlock -SaRows $group.SaRows)
                $backfilledBlocks++
                $verifiedGroups += $group
                Set-AutoImportBackfillStatus -Worksheet $auto -StatusColumns $statusColumns -Rows $group.AuditRows -Status '已回写' -TargetSheetName $decision.TargetBlock.SheetName -TargetBlockCell $targetBlockCell -Remark $decision.Reason
                continue
            }

            if ($decision.Action -eq 'QueuePending' -and $decision.Reason -eq '阶段一不自动创建新实验块') {
                $plan = New-ExperimentBackfillPlan -Metadata $metadata -Decision $decision -Group $group -Blocks $blocks -Workbook $workbook -Config $Config
                if ($plan.Action -eq 'PreviewNewBlock') {
                    $pendingRows += Append-NewBlockPreviewRecord -Worksheet $pending -Plan $plan -SaRows $group.SaRows
                    $previewNewBlocks++
                    Set-AutoImportBackfillStatus -Worksheet $auto -StatusColumns $statusColumns -Rows $group.AuditRows -Status '待确认' -TargetSheetName $plan.TargetSheetName -TargetBlockCell $plan.TargetRangeAddress -Remark '阶段二预览：可自动建块'
                    continue
                }

                if ($plan.Action -eq 'CreateNewBlock') {
                    $createResult = New-ExperimentBlockFromTemplate -Workbook $workbook -Plan $plan -SaRows $group.SaRows -Blocks $blocks
                    if ($createResult.Created) {
                        $blocks = @(Get-ExperimentBackfillBlocksFromWorkbook -Workbook $workbook -SheetNames $Config.TargetSheets)
                        $createdBlock = $blocks | Where-Object { $_.SheetName -eq $plan.TargetSheetName -and $_.ExperimentNo -eq $metadata.ExperimentNo } | Select-Object -First 1
                        if ($null -eq $createdBlock) {
                            throw ('阶段二自动建块后无法重新索引实验{0}，已取消保存。' -f $metadata.ExperimentNo)
                        }
                        $createdNewBlocks++
                        $verifiedGroups += $group
                        Set-AutoImportBackfillStatus -Worksheet $auto -StatusColumns $statusColumns -Rows $group.AuditRows -Status '已回写' -TargetSheetName $plan.TargetSheetName -TargetBlockCell $plan.TargetRangeAddress -Remark '阶段二自动新建实验块'
                        continue
                    }

                    $newBlockFailures++
                    $pendingRows += Append-PendingBackfillRecord -Worksheet $pending -Metadata $metadata -Decision ([pscustomobject]@{
                        Action          = 'QueuePending'
                        Reason          = $createResult.Reason
                        TargetSheetName = $plan.TargetSheetName
                        TargetBlock     = $plan.TemplateBlock
                        Confidence      = 0
                    }) -SaRows $group.SaRows
                    Set-AutoImportBackfillStatus -Worksheet $auto -StatusColumns $statusColumns -Rows $group.AuditRows -Status '待确认' -TargetSheetName $plan.TargetSheetName -TargetBlockCell $plan.TargetRangeAddress -Remark $createResult.Reason
                    continue
                }

                if ($plan.Action -eq 'QueuePending') {
                    $decision = [pscustomobject]@{
                        Action          = 'QueuePending'
                        Reason          = $plan.Reason
                        TargetSheetName = $plan.TargetSheetName
                        TargetBlock     = $plan.TemplateBlock
                        Confidence      = 0
                    }
                }
            }

            $pendingRows += Append-PendingBackfillRecord -Worksheet $pending -Metadata $metadata -Decision $decision -SaRows $group.SaRows
            $pendingRemark = $decision.Reason
            if (Test-BackfillDecisionIsCurrentParserIssue -Decision $decision) {
                $pendingRemark = '当前解析仍需确认：{0}' -f $decision.Reason
            }
            Set-AutoImportBackfillStatus -Worksheet $auto -StatusColumns $statusColumns -Rows $group.AuditRows -Status '待确认' -TargetSheetName $decision.TargetSheetName -TargetBlockCell '' -Remark $pendingRemark
        }

        $reviewColumns = Ensure-NewBlockPreviewColumns -Worksheet $pending
        foreach ($verifiedGroup in $verifiedGroups) {
            $metadata = Parse-StructuredExperimentFolderName -FolderPath $verifiedGroup.FolderPath -SlurryMappings $Config.SlurryMappings -ProcessDefaults $Config.ProcessDefaults
            $check = Resolve-ExperimentBackfillDecision -Metadata $metadata -Blocks $blocks
            if ($check.Action -ne 'BackfillExisting') { throw '写入后实验身份或目标位置无法唯一核验，已取消提交' }
            $snapshot = Get-ExperimentSaSnapshot -Worksheet $workbook.Worksheets.Item($check.TargetBlock.SheetName) -Block $check.TargetBlock
            if (-not (Test-ExperimentSaSnapshotComplete -Snapshot $snapshot -ExpectedRows $verifiedGroup.SaRows)) { throw "写入后样品数或 Sa 核验失败：实验$($metadata.ExperimentNo)" }
            for ($reviewRow = 2; $reviewRow -le [int] $pending.UsedRange.Rows.Count; $reviewRow++) {
                if ([string] $pending.Cells.Item($reviewRow, 3).Value2 -eq $verifiedGroup.FolderPath) {
                    $pending.Cells.Item($reviewRow, $reviewColumns['确认状态']).Value2 = '已处理'
                    $pending.Cells.Item($reviewRow, $reviewColumns['处理时间']).Value2 = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                }
            }
        }
        Save-ExperimentWorkbookCheckpoint -Workbook $workbook
        return [pscustomobject]@{
            BackupPath       = $backupPath
            GroupsProcessed  = $groups.Count
            BackfilledBlocks = $backfilledBlocks
            PendingRows      = $pendingRows
            NewBlockMode     = $Config.NewBlockMode
            PreviewNewBlocks = $previewNewBlocks
            CreatedNewBlocks = $createdNewBlocks
            NewBlockFailures = $newBlockFailures
        }
    }
    finally {
        Close-ExperimentExcelSession -Session $session -Workbook $workbook
    }
}
