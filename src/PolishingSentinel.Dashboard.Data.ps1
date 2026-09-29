Set-StrictMode -Version 2.0
Add-Type -AssemblyName System.IO.Compression
function Get-SentinelDashboardData {
    param([string]$ScriptRoot,[switch]$Force)
    $settings=Read-PolishingSentinelSettings $ScriptRoot
    $key=([IO.Path]::GetFullPath($settings.MasterWorkbookPath)) + '|' + ($settings.SlurryMappings | ConvertTo-Json -Compress)
    $path=Join-Path $ScriptRoot 'state/dashboard-summary.json'
    $cached=$null
    try {
        if (Test-Path -LiteralPath $path) {
            $candidate=Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($candidate.SchemaVersion -eq 3 -and $candidate.ConfigKey -eq $key) { $cached=$candidate }
        }
    } catch { $cached=$null }
    try {
        $file=Get-Item -LiteralPath $settings.MasterWorkbookPath -ErrorAction Stop
        $signature='{0}:{1}' -f $file.Length,$file.LastWriteTimeUtc.Ticks
        if (-not $Force -and $cached -and $cached.Signature -eq $signature) {
            return [pscustomobject]@{ Snapshot=$cached; Error=''; Cached=$true }
        }
        $snapshot=Read-SentinelDashboardWorkbook $settings.MasterWorkbookPath $settings.SlurryMappings
        $snapshot | Add-Member ConfigKey $key
        $directory=Split-Path $path
        [void][IO.Directory]::CreateDirectory($directory)
        $temp=Join-Path $directory ('.dashboard-'+[guid]::NewGuid().ToString('N')+'.tmp')
        try {
            [IO.File]::WriteAllText($temp,($snapshot | ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($true)))
            if ([IO.File]::Exists($path)) { [IO.File]::Replace($temp,$path,[Management.Automation.Language.NullString]::Value) }
            else { [IO.File]::Move($temp,$path) }
        } finally { if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) } }
        return [pscustomobject]@{ Snapshot=$snapshot; Error=''; Cached=$false }
    } catch {
        return [pscustomobject]@{ Snapshot=$cached; Error=$_.Exception.Message; Cached=$false }
    }
}

function Read-SentinelZipXml {
    param($Archive, [string]$Path)
    $entry = $Archive.GetEntry($Path)
    if (-not $entry) { throw "工作簿缺少部件：$Path" }
    $stream = $entry.Open()
    $options = New-Object Xml.XmlReaderSettings
    $options.DtdProcessing = 'Prohibit'
    $options.XmlResolver = $null
    $reader = [Xml.XmlReader]::Create($stream,$options)
    try {
        $document = New-Object Xml.XmlDocument
        $document.XmlResolver = $null
        $document.Load($reader)
        return ,$document
    } finally { $reader.Dispose(); $stream.Dispose() }
}

function Get-SentinelNumericValue {
    param($Value)
    $number = 0.0
    if ($null -ne $Value -and [double]::TryParse([string]$Value, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and
        -not [double]::IsNaN($number) -and -not [double]::IsInfinity($number) -and $number -ge 0) { return $number }
    return $null
}

function Read-SentinelSheetCells {
    param($Archive, [string]$Path, $Strings)
    $xml = Read-SentinelZipXml $Archive $Path
    $cells = @{}
    foreach ($node in $xml.SelectNodes('//*[local-name()="sheetData"]/*[local-name()="row"]/*[local-name()="c"]')) {
        $address = $node.GetAttribute('r')
        if ($address -notmatch '^([A-Z]+)(\d+)$') { continue }
        $row = [int]$Matches[2]; $column = 0
        foreach ($letter in $Matches[1].ToCharArray()) { $column = $column * 26 + [int]$letter - 64 }
        $type = $node.GetAttribute('t')
        $valueNode = $node.SelectSingleNode('*[local-name()="v"]')
        $text = if ($valueNode) { $valueNode.InnerText } else { '' }
        if ($type -eq 's') { $text = [string]$Strings[[int]$text] }
        elseif ($type -eq 'inlineStr') { $text = (@($node.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join '') }
        $numeric = if ($type -in @('', 'n')) { Get-SentinelNumericValue $text } else { $null }
        $cells["$row,$column"] = [pscustomobject]@{ Row=$row; Column=$column; Text=$text; Number=$numeric }
    }
    return $cells
}

function Convert-SentinelSheetExperiments {
    param($Cells, [string]$SheetName, [string]$DefaultKind)
    $titles = @($Cells.Values | Where-Object {
        $_.Text -match '^\s*(?:[A-Za-z][A-Za-z+\s]*\s+)?实验\s*\d+\s*.*抛光' -and $_.Text -notmatch '记录表'
    } | Sort-Object Row,Column)
    foreach ($title in $titles) {
        $numberMatch = [regex]::Match($title.Text,'实验\s*0*(\d+)')
        $kind = if ($title.Text -match '油基') { 'Oil' } elseif ($title.Text -match '水基') { 'Water' } else { $DefaultKind }
        $dateMatch = [regex]::Match($title.Text,'(20\d{2})[.\-/年](\d{1,2})[.\-/月](\d{1,2})')
        $date = $null
        if ($dateMatch.Success) {
            $parsed = [datetime]::MinValue
            $iso = '{0}-{1:D2}-{2:D2}' -f $dateMatch.Groups[1].Value,[int]$dateMatch.Groups[2].Value,[int]$dateMatch.Groups[3].Value
            if ([datetime]::TryParseExact($iso,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$parsed)) { $date=$parsed.ToString('yyyy-MM-dd') }
        }
        $header = $null; $batch = $null; $mean = $null; $samples = 0; $unit = ''; $derived = $false
        $headers = @($Cells.Values | Where-Object {
            $_.Row -ge $title.Row -and $_.Row -le $title.Row+8 -and $_.Column -ge $title.Column -and $_.Column -lt $title.Column+10 -and $_.Text -match '^Sa1'
        } | Sort-Object Row,Column)
        if ($headers.Count) {
            $header = $headers[0]
            $unit = if ($header.Text -match '(?i)(nm|纳米)') { 'nm' } elseif ($header.Text -match '(?i)([μµu]m|微米)') { 'um' } else { '' }
            $next = @($titles | Where-Object { $_.Column -eq $title.Column -and $_.Row -gt $title.Row } | Select-Object -First 1)
            $end = if ($next.Count) { $next[0].Row } else { ($Cells.Values | Measure-Object Row -Maximum).Maximum + 1 }
            $batches = @($Cells.Values | Where-Object {
                $_.Row -gt $header.Row -and $_.Row -lt $end -and $_.Column -ge $title.Column -and $_.Column -lt $title.Column+10 -and $_.Text -match '^单批均值'
            } | Sort-Object Row)
            if ($batches.Count) {
                $batch = $batches[0]
                $sampleMeans = @(
                    for ($r=$header.Row+1; $r -lt $batch.Row; $r++) {
                        $cell = $Cells["$r,$($header.Column+3)"]
                        $value = if ($cell) { $cell.Number } else { $null }
                        if ($null -eq $value) {
                            $raw = @(0..2 | ForEach-Object { $c=$Cells["$r,$($header.Column+$_)"]; if ($c -and $null -ne $c.Number) { $c.Number } })
                            if ($raw.Count) { $value=($raw | Measure-Object -Average).Average }
                        }
                        if ($null -ne $value) { $value }
                    }
                )
                $samples=$sampleMeans.Count
                $meanCell = $Cells["$($batch.Row),$($header.Column+3)"]
                if ($meanCell) { $mean=$meanCell.Number }
                if ($null -eq $mean -and $samples) { $mean=($sampleMeans | Measure-Object -Average).Average; $derived=$true }
                if ($unit -eq 'um' -and $null -ne $mean) { $mean=$mean*1000 }
                if (-not $unit) { $mean=$null }
            }
        }
        $fields=@{}
        foreach ($r in $title.Row..($title.Row+4)) {
            for ($c=$title.Column; $c -lt $title.Column+9; $c++) {
                $label=$Cells["$r,$c"]; $value=$Cells["$r,$($c+1)"]
                if ($label -and $value -and $label.Text -in @('抛光液','球胚','抛光垫','盘体','滚道','间距','充盈度','压力','配球')) { $fields[$label.Text]=$value.Text }
            }
        }
        if ($batch) {
            $note=$Cells["$($batch.Row+1),$($title.Column+1)"]
            if ($note) { $fields['备注']=$note.Text }
        }
        [pscustomobject]@{
            Fields=$fields
            Kind=$kind; ExperimentNo=$numberMatch.Groups[1].Value; Date=$date; Sa=$mean
            Samples=$samples; Sheet=$SheetName; Row=$title.Row; Column=$title.Column; Title=$title.Text
            Unit=$unit; Derived=$derived; Duplicate=$false
        }
    }
}

function Read-SentinelDashboardWorkbook {
    param([string]$Path, $Mappings)
    $before = Get-Item -LiteralPath $Path -ErrorAction Stop
    $signature = '{0}:{1}' -f $before.Length,$before.LastWriteTimeUtc.Ticks
    $memory = New-Object IO.MemoryStream
    $archive=$null
    try {
        $source = [IO.File]::Open($Path,'Open','Read',([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        try { $source.CopyTo($memory) } finally { $source.Dispose() }
        $after = Get-Item -LiteralPath $Path -ErrorAction Stop
        if ($signature -ne ('{0}:{1}' -f $after.Length,$after.LastWriteTimeUtc.Ticks)) { throw '主表正在更新，稍后重试统计' }
        $memory.Position=0
        $archive = [IO.Compression.ZipArchive]::new($memory,[IO.Compression.ZipArchiveMode]::Read,$true)
        $strings = @()
        if ($archive.GetEntry('xl/sharedStrings.xml')) {
            $xml=Read-SentinelZipXml $archive 'xl/sharedStrings.xml'
            $strings=@($xml.SelectNodes('//*[local-name()="si"]') | ForEach-Object { (@($_.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText }) -join '') })
        }
        $book=Read-SentinelZipXml $archive 'xl/workbook.xml'
        $rels=Read-SentinelZipXml $archive 'xl/_rels/workbook.xml.rels'
        $targets=@{}
        foreach ($rel in $rels.DocumentElement.ChildNodes) {
            if ($rel.GetAttribute('TargetMode') -eq 'External') { continue }
            $uri=[uri]::new([uri]'http://workbook/xl/workbook.xml',$rel.GetAttribute('Target'))
            $targets[$rel.GetAttribute('Id')]=$uri.AbsolutePath.TrimStart('/')
        }
        $sheetKinds=@{}
        foreach ($pair in @(@('硅溶胶','Water'),@('金刚石','Oil'))) {
            $property=$Mappings.PSObject.Properties[$pair[0]]
            if ($property -and $property.Value) { $sheetKinds[[string]$property.Value]=$pair[1] }
        }
        if (-not $sheetKinds.Count) { throw '未配置水基或油基目标工作表' }
        $records=@(); $tasks=@(); $found=@{}; $unnumbered=0; $taskUnnumbered=0; $sources=@()
        foreach ($sheet in $book.SelectNodes('//*[local-name()="sheet"]')) {
            $name=$sheet.GetAttribute('name')
            if (-not $sheetKinds.ContainsKey($name) -and $name -ne '加工任务') { continue }
            $id=$sheet.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            if (-not $targets.ContainsKey($id)) { throw "工作表关系无效：$name" }
            $cells=Read-SentinelSheetCells $archive $targets[$id] $strings
            $defaultKind=if ($sheetKinds.ContainsKey($name)) { $sheetKinds[$name] } else { 'Unknown' }
            $sheetRecords=@(Convert-SentinelSheetExperiments $cells $name $defaultKind)
            $sheetUnnumbered=@($cells.Values | Where-Object {
                $_.Text -match '^\s*(?:[A-Za-z][A-Za-z+\s]*\s+)?实验\s*(?:[Xx]+|待定|\?+)\s*.*抛光' -and $_.Text -notmatch '记录表'
            }).Count
            if ($name -eq '加工任务') { $tasks+=$sheetRecords; $taskUnnumbered+=$sheetUnnumbered }
            else { $records+=$sheetRecords; $unnumbered+=$sheetUnnumbered }
            $sources+=$name
            $found[$name]=$true
        }
        $missing=@($sheetKinds.Keys | Where-Object { -not $found.ContainsKey($_) })
        if ($missing.Count) { throw ('主表缺少已配置的工作表：' + ($missing -join '、')) }
        $unique=@()
        foreach ($group in @($records | Group-Object { $_.Kind+':'+$_.ExperimentNo })) {
            $entry=$group.Group[0]
            if ($group.Count -gt 1) { $entry.Duplicate=$true; $entry.Sa=$null }
            $unique+=$entry
        }
        [pscustomobject]@{
            SchemaVersion=3; SourcePath=[IO.Path]::GetFullPath($Path); Signature=$signature
            UpdatedAt=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Records=@($unique)
            DuplicateCount=@($unique | Where-Object Duplicate).Count
            UnnumberedCount=$unnumbered; SourceSheets=@($sources)
            TaskRecords=@($tasks); TaskUnnumberedCount=$taskUnnumbered
        }
    } finally { if ($archive) { $archive.Dispose() }; $memory.Dispose() }
}

function Get-SentinelDashboardSummary {
    param($Snapshot, [datetime]$Now=(Get-Date))
    $records=@($Snapshot.Records)
    $monday=$Now.Date.AddDays(-(([int]$Now.DayOfWeek+6)%7))
    $week=@($records | Where-Object { $_.Date -and -not $_.Duplicate -and [datetime]$_.Date -ge $monday -and [datetime]$_.Date -le $Now.Date })
    [pscustomobject]@{
        Total=$records.Count; Water=@($records | Where-Object Kind -eq Water).Count
        Oil=@($records | Where-Object Kind -eq Oil).Count; Week=$week.Count
        Missing=@($records | Where-Object { $null -eq $_.Sa }).Count
    }
}

function Get-SentinelDashboardPoints {
    param($Snapshot,[string]$Range='All',[string]$Kind='All',[datetime]$Now=(Get-Date))
    $start=if ($Range -eq 'Week') { $Now.Date.AddDays(-(([int]$Now.DayOfWeek+6)%7)) } elseif ($Range -eq 'Month') { $Now.Date.AddDays(-29) } else { [datetime]::MinValue }
    @($Snapshot.Records | Where-Object {
        $_.Date -and [datetime]$_.Date -ge $start -and [datetime]$_.Date -le $Now.Date -and ($Kind -eq 'All' -or $_.Kind -eq $Kind)
    } | Sort-Object Date,{[int]$_.ExperimentNo})
}
