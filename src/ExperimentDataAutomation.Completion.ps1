Set-StrictMode -Version 2.0

function Test-ExperimentSaNumberEqual {
    param([AllowNull()] $Actual, [AllowNull()] $Expected)
    if ([string]::IsNullOrWhiteSpace([string] $Expected)) { return [string]::IsNullOrWhiteSpace([string] $Actual) }
    $actualNumber = ConvertTo-NullableDouble ([string] $Actual)
    $expectedNumber = ConvertTo-NullableDouble ([string] $Expected)
    return ($null -ne $actualNumber -and $null -ne $expectedNumber -and
        -not [double]::IsNaN($actualNumber) -and -not [double]::IsInfinity($actualNumber) -and
        [Math]::Abs($actualNumber - $expectedNumber) -lt 0.000001)
}

function Get-ExperimentSaSnapshot {
    param([object] $Worksheet, [pscustomobject] $Block)
    if ((Get-ExperimentObjectProperty $Block 'SaHeaderRow' 1) -le 0) { return [pscustomobject]@{ SaRows = @(); BatchMean = $null } }
    $rows = @()
    for ($r = $Block.SampleStartRow; $r -lt $Block.BatchAverageRow; $r++) {
        $values = @(0..3 | ForEach-Object { $Worksheet.Cells.Item($r, ($Block.SaStartColumn + $_)).Value2 })
        if (@($values | Where-Object { -not [string]::IsNullOrWhiteSpace([string] $_) }).Count -eq 0) { continue }
        $rows += [pscustomobject]@{
            SampleIndex = $r - $Block.SampleStartRow + 1
            Sa1 = $values[0]; Sa2 = $values[1]; Sa3 = $values[2]; SaMean = $values[3]
        }
    }
    [pscustomobject]@{ SaRows = $rows; BatchMean = $Worksheet.Cells.Item($Block.BatchAverageRow, $Block.BatchAverageColumn).Value2; SampleCapacity = $Block.BatchAverageRow - $Block.SampleStartRow }
}

function Test-ExperimentSaSnapshotComplete {
    param([pscustomobject] $Snapshot, [AllowEmptyCollection()] [object[]] $ExpectedRows = @())
    $measured = @($ExpectedRows | Where-Object { $null -ne $_.SaMean })
    if ($ExpectedRows.Count -eq 0 -or $Snapshot.SaRows.Count -ne $measured.Count) { return $false }
    $indices = @{}
    foreach ($expected in $ExpectedRows) {
        $index = [int] $expected.SampleIndex
        if ($Snapshot.PSObject.Properties['SampleCapacity'] -and $index -gt $Snapshot.SampleCapacity) { return $false }
        if ($index -lt 1 -or $indices.ContainsKey($index)) { return $false }
        $hasValues = @(@($expected.Sa1, $expected.Sa2, $expected.Sa3) | Where-Object { $null -ne $_ }).Count -gt 0
        if ($hasValues -ne ($null -ne $expected.SaMean)) { return $false }
        $indices[$index] = $true
        $actual = @($Snapshot.SaRows | Where-Object { [int] $_.SampleIndex -eq $index })
        if (-not $hasValues) { if ($actual.Count -ne 0) { return $false }; continue }
        if ($actual.Count -ne 1) { return $false }
        foreach ($field in @('Sa1', 'Sa2', 'Sa3', 'SaMean')) {
            if (-not (Test-ExperimentSaNumberEqual -Actual $actual[0].$field -Expected $expected.$field)) { return $false }
        }
    }
    $mean = if ($measured.Count) { [Math]::Round(($measured | Measure-Object -Property SaMean -Average).Average, 6) } else { $null }
    return (Test-ExperimentSaNumberEqual -Actual $Snapshot.BatchMean -Expected $mean)
}
