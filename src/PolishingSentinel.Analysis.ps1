function Get-SentinelAnalysisFactors {
    return @('浓度','粒径','时间','球胚','抛光垫','pH','滚道','间距','充盈度','压力','配球')
}

function Get-SentinelAnalysisFields {
    return @((Get-SentinelAnalysisFactors) + @('备注','盘体'))
}
function Get-SentinelRecordedPH {
    param([string[]]$Texts)
    $values=@(); $invalid=$false
    foreach ($text in $Texts) {
        $normalized=([string]$text).Normalize([Text.NormalizationForm]::FormKC)
        $markers=[regex]::Matches($normalized,'(?i)(?<![a-z])pH')
        $matchesPH=[regex]::Matches($normalized,'(?i)(?<![a-z])pH\s*(?:[=:]\s*)?(?<value>\d+(?:\.\d+)?)(?![\d.])')
        if ($markers.Count -ne $matchesPH.Count) {$invalid=$true}
        foreach ($m in $matchesPH) {
            $value=[double]::Parse($m.Groups['value'].Value,[Globalization.CultureInfo]::InvariantCulture)
            $tail=$normalized.Substring($m.Index+$m.Length)
            if ($value -gt 14 -or $tail -match '^\s*(?:[-~～至到]\s*\d|±)') {$invalid=$true}
            $values+=$value
        }
    }
    $unique=@($values | Select-Object -Unique)
    if ($invalid -or $unique.Count -ne 1) {return '未知'}
    return $unique[0].ToString('G',[Globalization.CultureInfo]::InvariantCulture)
}
function Remove-SentinelRecordedPHText {
    param([string]$Text)
    $text=$Text.Normalize([Text.NormalizationForm]::FormKC)
    return (($text -replace '(?i)(?<![a-z])pH\s*(?:[=:]\s*)?\d+(?:\.\d+)?','') -replace '\(\s*\)','').Trim(' ',',',';','、')
}

function Convert-SentinelAnalysisRecords {
    param($Snapshot)
    foreach ($record in @($Snapshot.Records | Where-Object Kind -eq 'Water')) {
        $f=@{}; foreach ($key in Get-SentinelAnalysisFields) { $f[$key]='未知' }
        foreach ($key in $record.Fields.Keys) {
            if ($f.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$record.Fields[$key])) { $f[$key]=([string]$record.Fields[$key]).Trim() }
        }
        $slurry=[string]$record.Fields['抛光液']
        $f['pH']=Get-SentinelRecordedPH @($slurry,[string]$record.Fields['备注'],$record.Title)
        if ($slurry -match '^\s*(\d+(?:\.\d+)?)\s*%') { $f['浓度']=$Matches[1]+'%' }
        if ($slurry -match '^\s*\d+(?:\.\d+)?\s*%\s*-?\s*(\d+(?:\.\d+)?)\s*(nm|[μµu]m)') {
            $nm=[double]::Parse($Matches[1],[Globalization.CultureInfo]::InvariantCulture)
            if ($Matches[2] -ne 'nm') { $nm*=1000 }
            $f['粒径']=$nm.ToString('G',[Globalization.CultureInfo]::InvariantCulture)+'nm'
        }
        if ($record.Title -match '(\d+(?:\.\d+)?)\s*(h|小时|min|分钟)\s*$') {
            $hours=[double]::Parse($Matches[1],[Globalization.CultureInfo]::InvariantCulture)
            if ($Matches[2] -in @('min','分钟')) { $hours/=60 }
            $f['时间']=$hours.ToString('G',[Globalization.CultureInfo]::InvariantCulture)+'h'
        }
        # Keep unparsed slurry details in matching: supplier/additives cannot disappear.
        $context=([regex]'^\s*\d+(?:\.\d+)?\s*%').Replace($slurry,'',1)
        $context=([regex]'^\s*-?\s*\d+(?:\.\d+)?\s*(nm|[μµu]m)').Replace($context,'',1).Trim()
        [pscustomobject]@{ No=$record.ExperimentNo; Date=$record.Date; Sa=$record.Sa; Samples=$record.Samples; Duplicate=$record.Duplicate; Factors=$f; Context=$context; Sheet=$record.Sheet; Row=$record.Row; Title=$record.Title }
    }
}

function Get-SentinelFactorEffects {
    param([object[]]$Records,[string]$Factor)
    $valid=@($Records | Where-Object { -not $_.Duplicate -and $null -ne $_.Sa -and $_.Factors[$Factor] -ne '未知' })
    foreach ($group in @($valid | Group-Object { $_.Factors[$Factor] } | Sort-Object { if ($_.Name -match '^\d+(?:\.\d+)?') {[double]::Parse($Matches[0],[Globalization.CultureInfo]::InvariantCulture)} else {[double]::MaxValue} },Name)) {
        $mean=($group.Group | Measure-Object Sa -Average).Average
        $ordered=@($group.Group | ForEach-Object {[double]$_.Sa} | Sort-Object)
        $middle=[int][Math]::Floor($ordered.Count/2)
        $median=if ($ordered.Count % 2) {$ordered[$middle]} else {($ordered[$middle-1]+$ordered[$middle])/2.0}
        $sd=$null
        if ($group.Count -gt 1) { $sum=0.0; foreach ($r in $group.Group) { $sum+=[Math]::Pow($r.Sa-$mean,2) }; $sd=[Math]::Sqrt($sum/($group.Count-1)) }
        [pscustomobject]@{ 水平=$group.Name; 实验数=$group.Count; '平均Sa(nm)'=[Math]::Round($mean,4); '中位Sa(nm)'=[Math]::Round($median,4); '最小Sa(nm)'=[Math]::Round($ordered[0],4); '最大Sa(nm)'=[Math]::Round($ordered[-1],4); '实验间SD(nm)'=if ($null -ne $sd) {[Math]::Round($sd,4)} else {'不足2次'}; 实验号=($group.Group.No -join ', ') }
    }
}

function Get-SentinelControlCandidates {
    param([object[]]$Records,$Target,[string]$Factor)
    if (-not $Target) { return }
    foreach ($candidate in $Records) {
        if ($candidate.No -eq $Target.No -or $candidate.Duplicate -or $null -eq $candidate.Sa) { continue }
        $different=@(); $unknown=@()
        foreach ($key in Get-SentinelAnalysisFields) {
            if ($key -eq $Factor) { continue }
            if ($Target.Factors[$key] -eq '未知' -or $candidate.Factors[$key] -eq '未知') { $unknown+=$key }
            else {
                $left=$Target.Factors[$key]; $right=$candidate.Factors[$key]
                if ($key -eq '备注' -and $Target.Factors['pH'] -ne '未知' -and $candidate.Factors['pH'] -ne '未知') {
                    $left=Remove-SentinelRecordedPHText $left; $right=Remove-SentinelRecordedPHText $right
                }
                if ($left -cne $right) { $different+=$key }
            }
        }
        $leftContext=$Target.Context; $rightContext=$candidate.Context
        if ($Target.Factors['pH'] -ne '未知' -and $candidate.Factors['pH'] -ne '未知') {
            $leftContext=Remove-SentinelRecordedPHText $leftContext; $rightContext=Remove-SentinelRecordedPHText $rightContext
        }
        if ($leftContext -cne $rightContext) { $different+='抛光液其他成分/原文' }
        $contrast=$Target.Factors[$Factor] -ne '未知' -and $candidate.Factors[$Factor] -ne '未知' -and $Target.Factors[$Factor] -cne $candidate.Factors[$Factor]
        $strict=$contrast -and $different.Count -eq 0 -and $unknown.Count -eq 0 -and -not $Target.Duplicate -and $null -ne $Target.Sa
        [pscustomobject]@{ 实验号=$candidate.No; 日期=$candidate.Date; 目标因素=$candidate.Factors[$Factor]; 'Sa(nm)'=$candidate.Sa; 判定=if ($strict) {'条件一致候选，待人工确认'} else {'仅供核对，不是合格对照'}; 其他差异=($different -join '、'); 未知项=($unknown -join '、'); 因素有差异=$contrast; 优先级=($different.Count*10+$unknown.Count) }
    }
}

function Get-SentinelFactorRanking {
    param([object[]]$Records)
    foreach ($factor in Get-SentinelAnalysisFactors) {
        $valid=@($Records | Where-Object { -not $_.Duplicate -and $null -ne $_.Sa -and $_.Factors[$factor] -ne '未知' })
        $groups=@($valid | Group-Object { $_.Factors[$factor] })
        $reason=''; $eta=$null
        if ($valid.Count -lt 10) { $reason='有效实验少于10次' }
        elseif ($groups.Count -lt 2) { $reason='不足两个因素水平' }
        elseif (@($groups | Where-Object Count -lt 3).Count) { $reason='部分水平不足3次实验' }
        else {
            $mean=($valid | Measure-Object Sa -Average).Average; $total=0.0; $between=0.0
            foreach ($r in $valid) { $total+=[Math]::Pow($r.Sa-$mean,2) }
            foreach ($g in $groups) { $between+=$g.Count*[Math]::Pow(($g.Group | Measure-Object Sa -Average).Average-$mean,2) }
            if ($total -gt 0) { $eta=[Math]::Round($between/$total,4); $reason='未调整混杂，仅描述性关联' } else { $reason='Sa无变异' }
        }
        [pscustomobject]@{ 因素=$factor; '关联强度η²'=$eta; 有效实验=$valid.Count; 水平数=$groups.Count; 说明=$reason }
    }
}
