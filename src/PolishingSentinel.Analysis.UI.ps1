function Set-SentinelAnalysisGrid {
    param($Grid,[object[]]$Rows)
    $old=$Grid.DataSource
    $Grid.DataSource=$null
    if ($old -is [Data.DataTable]) { $old.Dispose() }
    $table=New-Object Data.DataTable
    if ($Rows.Count) {
        foreach ($p in $Rows[0].PSObject.Properties) { [void]$table.Columns.Add($p.Name) }
        foreach ($row in $Rows) {
            $values=@($row.PSObject.Properties | ForEach-Object { [string]$_.Value })
            [void]$table.Rows.Add([object[]]$values)
        }
    }
    $Grid.DataSource=$table
}

function New-SentinelAnalysisPage {
    param($Form,[string]$ScriptRoot)
    $page=New-Object Windows.Forms.TabPage
    $page.Text='数据分析'; $page.Name='AnalysisPage'; $page.BackColor=[Drawing.Color]::FromArgb(244,246,249)
    $page.Font=New-Object Drawing.Font('Microsoft YaHei UI',13)
    $layout=New-Object Windows.Forms.TableLayoutPanel
    $layout.Dock='Fill'; $layout.Padding=New-Object Windows.Forms.Padding(12); $layout.ColumnCount=1; $layout.RowCount=4
    foreach ($h in @(46,58)) { [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute',$h))) }
    [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent',100)))
    [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute',56)))
    $tools=New-Object Windows.Forms.FlowLayoutPanel; $tools.Dock='Fill'
    $factor=New-Object Windows.Forms.ComboBox; $factor.Width=130; $factor.DropDownStyle='DropDownList'
    $factor.Items.AddRange([object[]](Get-SentinelAnalysisFactors)); $factor.SelectedIndex=0
    $experiment=New-Object Windows.Forms.ComboBox; $experiment.Width=140; $experiment.DropDownStyle='DropDownList'
    $date=New-Object Windows.Forms.ComboBox; $date.Width=140; $date.DropDownStyle='DropDownList'
    $date.Items.AddRange([object[]]@('全部日期','近90天','本年')); $date.SelectedIndex=0
    $refresh=New-Object Windows.Forms.Button; $refresh.Text='刷新分析'; $refresh.Width=130; $refresh.Height=34
    $tools.Controls.AddRange([Windows.Forms.Control[]]@($factor,$experiment,$date,$refresh))
    $status=New-Object Windows.Forms.Label; $status.Dock='Fill'; $status.Text='尚未读取水基实验'; $status.AutoEllipsis=$true
    $tabs=New-Object Windows.Forms.TabControl; $tabs.Dock='Fill'
    $grids=@{}
    foreach ($name in @('数据核对','单因素效应','对照组匹配','因素关联排名')) {
        $tab=New-Object Windows.Forms.TabPage; $tab.Text=$name
        $grid=New-Object Windows.Forms.DataGridView; $grid.Dock='Fill'; $grid.ReadOnly=$true
        $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.AutoSizeColumnsMode='DisplayedCells'
        $grid.RowHeadersVisible=$false; $grid.BackgroundColor=[Drawing.Color]::White; $grid.AutoSizeRowsMode='AllCells'
        $grid.DefaultCellStyle.Font=New-Object Drawing.Font('Microsoft YaHei UI',11)
        Set-SentinelGridTypography $grid
        $grid.Add_FontChanged({param($sender,$e) Set-SentinelGridTypography $sender})
        $tab.Controls.Add($grid); [void]$tabs.TabPages.Add($tab); $grids[$name]=$grid
    }
    $chart=New-Object Windows.Forms.DataVisualization.Charting.Chart
    $chart.Dock='Top'; $chart.Height=180; $chart.BackColor=[Drawing.Color]::White
    $area=New-Object Windows.Forms.DataVisualization.Charting.ChartArea('Effects')
    $area.AxisY.Title='实验 Sa (nm)'; $area.AxisY.Minimum=0; $area.AxisX.Interval=1
    $area.AxisX.MajorGrid.Enabled=$false; $area.AxisY.MajorGrid.LineColor=[Drawing.Color]::FromArgb(232,237,243)
    $chart.ChartAreas.Add($area); $tabs.TabPages[1].Controls.Add($chart)
    $legend=[Windows.Forms.DataVisualization.Charting.Legend]::new('统计量'); $legend.Docking='Top'; $legend.Alignment='Far'; $chart.Legends.Add($legend)
    $warning=New-Object Windows.Forms.Label; $warning.Dock='Fill'; $warning.Text='统计单位：实验；Sa单位：nm。单因素分组与η²未校正其他因素，不能作为因果结论。未知参数不自动补齐；对照候选须人工核对。'
    $layout.Controls.Add($tools,0,0); $layout.Controls.Add($status,0,1); $layout.Controls.Add($tabs,0,2); $layout.Controls.Add($warning,0,3); $page.Controls.Add($layout)
    $state=[pscustomobject]@{Records=@(); Worker=$null; Pending=$null; Timer=(New-Object Windows.Forms.Timer); Busy=$false; Closed=$false; Update=$null; Load=$null}
    $update={
        if ($state.Busy) { return }
        $cutoff=if ($date.SelectedIndex -eq 1) {(Get-Date).AddDays(-90).Date} elseif ($date.SelectedIndex -eq 2) {[datetime]::new((Get-Date).Year,1,1)} else {$null}
        $records=@($state.Records | Where-Object { $null -eq $cutoff -or ($_.Date -and [datetime]$_.Date -ge $cutoff) } | Sort-Object {[int]$_.No})
        $selected=[string]$experiment.SelectedItem
        $state.Busy=$true; $experiment.Items.Clear()
        foreach ($r in @($records | Sort-Object {[int]$_.No})) { [void]$experiment.Items.Add('实验 '+$r.No) }
        if ($experiment.Items.Contains($selected)) { $experiment.SelectedItem=$selected } elseif ($experiment.Items.Count) { $experiment.SelectedIndex=0 }
        $state.Busy=$false
        $f=[string]$factor.SelectedItem
        $rows=@(foreach ($r in $records) {
            $row=[ordered]@{实验号=$r.No; 日期=$r.Date; 'Sa(nm)'=if ($null -ne $r.Sa) {[Math]::Round($r.Sa,4)} else {$null}; 有效样品=$r.Samples; 状态=if ($r.Duplicate){'重号，排除'} elseif ($null -eq $r.Sa){'缺少Sa'} else {'可分析'}; 来源=($r.Sheet+' 行'+$r.Row)}
            foreach ($key in Get-SentinelAnalysisFields) { $row[$key]=$r.Factors[$key] }
            [pscustomobject]$row
        })
        Set-SentinelAnalysisGrid $grids['数据核对'] $rows
        $effects=@(Get-SentinelFactorEffects $records $f)
        Set-SentinelAnalysisGrid $grids['单因素效应'] $effects
        $chart.Series.Clear()
        foreach ($definition in @(@('平均Sa','平均Sa(nm)',[Drawing.Color]::FromArgb(37,99,235)),@('中位Sa','中位Sa(nm)',[Drawing.Color]::FromArgb(13,148,136)))) {
            $series=New-Object Windows.Forms.DataVisualization.Charting.Series($definition[0])
            $series.ChartType='Column'; $series.Color=$definition[2]
            foreach ($effect in $effects) {
                $i=$series.Points.AddXY([string]$effect.水平,[double]$effect.($definition[1]))
                $series.Points[$i].ToolTip='实验数 '+$effect.实验数+'；平均Sa '+$effect.'平均Sa(nm)'+'；中位Sa '+$effect.'中位Sa(nm)'+' nm'
            }
            $chart.Series.Add($series)
        }
        $target=$records | Where-Object No -eq (([string]$experiment.SelectedItem) -replace '^实验\s*','') | Select-Object -First 1
        Set-SentinelAnalysisGrid $grids['对照组匹配'] @(Get-SentinelControlCandidates $records $target $f | Sort-Object 优先级)
        Set-SentinelAnalysisGrid $grids['因素关联排名'] @(Get-SentinelFactorRanking $records | Sort-Object 关联强度η² -Descending)
    }.GetNewClosure()
    $state.Update=$update
    [void]$tabs.TabPages.Add((New-SentinelAIPage -Form $Form -AnalysisState $state -ScriptRoot $ScriptRoot))
    $moduleRoot=$script:PolishingSentinelUiRoot
    $load={
        if ($state.Pending -or $state.Closed) { return }
        $refresh.Enabled=$false; $status.Text='正在读取水基实验…'
        $state.Worker=[PowerShell]::Create()
        [void]$state.Worker.AddScript({ param($root,$configRoot)
            $ErrorActionPreference='Stop'
            . (Join-Path $root 'PolishingSentinel.Config.ps1')
            . (Join-Path $root 'PolishingSentinel.Dashboard.Data.ps1')
            . (Join-Path $root 'PolishingSentinel.Analysis.ps1')
            $settings=Read-PolishingSentinelSettings $configRoot
            $snapshot=Read-SentinelDashboardWorkbook $settings.MasterWorkbookPath $settings.SlurryMappings
            [pscustomobject]@{Records=@(Convert-SentinelAnalysisRecords $snapshot); Stamp=$snapshot.UpdatedAt}
        }.ToString()).AddArgument($moduleRoot).AddArgument($ScriptRoot)
        $state.Pending=$state.Worker.BeginInvoke(); $state.Timer.Start()
    }.GetNewClosure()
    $state.Load=$load
    $state.Timer.Interval=250
    $state.Timer.Add_Tick({
        if (-not $state.Pending -or -not $state.Pending.IsCompleted) { return }
        try {
            $result=@($state.Worker.EndInvoke($state.Pending))[-1]
            if ($state.Worker.HadErrors) { throw ($state.Worker.Streams.Error | Out-String) }
            $state.Records=@($result.Records)
            & $state.Update
            $status.Text='水基实验 {0} 次 · 读取时间 {1} · 不包含油基实验及加工任务' -f $state.Records.Count,$result.Stamp
        } catch { $status.Text='读取失败，当前结果未更新：'+$_.Exception.Message }
        finally { $state.Worker.Dispose(); $state.Worker=$null; $state.Pending=$null; $state.Timer.Stop(); $refresh.Enabled=$true }
    }.GetNewClosure())
    foreach ($control in @($factor,$experiment,$date)) { $control.Add_SelectedIndexChanged($update) }
    $refresh.Add_Click($load)
    $page.Add_Enter({ if (-not $state.Records.Count) { & $state.Load } }.GetNewClosure())
    $Form.Add_Disposed({ $state.Closed=$true; $state.Timer.Stop(); $state.Timer.Dispose(); if ($state.Worker) {$state.Worker.Stop();$state.Worker.Dispose()} }.GetNewClosure())
    $page.Tag=[pscustomobject]@{State=$state; Grids=$grids; Factor=$factor; Experiment=$experiment; Status=$status}
    return $page
}
