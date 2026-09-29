Add-Type -AssemblyName System.Windows.Forms.DataVisualization

function Set-SentinelDashboardChart {
    param($View)
    Hide-SentinelChartHover $View.Hover
    $View.Chart.Series.Clear()
    if (-not $View.Snapshot) { $View.Empty.Text='统计数据尚未就绪'; $View.Empty.Visible=$true; $View.Quality.Text=''; return }
    $range=@('All','Month','Week')[$View.Range.SelectedIndex]
    $kind=@('All','Water','Oil')[$View.Kind.SelectedIndex]
    $points=@(Get-SentinelDashboardPoints $View.Snapshot $range $kind)
    $valid=0
    foreach ($type in @('Water','Oil')) {
        $segment=$null; $segmentIndex=0
        foreach ($record in @($points | Where-Object Kind -eq $type)) {
            if ($null -eq $record.Sa) { $segment=$null; continue }
            if (-not $segment) {
                $segment=New-Object Windows.Forms.DataVisualization.Charting.Series($type+$segmentIndex)
                $segment.ChartType='Line'
                $segment.XValueType='DateTime'
                $segment.BorderWidth=2
                $segment.MarkerStyle=if ($type -eq 'Water') { 'Circle' } else { 'Diamond' }
                $segment.MarkerSize=6
                $segment.Color=if ($type -eq 'Water') { [Drawing.Color]::FromArgb(37,99,235) } else { [Drawing.Color]::FromArgb(13,148,136) }
                $segment.LegendText=if ($type -eq 'Water') { '水基抛光' } else { '油基抛光' }
                $segment.IsVisibleInLegend=$segmentIndex -eq 0
                $segmentIndex++
                $View.Chart.Series.Add($segment)
            }
            $index=$segment.Points.AddXY(([datetime]$record.Date).ToOADate(),[double]$record.Sa)
            $origin=if ($record.Derived) { '样品均值汇总' } else { '主表单批均值' }
            $segment.Points[$index].Tag="实验$($record.ExperimentNo) · $($record.Date)`nSa：$([Math]::Round($record.Sa,4)) nm`n有效样品：$($record.Samples) · $origin"
            $valid++
        }
    }
    $View.Chart.ChartAreas[0].RecalculateAxesScale()
    $View.Empty.Text='所选范围暂无有效 Sa 数据'
    $View.Empty.Visible=$valid -eq 0
    $View.Chart.Invalidate()
    $unknown=@($View.Snapshot.Records | Where-Object { -not $_.Date }).Count
    $View.Quality.Text='趋势点 {0} 个  ·  缺少 Sa {1} 项  ·  日期缺失 {2} 项  ·  重号 {3} 项  ·  未编号待核对 {4} 项' -f $valid,(Get-SentinelDashboardSummary $View.Snapshot).Missing,$unknown,$View.Snapshot.DuplicateCount,$View.Snapshot.UnnumberedCount
}

function Start-SentinelDashboardRefresh {
    param($Form,$View,[string]$ScriptRoot)
    $controller=[pscustomobject]@{ Timer=(New-Object Windows.Forms.Timer); Worker=$null; Pending=$null; Next=[datetime]::MinValue; Force=$false; Closed=$false; Refresh=$null }
    $module=Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.Dashboard.Data.ps1'
    $config=Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.Config.ps1'
    $tick={
        if ($controller.Closed) { return }
        if ($controller.Pending) {
            if (-not $controller.Pending.IsCompleted) { return }
            try {
                $result=@($controller.Worker.EndInvoke($controller.Pending))[-1]
                if ($controller.Worker.HadErrors) { throw ($controller.Worker.Streams.Error | Out-String) }
                $View.Snapshot=$result.Snapshot
                if ($result.Snapshot) {
                    $summary=Get-SentinelDashboardSummary $result.Snapshot
                    foreach ($key in $View.Metrics.Keys) { $View.Metrics[$key].Text=[string]$summary.$key }
                    $View.Stamp.Text='统计更新：'+$result.Snapshot.UpdatedAt
                    $View.Tasks.Text='加工任务（独立分类）   已编号 {0} 条   /   未编号待核对 {1} 条' -f @($result.Snapshot.TaskRecords).Count,$result.Snapshot.TaskUnnumberedCount
                } else {
                    foreach ($label in $View.Metrics.Values) { $label.Text='--' }
                    $View.Stamp.Text='暂无可用统计'
                    $View.Tasks.Text='加工任务（独立分类）   --'
                }
                $View.Error.Text=if ($result.Error) { '统计未更新：'+$result.Error } else { '' }
                $View.Error.Visible=[bool]$result.Error
                Set-SentinelDashboardChart $View
            } catch {
                $View.Error.Text='统计读取失败：'+$_.Exception.Message
                $View.Error.Visible=$true
            } finally {
                $controller.Worker.Dispose(); $controller.Worker=$null; $controller.Pending=$null
                $controller.Next=(Get-Date).AddSeconds(15)
                $View.Refresh.Enabled=$true
            }
        }
        if ((Get-Date) -lt $controller.Next -and -not $controller.Force) { return }
        $View.Refresh.Enabled=$false
        $controller.Worker=[PowerShell]::Create()
        [void]$controller.Worker.AddScript({
            param($module,$config,$root,$force)
            $ErrorActionPreference='Stop'
            . $config
            . $module
            Get-SentinelDashboardData -ScriptRoot $root -Force:$force
        }.ToString()).AddArgument($module).AddArgument($config).AddArgument($ScriptRoot).AddArgument($controller.Force)
        $controller.Force=$false
        $controller.Pending=$controller.Worker.BeginInvoke()
    }.GetNewClosure()
    $controller.Refresh=$tick
    $controller.Timer.Interval=1000
    $controller.Timer.Add_Tick($tick)
    $View.Refresh.Add_Click({ $controller.Force=$true; & $controller.Refresh }.GetNewClosure())
    $Form.Add_Shown({ & $controller.Refresh; $controller.Timer.Start() }.GetNewClosure())
    $Form.Add_Disposed({
        $controller.Closed=$true
        $controller.Timer.Stop(); $controller.Timer.Dispose()
        if ($controller.Worker) { $controller.Worker.Stop(); $controller.Worker.Dispose(); $controller.Worker=$null }
        $controller.Pending=$null; $controller.Refresh=$null
    }.GetNewClosure())
    $View.Controller=$controller
}

function Initialize-SentinelDashboard {
    param($Form,$Tabs,$Overview,[string]$ScriptRoot)
    $theme=Get-SentinelUiTheme
    $state=$Form.Controls.Find('StatusStateLabel',$true)[0]
    $message=$Form.Controls.Find('StatusMessageLabel',$true)[0]
    $updated=$Form.Controls.Find('StatusUpdatedAtLabel',$true)[0]
    $health=$Form.Controls.Find('StatusHealthLabel',$true)[0]
    $scan=$Form.Controls.Find('RunScanNowButton',$true)[0]
    $refreshStatus=$Form.Controls.Find('RefreshStatusButton',$true)[0]
    $old=@($Overview.Controls)
    $Overview.Controls.Clear()
    $Overview.Padding=New-Object Windows.Forms.Padding(22,12,22,16)
    $Overview.AutoScroll=$true
    $layout=New-Object Windows.Forms.TableLayoutPanel
    $layout.Dock='Fill'; $layout.ColumnCount=1; $layout.RowCount=6
    [void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent',100)))
    foreach ($height in @(138,112,36,44)) { [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute',$height))) }
    [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent',100)))
    [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute',36)))
    $layout.MinimumSize=New-Object Drawing.Size(740,490)
    $Overview.Controls.Add($layout)
    $statusBand=New-Object Windows.Forms.TableLayoutPanel
    $statusBand.Dock='Fill'; $statusBand.ColumnCount=1; $statusBand.RowCount=5
    [void]$statusBand.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent',100)))
    foreach ($label in @($message,$health)) { $label.Dock='Fill'; $label.AutoSize=$false; $label.Height=24; $label.AutoEllipsis=$true }
    $errorLabel=New-Object Windows.Forms.Label
    $errorLabel.Dock='Fill'; $errorLabel.ForeColor=[Drawing.Color]::Firebrick; $errorLabel.Visible=$false
    $errorLabel.AutoEllipsis=$true
    $progressText=New-Object Windows.Forms.Label
    $progressText.Name='ProcessingProgressText'; $progressText.Dock='Fill'; $progressText.AutoEllipsis=$true
    $progressText.ForeColor=$theme.Primary
    $progressBar=New-Object SentinelGradientProgress
    $progressBar.Name='ProcessingProgressBar'; $progressBar.Dock='Fill'; $progressBar.Maximum=100
    $statusBand.Controls.Add($message,0,0); $statusBand.Controls.Add($health,0,1)
    $statusBand.Controls.Add($progressText,0,2); $statusBand.Controls.Add($progressBar,0,3); $statusBand.Controls.Add($errorLabel,0,4)
    foreach ($height in @(27,27,28,28,26)) { [void]$statusBand.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute',$height))) }
    $layout.Controls.Add($statusBand,0,0)
    $metrics=New-Object Windows.Forms.TableLayoutPanel
    $metrics.Dock='Fill'; $metrics.ColumnCount=4
    $values=@{}
    $definitions=@(@('Total','实验总数','全部工艺'),@('Oil','油基实验','油基抛光'),@('Water','水基实验','水基抛光'),@('Week','本周实验','周一至今日'))
    $colors=@([Drawing.Color]::FromArgb(37,99,235),[Drawing.Color]::FromArgb(13,148,136),[Drawing.Color]::FromArgb(37,99,235),[Drawing.Color]::FromArgb(217,119,6))
    for ($i=0;$i -lt 4;$i++) {
        [void]$metrics.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent',25)))
        $panel=New-Object Windows.Forms.Panel
        $panel.Dock='Fill'; $panel.BackColor=[Drawing.Color]::White
        $panel.Margin=New-Object Windows.Forms.Padding(0,4,12,8)
        $accent=New-Object Windows.Forms.Panel
        $accent.Dock='Left'; $accent.Width=3; $accent.BackColor=$colors[$i]
        $title=New-SentinelLabel -Text $definitions[$i][1] -X 18 -Y 12 -Width 170
        $title.ForeColor=$theme.MutedInk
        $titleStyle=if ($i -lt 3) { [Drawing.FontStyle]::Bold } else { [Drawing.FontStyle]::Regular }
        $title.Font=New-Object Drawing.Font('Microsoft YaHei UI',15,$titleStyle)
        $title.Height=32
        $value=New-SentinelLabel -Text '--' -X 18 -Y 45 -Width 170 -Name ('Metric'+$definitions[$i][0])
        $value.Height=42; $value.Font=New-Object Drawing.Font('Segoe UI',24,[Drawing.FontStyle]::Bold)
        $value.ForeColor=$theme.Ink
        $panel.Controls.AddRange([Windows.Forms.Control[]]@($title,$value,$accent))
        $metrics.Controls.Add($panel,$i,0)
        $values[$definitions[$i][0]]=$value
    }
    $layout.Controls.Add($metrics,0,1)
    $tasks=New-Object Windows.Forms.Label
    $tasks.Name='DashboardTasksSummary'; $tasks.Dock='Fill'; $tasks.TextAlign='MiddleLeft'
    $tasks.ForeColor=$theme.MutedInk; $tasks.Text='加工任务（独立分类）   --'
    $layout.Controls.Add($tasks,0,2)
    $chartTools=New-Object Windows.Forms.TableLayoutPanel
    $chartTools.Dock='Fill'; $chartTools.ColumnCount=5
    [void]$chartTools.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent',100)))
    foreach ($width in @(140,130,42,8)) { [void]$chartTools.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute',$width))) }
    $chartTitle=New-SentinelLabel -Text 'Sa 批次均值趋势' -X 0 -Y 0 -Width 230
    $chartTitle.Dock='Fill'; $chartTitle.TextAlign='MiddleLeft'; $chartTitle.Font=New-Object Drawing.Font('Microsoft YaHei UI',15,[Drawing.FontStyle]::Bold)
    $kind=New-Object Windows.Forms.ComboBox
    $kind.Name='DashboardKind'; $kind.DropDownStyle='DropDownList'; $kind.Dock='Fill'
    $kind.Items.AddRange([object[]]@('全部工艺','水基抛光','油基抛光')); $kind.SelectedIndex=0
    $range=New-Object Windows.Forms.ComboBox
    $range.Name='DashboardRange'; $range.DropDownStyle='DropDownList'; $range.Dock='Fill'
    $range.Items.AddRange([object[]]@('全部日期','近30天','本周')); $range.SelectedIndex=0
    $refresh=New-SentinelButton -Text '↻' -X 0 -Y 0 -Width 34 -Name 'RefreshDashboardButton'
    Set-SentinelButtonStyle $refresh $theme
    foreach ($combo in @($kind,$range)) { $combo.FlatStyle='Flat'; $combo.BackColor=[Drawing.Color]::White; $combo.ForeColor=$theme.Ink }
    $tip=New-Object Windows.Forms.ToolTip
    $tip.SetToolTip($refresh,'重新读取主表统计')
    $chartTools.Controls.Add($chartTitle,0,0); $chartTools.Controls.Add($kind,1,0)
    $chartTools.Controls.Add($range,2,0); $chartTools.Controls.Add($refresh,3,0)
    $layout.Controls.Add($chartTools,0,3)
    $chartHost=New-Object Windows.Forms.Panel
    $chartHost.Dock='Fill'; $chartHost.BackColor=[Drawing.Color]::White
    $chart=New-Object Windows.Forms.DataVisualization.Charting.Chart
    $chart.Name='SaTrendChart'; $chart.Dock='Fill'; $chart.BackColor=[Drawing.Color]::White
    $area=New-Object Windows.Forms.DataVisualization.Charting.ChartArea('Sa')
    $area.BackColor=[Drawing.Color]::White
    $area.AxisX.LabelStyle.Format='yy-MM-dd'
    $area.AxisX.MajorGrid.Enabled=$false
    $area.AxisX.LineColor=[Drawing.Color]::LightGray
    $area.AxisY.LineColor=[Drawing.Color]::Transparent
    $area.AxisY.MajorGrid.LineColor=[Drawing.Color]::FromArgb(232,237,243)
    $area.AxisY.Title='Sa (nm)'; $area.AxisY.Minimum=0
    $area.AxisY.LabelStyle.Format='0.###'
    $area.AxisX.Title='实验日期'
    $area.AxisX.LabelStyle.Font=New-Object Drawing.Font('Segoe UI',10)
    $area.AxisY.LabelStyle.Font=New-Object Drawing.Font('Segoe UI',10)
    $area.AxisX.TitleFont=New-Object Drawing.Font('Microsoft YaHei UI',15)
    $area.AxisY.TitleFont=New-Object Drawing.Font('Microsoft YaHei UI',15)
    $chart.ChartAreas.Add($area)
    $legend=New-Object Windows.Forms.DataVisualization.Charting.Legend('工艺')
    $legend.Docking='Top'; $legend.Alignment='Far'; $legend.Font=New-Object Drawing.Font('Microsoft YaHei UI',15)
    $chart.Legends.Add($legend)
    $empty=New-Object Windows.Forms.Label
    $empty.Text='正在读取实验统计…'; $empty.AutoSize=$true; $empty.BackColor=[Drawing.Color]::White
    $empty.Location=New-Object Drawing.Point(60,100); $empty.ForeColor=$theme.MutedInk
    $chartHost.Controls.Add($chart); $chartHost.Controls.Add($empty); $empty.BringToFront()
    $hover=New-Object Windows.Forms.Label
    $hover.Name='SaPointPopup'; $hover.AutoSize=$true; $hover.MaximumSize=New-Object Drawing.Size(420,0)
    $hover.Padding=New-Object Windows.Forms.Padding(12)
    $hover.Font=New-Object Drawing.Font('Microsoft YaHei UI',12,[Drawing.FontStyle]::Bold)
    $hover.BackColor=[Drawing.Color]::FromArgb(15,43,74); $hover.ForeColor=[Drawing.Color]::White
    $hover.Visible=$false; $chart.Controls.Add($hover)
    $chart.Add_MouseMove({ param($sender,$e) Update-SentinelChartHover $chart $hover $e.X $e.Y }.GetNewClosure())
    $chart.Add_MouseLeave({ Hide-SentinelChartHover $hover }.GetNewClosure())
    $chart.Add_Resize({ Hide-SentinelChartHover $hover }.GetNewClosure())
    $layout.Controls.Add($chartHost,0,4)
    $quality=New-Object Windows.Forms.Label
    $quality.Dock='Fill'; $quality.ForeColor=$theme.MutedInk; $quality.TextAlign='MiddleLeft'
    $layout.Controls.Add($quality,0,5)

    $shell=New-Object Windows.Forms.Panel
    $shell.Name='DashboardShell'; $shell.Dock='Fill'
    $sidebar=New-Object Windows.Forms.Panel
    $sidebar.Name='Sidebar'; $sidebar.Dock='Left'; $sidebar.Width=184
    $sidebar.BackColor=[Drawing.Color]::FromArgb(5,24,41)
    $content=New-Object Windows.Forms.Panel
    $content.Dock='Fill'
    $header=New-Object Windows.Forms.Panel
    $header.Dock='Top'; $header.Height=66; $header.BackColor=[Drawing.Color]::White
    $pageTitle=New-SentinelLabel -Text '实验总览' -X 24 -Y 18 -Width 250
    $pageTitle.Font=New-Object Drawing.Font('Microsoft YaHei UI',16,[Drawing.FontStyle]::Bold); $pageTitle.Height=36
    $header.Controls.Add($pageTitle)
    $actions=New-Object Windows.Forms.FlowLayoutPanel
    $actions.Dock='Right'; $actions.Width=310; $actions.Padding=New-Object Windows.Forms.Padding(8,16,0,0)
    foreach ($button in @($refreshStatus,$scan)) { $button.Width=136; $button.Height=36 }
    $actions.Controls.Add($refreshStatus); $actions.Controls.Add($scan)
    $scan.BackColor=$theme.Primary; $scan.ForeColor=[Drawing.Color]::White
    $header.Controls.Add($actions)
    $clip=New-Object Windows.Forms.Panel
    $clip.Dock='Fill'
    $Tabs.Dock='None'; $Tabs.Location=New-Object Drawing.Point(-4,-40)
    $Tabs.Size=New-Object Drawing.Size(($clip.Width+8),($clip.Height+44))
    $Tabs.Anchor='Top,Bottom,Left,Right'
    $clip.Controls.Add($Tabs)
    $clip.Add_Resize({
        $Tabs.Bounds=New-Object Drawing.Rectangle(-4,-40,($clip.ClientSize.Width+8),($clip.ClientSize.Height+44))
    }.GetNewClosure())
    $settingsNav=New-Object Windows.Forms.FlowLayoutPanel
    $settingsNav.Name='SettingsNavigation'; $settingsNav.Dock='Top'; $settingsNav.Height=56
    $settingsNav.Padding=New-Object Windows.Forms.Padding(22,8,0,4)
    $settingsNav.WrapContents=$false; $settingsNav.Visible=$false
    $settingsState=[pscustomobject]@{ LastIndex=1 }
    $settingsButtons=@()
    $settingsNames=@('路径设置','命名约定','匹配规则','固定工艺参数')
    for ($i=1;$i -le 4;$i++) {
        $button=New-Object Windows.Forms.Button
        $button.Name='SettingsNav'+$i; $button.Text=$settingsNames[$i-1]; $button.Tag=$i
        $button.Size=New-Object Drawing.Size(180,40)
        $button.Font=New-Object Drawing.Font('Microsoft YaHei UI',13)
        $button.FlatStyle='Flat'; $button.FlatAppearance.BorderSize=1
        $button.Cursor=[Windows.Forms.Cursors]::Hand
        $button.Add_Click({ param($sender,$e) $Tabs.SelectedIndex=[int]$sender.Tag }.GetNewClosure())
        $settingsNav.Controls.Add($button); $settingsButtons+=$button
    }
    $content.Controls.Add($clip); $content.Controls.Add($settingsNav); $content.Controls.Add($header)
    $shell.Controls.Add($content); $shell.Controls.Add($sidebar)
    $Form.Controls.Add($shell)
    $brand=New-SentinelLabel -Text '抛光哨站' -X 22 -Y 24 -Width 158
    $brand.Font=New-Object Drawing.Font('Microsoft YaHei UI',16,[Drawing.FontStyle]::Bold)
    $brand.Height=38; $brand.ForeColor=[Drawing.Color]::White
    $sidebar.Controls.Add($brand)
    $caption=New-SentinelLabel -Text '实验数据控制台' -X 24 -Y 64 -Width 154
    $caption.ForeColor=[Drawing.Color]::FromArgb(133,161,184)
    $sidebar.Controls.Add($caption)
    $nav=@()
    $names=@('实验总览','路径设置','命名约定','匹配规则','固定工艺参数','待确认与日志','自启动','数据分析')
    $navIndices=@(0,7,1,5,6)
    for ($position=0;$position -lt $navIndices.Count;$position++) {
        $i=$navIndices[$position]
        $button=New-Object Windows.Forms.Button
        $button.Name='Nav'+$i; $button.Text=if ($i -eq 1) { '软件配置' } else { $names[$i] }; $button.Tag=$i
        $button.FlatStyle='Flat'; $button.FlatAppearance.BorderSize=0
        $button.TextAlign='MiddleLeft'; $button.Padding=New-Object Windows.Forms.Padding(22,0,0,0)
        $button.Location=New-Object Drawing.Point(0,(116+$position*46))
        $button.Size=New-Object Drawing.Size(184,44)
        $button.Font=New-Object Drawing.Font('Microsoft YaHei UI',15)
        $button.ForeColor=[Drawing.Color]::FromArgb(198,215,230)
        $button.Cursor=[Windows.Forms.Cursors]::Hand
        $button.Add_Click({
            param($sender,$e)
            $Tabs.SelectedIndex=if ([int]$sender.Tag -eq 1) { $settingsState.LastIndex } else { [int]$sender.Tag }
        }.GetNewClosure())
        $sidebar.Controls.Add($button); $nav+=$button
    }
    $updateNav={
        $isSettings=$Tabs.SelectedIndex -in 1..4
        if ($isSettings) { $settingsState.LastIndex=$Tabs.SelectedIndex }
        $settingsNav.Visible=$isSettings
        for ($j=0;$j -lt $nav.Count;$j++) {
            $selected=($Tabs.SelectedIndex -eq [int]$nav[$j].Tag) -or ($isSettings -and [int]$nav[$j].Tag -eq 1)
            $nav[$j].BackColor=if ($selected) { $theme.Primary } else { $sidebar.BackColor }
        }
        foreach ($button in $settingsButtons) {
            $selected=$Tabs.SelectedIndex -eq [int]$button.Tag
            $button.BackColor=if ($selected) { $theme.Primary } else { [Drawing.Color]::White }
            $button.ForeColor=if ($selected) { [Drawing.Color]::White } else { $theme.Ink }
            $button.FlatAppearance.BorderColor=if ($selected) { $theme.Primary } else { [Drawing.Color]::LightGray }
        }
        $pageTitle.Text=if ($isSettings) { '软件配置' } else { $names[$Tabs.SelectedIndex] }
    }.GetNewClosure()
    $Tabs.Add_SelectedIndexChanged($updateNav); & $updateNav
    $footer=New-Object Windows.Forms.Panel
    $footer.Dock='Bottom'; $footer.Height=176
    $state.Location=New-Object Drawing.Point(20,12); $state.Size=New-Object Drawing.Size(158,24)
    $state.ForeColor=[Drawing.Color]::FromArgb(198,215,230)
    $updated.Location=New-Object Drawing.Point(20,46); $updated.Size=New-Object Drawing.Size(158,54)
    $updated.ForeColor=[Drawing.Color]::FromArgb(167,190,210)
    $stamp=New-SentinelLabel -Text '统计尚未更新' -X 20 -Y 104 -Width 158
    $stamp.Height=70; $stamp.ForeColor=[Drawing.Color]::FromArgb(167,190,210)
    $footer.Controls.AddRange([Windows.Forms.Control[]]@($state,$updated,$stamp))
    $sidebar.Controls.Add($footer)
    foreach ($control in $old) { $control.Dispose() }
    $view=[pscustomobject]@{
        Snapshot=$null; Metrics=$values; Chart=$chart; Range=$range; Kind=$kind; Empty=$empty; Quality=$quality
        Stamp=$stamp; Error=$errorLabel; Refresh=$refresh; Controller=$null; Tasks=$tasks
        Hover=$hover; ProgressBar=$progressBar; ProgressText=$progressText
    }
    $Overview.Tag=$view
    $range.Add_SelectedIndexChanged({ Set-SentinelDashboardChart $view }.GetNewClosure())
    $kind.Add_SelectedIndexChanged({ Set-SentinelDashboardChart $view }.GetNewClosure())
    $Form.Add_Disposed({ $tip.Dispose() }.GetNewClosure())
    Start-SentinelDashboardRefresh $Form $view $ScriptRoot
}
