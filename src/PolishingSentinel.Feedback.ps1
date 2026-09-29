function Set-SentinelProcessingProgress {
    param($Status,$Bar,$Label)
    $state=[string]$Status.State
    $data=if ($Status.PSObject.Properties['Data']) { $Status.Data } else { $null }
    $counts=@{}
    foreach ($name in @('TotalFolders','FinishedFolders','CompletedFolders','FailedFolders','DeferredFolders','PendingRows')) {
        $value=0
        if ($data -is [Collections.IDictionary]) { $raw=$data[$name] }
        elseif ($data -and $data.PSObject.Properties[$name]) { $raw=$data.$name }
        else { $raw=$null }
        if ($null -ne $raw) { [void][int]::TryParse([string]$raw,[ref]$value) }
        $counts[$name]=[Math]::Max(0,$value)
    }
    $active=$state -in @('Scanning','CheckingUsb','Recovering','ValidatingSa','RunningMatlab','ExtractingSa','WritingExcel','Processing')
    $marquee=$active -and $counts.TotalFolders -eq 0
    $style=if ($marquee) { 'Marquee' } else { 'Continuous' }
    $speed=if ($marquee) { 30 } else { 0 }
    if ([string]$Bar.Style -ne $style) { $Bar.Style=$style }
    if ($Bar.MarqueeAnimationSpeed -ne $speed) { $Bar.MarqueeAnimationSpeed=$speed }
    $value=0
    if ($counts.TotalFolders -gt 0 -and $state -ne 'DryRun') {
        $value=[int][Math]::Floor(100.0*[Math]::Min($counts.TotalFolders,$counts.FinishedFolders)/$counts.TotalFolders)
    }
    # Reassigning 0 then 100 every poll restarts the native progress animation.
    if ($Bar.Value -ne $value) { $Bar.Value=$value }
    if ($state -eq 'DryRun') { $Label.Text='预检查完成，未执行数据处理'; return }
    if ($counts.TotalFolders -gt 0) {
        $finished=[Math]::Min($counts.TotalFolders,$counts.FinishedFolders)
        $prefix=if ($active) { '本轮进度' } else { '本轮结果' }
        $Label.Text='{0} {1}/{2}（{3}%） · Sa已保存 {4} · 失败 {5} · 处理暂缓 {6} · 回写待确认 {7}' -f $prefix,$finished,$counts.TotalFolders,$Bar.Value,$counts.CompletedFolders,$counts.FailedFolders,$counts.DeferredFolders,$counts.PendingRows
        if ($state -in @('Exited','Stale','TimedOut','Error')) { $Label.Text='已中断 / 异常 · '+$Label.Text }
    } elseif ($active) {
        $Label.Text='正在扫描、核对或准备数据，处理总量尚未确定'
    } else {
        $Label.Text=if ($state -eq 'Completed') { '本轮没有新增处理任务' } elseif ($state -eq 'WaitingForUsb') { '等待 U 盘，暂无处理任务' } else { '暂无可用处理进度' }
        if ($counts.PendingRows -gt 0) { $Label.Text += ' · 回写待确认 ' + $counts.PendingRows }
    }
}

function Hide-SentinelChartHover {
    param($Popup)
    $Popup.Visible=$false
    if ($Popup.Tag) { $Popup.Tag.MarkerSize=6; $Popup.Tag=$null }
}

function Update-SentinelChartHover {
    param($Chart,$Popup,[int]$X,[int]$Y)
    if (-not $Chart.Series.Count) { Hide-SentinelChartHover $Popup; return }
    $area=$Chart.ChartAreas[0]
    $closest=$null; $distance=169.0
    foreach ($series in $Chart.Series) {
        foreach ($point in $series.Points) {
            if ($point.IsEmpty) { continue }
            $px=$area.AxisX.ValueToPixelPosition($point.XValue)
            $py=$area.AxisY.ValueToPixelPosition($point.YValues[0])
            $d=($px-$X)*($px-$X)+($py-$Y)*($py-$Y)
            if ($d -le $distance) { $distance=$d; $closest=$point }
        }
    }
    if (-not $closest) { Hide-SentinelChartHover $Popup; return }
    if ($Popup.Tag -ne $closest) {
        Hide-SentinelChartHover $Popup
        $Popup.Tag=$closest; $closest.MarkerSize=10
        $Popup.Text=[string]$closest.Tag
    }
    $left=$X+18; $top=$Y+18
    if ($left+$Popup.Width -gt $Chart.ClientSize.Width-8) { $left=$X-$Popup.Width-18 }
    if ($top+$Popup.Height -gt $Chart.ClientSize.Height-8) { $top=$Y-$Popup.Height-18 }
    $Popup.Location=New-Object Drawing.Point(([Math]::Max(8,$left)),([Math]::Max(8,$top)))
    $Popup.Visible=$true; $Popup.BringToFront()
}
