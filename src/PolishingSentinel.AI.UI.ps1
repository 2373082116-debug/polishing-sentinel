function New-SentinelAIPage {
    param($Form,$AnalysisState,[string]$ScriptRoot)
    $page=[Windows.Forms.TabPage]::new('AI工艺建议'); $page.Name='AIAdvicePage'
    $layout=[Windows.Forms.TableLayoutPanel]::new(); $layout.Dock='Fill'; $layout.ColumnCount=1; $layout.RowCount=5
    foreach ($height in @(90,84,44,48)) {[void]$layout.RowStyles.Add([Windows.Forms.RowStyle]::new('Absolute',$height))}
    [void]$layout.RowStyles.Add([Windows.Forms.RowStyle]::new('Percent',100))
    $settings=Read-SentinelAISettings $ScriptRoot
    $config=[Windows.Forms.FlowLayoutPanel]::new(); $config.Dock='Fill'; $config.AutoScroll=$true
    $fields=@{}
    foreach ($definition in @(@('Endpoint','HTTPS地址',310,$settings.Endpoint),@('Model','模型',190,$settings.Model),@('Key','API Key',190,''),@('Timeout','超时秒',80,[string]$settings.Timeout))) {
        $group=[Windows.Forms.Panel]::new(); $group.Width=[int]$definition[2]; $group.Height=68
        $label=[Windows.Forms.Label]::new(); $label.Text=$definition[1]; $label.Dock='Top'; $label.Height=26
        $input=if ($definition[0] -eq 'Model') { [Windows.Forms.ComboBox]::new() } else { [Windows.Forms.TextBox]::new() }
        $input.Dock='Bottom'; $input.Text=$definition[3]; $input.Name='AI'+$definition[0]
        if ($definition[0] -eq 'Key') {
            $input.UseSystemPasswordChar=$true
            $input.Dock='None'; $input.Location=[Drawing.Point]::new(0,36); $input.Width=114
            $keyLabel=$label
            $keyLock=[Windows.Forms.Button]::new(); $keyLock.Name='AILockKey'
            $keyLock.Text='锁定'; $keyLock.Location=[Drawing.Point]::new(118,34); $keyLock.Size=[Drawing.Size]::new(72,34)
            $group.Controls.Add($keyLock)
        }
        $group.Controls.Add($input); $group.Controls.Add($label); $config.Controls.Add($group); $fields[$definition[0]]=$input
    }
    $inputs=[Windows.Forms.TableLayoutPanel]::new(); $inputs.Dock='Fill'; $inputs.ColumnCount=3; $inputs.RowCount=1
    [void]$inputs.RowStyles.Add([Windows.Forms.RowStyle]::new('Percent',100))
    foreach ($width in @(25,35,40)) {[void]$inputs.ColumnStyles.Add([Windows.Forms.ColumnStyle]::new('Percent',$width))}
    foreach ($definition in @(@('Numbers','实验号（空白为全部）',''),@('Goal','分析目标','降低最终Sa，提出下一轮验证实验'),@('Constraints','工艺约束',''))) {
        $panel=[Windows.Forms.Panel]::new(); $panel.Dock='Fill'
        $label=[Windows.Forms.Label]::new(); $label.Text=$definition[1]; $label.Dock='Top'; $label.Height=26
        $box=[Windows.Forms.TextBox]::new(); $box.Multiline=$true; $box.Dock='Fill'; $box.Height=43; $box.Text=$definition[2]; $box.Name='AI'+$definition[0]
        $panel.Controls.Add($box); $panel.Controls.Add($label); $inputs.Controls.Add($panel); $fields[$definition[0]]=$box
    }
    $actions=[Windows.Forms.FlowLayoutPanel]::new(); $actions.Dock='Fill'; $actions.AutoScroll=$true
    $buttons=@{}
    foreach ($name in @('保存接口','测试连接','生成预览','确认发送','取消','分析历史','清除Key')) {
        $button=[Windows.Forms.Button]::new(); $button.Text=$name; $button.Width=104; $button.Height=34
        $actions.Controls.Add($button); $buttons[$name]=$button
    }
    $buttons['确认发送'].Enabled=$false; $buttons['取消'].Enabled=$false
    $status=[Windows.Forms.Label]::new(); $status.Dock='Fill'; $status.Text='仅手动发送；实验数据将离开本机，调用可能计费。Key仅限当前Windows用户解密。'
    $content=[Windows.Forms.RichTextBox]::new(); $content.ReadOnly=$true; $content.DetectUrls=$false; $content.WordWrap=$true; $content.ScrollBars='Vertical'; $content.Dock='Fill'; $content.Font=[Drawing.Font]::new('Microsoft YaHei UI',13)
    $outputTabs=[Windows.Forms.TabControl]::new(); $outputTabs.Dock='Fill'
    $summaryPage=[Windows.Forms.TabPage]::new('发送摘要')
    $answerPage=[Windows.Forms.TabPage]::new('AI分析结果')
    $rawPage=[Windows.Forms.TabPage]::new('完整请求')
    $summary=[Windows.Forms.TextBox]::new(); $summary.Multiline=$true; $summary.ReadOnly=$true; $summary.Dock='Fill'; $summary.ScrollBars='Vertical'
    $raw=[Windows.Forms.TextBox]::new(); $raw.Multiline=$true; $raw.ReadOnly=$true; $raw.Dock='Fill'; $raw.ScrollBars='Both'
    $summaryPage.Controls.Add($summary); $answerPage.Controls.Add($content); $rawPage.Controls.Add($raw)
    $outputTabs.TabPages.AddRange([Windows.Forms.TabPage[]]@($summaryPage,$answerPage,$rawPage))
    $readingBar=[Windows.Forms.FlowLayoutPanel]::new(); $readingBar.Dock='Top'; $readingBar.Height=40
    $focusButton=[Windows.Forms.Button]::new(); $focusButton.Text='专注阅读'; $focusButton.Size=[Drawing.Size]::new(130,34)
    $windowButton=[Windows.Forms.Button]::new(); $windowButton.Text='独立窗口'; $windowButton.Size=[Drawing.Size]::new(130,34)
    $readingBar.Controls.AddRange([Windows.Forms.Control[]]@($focusButton,$windowButton))
    $outputHost=[Windows.Forms.Panel]::new(); $outputHost.Dock='Fill'; $outputHost.Controls.Add($outputTabs); $outputHost.Controls.Add($readingBar)
    $reader=[pscustomobject]@{Focused=$false;Window=$null;Text=$null}
    $setReading={param([bool]$enabled)
        $reader.Focused=$enabled; $layout.SuspendLayout()
        $heights=@(90,84,44)
        for ($i=0;$i -lt 3;$i++) {$layout.RowStyles[$i].Height=if ($enabled) {0} else {$heights[$i]}}
        foreach ($control in @($config,$inputs,$actions)) {$control.Visible=-not $enabled}
        $focusButton.Text=if ($enabled) {'返回设置'} else {'专注阅读'}
        $layout.ResumeLayout()
    }.GetNewClosure()
    $focusButton.Add_Click({& $setReading (-not $reader.Focused)}.GetNewClosure())
    $windowButton.Add_Click({
        if ($reader.Window -and -not $reader.Window.IsDisposed) {$reader.Window.Activate();return}
        $reader.Window=[Windows.Forms.Form]::new(); $reader.Window.Text='AI分析结果 · 抛光哨站'
        $reader.Window.Size=[Drawing.Size]::new(1000,760); $reader.Window.MinimumSize=[Drawing.Size]::new(640,480); $reader.Window.StartPosition='CenterParent'
        $reader.Window.Padding=[Windows.Forms.Padding]::new(20)
        $reader.Text=[Windows.Forms.RichTextBox]::new(); $reader.Text.Dock='Fill'; $reader.Text.ReadOnly=$true; $reader.Text.DetectUrls=$false
        $reader.Text.Font=[Drawing.Font]::new('Microsoft YaHei UI',13); $reader.Text.Text=$content.Text
        $reader.Window.Controls.Add($reader.Text); $reader.Window.Show($Form)
    }.GetNewClosure())
    $content.Add_TextChanged({if ($reader.Window -and -not $reader.Window.IsDisposed) {$reader.Text.Text=$content.Text}}.GetNewClosure())
    $Form.Add_Disposed({if ($reader.Window) {$reader.Window.Dispose()}}.GetNewClosure())
    $layout.Controls.Add($config,0,0); $layout.Controls.Add($inputs,0,1); $layout.Controls.Add($actions,0,2); $layout.Controls.Add($status,0,3); $layout.Controls.Add($outputHost,0,4); $page.Controls.Add($layout)
    $state=[pscustomobject]@{Packet=$null;Worker=$null;Pending=$null;Cancel=$null;Timer=[Windows.Forms.Timer]::new();Closed=$false;Start=$null}
    $invalidate={ $state.Packet=$null; $buttons['确认发送'].Enabled=$false; $summary.Text='尚无有效发送预览，请生成预览。'; $raw.Clear() }.GetNewClosure()
    $syncKey={
        $saved=Read-SentinelAISettings $ScriptRoot
        $locked=[bool]$saved.KeyCipher -and $saved.Endpoint -eq $fields.Endpoint.Text.TrimEnd('/')
        $fields.Key.Clear(); $fields.Key.ReadOnly=$locked
        $keyLock.Text=if ($locked) {'更换'} else {'锁定'}
        $keyLabel.Text=if ($locked) {'API Key · 已保存'} else {'API Key · 未锁定'}
    }.GetNewClosure()
    & $syncKey
    foreach ($field in $fields.Values) {$field.Add_TextChanged($invalidate)}
    $buttons['保存接口'].Add_Click({
        try { Save-SentinelAISettings $ScriptRoot $fields.Endpoint.Text $fields.Model.Text $fields.Key.Text ([int]$fields.Timeout.Text); $saved=Read-SentinelAISettings $ScriptRoot; $fields.Endpoint.Text=$saved.Endpoint; $fields.Model.Text=$saved.Model; & $syncKey; & $invalidate; $status.Text='接口已保存，Key已加密锁定，重启后可直接使用。' }
        catch {$status.Text=$_.Exception.Message}
    }.GetNewClosure())
    $keyLock.Add_Click({
        if ($fields.Key.ReadOnly) {
            $fields.Key.ReadOnly=$false; $fields.Key.Clear(); $fields.Key.Focus()
            $keyLock.Text='锁定'; $keyLabel.Text='API Key · 更换中'; & $invalidate
            $status.Text='填写新 Key 后锁定；保存成功前原 Key 仍保留。'
        } elseif ([string]::IsNullOrWhiteSpace($fields.Key.Text)) { $status.Text='请先输入 Key。' }
        else { $buttons['保存接口'].PerformClick() }
    }.GetNewClosure())
    $fields.Endpoint.Add_TextChanged({
        $fields.Key.ReadOnly=$false; $fields.Key.Clear(); $keyLock.Text='锁定'; $keyLabel.Text='API Key · 地址已改'
    }.GetNewClosure())
    $buttons['清除Key'].Add_Click({
        if ([Windows.Forms.MessageBox]::Show('清除本机保存的接口与Key？分析历史不会删除。','清除凭据','YesNo','Question') -eq 'Yes') {
            $path=Join-Path $ScriptRoot 'config/ai-settings.json'; if ([IO.File]::Exists($path)) {[IO.File]::Delete($path)}
            & $syncKey; & $invalidate; $status.Text='本机凭据已清除。'
        }
    }.GetNewClosure())
    $buttons['生成预览'].Add_Click({
        try {
            $saved=Read-SentinelAISettings $ScriptRoot
            if ($saved.Endpoint -ne $fields.Endpoint.Text.TrimEnd('/') -or $saved.Model -ne $fields.Model.Text.Trim() -or $fields.Key.Text) {throw '请先保存接口设置。'}
            $state.Packet=New-SentinelAIPreview @($AnalysisState.Records) $fields.Numbers.Text $fields.Goal.Text $fields.Constraints.Text $saved
            $raw.Text=$state.Packet | ConvertTo-Json -Depth 18
            $data=$state.Packet.Payload.messages[1].content | ConvertFrom-Json
            $metrics=$state.Packet.Metrics
            $summary.Text="实验数量：$($state.Packet.ExperimentCount) 次（水基） · 缺少 Sa：$($metrics.Missing) · 重号：$($metrics.Duplicates)`r`n输入字符：$($metrics.BeforeCharacters) → $($metrics.AfterCharacters)，减少 $($metrics.ReductionPercent)%`r`n输入 token 粗估：$($metrics.EstimatedTokensLow)–$($metrics.EstimatedTokensHigh)（非实际账单，不含输出）`r`n模型：$($state.Packet.Model) · 服务：$($state.Packet.Endpoint)`r`n分析目标：$($data.Goal)`r`n工艺约束：$(if ($data.Constraints) {$data.Constraints} else {'未设置'})`r`n实验编号：$($metrics.Numbers -join '、')`r`n全部所选实验保留，公共参数提取一次；未知、异常值和备注未删除。数据将离开本机，可能计费。"
            $outputTabs.SelectedTab=$summaryPage
            $status.Text='预览已生成，请核对发送摘要，确认后才发送。'
            $buttons['确认发送'].Enabled=$true
        } catch {& $invalidate; $status.Text=$_.Exception.Message}
    }.GetNewClosure())
    $module=Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.AI.ps1'
    $start={ param([bool]$test)
        if ($state.Pending) {return}
        $saved=Read-SentinelAISettings $ScriptRoot
        if ($saved.Endpoint -ne $fields.Endpoint.Text.TrimEnd('/') -or $saved.Model -ne $fields.Model.Text.Trim() -or $fields.Key.Text) { $status.Text='请先保存接口设置。'; return }
        if (-not $test) {
            if (-not $state.Packet) {return}
            $confirm='将向 '+$state.Packet.Endpoint+' 的 '+$state.Packet.Model+' 发送预览中的 '+$state.Packet.ExperimentCount+' 次实验数据，可能产生费用。继续？'
            if ([Windows.Forms.MessageBox]::Show($confirm,'确认数据发送','YesNo','Warning') -ne 'Yes') {return}
        }
        $state.Cancel=[Threading.CancellationTokenSource]::new()
        $state.Worker=[PowerShell]::Create()
        [void]$state.Worker.AddScript({param($module,$root,$packet,$test,$token) $ErrorActionPreference='Stop'; . $module; Invoke-SentinelAI -Root $root -Packet $packet -TestConnection:$test -Token $token}.ToString()).AddArgument($module).AddArgument($ScriptRoot).AddArgument($state.Packet).AddArgument($test).AddArgument($state.Cancel.Token)
        foreach ($button in $buttons.Values) {$button.Enabled=$false}; $buttons['取消'].Enabled=$true
        foreach ($field in $fields.Values) {$field.Enabled=$false}
        $keyLock.Enabled=$false
        $status.Text=if ($test) {'正在读取可用模型，不发送实验数据…'} else {'正在等待AI回复；建议仅供验证，不自动执行。'}
        $state.Pending=$state.Worker.BeginInvoke(); $state.Timer.Start()
    }.GetNewClosure()
    $state.Start=$start
    $buttons['测试连接'].Add_Click({ & $state.Start $true }.GetNewClosure())
    $buttons['确认发送'].Add_Click({ & $state.Start $false }.GetNewClosure())
    $buttons['取消'].Add_Click({if ($state.Cancel) {$state.Cancel.Cancel();$status.Text='正在取消；已发送的请求可能仍计费。'}}.GetNewClosure())
    $buttons['分析历史'].Add_Click({
        $dialog=[Windows.Forms.OpenFileDialog]::new(); $dialog.InitialDirectory=Join-Path $ScriptRoot 'state/ai-history'; $dialog.Filter='AI分析历史 (*.json)|*.json'
        try { if ($dialog.ShowDialog() -eq 'OK') { $record=Get-Content -LiteralPath $dialog.FileName -Raw -Encoding UTF8 | ConvertFrom-Json; $content.Text="模型：$($record.Model)`r`n时间：$($record.Time)`r`n`r`n$($record.Answer)"; $outputTabs.SelectedTab=$answerPage; $status.Text='历史结果，仅供参考：'+$record.Time; & $invalidate } }
        catch {$status.Text='历史文件无法读取。'} finally {$dialog.Dispose()}
    }.GetNewClosure())
    $state.Timer.Interval=250
    $state.Timer.Add_Tick({
        if (-not $state.Pending -or -not $state.Pending.IsCompleted) {return}
        try {
            $result=@($state.Worker.EndInvoke($state.Pending))[-1]
            if ($state.Worker.HadErrors) {
                $status.Text=Get-SentinelAISafeError $state.Worker.Streams.Error[0].Exception
                return
            }
            if ($null -ne $result.PSObject.Properties['Models']) {
                $current=$fields.Model.Text
                $fields.Model.Items.Clear()
                $fields.Model.Items.AddRange([object[]]$result.Models)
                $fields.Model.Text=$current
                $status.Text=$result.Text
                return
            }
            $content.Text=$result.Text; $outputTabs.SelectedTab=$answerPage; $status.Text=if ($result.HistoryPath) {'分析完成，历史已保存在本机。建议未经实验验证。'} else {'请求已结束。'}
            if ($null -eq $result.PSObject.Properties['Models']) {& $setReading $true}
        } catch {$status.Text=if ($state.Cancel.IsCancellationRequested) {'已取消；未自动重试。'} else {Get-SentinelAISafeError $_.Exception}}
        finally {
            $state.Worker.Dispose(); $state.Worker=$null; $state.Pending=$null; $state.Cancel.Dispose(); $state.Cancel=$null; $state.Timer.Stop()
            foreach ($field in $fields.Values) {$field.Enabled=$true}; foreach ($button in $buttons.Values) {$button.Enabled=$true}
            $keyLock.Enabled=$true
            $buttons['取消'].Enabled=$false; $state.Packet=$null; $buttons['确认发送'].Enabled=$false
        }
    }.GetNewClosure())
    $Form.Add_Disposed({$state.Closed=$true;$state.Timer.Stop();$state.Timer.Dispose();if ($state.Cancel) {$state.Cancel.Cancel()};if ($state.Worker) {$state.Worker.Stop();$state.Worker.Dispose()};if ($state.Cancel) {$state.Cancel.Dispose()}}.GetNewClosure())
    $page.Tag=[pscustomobject]@{Fields=$fields;Buttons=$buttons;State=$state;Status=$status;Content=$content;KeyLock=$keyLock;Summary=$summary;Raw=$raw;OutputTabs=$outputTabs;FocusButton=$focusButton;WindowButton=$windowButton;Reader=$reader}
    $latest=Read-SentinelAILatestResult $ScriptRoot
    if ($latest) {
        $content.Text="模型：$($latest.Model)`r`n时间：$($latest.Time)`r`n`r`n$($latest.Answer)"
        $outputTabs.SelectedTab=$answerPage
        $status.Text='已恢复上次分析结果（历史数据，未重新分析）。'
        & $setReading $true
    }
    return $page
}
