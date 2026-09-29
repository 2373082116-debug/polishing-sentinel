Set-StrictMode -Version 2.0

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
. (Join-Path $PSScriptRoot 'PolishingSentinel.Progress.ps1')

$script:PolishingSentinelUiRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.Config.ps1')
. (Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.Status.ps1')
. (Join-Path $script:PolishingSentinelUiRoot 'ExperimentDataAutomation.Core.ps1')
. (Join-Path $script:PolishingSentinelUiRoot 'ExperimentDataAutomation.Backfill.ps1')

function Get-SentinelUiTheme {
    [pscustomobject]@{
        WindowBack      = [System.Drawing.Color]::FromArgb(244, 246, 249)
        Surface         = [System.Drawing.Color]::FromArgb(246, 251, 255)
        SurfaceAlt      = [System.Drawing.Color]::FromArgb(234, 246, 253)
        InputBack       = [System.Drawing.Color]::FromArgb(238, 248, 254)
        Primary         = [System.Drawing.Color]::FromArgb(37, 99, 235)
        PrimaryDark     = [System.Drawing.Color]::FromArgb(9, 43, 84)
        Border          = [System.Drawing.Color]::FromArgb(91, 141, 190)
        BorderSoft      = [System.Drawing.Color]::FromArgb(174, 209, 232)
        Ink             = [System.Drawing.Color]::FromArgb(15, 43, 74)
        MutedInk        = [System.Drawing.Color]::FromArgb(58, 91, 124)
        White           = [System.Drawing.Color]::FromArgb(255, 255, 255)
        Font            = New-Object System.Drawing.Font('Microsoft YaHei UI', 13)
        StrongFont      = New-Object System.Drawing.Font('Microsoft YaHei UI', 13, [System.Drawing.FontStyle]::Bold)
    }
}

function Enter-PolishingSentinelSingleInstance {
    param([string] $Name = 'Local\PolishingSentinel.ControlPanel')

    $createdNew = $false
    $mutex = New-Object System.Threading.Mutex($true, $Name, [ref] $createdNew)
    if (-not $createdNew) {
        $mutex.Dispose()
        return [pscustomobject]@{
            Acquired = $false
            Mutex    = $null
            Name     = $Name
        }
    }

    return [pscustomobject]@{
        Acquired = $true
        Mutex    = $mutex
        Name     = $Name
    }
}

function Exit-PolishingSentinelSingleInstance {
    param([AllowNull()] [pscustomobject] $Handle)

    if ($null -eq $Handle -or -not $Handle.Acquired -or $null -eq $Handle.Mutex) {
        return
    }

    try {
        $Handle.Mutex.ReleaseMutex()
    }
    catch {
    }
    finally {
        $Handle.Mutex.Dispose()
    }
}

function Set-SentinelButtonStyle {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.Button] $Button,
        [Parameter(Mandatory = $true)] [pscustomobject] $Theme
    )

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.BackColor = $Theme.SurfaceAlt
    $Button.ForeColor = $Theme.Primary
    $Button.Font = $Theme.StrongFont
    $Button.UseVisualStyleBackColor = $false
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = $Theme.Primary
    $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(218, 237, 249)
    $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(194, 222, 242)
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
}

function Set-SentinelTextBoxStyle {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.TextBox] $TextBox,
        [Parameter(Mandatory = $true)] [pscustomobject] $Theme
    )

    $TextBox.BackColor = $Theme.InputBack
    $TextBox.ForeColor = $Theme.Ink
    $TextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
}

function Set-SentinelListViewStyle {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.ListView] $ListView,
        [Parameter(Mandatory = $true)] [pscustomobject] $Theme
    )

    $ListView.BackColor = $Theme.InputBack
    $ListView.ForeColor = $Theme.Ink
    $ListView.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $ListView.GridLines = $true
}

function Set-SentinelGroupBoxStyle {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.GroupBox] $GroupBox,
        [Parameter(Mandatory = $true)] [pscustomobject] $Theme
    )

    $GroupBox.BackColor = $Theme.Surface
    $GroupBox.ForeColor = $Theme.Primary
    $GroupBox.Font = $Theme.StrongFont
    $GroupBox.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
}

function Set-SentinelTabStyle {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.TabControl] $TabControl,
        [Parameter(Mandatory = $true)] [pscustomobject] $Theme
    )

    $TabControl.Appearance = [System.Windows.Forms.TabAppearance]::FlatButtons
    $TabControl.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
    $TabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
    $TabControl.ItemSize = New-Object System.Drawing.Size(126, 34)
    $TabControl.BackColor = $Theme.WindowBack
    $TabControl.Font = $Theme.StrongFont
    $TabControl.Add_DrawItem({
        param($sender, $eventArgs)

        $page = $sender.TabPages[$eventArgs.Index]
        $selected = ($eventArgs.Index -eq $sender.SelectedIndex)
        $backColor = $Theme.SurfaceAlt
        $foreColor = $Theme.Primary
        if ($selected) {
            $backColor = $Theme.Primary
            $foreColor = $Theme.White
        }

        $brush = New-Object System.Drawing.SolidBrush($backColor)
        try {
            $eventArgs.Graphics.FillRectangle($brush, $eventArgs.Bounds)
        }
        finally {
            $brush.Dispose()
        }

        $pen = New-Object System.Drawing.Pen($Theme.Border)
        try {
            $eventArgs.Graphics.DrawRectangle($pen, $eventArgs.Bounds)
        }
        finally {
            $pen.Dispose()
        }

        $flags = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis
        [System.Windows.Forms.TextRenderer]::DrawText($eventArgs.Graphics, $page.Text, $sender.Font, $eventArgs.Bounds, $foreColor, $flags)
    }.GetNewClosure())
}

function Set-SentinelGridTypography {
    param([Windows.Forms.DataGridView]$Grid)
    $Grid.ColumnHeadersDefaultCellStyle.Font=$Grid.Font
    $Grid.ColumnHeadersDefaultCellStyle.WrapMode='True'
    $Grid.ColumnHeadersDefaultCellStyle.Padding=[Windows.Forms.Padding]::new(4,6,4,6)
    $Grid.ColumnHeadersDefaultCellStyle.Alignment='MiddleLeft'
    $Grid.ColumnHeadersHeightSizeMode='AutoSize'
    $Grid.DefaultCellStyle.Padding=[Windows.Forms.Padding]::new(3,4,3,4)
    $Grid.DefaultCellStyle.Alignment='MiddleLeft'
    $Grid.AutoSizeRowsMode='AllCells'
    $font=if ($Grid.DefaultCellStyle.Font) {$Grid.DefaultCellStyle.Font} else {$Grid.Font}
    $Grid.RowTemplate.MinimumHeight=[Math]::Max(32,[Windows.Forms.TextRenderer]::MeasureText('实验Ag',$font).Height+10)
}

function Set-SentinelControlTheme {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.Control] $Control,
        [Parameter(Mandatory = $true)] [pscustomobject] $Theme
    )

    $Control.Font = $Theme.Font

    if ($Control -is [System.Windows.Forms.Form]) {
        $Control.BackColor = $Theme.WindowBack
        $Control.ForeColor = $Theme.Ink
    }
    elseif ($Control -is [System.Windows.Forms.TabControl]) {
        Set-SentinelTabStyle -TabControl $Control -Theme $Theme
    }
    elseif ($Control -is [System.Windows.Forms.TabPage]) {
        $Control.BackColor = $Theme.WindowBack
        $Control.ForeColor = $Theme.Ink
    }
    elseif ($Control -is [System.Windows.Forms.GroupBox]) {
        Set-SentinelGroupBoxStyle -GroupBox $Control -Theme $Theme
    }
    elseif ($Control -is [System.Windows.Forms.Button]) {
        Set-SentinelButtonStyle -Button $Control -Theme $Theme
    }
    elseif ($Control -is [System.Windows.Forms.TextBox]) {
        Set-SentinelTextBoxStyle -TextBox $Control -Theme $Theme
    }
    elseif ($Control -is [System.Windows.Forms.ListView]) {
        Set-SentinelListViewStyle -ListView $Control -Theme $Theme
    }
    elseif ($Control -is [System.Windows.Forms.DataGridView]) {
        Set-SentinelGridTypography $Control
    }
    elseif ($Control -is [System.Windows.Forms.ComboBox]) {
        $Control.BackColor = $Theme.InputBack
        $Control.ForeColor = $Theme.Ink
        $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    }
    elseif ($Control -is [System.Windows.Forms.CheckBox]) {
        $Control.BackColor = $Theme.WindowBack
        $Control.ForeColor = $Theme.Ink
    }
    elseif ($Control -is [System.Windows.Forms.Label]) {
        $Control.BackColor = [System.Drawing.Color]::Transparent
        $Control.ForeColor = $Theme.Ink
    }

    foreach ($child in $Control.Controls) {
        Set-SentinelControlTheme -Control $child -Theme $Theme
    }
}

function Get-SentinelObjectProperty {
    param(
        [AllowNull()] [object] $Object,
        [Parameter(Mandatory = $true)] [string] $Name,
        [string] $Default = ''
    )

    if ($null -eq $Object) {
        return $Default
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $Default
    }

    return [string] $property.Value
}

function New-SentinelLabel {
    param(
        [string] $Text,
        [int] $X,
        [int] $Y,
        [int] $Width = 140,
        [int] $Height = 24,
        [string] $Name = ''
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, $Height)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $label.AutoEllipsis = $true
    if (-not [string]::IsNullOrWhiteSpace($Name)) {
        $label.Name = $Name
    }
    return $label
}

function New-SentinelTextBox {
    param(
        [string] $Text,
        [int] $X,
        [int] $Y,
        [int] $Width = 620,
        [string] $Name = ''
    )

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $Text
    $textBox.Location = New-Object System.Drawing.Point($X, $Y)
    $textBox.Size = New-Object System.Drawing.Size($Width, 24)
    if (-not [string]::IsNullOrWhiteSpace($Name)) {
        $textBox.Name = $Name
    }
    return $textBox
}

function New-SentinelButton {
    param(
        [string] $Text,
        [int] $X,
        [int] $Y,
        [int] $Width = 110,
        [string] $Name = ''
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, 30)
    if (-not [string]::IsNullOrWhiteSpace($Name)) {
        $button.Name = $Name
    }
    return $button
}

function Add-SentinelSettingRow {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.Control] $Parent,
        [Parameter(Mandatory = $true)] [hashtable] $Controls,
        [Parameter(Mandatory = $true)] [string] $Key,
        [Parameter(Mandatory = $true)] [string] $Label,
        [AllowNull()] [object] $Value,
        [Parameter(Mandatory = $true)] [int] $Y
    )

    $Parent.Controls.Add((New-SentinelLabel -Text $Label -X 20 -Y $Y -Width 150))
    $box = New-SentinelTextBox -Text ([string] $Value) -X 180 -Y $Y -Width 720 -Name ('PathSetting_{0}' -f $Key)
    $Parent.Controls.Add($box)
    $Controls[$Key] = $box
}

function Set-SentinelStatusLabels {
    param(
        [Parameter(Mandatory = $true)] [hashtable] $Labels,
        [Parameter(Mandatory = $true)] [pscustomobject] $Status
    )

    $Labels.State.Text = Get-SentinelObjectProperty -Object $Status -Name 'State'
    $Labels.Message.Text = Get-SentinelObjectProperty -Object $Status -Name 'Message'
    $Labels.UpdatedAt.Text = Get-SentinelObjectProperty -Object $Status -Name 'UpdatedAt'
    if ($Labels.ContainsKey('ProgressBar')) {
        Set-SentinelProcessingProgress -Status $Status -Bar $Labels.ProgressBar -Label $Labels.ProgressText
    }
    if ($Labels.ContainsKey('Health')) {
        $healthText = Get-SentinelObjectProperty -Object $Status -Name 'HealthText'
        $elapsed = Get-SentinelObjectProperty -Object $Status -Name 'StageElapsedSeconds'
        $terminal = $Status.State -in @('Completed','PartialSuccess','Error','DryRun','TimedOut','NeedsConfirmation','RetryPending','Exited')
        $Labels.Health.Text = if ($healthText -and $terminal) { "$healthText；本轮已结束" } elseif ($healthText) { "$healthText；阶段耗时 ${elapsed} 秒" } else { '尚无可核验的后台健康信息' }
    }
}

function Update-SentinelPathCheckList {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.ListView] $ListView,
        [Parameter(Mandatory = $true)] [pscustomobject] $Settings
    )

    $ListView.Items.Clear()
    foreach ($check in @(Test-PolishingSentinelPathSet -Settings $Settings)) {
        $existsText = '缺失'
        if ($check.Exists) {
            $existsText = '存在'
        }

        $item = New-Object System.Windows.Forms.ListViewItem([string] $check.Label)
        [void] $item.SubItems.Add($existsText)
        [void] $item.SubItems.Add([string] $check.Path)
        [void] $ListView.Items.Add($item)
    }
}

function New-SentinelPlaceholderPage {
    param([Parameter(Mandatory = $true)] [string] $Text)

    $page = New-Object System.Windows.Forms.TabPage
    $page.Text = $Text
    $page.AutoScroll = $true
    return $page
}

function Get-SentinelPowerShellExe {
    $powershellExe = Join-Path $PSHOME 'powershell.exe'
    if (Test-Path -LiteralPath $powershellExe) {
        return $powershellExe
    }

    return 'powershell.exe'
}

function New-SentinelPowerShellFileArguments {
    param(
        [Parameter(Mandatory = $true)] [string] $ScriptPath,
        [string[]] $ExtraArguments = @()
    )

    $quotedScriptPath = '"{0}"' -f $ScriptPath
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $quotedScriptPath)
    if ($ExtraArguments.Count -gt 0) {
        $arguments += $ExtraArguments
    }

    return ($arguments -join ' ')
}

function Get-SentinelProcessDefaultsText {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Settings)

    $water = $Settings.ProcessDefaults.Water
    $oil = $Settings.ProcessDefaults.Oil
    $fieldDefinitions = @(Get-ExperimentProcessFieldDefinitions)
    $groupDefinitions = @(
        [pscustomobject]@{ Title = '探索实验'; Group = $water },
        [pscustomobject]@{ Title = '油基抛光'; Group = $oil }
    )

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($groupDefinition in $groupDefinitions) {
        if ($lines.Count -gt 0) {
            $lines.Add('')
        }
        $lines.Add($groupDefinition.Title)
        foreach ($field in $fieldDefinitions) {
            $value = Get-SentinelObjectProperty -Object $groupDefinition.Group -Name $field.Key -Default ''
            $lines.Add(('{0}：{1}' -f $field.Label, $value))
        }
    }

    return ($lines -join [Environment]::NewLine)
}

function ConvertFrom-SentinelProcessDefaultsText {
    param([Parameter(Mandatory = $true)] [string] $Text)

    $result = [pscustomobject]@{
        Water = [pscustomobject]@{}
        Oil   = [pscustomobject]@{}
    }
    $fieldDefinitions = @(Get-ExperimentProcessFieldDefinitions)
    foreach ($groupName in @('Water', 'Oil')) {
        $group = [pscustomobject]@{}
        foreach ($field in $fieldDefinitions) {
            $group | Add-Member -NotePropertyName $field.Key -NotePropertyValue '' -Force
        }
        $result | Add-Member -NotePropertyName $groupName -NotePropertyValue $group -Force
    }
    $fieldByLabel = @{}
    foreach ($field in $fieldDefinitions) {
        $fieldByLabel[$field.Label] = $field
    }
    $current = $null

    foreach ($rawLine in ($Text -split '\r?\n')) {
        $line = $rawLine.Trim()
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        if ($line -eq '探索实验') {
            $current = $result.Water
            continue
        }
        if ($line -eq '油基抛光') {
            $current = $result.Oil
            continue
        }
        if ($null -eq $current) {
            continue
        }

        $separatorIndex = $line.IndexOf('：')
        if ($separatorIndex -lt 0) {
            $separatorIndex = $line.IndexOf(':')
        }
        if ($separatorIndex -lt 0) {
            continue
        }

        $key = $line.Substring(0, $separatorIndex).Trim()
        $value = $line.Substring($separatorIndex + 1).Trim()
        if ($fieldByLabel.ContainsKey($key)) {
            $field = $fieldByLabel[$key]
            $current | Add-Member -NotePropertyName $field.Key -NotePropertyValue $value -Force
        }
    }

    foreach ($groupName in @('Water', 'Oil')) {
        $group = $result.$groupName
        foreach ($field in $fieldDefinitions) {
            if ($field.Required -and [string]::IsNullOrWhiteSpace([string] $group.$($field.Key))) {
                throw "固定工艺参数缺少 $groupName.$($field.Key)"
            }
        }
    }

    return $result
}

function ConvertFrom-SentinelSlurryMappingsText {
    param([Parameter(Mandatory = $true)] [string] $Text)

    $mappings = [ordered]@{}
    foreach ($rawLine in ($Text -split '\r?\n')) {
        $line = $rawLine.Trim()
        if ([string]::IsNullOrWhiteSpace($line) -or $line -notmatch '->') {
            continue
        }

        $parts = @($line -split '->', 2)
        if ($parts.Count -ne 2) {
            continue
        }

        $slurry = $parts[0].Trim()
        $sheetName = $parts[1].Trim()
        if ([string]::IsNullOrWhiteSpace($slurry) -or [string]::IsNullOrWhiteSpace($sheetName)) {
            continue
        }

        $mappings[$slurry] = $sheetName
    }

    if ($mappings.Count -eq 0) {
        throw '匹配规则中没有有效的“抛光液 -> 工作表”映射'
    }

    $result = [pscustomobject]@{}
    foreach ($key in $mappings.Keys) {
        $result | Add-Member -NotePropertyName $key -NotePropertyValue $mappings[$key] -Force
    }

    return $result
}

function Get-SentinelMatchingRulesText {
    param([Parameter(Mandatory = $true)] [pscustomobject] $Settings)

    $mappingLines = New-Object System.Collections.Generic.List[string]
    $mappings = $Settings.SlurryMappings
    if ($mappings -is [System.Collections.IDictionary]) {
        foreach ($key in $mappings.Keys) {
            $mappingLines.Add(('{0} -> {1}' -f $key, $mappings[$key]))
        }
    }
    else {
        foreach ($property in $mappings.PSObject.Properties) {
            $mappingLines.Add(('{0} -> {1}' -f $property.Name, $property.Value))
        }
    }

    if ($mappingLines.Count -eq 0) {
        $mappingLines.Add('金刚石 -> 油基抛光')
        $mappingLines.Add('硅溶胶 -> 探索实验')
    }

    $newBlockMode = Get-SentinelObjectProperty -Object $Settings -Name 'NewBlockMode' -Default 'Preview'

    return (@(
        '映射说明（可修改“抛光液 -> 工作表名称”行）'
    ) + @($mappingLines) + @(
        '',
        '阶段二新建块模式',
        ('NewBlockMode = {0}（Preview=只写待确认预览；Auto=高置信自动建块；Disabled=关闭阶段二）' -f $newBlockMode),
        '',
        '冲突策略',
        ('冲突、新块、不唯一：进入 {0}' -f $Settings.PendingSheetName)
    ) -join [Environment]::NewLine)
}

function Start-SentinelStatusAutoRefresh {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.Form] $Form,
        [Parameter(Mandatory = $true)] [hashtable] $Labels,
        [Parameter(Mandatory = $true)] [System.Windows.Forms.Label] $StartupLabel,
        [Parameter(Mandatory = $true)] [string] $ScriptRoot,
        [Parameter(Mandatory = $true)] [string] $TaskName,
        [scriptblock] $StartupQuery = {
            param($ScriptRoot, $TaskName, $StatusModule)
            $ErrorActionPreference = 'Stop'
            . $StatusModule
            $task = Get-PolishingSentinelTaskStatus -TaskName $TaskName
            $processes = @(Get-PolishingSentinelMonitorProcesses -ScriptRoot $ScriptRoot)
            $taskText = if ($task.Exists) { $task.State } else { '未安装' }
            '计划任务：{0}；监控进程：{1} 个' -f $taskText, $processes.Count
        }
    )

    $controller = [pscustomobject]@{
        Timer = New-Object System.Windows.Forms.Timer
        Worker = $null
        Pending = $null
        Disposed = $false
        Refresh = $null
    }
    $statusModule = Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.Status.ps1'
    $readStatus = (Get-Command Read-PolishingSentinelStatus).ScriptBlock
    $setLabels = (Get-Command Set-SentinelStatusLabels).ScriptBlock
    $refresh = {
        if ($controller.Disposed) { return }
        try {
            $latest = & $readStatus -ScriptRoot $ScriptRoot
            if ($null -eq $latest) { throw '运行状态文件为空' }
            & $setLabels -Labels $Labels -Status $latest
        }
        catch {
            $Labels.Message.Text = '运行状态读取失败，将自动重试：{0}' -f $_.Exception.Message
            if ($Labels.ContainsKey('ProgressBar')) {
                Set-SentinelProcessingProgress -Status ([pscustomobject]@{ State='Unknown'; Data=@{} }) -Bar $Labels.ProgressBar -Label $Labels.ProgressText
                $Labels.ProgressText.Text='进度读取失败，等待状态恢复'
            }
        }

        # Only one background query may be in flight; slow CIM calls never block the UI timer.
        if ($null -ne $controller.Pending) {
            if (-not $controller.Pending.IsCompleted) { return }
            try {
                $result = $controller.Worker.EndInvoke($controller.Pending)
                if ($controller.Worker.HadErrors) { throw ($controller.Worker.Streams.Error | Out-String) }
                $StartupLabel.Text = [string] ($result | Select-Object -Last 1)
            }
            catch {
                $StartupLabel.Text = '计划任务/监控进程状态读取失败，将自动重试：{0}' -f $_.Exception.Message
            }
            finally {
                $controller.Worker.Dispose()
                $controller.Worker = $null
                $controller.Pending = $null
            }
        }
        try {
            $controller.Worker = [PowerShell]::Create()
            [void] $controller.Worker.AddScript($StartupQuery.ToString()).AddArgument($ScriptRoot).AddArgument($TaskName).AddArgument($statusModule)
            $controller.Pending = $controller.Worker.BeginInvoke()
        }
        catch {
            $StartupLabel.Text = '状态查询启动失败，将自动重试：{0}' -f $_.Exception.Message
            if ($null -ne $controller.Worker) { $controller.Worker.Dispose() }
            $controller.Worker = $null
            $controller.Pending = $null
        }
    }.GetNewClosure()
    $controller.Refresh = $refresh
    $controller.Timer.Interval = 1000
    $controller.Timer.Add_Tick($refresh)
    $Form.Add_Shown($refresh)
    $Form.Add_Disposed({
        $controller.Disposed = $true
        $controller.Timer.Stop()
        $controller.Timer.Dispose()
        if ($null -ne $controller.Worker) {
            $controller.Worker.Stop()
            $controller.Worker.Dispose()
            $controller.Worker = $null
            $controller.Pending = $null
        }
        $controller.Refresh = $null
    }.GetNewClosure())
    $StartupLabel.Tag = $controller
    $controller.Timer.Start()
    return $controller
}

function Set-SentinelStartupStatusLabel {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.Label] $Label,
        [Parameter(Mandatory = $true)] [pscustomobject] $Settings,
        [Parameter(Mandatory = $true)] [string] $ScriptRoot
    )

    try {
        $task = Get-PolishingSentinelTaskStatus -TaskName $Settings.TaskName
        $processes = @(Get-PolishingSentinelMonitorProcesses -ScriptRoot $ScriptRoot)
        $taskText = '未安装'
        if ($task.Exists) {
            $taskText = $task.State
        }

        $Label.Text = '计划任务：{0}；监控进程：{1} 个' -f $taskText, $processes.Count
    }
    catch {
        $Label.Text = '计划任务/监控进程状态读取失败：{0}' -f $_.Exception.Message
    }
}

function Add-PolishingSentinelNamingTab {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.TabPage] $Page,
        [Parameter(Mandatory = $true)] [pscustomobject] $Settings,
        [Parameter(Mandatory = $true)] [string] $ScriptRoot
    )

    $Page.Controls.Add((New-SentinelLabel -Text '命名约定' -X 20 -Y 24 -Width 120))
    $conventionBox = New-SentinelTextBox -Text $Settings.FolderNamingConvention -X 150 -Y 24 -Width 780 -Name 'NamingConventionBox'
    $conventionBox.ReadOnly = $true
    $Page.Controls.Add($conventionBox)

    $Page.Controls.Add((New-SentinelLabel -Text '测试文件夹名' -X 20 -Y 70 -Width 120))
    $folderTestBox = New-SentinelTextBox -Text '72-2026.7.8-10%-50nm-金刚石-盛金灵原球20颗-2h-植绒-备注' -X 150 -Y 70 -Width 780 -Name 'FolderNameTestBox'
    $Page.Controls.Add($folderTestBox)

    $parseResult = New-Object System.Windows.Forms.TextBox
    $parseResult.Name = 'NamingParseResult'
    $parseResult.Multiline = $true
    $parseResult.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $parseResult.Location = New-Object System.Drawing.Point(150, 110)
    $parseResult.Size = New-Object System.Drawing.Size(780, 360)
    $parseResult.ReadOnly = $true
    $Page.Controls.Add($parseResult)

    $parseButton = New-SentinelButton -Text '解析' -X 940 -Y 68 -Width 80 -Name 'NamingParseButton'
    $parseButton.Add_Click({
        try {
            $folderName = $folderTestBox.Text.Trim()
            $folderPath = [System.IO.Path]::Combine('F:\实验数据\抛光实验\测试', $folderName)
            $meta = Parse-StructuredExperimentFolderName -FolderPath $folderPath -SlurryMappings $Settings.SlurryMappings -ProcessDefaults $Settings.ProcessDefaults
            $issues = '无'
            if ($meta.Issues.Count -gt 0) {
                $issues = $meta.Issues -join '；'
            }

            $parseResult.Text = @(
                ('实验号：{0}' -f $meta.ExperimentNo),
                ('日期：{0}' -f $meta.ExperimentDate),
                ('浓度：{0}' -f $meta.Concentration),
                ('粒径：{0}' -f $meta.ParticleSize),
                ('抛光液类型：{0}' -f $meta.SlurryType),
                ('目标工作表：{0}' -f $meta.TargetSheetName),
                ('球胚：{0}' -f $meta.BallSource),
                ('抛光时间：{0}' -f $meta.Duration),
                ('抛光垫：{0}' -f $meta.PadType),
                ('备注：{0}' -f $meta.Note),
                ('问题：{0}' -f $issues)
            ) -join [Environment]::NewLine
        }
        catch {
            $parseResult.Text = '解析失败：{0}' -f $_.Exception.Message
        }
    }.GetNewClosure())
    $Page.Controls.Add($parseButton)
    $fitNaming={
        $width=[Math]::Max(320,$Page.ClientSize.Width-174)
        $conventionBox.Width=$width
        $folderTestBox.Width=[Math]::Max(220,$width-92)
        $parseButton.Left=150+$width-80
        $parseResult.Width=$width
        $parseResult.Height=[Math]::Max(180,$Page.ClientSize.Height-134)
    }.GetNewClosure()
    $Page.Add_Resize($fitNaming)
    & $fitNaming
}

. (Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.SettingsForms.ps1')
. (Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.Dashboard.Data.ps1')
. (Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.Dashboard.UI.ps1')
. (Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.Analysis.ps1')
. (Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.Analysis.UI.ps1')
. (Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.AI.ps1')
. (Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.AI.UI.ps1')
. (Join-Path $script:PolishingSentinelUiRoot 'PolishingSentinel.Feedback.ps1')

function New-SentinelSaRequestDialog {
    param([Parameter(Mandatory = $true)] [object[]] $Requests)

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = '处理编号冲突'
    $dialog.StartPosition = 'CenterParent'
    $dialog.ClientSize = New-Object System.Drawing.Size(800, 430)
    $dialog.MinimumSize = New-Object System.Drawing.Size(816, 469)
    $list = New-Object System.Windows.Forms.ListView
    $list.Name = 'SaRequestsList'
    $list.View = 'Details'
    $list.FullRowSelect = $true
    $list.MultiSelect = $false
    $list.HideSelection = $false
    $list.Location = New-Object System.Drawing.Point(20, 20)
    $list.Size = New-Object System.Drawing.Size(760, 180)
    $list.Anchor = 'Top, Left, Right'
    [void] $list.Columns.Add('实验目录', 440)
    [void] $list.Columns.Add('请求文件', 300)
    foreach ($request in $Requests) {
        $item = New-Object System.Windows.Forms.ListViewItem([string] $request.FolderPath)
        [void] $item.SubItems.Add([string] $request.RequestPath)
        $item.Tag = $request
        [void] $list.Items.Add($item)
    }
    $details = New-Object System.Windows.Forms.TextBox
    $details.Name = 'SaConflictDetails'
    $details.Multiline = $true
    $details.ReadOnly = $true
    $details.ScrollBars = 'Both'
    $details.Location = New-Object System.Drawing.Point(20, 215)
    $details.Size = New-Object System.Drawing.Size(760, 150)
    $details.Anchor = 'Top, Bottom, Left, Right'
    $select = New-SentinelButton -Text '选择此请求' -X 540 -Y 385 -Width 120 -Name 'SelectSaRequestButton'
    $select.Anchor = 'Bottom, Right'
    $select.Enabled = $false
    $cancel = New-SentinelButton -Text '取消' -X 680 -Y 385 -Width 100
    $cancel.Anchor = 'Bottom, Right'
    $cancel.DialogResult = 'Cancel'
    $dialog.CancelButton = $cancel
    $list.Add_SelectedIndexChanged({
        $select.Enabled = $list.SelectedItems.Count -eq 1
        $details.Text = ''
        if ($select.Enabled) {
            $request = $list.SelectedItems[0].Tag
            $details.Text = (@($request.FolderPath, $request.RequestPath) + @($request.Conflicts)) -join [Environment]::NewLine
        }
    }.GetNewClosure())
    $select.Add_Click({
        if ($list.SelectedItems.Count -eq 1) {
            $dialog.Tag = $list.SelectedItems[0].Tag
            $dialog.DialogResult = 'OK'
        }
    }.GetNewClosure())
    $dialog.Controls.AddRange(@($list, $details, $select, $cancel))
    Set-SentinelControlTheme -Control $dialog -Theme (Get-SentinelUiTheme)
    return $dialog
}

function Show-SentinelSaRequestDialog {
    param([Parameter(Mandatory = $true)] [object[]] $Requests)
    $dialog = New-SentinelSaRequestDialog -Requests $Requests
    try {
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.Tag }
    }
    finally { $dialog.Dispose() }
}

function Confirm-SentinelSaRequest {
    param([Parameter(Mandatory = $true)] [object] $Request)
    $message = (@($Request.FolderPath, $Request.RequestPath) + @($Request.Conflicts) + @('', '冲突位置留空，其他数据继续，是否继续？')) -join [Environment]::NewLine
    return ([System.Windows.Forms.MessageBox]::Show($message, '处理编号冲突',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2) -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Invoke-SentinelSaConflictReview {
    param([Parameter(Mandatory = $true)] [string] $ScriptRoot)
    $settings = Read-PolishingSentinelSettings -ScriptRoot $ScriptRoot
    if (-not $settings.PSObject.Properties['SaEngine']) {
        $settings | Add-Member -NotePropertyName SaEngine -NotePropertyValue 'BuiltIn'
    }
    $config = ConvertTo-ExperimentAutomationConfigFromSettings -ScriptRoot $ScriptRoot -Settings $settings
    $requests = @(Get-ExperimentSaConfirmations -Config $config)
    if ($requests.Count -eq 0) { return '没有待处理的编号冲突。' }
    $request = Show-SentinelSaRequestDialog -Requests $requests
    if ($null -eq $request) { return '未选择请求，保持暂停。' }
    if (-not (Confirm-SentinelSaRequest -Request $request)) { return '未确认继续，保持暂停。' }
    Resolve-ExperimentSaConfirmation -Config $config -RequestPath $request.RequestPath -Decision ContinueWithBlanks -ExpectedFingerprint $request.Fingerprint | Out-Null
    return '已确认冲突位置留空，下次扫描时继续处理。'
}

function Add-PolishingSentinelLogsTab {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.TabPage] $Page,
        [Parameter(Mandatory = $true)] [pscustomobject] $Settings,
        [Parameter(Mandatory = $true)] [string] $ScriptRoot
    )

    $logBox = New-Object System.Windows.Forms.TextBox
    $logBox.Name = 'RecentLogText'
    $logBox.Multiline = $true
    $logBox.ReadOnly = $true
    $logBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $logBox.Location = New-Object System.Drawing.Point(30, 30)
    $logBox.Size = New-Object System.Drawing.Size(880, 420)
    $Page.Controls.Add($logBox)

    $refreshLogButton = New-SentinelButton -Text '刷新日志' -X 30 -Y 470 -Width 110 -Name 'RefreshLogButton'
    $refreshLogButton.Add_Click({
        $logPath = Join-Path $Settings.LogDirectory 'automation.log'
        if (Test-Path -LiteralPath $logPath) {
            $logBox.Text = (Get-Content -LiteralPath $logPath -Tail 100 -Encoding UTF8) -join [Environment]::NewLine
            $logBox.SelectionStart = $logBox.Text.Length
            $logBox.ScrollToCaret()
        }
        else {
            $logBox.Text = '尚未找到日志文件。'
        }
    }.GetNewClosure())
    $Page.Controls.Add($refreshLogButton)

    $openBackupButton = New-SentinelButton -Text '打开备份目录' -X 150 -Y 470 -Width 130 -Name 'OpenBackupDirectoryButton'
    $openBackupButton.Add_Click({
        try {
            if (-not (Test-Path -LiteralPath $Settings.BackupDirectory)) {
                New-Item -ItemType Directory -Path $Settings.BackupDirectory -Force | Out-Null
            }
            Start-Process -FilePath $Settings.BackupDirectory | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(("打开备份目录失败：{0}" -f $_.Exception.Message), '抛光哨站') | Out-Null
        }
    }.GetNewClosure())
    $Page.Controls.Add($openBackupButton)

    $conflictStatus = New-SentinelLabel -Text '' -X 30 -Y 512 -Width 880 -Name 'SaConflictStatusLabel'
    $Page.Controls.Add($conflictStatus)
    $conflictButton = New-SentinelButton -Text '处理编号冲突' -X 300 -Y 470 -Width 140 -Name 'ResolveSaConflictsButton'
    $conflictButton.Add_Click({
        try { $conflictStatus.Text = Invoke-SentinelSaConflictReview -ScriptRoot $ScriptRoot }
        catch { $conflictStatus.Text = '处理编号冲突失败：{0}' -f $_.Exception.Message }
    }.GetNewClosure())
    $Page.Controls.Add($conflictButton)
}

function Add-PolishingSentinelStartupTab {
    param(
        [Parameter(Mandatory = $true)] [System.Windows.Forms.TabPage] $Page,
        [Parameter(Mandatory = $true)] [pscustomobject] $Settings,
        [Parameter(Mandatory = $true)] [string] $ScriptRoot
    )

    $taskStatusLabel = New-SentinelLabel -Text '' -X 30 -Y 30 -Width 760 -Name 'StartupStatusLabel'
    $Page.Controls.Add($taskStatusLabel)

    $newPowerShellArguments = (Get-Command New-SentinelPowerShellFileArguments -CommandType Function -ErrorAction Stop).ScriptBlock

    $refreshStartup = {
        if ($null -ne $taskStatusLabel.Tag) { & $taskStatusLabel.Tag.Refresh }
    }.GetNewClosure()

    $refreshStartupButton = New-SentinelButton -Text '刷新' -X 30 -Y 70 -Width 90 -Name 'RefreshStartupStatusButton'
    $refreshStartupButton.Add_Click($refreshStartup)
    $Page.Controls.Add($refreshStartupButton)

    $powershellExe = Get-SentinelPowerShellExe

    $installButton = New-SentinelButton -Text '安装开机自启动' -X 30 -Y 120 -Width 180 -Name 'InstallStartupTaskButton'
    $installButton.Add_Click({
        $scriptPath = Join-Path $ScriptRoot 'Install-ExperimentDataAutomationTask.ps1'
        Start-Process -FilePath $powershellExe -ArgumentList (& $newPowerShellArguments -ScriptPath $scriptPath) -WindowStyle Hidden -Wait | Out-Null
        & $refreshStartup
    }.GetNewClosure())
    $Page.Controls.Add($installButton)

    $uninstallButton = New-SentinelButton -Text '取消开机自启动' -X 230 -Y 120 -Width 180 -Name 'UninstallStartupTaskButton'
    $uninstallButton.Add_Click({
        $scriptPath = Join-Path $ScriptRoot 'Uninstall-ExperimentDataAutomationTask.ps1'
        Start-Process -FilePath $powershellExe -ArgumentList (& $newPowerShellArguments -ScriptPath $scriptPath) -WindowStyle Hidden -Wait | Out-Null
        & $refreshStartup
    }.GetNewClosure())
    $Page.Controls.Add($uninstallButton)

    $startMonitorButton = New-SentinelButton -Text '启动后台监控' -X 30 -Y 170 -Width 180 -Name 'StartMonitorButton'
    $startMonitorButton.Add_Click({
        $scriptPath = Join-Path $ScriptRoot 'Start-ExperimentDataAutomation.ps1'
        Start-Process -FilePath $powershellExe -ArgumentList (& $newPowerShellArguments -ScriptPath $scriptPath) -WindowStyle Hidden | Out-Null
        $taskStatusLabel.Text = '已请求启动后台监控。'
    }.GetNewClosure())
    $Page.Controls.Add($startMonitorButton)

    $stopMonitorButton = New-SentinelButton -Text '停止后台监控' -X 230 -Y 170 -Width 180 -Name 'StopMonitorButton'
    $stopMonitorButton.Add_Click({
        $scriptPath = Join-Path $ScriptRoot 'Stop-ExperimentDataAutomation.ps1'
        Start-Process -FilePath $powershellExe -ArgumentList (& $newPowerShellArguments -ScriptPath $scriptPath) -WindowStyle Hidden -Wait | Out-Null
        & $refreshStartup
    }.GetNewClosure())
    $Page.Controls.Add($stopMonitorButton)

    $taskStatusLabel.Text = '正在读取计划任务和监控进程状态…'
}

function New-PolishingSentinelMainForm {
    param([Parameter(Mandatory = $true)] [string] $ScriptRoot)

    $settings = Read-PolishingSentinelSettings -ScriptRoot $ScriptRoot
    if (-not $settings.PSObject.Properties['SaEngine']) {
        $settings | Add-Member -NotePropertyName SaEngine -NotePropertyValue 'BuiltIn'
    }
    try {
        $status = Read-PolishingSentinelStatus -ScriptRoot $ScriptRoot
        if ($null -eq $status) { throw '运行状态文件为空' }
    }
    catch {
        $status = [pscustomobject]@{ State = 'Unknown'; Message = '运行状态读取失败，将自动重试'; UpdatedAt = '' }
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = '抛光哨站'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(1280, 820)
    $form.MinimumSize = New-Object System.Drawing.Size(1180, 720)
    $form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 13)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Name = 'MainTabs'
    $tabs.Dock = [System.Windows.Forms.DockStyle]::Fill
    $form.Controls.Add($tabs)

    $overview = New-SentinelPlaceholderPage -Text '总览'
    $paths = New-SentinelPlaceholderPage -Text '路径设置'
    $naming = New-SentinelPlaceholderPage -Text '命名约定'
    $matching = New-SentinelPlaceholderPage -Text '匹配规则'
    $process = New-SentinelPlaceholderPage -Text '固定工艺参数'
    $logs = New-SentinelPlaceholderPage -Text '待确认与日志'
    $startup = New-SentinelPlaceholderPage -Text '自启动'

    [void] $tabs.TabPages.Add($overview)
    [void] $tabs.TabPages.Add($paths)
    [void] $tabs.TabPages.Add($naming)
    [void] $tabs.TabPages.Add($matching)
    [void] $tabs.TabPages.Add($process)
    [void] $tabs.TabPages.Add($logs)
    [void] $tabs.TabPages.Add($startup)

    $statusGroup = New-Object System.Windows.Forms.GroupBox
    $statusGroup.Text = '运行状态'
    $statusGroup.Location = New-Object System.Drawing.Point(20, 20)
    $statusGroup.Size = New-Object System.Drawing.Size(1020, 186)
    $overview.Controls.Add($statusGroup)

    $statusLabels = @{
        State     = New-SentinelLabel -Text '' -X 150 -Y 28 -Width 820 -Name 'StatusStateLabel'
        Message   = New-SentinelLabel -Text '' -X 150 -Y 64 -Width 820 -Name 'StatusMessageLabel'
        UpdatedAt = New-SentinelLabel -Text '' -X 150 -Y 100 -Width 360 -Name 'StatusUpdatedAtLabel'
        Health    = New-SentinelLabel -Text '' -X 150 -Y 136 -Width 820 -Name 'StatusHealthLabel'
    }

    $statusGroup.Controls.Add((New-SentinelLabel -Text '当前状态' -X 24 -Y 28 -Width 100))
    $statusGroup.Controls.Add((New-SentinelLabel -Text '状态说明' -X 24 -Y 64 -Width 100))
    $statusGroup.Controls.Add((New-SentinelLabel -Text '更新时间' -X 24 -Y 100 -Width 100))
    $statusGroup.Controls.Add((New-SentinelLabel -Text '后台健康' -X 24 -Y 136 -Width 100))
    $statusGroup.Controls.Add($statusLabels.State)
    $statusGroup.Controls.Add($statusLabels.Message)
    $statusGroup.Controls.Add($statusLabels.UpdatedAt)
    $statusGroup.Controls.Add($statusLabels.Health)
    Set-SentinelStatusLabels -Labels $statusLabels -Status $status

    $overviewActions = New-Object System.Windows.Forms.GroupBox
    $overviewActions.Text = '操作'
    $overviewActions.Location = New-Object System.Drawing.Point(20, 226)
    $overviewActions.Size = New-Object System.Drawing.Size(1020, 90)
    $overview.Controls.Add($overviewActions)

    $refreshButton = New-SentinelButton -Text '刷新状态' -X 24 -Y 34 -Width 110 -Name 'RefreshStatusButton'
    $refreshButton.Add_Click({
        $startupStatusLabel = $form.Controls.Find('StartupStatusLabel', $true)[0]
        if ($null -ne $startupStatusLabel.Tag) { & $startupStatusLabel.Tag.Refresh }
    }.GetNewClosure())
    $overviewActions.Controls.Add($refreshButton)

    $scanButton = New-SentinelButton -Text '立即扫描' -X 150 -Y 34 -Width 110 -Name 'RunScanNowButton'
    $scanButton.Add_Click({
        try {
            $powershellExe = Join-Path $PSHOME 'powershell.exe'
            if (-not (Test-Path -LiteralPath $powershellExe)) {
                $powershellExe = 'powershell'
            }

            Start-Process -FilePath $powershellExe -ArgumentList @(
                '-NoProfile',
                '-ExecutionPolicy',
                'Bypass',
                '-File',
                (Join-Path $ScriptRoot 'Start-ExperimentDataAutomation.ps1'),
                '-RunOnce'
            ) -WindowStyle Hidden | Out-Null

            $statusLabels.Message.Text = '已启动一次扫描。'
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(("启动扫描失败：{0}" -f $_.Exception.Message), '抛光哨站') | Out-Null
        }
    }.GetNewClosure())
    $overviewActions.Controls.Add($scanButton)

    $settingsGroup = New-Object System.Windows.Forms.GroupBox
    $settingsGroup.Text = '路径设置'
    $settingsGroup.Location = New-Object System.Drawing.Point(20, 20)
    $settingsGroup.Size = New-Object System.Drawing.Size(920, 330)
    $paths.Controls.Add($settingsGroup)

    $settingsControls = @{}
    $y = 30
    Add-SentinelSettingRow -Parent $settingsGroup -Controls $settingsControls -Key 'VolumeLabel' -Label 'U盘卷标' -Value $settings.VolumeLabel -Y $y
    $y += 32
    Add-SentinelSettingRow -Parent $settingsGroup -Controls $settingsControls -Key 'ExpectedDriveLetter' -Label '预期盘符' -Value $settings.ExpectedDriveLetter -Y $y
    $y += 32
    Add-SentinelSettingRow -Parent $settingsGroup -Controls $settingsControls -Key 'WaterSourceRelativePath' -Label '水基目录' -Value $settings.WaterSourceRelativePath -Y $y
    $y += 32
    Add-SentinelSettingRow -Parent $settingsGroup -Controls $settingsControls -Key 'OilSourceRelativePath' -Label '油基目录' -Value $settings.OilSourceRelativePath -Y $y
    $y += 32
    Add-SentinelSettingRow -Parent $settingsGroup -Controls $settingsControls -Key 'MasterWorkbookPath' -Label '主表' -Value $settings.MasterWorkbookPath -Y $y
    $y += 32
    $settingsGroup.Controls.Add((New-SentinelLabel -Text 'Sa处理方式' -X 20 -Y $y -Width 150))
    $saEngineCombo = New-Object System.Windows.Forms.ComboBox
    $saEngineCombo.Name = 'SaEngineCombo'
    $saEngineCombo.DropDownStyle = 'DropDownList'
    $saEngineCombo.Location = New-Object System.Drawing.Point(180, $y)
    $saEngineCombo.Size = New-Object System.Drawing.Size(300, 26)
    [void] $saEngineCombo.Items.Add('内置（无需 MATLAB）')
    [void] $saEngineCombo.Items.Add('MATLAB（兼容）')
    $saEngineCombo.SelectedIndex = if ($settings.SaEngine -eq 'Matlab') { 1 } else { 0 }
    $settingsGroup.Controls.Add($saEngineCombo)
    $y += 32
    Add-SentinelSettingRow -Parent $settingsGroup -Controls $settingsControls -Key 'MatlabScriptPath' -Label 'MATLAB 程序' -Value $settings.MatlabScriptPath -Y $y
    $y += 32
    Add-SentinelSettingRow -Parent $settingsGroup -Controls $settingsControls -Key 'MatlabExePath' -Label 'MATLAB exe' -Value $settings.MatlabExePath -Y $y
    $updateSaEngine = {
        $settingsControls['MatlabScriptPath'].Enabled = $saEngineCombo.SelectedIndex -eq 1
        $settingsControls['MatlabExePath'].Enabled = $saEngineCombo.SelectedIndex -eq 1
    }.GetNewClosure()
    $saEngineCombo.Add_SelectedIndexChanged($updateSaEngine)
    & $updateSaEngine

    $saveStatusLabel = New-SentinelLabel -Text '' -X 310 -Y 294 -Width 560 -Name 'PathSettingsSaveStatusLabel'
    $settingsGroup.Controls.Add($saveStatusLabel)

    $pathCheckGroup = New-Object System.Windows.Forms.GroupBox
    $pathCheckGroup.Text = '路径检查'
    $pathCheckGroup.Location = New-Object System.Drawing.Point(20, 360)
    $pathCheckGroup.Size = New-Object System.Drawing.Size(920, 180)
    $paths.Controls.Add($pathCheckGroup)

    $pathCheckList = New-Object System.Windows.Forms.ListView
    $pathCheckList.Name = 'PathCheckList'
    $pathCheckList.View = [System.Windows.Forms.View]::Details
    $pathCheckList.FullRowSelect = $true
    $pathCheckList.GridLines = $true
    $pathCheckList.Location = New-Object System.Drawing.Point(20, 28)
    $pathCheckList.Size = New-Object System.Drawing.Size(880, 140)
    [void] $pathCheckList.Columns.Add('项目', 160)
    [void] $pathCheckList.Columns.Add('状态', 80)
    [void] $pathCheckList.Columns.Add('路径', 700)
    $pathCheckGroup.Controls.Add($pathCheckList)
    Update-SentinelPathCheckList -ListView $pathCheckList -Settings $settings

    $saveButton = New-SentinelButton -Text '保存设置' -X 180 -Y 290 -Width 110 -Name 'SavePathSettingsButton'
    $saveButton.Add_Click({
        try {
            $current = Read-PolishingSentinelSettings -ScriptRoot $ScriptRoot
            foreach ($key in $settingsControls.Keys) {
                $current | Add-Member -NotePropertyName $key -NotePropertyValue $settingsControls[$key].Text.Trim() -Force
            }
            $engine = if ($saEngineCombo.SelectedIndex -eq 1) { 'Matlab' } else { 'BuiltIn' }
            $current | Add-Member -NotePropertyName SaEngine -NotePropertyValue $engine -Force

            Save-PolishingSentinelSettings -ScriptRoot $ScriptRoot -Settings $current | Out-Null
            Update-SentinelPathCheckList -ListView $pathCheckList -Settings $current
            $saveStatusLabel.Text = ('已保存：{0}' -f (Get-Date).ToString('HH:mm:ss'))
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(("保存设置失败：{0}" -f $_.Exception.Message), '抛光哨站') | Out-Null
        }
    }.GetNewClosure())
    $settingsGroup.Controls.Add($saveButton)

    Add-PolishingSentinelNamingTab -Page $naming -Settings $settings -ScriptRoot $ScriptRoot
    Add-PolishingSentinelMatchingTab -Page $matching -Settings $settings -ScriptRoot $ScriptRoot
    Add-PolishingSentinelProcessTab -Page $process -Settings $settings -ScriptRoot $ScriptRoot
    $tabs.Add_Deselecting({
        param($sender, $eventArgs)
        if (-not (Confirm-SentinelEditorLeave $eventArgs.TabPage)) { $eventArgs.Cancel = $true }
    }.GetNewClosure())
    $form.Add_FormClosing({
        param($sender, $eventArgs)
        foreach ($page in @($matching, $process)) {
            if (-not (Confirm-SentinelEditorLeave $page)) { $eventArgs.Cancel = $true; break }
        }
    }.GetNewClosure())
    Add-PolishingSentinelLogsTab -Page $logs -Settings $settings -ScriptRoot $ScriptRoot
    Add-PolishingSentinelStartupTab -Page $startup -Settings $settings -ScriptRoot $ScriptRoot

    $startupStatusLabel = $form.Controls.Find('StartupStatusLabel', $true)[0]
    Start-SentinelStatusAutoRefresh -Form $form -Labels $statusLabels -StartupLabel $startupStatusLabel -ScriptRoot $ScriptRoot -TaskName $settings.TaskName | Out-Null

    Set-SentinelControlTheme -Control $form -Theme (Get-SentinelUiTheme)
    [void]$tabs.TabPages.Add((New-SentinelAnalysisPage -Form $form -ScriptRoot $ScriptRoot))
    Initialize-SentinelDashboard -Form $form -Tabs $tabs -Overview $overview -ScriptRoot $ScriptRoot
    $statusLabels.ProgressBar=$overview.Tag.ProgressBar
    $statusLabels.ProgressText=$overview.Tag.ProgressText
    Set-SentinelStatusLabels -Labels $statusLabels -Status $status

    return $form
}

function Show-PolishingSentinelWindow {
    param([Parameter(Mandatory = $true)] [string] $ScriptRoot)

    $singleInstance = Enter-PolishingSentinelSingleInstance
    if (-not $singleInstance.Acquired) {
        return
    }

    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form = $null
    try {
        $form = New-PolishingSentinelMainForm -ScriptRoot $ScriptRoot
        [void] [System.Windows.Forms.Application]::Run($form)
    }
    finally {
        if ($form) {
            $form.Dispose()
        }
        Exit-PolishingSentinelSingleInstance -Handle $singleInstance
    }
}
