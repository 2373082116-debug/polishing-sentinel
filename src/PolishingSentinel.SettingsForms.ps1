Set-StrictMode -Version 2.0

function Get-SentinelEditorDraft {
    param($Editor)
    if ($Editor.Kind -eq 'Matching') {
        $rows = @(
            foreach ($row in $Editor.Grid.Rows) {
                if (-not $row.IsNewRow) {
                    $values = @([string]$row.Cells[0].Value, [string]$row.Cells[1].Value)
                    if ($Editor.Grid.IsCurrentCellInEditMode -and $Editor.Grid.EditingControl -and $Editor.Grid.CurrentCell.RowIndex -eq $row.Index) {
                        $values[$Editor.Grid.CurrentCell.ColumnIndex] = $Editor.Grid.EditingControl.Text
                    }
                    [pscustomobject]@{ Slurry = $values[0]; Sheet = $values[1] }
                }
            }
        )
        return [pscustomobject]@{ Enabled = $Editor.Check.Checked; Mode = $Editor.Mode.SelectedIndex; Rows = $rows }
    }
    $values = [ordered]@{}
    foreach ($key in @($Editor.Inputs.Keys | Sort-Object)) { $values[$key] = $Editor.Inputs[$key].Text }
    return [pscustomobject]$values
}

function Update-SentinelEditorState {
    param($Editor)
    if ($Editor.Loading) { return }
    $draft = Get-SentinelEditorDraft $Editor | ConvertTo-Json -Depth 10 -Compress
    $Editor.Dirty = $draft -cne $Editor.Baseline
    $Editor.Save.Enabled = $Editor.Dirty
    $Editor.Cancel.Enabled = $Editor.Dirty -or $Editor.Editing
    $Editor.Status.ForeColor = [Drawing.Color]::FromArgb(58,91,124)
    $Editor.Status.Text = if ($Editor.Dirty) { '有未保存修改' } elseif ($Editor.Editing) { '编辑中 · 未修改' } else { '未修改' }
}

function Update-SentinelProcessTitles {
    param($Editor)
    foreach ($group in @('Water','Oil')) {
        $slurry = if ($group -eq 'Water') { '硅溶胶' } else { '金刚石' }
        $name = Get-SentinelObjectProperty -Object $Editor.Settings.SlurryMappings -Name $slurry -Default '未配置'
        $prefix = if ($group -eq 'Water') { '水基' } else { '油基' }
        $Editor.Headings[$group].Text = '{0}  /  {1}' -f $prefix,$name
    }
}

function Reset-SentinelEditor {
    param($Editor)
    $Editor.Loading = $true
    try {
        $Editor.Errors.Clear()
        if ($Editor.Kind -eq 'Matching') {
            $Editor.Grid.Rows.Clear()
            $Editor.Check.Checked = [bool]$Editor.Settings.EnableBackfill
            $index = [array]::IndexOf(@('Auto','Preview','Disabled'), [string]$Editor.Settings.NewBlockMode)
            $Editor.Mode.SelectedIndex = if ($index -lt 0) { 1 } else { $index }
            $mappings = $Editor.Settings.SlurryMappings
            if ($mappings -is [Collections.IDictionary]) {
                foreach ($key in $mappings.Keys) { [void]$Editor.Grid.Rows.Add([string]$key,[string]$mappings[$key]) }
            } else {
                foreach ($property in $mappings.PSObject.Properties) {
                    [void]$Editor.Grid.Rows.Add([string]$property.Name,[string]$property.Value)
                }
            }
        } else {
            foreach ($key in $Editor.Inputs.Keys) {
                $group,$field = $key.Split('.')
                $Editor.Inputs[$key].Text = [string](Get-SentinelObjectProperty -Object $Editor.Settings.ProcessDefaults.$group -Name $field -Default '')
                $Editor.Inputs[$key].ReadOnly = $true
            }
            $Editor.Editing = $false
            $Editor.Edit.Enabled = $true
            Update-SentinelProcessTitles $Editor
        }
        $Editor.Baseline = Get-SentinelEditorDraft $Editor | ConvertTo-Json -Depth 10 -Compress
    } finally { $Editor.Loading = $false }
    Update-SentinelEditorState $Editor
}

function Save-SentinelEditor {
    param($Editor)
    try {
        $Editor.Errors.Clear()
        $valid = $true
        if ($Editor.Kind -eq 'Matching') {
            [void]$Editor.Grid.EndEdit()
            $mappings = [ordered]@{}
            $seen = @{}
            foreach ($row in $Editor.Grid.Rows) {
                if ($row.IsNewRow) { continue }
                $row.ErrorText = ''
                $slurry = ([string]$row.Cells[0].Value).Trim()
                $sheet = ([string]$row.Cells[1].Value).Trim()
                if (-not $slurry -and -not $sheet) { continue }
                $errorText = ''
                if (-not $slurry -or -not $sheet) { $errorText = '请填写抛光液和目标工作表' }
                elseif ($seen.ContainsKey($slurry)) { $errorText = '抛光液名称重复' }
                elseif ($sheet.Length -gt 31 -or $sheet.IndexOfAny([char[]]':\/?*[]') -ge 0 -or $sheet.StartsWith("'") -or $sheet.EndsWith("'")) {
                    $errorText = '工作表名称无效：最多31字，不能包含 : \\ / ? * [ ] 或首尾单引号'
                }
                if ($errorText) { $row.ErrorText = $errorText; $valid = $false }
                if ($slurry) { $seen[$slurry] = $true }
                if ($slurry -and -not $mappings.Contains($slurry)) { $mappings[$slurry] = $sheet }
            }
            if ($mappings.Count -eq 0) { throw '至少保留一条有效映射' }
            if (-not $valid) { throw '请修正映射表中标记的错误' }
        } else {
            foreach ($key in $Editor.Inputs.Keys) {
                if ([string]::IsNullOrWhiteSpace($Editor.Inputs[$key].Text)) {
                    $Editor.Errors.SetError($Editor.Inputs[$key], '此参数不能为空')
                    $valid = $false
                }
            }
            if (-not $valid) { throw '请填写标记为必填的工艺参数' }
        }
        # Merge into the latest configuration so other pages and extension fields survive.
        $current = Read-PolishingSentinelSettings -ScriptRoot $Editor.ScriptRoot
        if ($Editor.Kind -eq 'Matching') {
            $current.EnableBackfill = $Editor.Check.Checked
            $current.NewBlockMode = @('Auto','Preview','Disabled')[$Editor.Mode.SelectedIndex]
            $current.SlurryMappings = [pscustomobject]$mappings
        } else {
            foreach ($key in $Editor.Inputs.Keys) {
                $group,$field = $key.Split('.')
                $current.ProcessDefaults.$group | Add-Member -NotePropertyName $field -NotePropertyValue $Editor.Inputs[$key].Text.Trim() -Force
            }
        }
        Save-PolishingSentinelSettings -ScriptRoot $Editor.ScriptRoot -Settings $current | Out-Null
        $fields = if ($Editor.Kind -eq 'Matching') { @('EnableBackfill','NewBlockMode','SlurryMappings') } else { @('ProcessDefaults') }
        foreach ($field in $fields) {
            $Editor.Settings | Add-Member -NotePropertyName $field -NotePropertyValue $current.$field -Force
        }
        Reset-SentinelEditor $Editor
        $Editor.Status.Text = '已保存 · ' + (Get-Date -Format HH:mm:ss)
        return $true
    } catch {
        Update-SentinelEditorState $Editor
        $Editor.Status.ForeColor = [Drawing.Color]::Firebrick
        $Editor.Status.Text = '保存失败：' + $_.Exception.Message
        return $false
    }
}

function Confirm-SentinelEditorLeave {
    param($Page, [scriptblock]$Confirm)
    if ($null -eq $Page -or $null -eq $Page.Tag -or -not $Page.Tag.PSObject.Properties['Editor']) { return $true }
    $editor = $Page.Tag.Editor
    if ($editor.Kind -eq 'Matching') { [void]$editor.Grid.EndEdit() }
    Update-SentinelEditorState $editor
    if (-not $editor.Dirty) { return $true }
    $discard = if ($Confirm) { & $Confirm } else {
        [Windows.Forms.MessageBox]::Show($Page.FindForm(), '有未保存的修改，是否放弃？', $Page.Text,
            [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Warning,
            [Windows.Forms.MessageBoxDefaultButton]::Button2) -eq [Windows.Forms.DialogResult]::Yes
    }
    if (-not $discard) { return $false }
    Reset-SentinelEditor $editor
    return $true
}

function New-SentinelSettingsEditor {
    param($Page,$Settings,$ScriptRoot,$Kind)
    $Page.Padding = New-Object Windows.Forms.Padding(24)
    $Page.AutoScroll = $true
    $layout = New-Object Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.MinimumSize = New-Object Drawing.Size(720,400)
    $layout.ColumnCount = 1
    [void]$layout.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent',100)))
    $layout.RowCount = 2
    [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent',100)))
    [void]$layout.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute',100)))
    $Page.Controls.Add($layout)
    $footer = New-Object Windows.Forms.FlowLayoutPanel
    $footer.Dock = 'Fill'
    $footer.FlowDirection = 'LeftToRight'
    $save = New-SentinelButton -Text '保存修改' -X 0 -Y 0 -Width 120 -Name ("Save{0}SettingsButton" -f $Kind)
    $cancel = New-SentinelButton -Text '取消修改' -X 0 -Y 0 -Width 120 -Name ("Cancel{0}SettingsButton" -f $Kind)
    $status = New-Object Windows.Forms.Label
    $status.AutoSize = $false
    $status.Height = 48
    $status.Width = 670
    $status.Margin = New-Object Windows.Forms.Padding(3,10,3,0)
    $footer.Controls.Add($save)
    $footer.Controls.Add($cancel)
    $footer.SetFlowBreak($cancel,$true)
    $footer.Controls.Add($status)
    $layout.Controls.Add($footer,0,1)
    $errors = New-Object Windows.Forms.ErrorProvider
    $errors.BlinkStyle = 'NeverBlink'
    $errors.ContainerControl = $Page.FindForm()
    $editor = [pscustomobject]@{
        Kind=$Kind; Page=$Page; Settings=$Settings; ScriptRoot=$ScriptRoot; Layout=$layout; Footer=$footer
        Save=$save; Cancel=$cancel; Status=$status; Errors=$errors; Dirty=$false; Loading=$true
        Baseline=''; Editing=$false; Grid=$null; Check=$null; Mode=$null; Edit=$null
        Inputs=@{}; Headings=@{}
    }
    $Page.Tag = [pscustomobject]@{Editor=$editor}
    $save.Add_Click({ [void](Save-SentinelEditor $editor) }.GetNewClosure())
    $cancel.Add_Click({ Reset-SentinelEditor $editor }.GetNewClosure())
    $Page.Add_Disposed({ $errors.Dispose() }.GetNewClosure())
    return $editor
}

function Add-PolishingSentinelMatchingTab {
    param($Page,$Settings,$ScriptRoot)
    $editor = New-SentinelSettingsEditor $Page $Settings $ScriptRoot 'Matching'
    $body = New-Object Windows.Forms.TableLayoutPanel
    $body.Dock = 'Fill'
    $body.ColumnCount = 1
    [void]$body.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent',100)))
    $body.RowCount = 3
    foreach ($height in @(56,42)) { [void]$body.RowStyles.Add((New-Object Windows.Forms.RowStyle('Absolute',$height))) }
    [void]$body.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent',100)))
    $editor.Layout.Controls.Add($body,0,0)
    $options = New-Object Windows.Forms.FlowLayoutPanel
    $options.Dock = 'Fill'
    $check = New-Object Windows.Forms.CheckBox
    $check.Name = 'EnableBackfillCheckBox'
    $check.Text = '自动回写已有实验'
    $check.Size = New-Object Drawing.Size(220,30)
    $label = New-Object Windows.Forms.Label
    $label.Text = '新实验处理'
    $label.Size = New-Object Drawing.Size(100,30)
    $label.TextAlign = 'MiddleLeft'
    $mode = New-Object Windows.Forms.ComboBox
    $mode.Name = 'NewBlockModeComboBox'
    $mode.DropDownStyle = 'DropDownList'
    $mode.Width = 230
    $mode.Items.AddRange([object[]]@('自动新建','仅生成待确认预览','不新建'))
    $options.Controls.AddRange([Windows.Forms.Control[]]@($check,$label,$mode))
    $body.Controls.Add($options,0,0)
    $toolbar = New-Object Windows.Forms.FlowLayoutPanel
    $toolbar.Dock = 'Fill'
    $title = New-Object Windows.Forms.Label
    $title.Text = '抛光液与工作表'
    $title.Size = New-Object Drawing.Size(220,32)
    $title.TextAlign = 'MiddleLeft'
    $add = New-SentinelButton -Text '+' -X 0 -Y 0 -Width 36 -Name 'AddMappingButton'
    $remove = New-SentinelButton -Text '−' -X 0 -Y 0 -Width 36 -Name 'RemoveMappingButton'
    $tip = New-Object Windows.Forms.ToolTip
    $tip.SetToolTip($add,'新增映射')
    $tip.SetToolTip($remove,'删除选中映射')
    $toolbar.Controls.AddRange([Windows.Forms.Control[]]@($title,$add,$remove))
    $body.Controls.Add($toolbar,0,1)
    $grid = New-Object Windows.Forms.DataGridView
    $grid.Name = 'SlurryMappingsGrid'
    $grid.Dock = 'Fill'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.MultiSelect = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.BackgroundColor = [Drawing.Color]::FromArgb(246,251,255)
    $grid.BorderStyle = 'None'
    $grid.CellBorderStyle = 'SingleHorizontal'
    $grid.GridColor = [Drawing.Color]::FromArgb(210,225,236)
    $grid.RowTemplate.Height = 36
    $grid.ColumnHeadersHeight = 38
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersDefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(18,70,126)
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = [Drawing.Color]::White
    $grid.DefaultCellStyle.SelectionBackColor = [Drawing.Color]::FromArgb(207,231,246)
    $grid.DefaultCellStyle.SelectionForeColor = [Drawing.Color]::FromArgb(15,43,74)
    [void]$grid.Columns.Add('Slurry','抛光液')
    [void]$grid.Columns.Add('Sheet','目标工作表')
    $body.Controls.Add($grid,0,2)
    $editor.Grid=$grid; $editor.Check=$check; $editor.Mode=$mode
    $changed = { Update-SentinelEditorState $editor }.GetNewClosure()
    $check.Add_CheckedChanged($changed)
    $mode.Add_SelectedIndexChanged($changed)
    $grid.Add_CellValueChanged($changed)
    $grid.Add_RowsRemoved($changed)
    $grid.Add_EditingControlShowing({
        param($sender,$eventArgs)
        if ($null -eq $eventArgs.Control.Tag) {
            $eventArgs.Control.Add_TextChanged($editor.Grid.Tag)
            $eventArgs.Control.Tag = 'SentinelDraftTracking'
        }
    }.GetNewClosure())
    $grid.Tag = {
        Update-SentinelEditorState $editor
    }.GetNewClosure()
    $add.Add_Click({
        $index=$grid.Rows.Add('','')
        $grid.CurrentCell=$grid.Rows[$index].Cells[0]
        [void]$grid.BeginEdit($true)
        Update-SentinelEditorState $editor
    }.GetNewClosure())
    $remove.Add_Click({
        if ($grid.CurrentRow) { $grid.Rows.RemoveAt($grid.CurrentRow.Index) }
    }.GetNewClosure())
    $grid.Add_SelectionChanged({ $remove.Enabled=$null -ne $grid.CurrentRow }.GetNewClosure())
    $Page.Add_Disposed({ $tip.Dispose() }.GetNewClosure())
    Reset-SentinelEditor $editor
}

function Add-PolishingSentinelProcessTab {
    param($Page,$Settings,$ScriptRoot)
    $editor = New-SentinelSettingsEditor $Page $Settings $ScriptRoot 'Process'
    $body = New-Object Windows.Forms.TableLayoutPanel
    $body.Dock='Fill'
    $body.ColumnCount=2
    $body.RowCount=1
    [void]$body.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent',100)))
    foreach ($i in 1..2) { [void]$body.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent',50))) }
    $editor.Layout.Controls.Add($body,0,0)
    $column=0
    foreach ($group in @('Water','Oil')) {
        $fields=New-Object Windows.Forms.TableLayoutPanel
        $fields.Dock='Fill'
        $fields.AutoSize=$false
        $fields.ColumnCount=2
        $fields.RowCount=8
        foreach ($i in 1..7) { [void]$fields.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize'))) }
        [void]$fields.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent',100)))
        $fields.Padding=New-Object Windows.Forms.Padding(0,0,28,0)
        [void]$fields.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Absolute',90)))
        [void]$fields.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent',100)))
        $heading=New-Object Windows.Forms.Label
        $heading.Dock='Fill'
        $heading.AutoSize=$true
        $heading.Padding=New-Object Windows.Forms.Padding(0,0,0,14)
        $heading.Font=New-Object Drawing.Font('Microsoft YaHei UI',10,[Drawing.FontStyle]::Bold)
        $fields.Controls.Add($heading,0,0)
        $fields.SetColumnSpan($heading,2)
        $editor.Headings[$group]=$heading
        $row=1
        foreach ($definition in Get-ExperimentProcessFieldDefinitions) {
            $label=New-Object Windows.Forms.Label
            $label.Text=$definition.Label
            $label.Dock='Fill'
            $label.TextAlign='MiddleLeft'
            $input=New-Object Windows.Forms.TextBox
            $input.Name=$group + $definition.Key + 'Input'
            $input.Dock='Fill'
            $input.Margin=New-Object Windows.Forms.Padding(3,9,16,9)
            $fields.Controls.Add($label,0,$row)
            $fields.Controls.Add($input,1,$row)
            $editor.Inputs[$group+'.'+$definition.Key]=$input
            $input.Add_TextChanged({ Update-SentinelEditorState $editor }.GetNewClosure())
            $row++
        }
        $body.Controls.Add($fields,$column,0)
        $column++
    }
    $edit=New-SentinelButton -Text '编辑参数' -X 0 -Y 0 -Width 120 -Name 'UnlockProcessDefaultsButton'
    $editor.Edit=$edit
    $editor.Footer.Controls.Add($edit)
    $editor.Footer.Controls.SetChildIndex($edit,0)
    $editor.Save.Name='SaveProcessDefaultsButton'
    $edit.Add_Click({
        $editor.Editing=$true
        $edit.Enabled=$false
        foreach ($input in $editor.Inputs.Values) { $input.ReadOnly=$false }
        Update-SentinelEditorState $editor
        [void]$editor.Inputs['Water.Plate'].Focus()
    }.GetNewClosure())
    $Page.Add_Enter({ Update-SentinelProcessTitles $editor }.GetNewClosure())
    Reset-SentinelEditor $editor
}
