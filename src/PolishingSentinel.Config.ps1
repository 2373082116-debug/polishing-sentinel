Set-StrictMode -Version 2.0

function Get-PolishingSentinelSettingsPath {
    param([Parameter(Mandatory = $true)] [string] $ScriptRoot)

    return (Join-Path (Join-Path $ScriptRoot 'config') 'settings.json')
}

function New-PolishingSentinelDefaultSettings {
    param([Parameter(Mandatory = $true)] [string] $ScriptRoot)

    [pscustomobject]@{
        AppName                 = '抛光哨站'
        VolumeLabel             = '虎大子李'
        ExpectedDriveLetter     = 'F'
        WaterSourceRelativePath = '实验数据\抛光实验\水基抛光（光镜+白光）'
        OilSourceRelativePath   = '实验数据\抛光实验\油基抛光'
        MasterWorkbookPath      = 'D:\共享文件\实验数据\实验数据.xlsx'
        MatlabScriptPath        = 'D:\共享文件\matlab\program\baiguang.m'
        MatlabExePath           = 'D:\Program Files\MATLAB\R2025b\bin\matlab.exe'
        SaEngine                = 'BuiltIn'
        AutoSheetName           = '自动导入Sa数据'
        PendingSheetName        = '待确认补齐'
        SaOutputFileName        = '所有的Sa数据.xlsx'
        PollSeconds             = 5
        RescanSeconds           = 60
        MatlabTimeoutSeconds    = 1200
        ExcelTimeoutSeconds     = 600
        MaxExperimentAttempts   = 3
        RetryDelaySeconds       = 60
        EnableBackfill          = $true
        NewBlockMode            = 'Preview'
        EnableStartupTask       = $true
        TaskName                = 'PolishingSentinel'
        LogDirectory            = Join-Path $ScriptRoot 'logs'
        StateDirectory          = Join-Path $ScriptRoot 'state'
        TempDirectory           = Join-Path $ScriptRoot 'temp'
        BackupDirectory         = Join-Path $ScriptRoot 'backups'
        SlurryMappings          = [pscustomobject]@{
            '金刚石' = '油基抛光'
            '硅溶胶' = '探索实验'
        }
        ProcessDefaults         = [pscustomobject]@{
            Water = [pscustomobject]@{
                Plate    = 'PEEK'
                Track    = '4'
                Spacing  = '9'
                Fullness = '92%'
                Pressure = '单球0.1N'
                BallMatch = '0.8mm氧化锆球'
            }
            Oil = [pscustomobject]@{
                Plate    = '铸铁抛光盘'
                Track    = '3'
                Spacing  = '9'
                Fullness = '92%'
                Pressure = '单球0.1N'
                BallMatch = '0.8mm钢球'
            }
        }
        FolderNamingConvention  = '实验号-日期-浓度-粒径-抛光液类型-球胚-抛光时间-抛光垫类型-备注'
    }
}

function Merge-PolishingSentinelSettings {
    param(
        [Parameter(Mandatory = $true)] [pscustomobject] $Defaults,
        [Parameter(Mandatory = $true)] [pscustomobject] $Loaded
    )

    foreach ($property in $Loaded.PSObject.Properties) {
        if ($property.Name -eq 'ProcessDefaults') {
            if ($property.Value -isnot [pscustomobject]) { throw '固定工艺参数配置必须是对象，不能为 null 或文本。' }
            foreach ($group in $property.Value.PSObject.Properties) {
                if ($group.Name -in @('Water', 'Oil')) {
                    if ($group.Value -isnot [pscustomobject]) { throw ('工艺参数组 {0} 必须是对象。' -f $group.Name) }
                    foreach ($field in $group.Value.PSObject.Properties) {
                        $Defaults.ProcessDefaults.($group.Name) | Add-Member -NotePropertyName $field.Name -NotePropertyValue $field.Value -Force
                    }
                }
                else {
                    $Defaults.ProcessDefaults | Add-Member -NotePropertyName $group.Name -NotePropertyValue $group.Value -Force
                }
            }
            continue
        }
        $Defaults | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value -Force
    }
    return $Defaults
}

function Read-PolishingSentinelSettings {
    param([Parameter(Mandatory = $true)] [string] $ScriptRoot)

    $settingsPath = Get-PolishingSentinelSettingsPath -ScriptRoot $ScriptRoot
    $defaults = New-PolishingSentinelDefaultSettings -ScriptRoot $ScriptRoot

    if (-not (Test-Path -LiteralPath $settingsPath)) {
        [void] (Save-PolishingSentinelSettings -ScriptRoot $ScriptRoot -Settings $defaults)
        return $defaults
    }

    $json = Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($json)) {
        throw ('配置文件为空，已停止读取以避免使用错误的默认设置：{0}' -f $settingsPath)
    }

    $loaded = ConvertFrom-Json -InputObject $json -ErrorAction Stop
    if (-not $json.TrimStart().StartsWith('{') -or $loaded -isnot [pscustomobject]) {
        throw ('配置文件根节点必须是 JSON 对象：{0}' -f $settingsPath)
    }
    return (Merge-PolishingSentinelSettings -Defaults $defaults -Loaded $loaded)
}

function Save-PolishingSentinelSettings {
    param(
        [Parameter(Mandatory = $true)] [string] $ScriptRoot,
        [Parameter(Mandatory = $true)] [pscustomobject] $Settings
    )

    $settingsPath = Get-PolishingSentinelSettingsPath -ScriptRoot $ScriptRoot
    $settingsDirectory = Split-Path -Parent $settingsPath
    if (-not (Test-Path -LiteralPath $settingsDirectory)) {
        New-Item -ItemType Directory -Path $settingsDirectory -Force -ErrorAction Stop | Out-Null
    }

    $json = $Settings | ConvertTo-Json -Depth 8
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    $temporaryPath = Join-Path $settingsDirectory ('.settings-' + [guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [System.IO.File]::WriteAllText($temporaryPath, $json, $utf8Bom)
        # Publish a complete file; concurrent readers must never see partial JSON.
        if ([System.IO.File]::Exists($settingsPath)) {
            [System.IO.File]::Replace($temporaryPath, $settingsPath, [System.Management.Automation.Language.NullString]::Value)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $settingsPath)
        }
    }
    finally {
        if ([System.IO.File]::Exists($temporaryPath)) {
            try { [System.IO.File]::Delete($temporaryPath) } catch { }
        }
    }
    return $settingsPath
}

function ConvertTo-ExperimentAutomationConfigFromSettings {
    param(
        [Parameter(Mandatory = $true)] [string] $ScriptRoot,
        [Parameter(Mandatory = $true)] [pscustomobject] $Settings,
        [switch] $DryRun,
        [Nullable[int]] $PollSecondsOverride = $null
    )

    $pollSeconds = [int] $Settings.PollSeconds
    if ($null -ne $PollSecondsOverride) {
        $pollSeconds = [int] $PollSecondsOverride
    }

    $config = New-ExperimentAutomationConfig `
        -ScriptRoot $ScriptRoot `
        -VolumeLabel $Settings.VolumeLabel `
        -ExpectedDriveLetter $Settings.ExpectedDriveLetter `
        -MasterWorkbookPath $Settings.MasterWorkbookPath `
        -MatlabScriptPath $Settings.MatlabScriptPath `
        -MatlabExePath $Settings.MatlabExePath `
        -SourceRelativeFolders @($Settings.WaterSourceRelativePath, $Settings.OilSourceRelativePath) `
        -PollSeconds $pollSeconds `
        -DryRun:$DryRun

    $config.AutoSheetName = $Settings.AutoSheetName
    $config.SaOutputFileName = $Settings.SaOutputFileName
    $config.PendingSheetName = $Settings.PendingSheetName
    $config.EnableBackfill = [bool] $Settings.EnableBackfill
    $config.LogDirectory = $Settings.LogDirectory
    $config.StateDirectory = $Settings.StateDirectory
    $config.TempDirectory = $Settings.TempDirectory
    $config.BackupDirectory = $Settings.BackupDirectory
    $config.SlurryMappings = $Settings.SlurryMappings
    $config.ProcessDefaults = $Settings.ProcessDefaults
    if ($Settings.PSObject.Properties['SaEngine']) {
        if ($Settings.SaEngine -notin @('BuiltIn', 'Matlab')) { throw 'Sa 处理方式必须为 BuiltIn 或 Matlab。' }
        $config.SaEngine = $Settings.SaEngine
    }
    foreach ($spec in @(@('MatlabTimeoutSeconds', 1, 86400), @('ExcelTimeoutSeconds', 1, 86400), @('MaxExperimentAttempts', 1, 10), @('RetryDelaySeconds', 1, 86400))) {
        $name = $spec[0]
        if ($Settings.PSObject.Properties[$name]) {
            $value = [int]$Settings.$name
            if ($value -lt $spec[1] -or $value -gt $spec[2]) { throw "设置 $name 必须在 $($spec[1]) 至 $($spec[2]) 之间。" }
            $config.$name = $value
        }
    }
    $newBlockMode = [string] $Settings.NewBlockMode
    if ([string]::IsNullOrWhiteSpace($newBlockMode)) {
        $newBlockMode = 'Preview'
    }
    $config | Add-Member -NotePropertyName NewBlockMode -NotePropertyValue $newBlockMode -Force
    return $config
}
