$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$launcher = Join-Path $scriptRoot 'Launch-PolishingSentinel.ps1'
if (-not (Test-Path -LiteralPath $launcher)) {
    throw "未找到抛光哨站启动脚本：$launcher"
}

$desktop = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrWhiteSpace($desktop)) {
    throw '未能定位当前用户桌面目录。'
}

$powershellCandidates = @(
    (Join-Path $PSHOME 'pwsh.exe'),
    'C:\Program Files\PowerShell\7\pwsh.exe',
    'C:\Program Files\PowerShell\7-preview\pwsh.exe',
    (Join-Path $PSHOME 'powershell.exe'),
    'powershell.exe'
)
$powershellExe = $powershellCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($powershellExe)) {
    $powershellExe = 'powershell.exe'
}

$shortcutPath = Join-Path $desktop '抛光哨站.lnk'
$shell = $null
$shortcut = $null

try {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $powershellExe
    $shortcut.Arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $launcher
    $shortcut.WorkingDirectory = $scriptRoot
    $shortcut.WindowStyle = 7
    $shortcut.Description = '抛光哨站实验数据自动化控制台'
    $shortcut.IconLocation = '{0},0' -f $powershellExe
    $shortcut.Save()
}
finally {
    if ($shortcut) {
        [void] [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shortcut)
    }
    if ($shell) {
        [void] [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
    }
}

Write-Host "已创建桌面快捷方式：$shortcutPath"
