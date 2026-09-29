param(
    [string] $TaskName = 'PolishingSentinel'
)

$ErrorActionPreference = 'Stop'

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "已卸载任务计划：$TaskName"
}
else {
    Write-Host "未找到任务计划：$TaskName"
}
