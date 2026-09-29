$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptRoot 'PolishingSentinel.UI.ps1')

Show-PolishingSentinelWindow -ScriptRoot $scriptRoot
