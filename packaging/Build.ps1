param([string]$Source='', [Parameter(Mandatory=$true)][string]$Compiler)
$ErrorActionPreference='Stop'
if (-not $Source) {
    $Source=Join-Path $PSScriptRoot '../src'
    if (-not (Test-Path $Source)) {$Source=Join-Path $PSScriptRoot '../outputs/ExperimentDataAutomation'}
}
$root=Split-Path $PSScriptRoot
$stage=Join-Path $root 'stage'
if (Test-Path $stage) {throw 'Stage already exists. Use a fresh build directory to avoid stale files.'}
[void][IO.Directory]::CreateDirectory($stage)
Get-ChildItem -LiteralPath $Source -File -Filter '*.ps1' | Copy-Item -Destination $stage
Copy-Item (Join-Path $PSScriptRoot 'README.md') (Join-Path $stage 'README.md')
$csc=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
& $csc /nologo /target:winexe /reference:System.Windows.Forms.dll ('/out:'+(Join-Path $stage 'PolishingSentinel.exe')) (Join-Path $PSScriptRoot 'Launcher.cs')
if ($LASTEXITCODE) {throw 'Launcher compilation failed'}
& $Compiler (Join-Path $PSScriptRoot 'setup.iss')
if ($LASTEXITCODE) {throw 'Installer compilation failed'}
Get-ChildItem (Join-Path $root 'dist') -Filter '*.exe' | Get-FileHash -Algorithm SHA256
