#ifndef AppVersion
  #define AppVersion "0.9.0"
#endif
[Setup]
AppId={{9AC6F0AA-556D-4F5B-B5FA-CF189B6C3843}
AppName=抛光哨站
AppVersion={#AppVersion}
DefaultDirName={localappdata}\Programs\PolishingSentinel
DefaultGroupName=抛光哨站
PrivilegesRequired=lowest
MinVersion=10.0
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist
OutputBaseFilename=PolishingSentinel-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\PolishingSentinel.exe
CloseApplications=no
SetupLogging=yes
[Files]
Source: "..\stage\*.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\stage\PolishingSentinel.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\stage\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "Manage-Installation.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "Manage-Installation.ps1"; Flags: dontcopy
[Tasks]
Name: desktopicon; Description: "创建桌面快捷方式"; Flags: unchecked
[Icons]
Name: "{group}\抛光哨站"; Filename: "{app}\PolishingSentinel.exe"
Name: "{autodesktop}\抛光哨站"; Filename: "{app}\PolishingSentinel.exe"; Tasks: desktopicon
[Code]
function Manage(const Script, Mode: String): Boolean;
var Code: Integer;
begin
  Result := Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
    '-NoProfile -ExecutionPolicy Bypass -File "' + Script + '" -Mode ' + Mode + ' -AppRoot "' + ExpandConstant('{app}') + '"',
    '', SW_HIDE, ewWaitUntilTerminated, Code) and (Code = 0);
end;
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  ExtractTemporaryFile('Manage-Installation.ps1');
  if not Manage(ExpandConstant('{tmp}\Manage-Installation.ps1'), 'Check') then
    Result := '请先等待任务结束，停止后台监控并关闭抛光哨站，再继续安装。';
end;
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    if not Manage(ExpandConstant('{app}\Manage-Installation.ps1'), 'Install') then
      RaiseException('用户数据目录初始化失败，请查看安装日志。');
end;
function InitializeUninstall(): Boolean;
begin
  Result := Manage(ExpandConstant('{app}\Manage-Installation.ps1'), 'Check');
  if not Result then MsgBox('请先等待任务结束，停止后台监控并关闭抛光哨站，再卸载。', mbError, MB_OK);
end;
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
    if not Manage(ExpandConstant('{app}\Manage-Installation.ps1'), 'Remove') then
      RaiseException('卸载准备失败，用户数据未删除。');
end;
