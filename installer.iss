; Inno Setup 6 脚本：按用户安装，无需管理员权限
; 构建：安装 Inno Setup 6 后运行 build_inno.bat，或直接 iscc installer.iss

#define MyAppName "底稿未完成助手"
#define MyAppVersion "0.3.0"
#define MyAppPublisher "本地工具"
#define MyAppExeName "底稿未完成助手.exe"

[Setup]
AppId={{9C3B6E2A-1C7D-4F5A-9E2B-7A6D5C4B3A21}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=release
OutputBaseFilename={#MyAppName}-{#MyAppVersion}-安装程序
SetupIconFile=assets\app.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
AppMutex=Local\WorkpaperTodo_SingleInstance
CloseApplications=yes
RestartApplications=no
VersionInfoVersion=0.3.0.0

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务："; Flags: unchecked

[Files]
Source: "dist\{#MyAppName}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "docs\使用说明.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "CHANGELOG.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autoprograms}\卸载{#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动{#MyAppName}"; Flags: nowait postinstall skipifsilent

; 事项数据在 %LOCALAPPDATA%\WorkpaperTodo，卸载时不删除。
