; ── PPT Touch Controller 安装脚本 ──────────────────────────────
; Inno Setup 6 脚本
; 运行: "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup.iss

#define MyAppName "PPT Touch Controller"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "PPT Touch Controller"
#define MyAppURL "https://github.com/gengyangze-hub/ppt-touch-controller"
#define MyAppExeName "PPTTouchController.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer
OutputBaseFilename=PPT-Touch-Controller-Setup-v1.0.0
; SetupIconFile=src\resources\icons\app.ico  (未创建图标，暂时禁用)
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
; 安装器描述
VersionInfoVersion={#MyAppVersion}
VersionInfoDescription=PPT Touch Controller - 触控屏 PPT 翻页工具

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create &desktop shortcut"; GroupDescription: "Additional icons:"
Name: "fileassoc"; Description: "Associate .pptx files (double-click to open with touch controller)"; GroupDescription: "File association:"

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; 注意：如需自定义图标，将 app.ico 放到 src/resources/icons/ 目录
; 然后取消下面注释

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\卸载 PPT Touch Controller"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
; 文件关联: .pptx → PPTTouchController
Root: HKA; Subkey: "Software\Classes\.pptx\OpenWithProgids"; ValueType: string; ValueName: "PPTTouchController.pptx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: fileassoc
Root: HKA; Subkey: "Software\Classes\PPTTouchController.pptx"; ValueType: string; ValueData: "Microsoft PowerPoint 演示文稿 (触控控制器)"; Flags: uninsdeletekey; Tasks: fileassoc
Root: HKA; Subkey: "Software\Classes\PPTTouchController.pptx\DefaultIcon"; ValueType: string; ValueData: "{app}\{#MyAppExeName},0"; Tasks: fileassoc
Root: HKA; Subkey: "Software\Classes\PPTTouchController.pptx\shell\open\command"; ValueType: string; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Tasks: fileassoc

; 添加到「打开方式」菜单
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".pptx"; ValueData: ""; Flags: uninsdeletevalue
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\shell\open\command"; ValueType: string; ValueData: """{app}\{#MyAppExeName}"" ""%1"""

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即启动 PPT Touch Controller"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\*.log"
Type: dirifempty; Name: "{app}"
