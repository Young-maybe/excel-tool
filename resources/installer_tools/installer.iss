; Inno Setup 安装脚本（生成带向导的 Setup.exe）
; 功能：
; - 可选择安装路径
; - 安装后创建桌面快捷方式
; - 可选开始菜单快捷方式
; - 卸载入口

#define MyAppName "Excel处理工具"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "yangxinpeng"
#define MyAppExeName "ExcelTools.exe"
; Inno Setup 的 AppId 推荐使用 GUID（用于升级/卸载识别）
#define MyAppIdGuid "B5E6E2A9-6B3A-4E7C-9A7A-8D8E7B1C1B25"
; 输出文件名尽量用英文，避免部分系统/压缩软件出现乱码
#define MyOutputBase "Install_ExcelTools"

[Setup]
; 这里需要字面量大括号包住 GUID，所以用 {{ 和 }} 转义
AppId={{{#MyAppIdGuid}}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
OutputDir=..\..\dist_installer
OutputBaseFilename={#MyOutputBase}
Compression=lzma
SolidCompression=yes
WizardStyle=modern

; 图标（如果存在）
SetupIconFile=..\..\resources\app_icon.ico

; 语言：部分 Inno Setup 安装不包含 ChineseSimplified.isl（会导致编译失败）。
; 这里固定使用默认语言文件，保证“必定能编译出安装包”。
[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "额外任务"; Flags: unchecked

[Files]
; 注意：这里默认打包 PyInstaller 的 onedir 输出（dist\Excel处理工具\*）
Source: "..\..\dist\ExcelTools\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "运行 {#MyAppName}"; Flags: nowait postinstall skipifsilent


