; 进度跟踪报表工具 - Inno Setup 安装包脚本
; 前置：先由 build.bat 生成 dist\进度跟踪报表工具.exe
; 用法：安装 Inno Setup（https://jrsoftware.org/isdl.php）后执行：
;         iscc installer.iss
; 产物：installer\进度跟踪报表工具_setup.exe
;
; 注意：GitHub Actions 的 windows-latest runner 上 choco 装的 Inno Setup
; 默认不含 Chinese.isl，故此处不声明 [Languages]，使用默认语言。

[Setup]
AppName=进度跟踪报表工具
AppVersion=1.3
DefaultDirName={localappdata}\进度跟踪报表工具
DefaultGroupName=进度跟踪报表工具
OutputDir=installer
OutputBaseFilename=进度跟踪报表工具_setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\进度跟踪报表工具.exe

[Files]
Source: "dist\进度跟踪报表工具.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\进度跟踪报表工具"; Filename: "{app}\进度跟踪报表工具.exe"
Name: "{commondesktop}\进度跟踪报表工具"; Filename: "{app}\进度跟踪报表工具.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务"; Flags: unchecked

[Run]
Filename: "{app}\进度跟踪报表工具.exe"; Description: "启动 进度跟踪报表工具"; Flags: nowait postinstall skipifsilent
