; ===================================================================
; 三幸町自治会 プログラム配布リスト生成ツール インストーラースクリプト
; Inno Setup 6.x 用
; ===================================================================
#define MyAppName "三幸町自治会 プログラム配布リスト生成ツール"
#define MyAppVersion "1.7.0"
#define MyAppPublisher "久野耕司"
#define MyAppExeName "jichikai_program_distribution_tool.exe"

[Setup]
; AppId は一度生成したら変更しないでください（再インストール・更新の識別に使用）
AppId={{20968F67-9286-47C1-B2B4-81320B003E2C}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
; ユーザーが自由に書き込める場所（ユーザーごとの AppData\Local）にインストールする
DefaultDirName={localappdata}\JichikaiProgramDistributionTool
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; 出力されるインストーラーEXEのファイル名（拡張子なし）
OutputBaseFilename=jichikai_program_distribution_tool_setup
; アイコン
SetupIconFile=app_icon_final.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; ユーザーごとのフォルダにインストールするため管理者権限は不要
PrivilegesRequired=lowest
; 64bit専用
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにアイコンを作成する"; GroupDescription: "追加のアイコン:"; Flags: unchecked

[Files]
; 本体EXE
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; アイコン（ウィンドウアイコン表示用）
Source: "app_icon_final.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent