; デジタル板書（デスクトップ版）のインストーラー (Inno Setup 6)。tools/build-desktop.js から
;   ISCC /DAppVersion=x.y.z /DSourceDir=<dotnet build の出力> /O<出力先> Bansho.iss
; として呼ばれる。管理者権限なしで、そのユーザーの %LOCALAPPDATA%\Programs\Bansho に入る。
; 画面の設定・自動バックアップ (%LOCALAPPDATA%\Bansho\WebView2) はアンインストールでも消さない。

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef SourceDir
  #error SourceDir を指定してください
#endif

#define AppName "デジタル板書"
#define AppExe "Bansho.exe"

[Setup]
AppId={{6B1E2D7A-3C54-4F0B-9A8E-2D5C7B91E4A3}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=maezhir
VersionInfoVersion={#AppVersion}
VersionInfoProductName={#AppName}
VersionInfoDescription=デジタル板書 インストーラー
DefaultDirName={autopf}\Bansho
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; .NET Framework 4.8 が最初から入っている Windows 10 1903 以降
MinVersion=10.0.18362
OutputBaseFilename=Bansho-Setup-{#AppVersion}
SetupIconFile=..\bansho.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
ChangesAssociations=yes

[Languages]
Name: "ja"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにショートカットを作成する"; GroupDescription: "追加のショートカット:"
Name: "pptx"; Description: "PowerPoint ファイルの右クリック「プログラムから開く」に板書を加える"; GroupDescription: "追加の設定:"

[InstallDelete]
; 旧版の画面ファイルが残らないように入れ替える
Type: filesandordirs; Name: "{app}\app"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Excludes: "*.pdb"; Flags: ignoreversion recursesubdirs createallsubdirs
; WebView2 ランタイムがないときだけ取り出して実行する（tools/build-desktop.js が Microsoft から取得）
Source: "MicrosoftEdgeWebview2Setup.exe"; Flags: dontcopy

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon
; 右クリック →「送る」→「デジタル板書」で .board / .pptx を開けるようにする
Name: "{usersendto}\{#AppName}"; Filename: "{app}\{#AppExe}"

[Registry]
; .board を板書で開く（ダブルクリック・Enter）
Root: HKCU; Subkey: "Software\Classes\.board"; ValueType: string; ValueName: ""; ValueData: "Bansho.board"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Classes\.board\OpenWithProgids"; ValueType: string; ValueName: "Bansho.board"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Classes\Bansho.board"; ValueType: string; ValueName: ""; ValueData: "板書ファイル"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\Bansho.board\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExe}"",0"
Root: HKCU; Subkey: "Software\Classes\Bansho.board\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExe}"" ""%1"""
; .pptx は既定を変えず、「プログラムから開く」の候補に加えるだけ
Root: HKCU; Subkey: "Software\Classes\.pptx\OpenWithProgids"; ValueType: string; ValueName: "Bansho.pptx"; ValueData: ""; Flags: uninsdeletevalue; Tasks: pptx
Root: HKCU; Subkey: "Software\Classes\Bansho.pptx"; ValueType: string; ValueName: ""; ValueData: "PowerPoint プレゼンテーション"; Flags: uninsdeletekey; Tasks: pptx
Root: HKCU; Subkey: "Software\Classes\Bansho.pptx\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#AppExe}"" ""%1"""; Tasks: pptx
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExe}"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExe}\SupportedTypes"; ValueType: string; ValueName: ".board"; ValueData: ""
Root: HKCU; Subkey: "Software\Classes\Applications\{#AppExe}\SupportedTypes"; ValueType: string; ValueName: ".pptx"; ValueData: ""

[Run]
Filename: "{app}\{#AppExe}"; Description: "デジタル板書を起動する"; Flags: nowait postinstall skipifsilent

[Code]
// 画面の表示に使う WebView2 ランタイム（Windows 11 と、更新済みの Windows 10 には入っている）があるか
function HasWebView2: Boolean;
var
  v: String;
begin
  Result :=
    (RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', v) and (v <> '') and (v <> '0.0.0.0')) or
    (RegQueryStringValue(HKCU, 'Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', v) and (v <> '') and (v <> '0.0.0.0'));
end;

// .NET Framework 4.8 以降（Release 528040 以上）
function InitializeSetup: Boolean;
var
  r: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', r) and (r >= 528040);
  if not Result then
    SuppressibleMsgBox('このパソコンには .NET Framework 4.8 が入っていないため、インストールできません。' + #13#10 +
      'Windows Update で Windows を最新の状態にしてから、もう一度実行してください。', mbCriticalError, MB_OK, IDOK);
end;

// WebView2 ランタイムがなければ、同梱のセットアップで入れる（インターネット接続が要る）
procedure CurStepChanged(CurStep: TSetupStep);
var
  code: Integer;
begin
  if (CurStep = ssPostInstall) and not HasWebView2 then
  begin
    WizardForm.StatusLabel.Caption := '画面表示用の WebView2 ランタイムをインストールしています…';
    ExtractTemporaryFile('MicrosoftEdgeWebview2Setup.exe');
    Exec(ExpandConstant('{tmp}\MicrosoftEdgeWebview2Setup.exe'), '/silent /install', '', SW_HIDE, ewWaitUntilTerminated, code);
    if not HasWebView2 then
      SuppressibleMsgBox('画面表示用の「Microsoft Edge WebView2 ランタイム」をインストールできませんでした。' + #13#10#13#10 +
        'インターネットに接続した状態で、このインストーラーをもう一度実行してください。', mbError, MB_OK, IDOK);
  end;
end;
