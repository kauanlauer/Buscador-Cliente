#define MyAppName "Buscador Cliente HeadCargo"
#define MyAppVersion "1.2.0"
#define MyAppPublisher "HeadCargo"
#define MyAppExeName "Buscador Cliente HeadCargo.exe"
#define MyAppSourceExe "dist\Buscador Cliente HeadCargo.exe"
#define MyAppIconFile "logo_buscador.ico"
#define DefaultClientRoot "D:\OneDrive - headsoft.com.br\HeadSoft Home - Suporte\Pastinha Clientes\Acessos Clientes"

[Setup]
AppId={{5C519B03-7AC6-4C8D-88E2-4DD7739B9A31}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=Setup Buscador Cliente HeadCargo
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupIconFile={#MyAppIconFile}
UsePreviousTasks=yes

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar atalho na Area de Trabalho"; Flags: unchecked
Name: "autostart"; Description: "Iniciar com o Windows"; Flags: checkedonce

[Files]
Source: "{#MyAppSourceExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userstartup}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Parameters: "--minimized"; Tasks: autostart

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Executar {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
var
  ClientRootPage: TInputDirWizardPage;

function FindFrom(SubText, Text: string; StartPos: Integer): Integer;
var
  I: Integer;
begin
  Result := 0;
  for I := StartPos to Length(Text) - Length(SubText) + 1 do
  begin
    if Copy(Text, I, Length(SubText)) = SubText then
    begin
      Result := I;
      Exit;
    end;
  end;
end;

function GetConfigPath(): string;
begin
  Result := ExpandConstant('{userappdata}\BuscadorClienteHeadCargo\launcher_clientes_onedrive_config.json');
end;

function JsonEscape(Value: string): string;
begin
  Result := Value;
  StringChangeEx(Result, '\', '\\', True);
  StringChangeEx(Result, '"', '\"', True);
end;

procedure WriteAppConfig();
var
  ConfigDir, ConfigPath, ConfigText: string;
begin
  ConfigDir := ExpandConstant('{userappdata}\BuscadorClienteHeadCargo');
  ConfigPath := GetConfigPath();

  ForceDirectories(ConfigDir);

  if FileExists(ConfigPath) then
    exit;

  ConfigText :=
    '{'#13#10 +
    '  "root_path": "' + JsonEscape(ClientRootPage.Values[0]) + '",'#13#10 +
    '  "preferred_username": "",'#13#10 +
    '  "mark_save_password": false,'#13#10 +
    '  "encrypted_password": "",'#13#10 +
    '  "last_update_check": ""'#13#10 +
    '}';

  SaveStringToFile(ConfigPath, ConfigText, False);
end;

procedure InitializeWizard();
var
  DefaultPath: string;
begin
  DefaultPath := '{#DefaultClientRoot}';
  ClientRootPage := CreateInputDirPage(
    wpSelectDir,
    'Pasta dos clientes',
    'Selecione a Pastinha Clientes\Acessos Clientes',
    'Informe a pasta que contem os clientes para o buscador funcionar nessa maquina.',
    False,
    ''
  );
  ClientRootPage.Add('');
  ClientRootPage.Values[0] := DefaultPath;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if CurPageID = ClientRootPage.ID then
  begin
    if (Trim(ClientRootPage.Values[0]) = '') or (not DirExists(ClientRootPage.Values[0])) then
    begin
      MsgBox('Selecione uma pasta valida para os clientes.', mbError, MB_OK);
      Result := False;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    WriteAppConfig();
end;
