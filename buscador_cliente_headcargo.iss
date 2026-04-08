#define MyAppName "Buscador Cliente HeadCargo"
#define MyAppVersion "1.2.3"
#define MyAppPublisher "HeadCargo"
#define MyAppExeName "Buscador Cliente HeadCargo.exe"
#define MyAppSourceExe "dist\Buscador Cliente HeadCargo.exe"
#define MyAppIconFile "logo_buscador.ico"

[Setup]
AppId={{5C519B03-7AC6-4C8D-88E2-4DD7739B9A31}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/kauanlauer/Buscador-Cliente
AppSupportURL=https://github.com/kauanlauer/Buscador-Cliente
AppUpdatesURL=https://github.com/kauanlauer/Buscador-Cliente/releases
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
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=Instalador do Buscador Cliente HeadCargo
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
VersionInfoVersion={#MyAppVersion}.0

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
  PreferredUsernamePage: TInputQueryWizardPage;

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

function JsonUnescape(Value: string): string;
var
  Placeholder: string;
begin
  Placeholder := #1;
  Result := Value;
  StringChangeEx(Result, '\\', Placeholder, True);
  StringChangeEx(Result, '\"', '"', True);
  StringChangeEx(Result, Placeholder, '\', True);
end;

function JsonEscape(Value: string): string;
begin
  Result := Value;
  StringChangeEx(Result, '\', '\\', True);
  StringChangeEx(Result, '"', '\"', True);
end;

function ReadConfigStringValue(KeyName: string): string;
var
  ConfigPath, ConfigText: string;
  ConfigBytes: AnsiString;
  KeyPos, ColonPos, StartPos, EndPos: Integer;
begin
  Result := '';
  ConfigPath := GetConfigPath();
  if not FileExists(ConfigPath) then
    exit;
  if not LoadStringFromFile(ConfigPath, ConfigBytes) then
    exit;
  ConfigText := UTF8Decode(ConfigBytes);

  KeyPos := FindFrom('"' + KeyName + '"', ConfigText, 1);
  if KeyPos = 0 then
    exit;

  ColonPos := FindFrom(':', ConfigText, KeyPos + Length(KeyName) + 2);
  if ColonPos = 0 then
    exit;

  StartPos := ColonPos + 1;
  while (StartPos <= Length(ConfigText)) and ((ConfigText[StartPos] = ' ') or (ConfigText[StartPos] = #9) or (ConfigText[StartPos] = #10) or (ConfigText[StartPos] = #13)) do
    StartPos := StartPos + 1;

  if (StartPos > Length(ConfigText)) or (ConfigText[StartPos] <> '"') then
    exit;

  StartPos := StartPos + 1;
  EndPos := StartPos;
  while EndPos <= Length(ConfigText) do
  begin
    if (ConfigText[EndPos] = '"') and ((EndPos = StartPos) or (ConfigText[EndPos - 1] <> '\')) then
      break;
    EndPos := EndPos + 1;
  end;

  if EndPos > Length(ConfigText) then
    exit;

  Result := JsonUnescape(Copy(ConfigText, StartPos, EndPos - StartPos));
end;

function UpdateConfigStringValue(ConfigText, KeyName, NewValue: string): string;
var
  KeyPos, ColonPos, StartPos, EndPos: Integer;
begin
  Result := ConfigText;
  KeyPos := FindFrom('"' + KeyName + '"', ConfigText, 1);
  if KeyPos = 0 then
    exit;

  ColonPos := FindFrom(':', ConfigText, KeyPos + Length(KeyName) + 2);
  if ColonPos = 0 then
    exit;

  StartPos := ColonPos + 1;
  while (StartPos <= Length(ConfigText)) and ((ConfigText[StartPos] = ' ') or (ConfigText[StartPos] = #9) or (ConfigText[StartPos] = #10) or (ConfigText[StartPos] = #13)) do
    StartPos := StartPos + 1;

  if (StartPos > Length(ConfigText)) or (ConfigText[StartPos] <> '"') then
    exit;

  StartPos := StartPos + 1;
  EndPos := StartPos;
  while EndPos <= Length(ConfigText) do
  begin
    if (ConfigText[EndPos] = '"') and ((EndPos = StartPos) or (ConfigText[EndPos - 1] <> '\')) then
      break;
    EndPos := EndPos + 1;
  end;

  if EndPos > Length(ConfigText) then
    exit;

  Result :=
    Copy(ConfigText, 1, StartPos - 1) +
    JsonEscape(NewValue) +
    Copy(ConfigText, EndPos, Length(ConfigText) - EndPos + 1);
end;

procedure WriteAppConfig();
var
  ConfigDir, ConfigPath, ConfigText: string;
  ConfigBytes: AnsiString;
begin
  ConfigDir := ExpandConstant('{userappdata}\BuscadorClienteHeadCargo');
  ConfigPath := GetConfigPath();

  ForceDirectories(ConfigDir);

  if FileExists(ConfigPath) then
  begin
    if LoadStringFromFile(ConfigPath, ConfigBytes) then
    begin
      ConfigText := UTF8Decode(ConfigBytes);
      ConfigText := UpdateConfigStringValue(ConfigText, 'root_path', ClientRootPage.Values[0]);
      ConfigText := UpdateConfigStringValue(ConfigText, 'preferred_username', Trim(PreferredUsernamePage.Values[0]));
      SaveStringToFile(ConfigPath, UTF8Encode(ConfigText), False);
      exit;
    end;
  end;

  ConfigText :=
    '{'#13#10 +
    '  "root_path": "' + JsonEscape(ClientRootPage.Values[0]) + '",'#13#10 +
    '  "preferred_username": "' + JsonEscape(Trim(PreferredUsernamePage.Values[0])) + '",'#13#10 +
    '  "mark_save_password": false,'#13#10 +
    '  "encrypted_password": "",'#13#10 +
    '  "last_update_check": "",'#13#10 +
    '  "recent_clients": []'#13#10 +
    '}';

  SaveStringToFile(ConfigPath, UTF8Encode(ConfigText), False);
end;

procedure InitializeWizard();
var
  StoredRootPath, StoredUsername: string;
begin
  StoredRootPath := ReadConfigStringValue('root_path');
  StoredUsername := ReadConfigStringValue('preferred_username');
  ClientRootPage := CreateInputDirPage(
    wpSelectDir,
    'Pasta dos clientes',
    'Selecione a pasta da Pastinha Clientes no OneDrive',
    'Informe a pasta onde fica a Pastinha Clientes/Acessos Clientes desta maquina.',
    False,
    ''
  );
  ClientRootPage.Add('Pasta dos clientes:');
  ClientRootPage.Values[0] := StoredRootPath;

  PreferredUsernamePage := CreateInputQueryPage(
    ClientRootPage.ID,
    'Usuario padrao',
    'Informe o usuario padrao do sistema',
    'Esse usuario sera gravado nas configuracoes iniciais do buscador e usado no preenchimento do login do HeadCargo.'
  );
  PreferredUsernamePage.Add('Usuario padrao:', False);
  PreferredUsernamePage.Values[0] := StoredUsername;
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
  if CurPageID = PreferredUsernamePage.ID then
  begin
    if Trim(PreferredUsernamePage.Values[0]) = '' then
    begin
      MsgBox('Informe o usuario padrao do sistema.', mbError, MB_OK);
      Result := False;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    WriteAppConfig();
end;
