; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "Lanya CorelVBA Test Version"
#define MyAppVersion "2023.7.5"
#define MyAppPublisher "lyvba.com"
#define MyAppURL "https://lyvba.com/"
#define MyAppExeName "GMS"
#define MyAppAssocName MyAppName + ""
#define MyAppAssocExt ".myp"
#define MyAppAssocKey StringChange(MyAppAssocName, " ", "") + MyAppAssocExt

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{0006790C-7107-4C59-A557-7F2EEDB64AFB}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

ChangesAssociations=yes
DisableProgramGroupPage=yes
; Uncomment the following line to run in non administrative install mode (install for current user only.)
;PrivilegesRequired=lowest
OutputDir=C:\app\CorelVBA
OutputBaseFilename=Lanya_CorelVBA
SetupIconFile=C:\app\CorelVBA\GMS\LYVBA\LOGO.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UsePreviousAppDir=no

DefaultDirName={code:GetInstallDir}

[Code]
function GetInstallDir(Param: String): String;
var
  InstallDir: String;
begin
  // 从注册表中读取安装目录
  if RegQueryStringValue(HKLM64, 'SOFTWARE\Corel\Setup\CorelDRAW Graphics Suite 2023', 'Destination', InstallDir) then
  begin
    Result := ExtractFilePath(InstallDir) + 'Draw\GMS';
  end 

  else if RegQueryStringValue(HKLM64, 'SOFTWARE\Corel\Setup\CorelDRAW Graphics Suite 2022', 'Destination', InstallDir) then
  begin
    Result := ExtractFilePath(InstallDir) + 'Draw\GMS';
  end 

  else if RegQueryStringValue(HKLM64, 'SOFTWARE\Corel\Setup\CorelDRAW Graphics Suite 2021', 'Destination', InstallDir) then
  begin
    Result := ExtractFilePath(InstallDir) + 'Draw\GMS';
  end 

  else if RegQueryStringValue(HKLM64, 'SOFTWARE\Corel\Setup\CorelDRAW Graphics Suite 2020', 'Destination', InstallDir) then
  begin
    Result := ExtractFilePath(InstallDir) + 'Draw\GMS';
  end 

  else if RegQueryStringValue(HKLM64, 'SOFTWARE\Corel\Setup\CorelDRAW Graphics Suite 2019', 'Destination', InstallDir) then
  begin
    Result := ExtractFilePath(InstallDir) + 'Draw\GMS';
  end 

  else if RegQueryStringValue(HKLM64, 'SOFTWARE\Corel\Setup\CorelDRAW Graphics Suite 2018', 'Destination', InstallDir) then
  begin
    Result := ExtractFilePath(InstallDir) + 'Draw\GMS';
  end 

  else if RegQueryStringValue(HKLM64, 'SOFTWARE\Corel\Setup\CorelDRAW Graphics Suite 16', 'Destination', InstallDir) then
  begin
    Result := ExtractFilePath(InstallDir) + 'Draw\GMS';
  end 

  else
  begin
    // 如果读取失败，则使用默认安装目录
    Result := ExpandConstant('C:\Program Files\Corel\CorelDRAW Graphics Suite 2020\Draw\GMS');
  end;
end;

[Languages]
Name: en; MessagesFile: "compiler:Default.isl"
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: nl; MessagesFile: "compiler:Languages\Dutch.isl"
Name: de; MessagesFile: "compiler:Languages\German.isl"


[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\app\CorelVBA\GMS\LYVBA.gms"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\app\CorelVBA\GMS\Adobe_Illustrator.gms"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\app\CorelVBA\GMS\ColorMark.cdr"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files
Source: "C:\app\CorelVBA\GMS\LYVBA\*"; DestDir: "{app}\LYVBA\"; Flags: ignoreversion
Source: "C:\app\CorelVBA\GMS\LYVBA\100\*"; DestDir: "{app}\LYVBA\100\"; Flags: ignoreversion
Source: "C:\app\CorelVBA\GMS\LYVBA\125\*"; DestDir: "{app}\LYVBA\125\"; Flags: ignoreversion
Source: "C:\app\CorelVBA\GMS\LYVBA\150\*"; DestDir: "{app}\LYVBA\150\"; Flags: ignoreversion
Source: "C:\app\CorelVBA\GMS\LYVBA\175\*"; DestDir: "{app}\LYVBA\175\"; Flags: ignoreversion
Source: "C:\app\CorelVBA\GMS\LYVBA\200\*"; DestDir: "{app}\LYVBA\200\"; Flags: ignoreversion

Source: "C:\app\CorelVBA\TSP\*"; DestDir: "C:\TSP\"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
;Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
;Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: shellexec postinstall skipifsilent

