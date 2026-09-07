#define MyAppName "Turnitin Rubric Converter"
#define MyAppPublisher "amdgeo"
#define MyAppURL "https://github.com/amdgeo/turnitin_rubric_converter"
#ifndef MyAppVersion
  #define MyAppVersion "development"
#endif

[Setup]
AppId={{E2CA29DE-AF42-49E3-AB50-7F152FC993F2}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE
OutputDir=..\dist
OutputBaseFilename=Turnitin-Rubric-Converter-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UninstallDisplayName={#MyAppName}

[Files]
Source: "..\dist\rubric-converter.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\rubric-converter.exe"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\rubric-converter.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Run]
Filename: "{app}\rubric-converter.exe"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
