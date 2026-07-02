; Inno Setup script for Dynamics Desk
; Builds a per-user installer (no admin needed) so the in-place auto-updater keeps working.
; Version is passed in by the build script: ISCC /DAppVersion=1.1.1 installer.iss

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#define AppName "Dynamics Desk"
#define AppExe "Dynamics Desk.exe"
#define SourceDir "dist-packaged\Dynamics Desk-win32-x64"

[Setup]
AppId={{8F3B2A41-6C9D-4E27-9B0A-DD1A2C7F5E10}}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=Dynamics Desk
DefaultDirName={localappdata}\Programs\Dynamics Desk
DisableProgramGroupPage=yes
DefaultGroupName={#AppName}
PrivilegesRequired=lowest
OutputDir=dist-installer
OutputBaseFilename=Dynamics-Desk-Setup-v{#AppVersion}
SetupIconFile=assets\icon.ico
UninstallDisplayIcon={app}\{#AppExe}
WizardStyle=modern
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
RestartApplications=no

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent
