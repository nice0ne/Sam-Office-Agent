; Script Inno Setup untuk Sam Office Agent
; Dibuat untuk kompilasi installer standalone non-elevated (User Mode)

[Setup]
AppId={{d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23}}
AppName=Sam Office Agent
AppVersion=1.0.0
AppVerName=Sam Office Agent 1.0.0
AppPublisher=Sam Office Agent Team
DefaultDirName={localappdata}\SamOfficeAgent
DefaultGroupName=Sam Office Agent
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputBaseFilename=SamOfficeAgent-Setup
OutputDir=..\release
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
CloseApplications=force
CloseApplicationsFilter=SamTrayServer.exe
UninstallDisplayIcon={app}\SamTrayServer.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "autostart"; Description: "Jalankan Sam Office Agent secara otomatis saat Windows mulai"; GroupDescription: "Pengaturan Tambahan:"

[Files]
Source: "..\dist-release\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Sam Office Agent"; Filename: "{app}\SamTrayServer.exe"
Name: "{group}\{cm:UninstallProgram,Sam Office Agent}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Sam Office Agent"; Filename: "{app}\SamTrayServer.exe"; Tasks: desktopicon
Name: "{userstartup}\Sam Office Agent.lnk"; Filename: "{app}\SamTrayServer.exe"; Tasks: autostart

[Registry]
; Registrasi manifes Office Add-in ke WEF Developer Registry
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; ValueType: string; ValueName: "d2c3e1b7-4a5f-4d9a-9e12-8b7a4c9f1a23"; ValueData: "{app}\manifest.xml"; Flags: uninsdeletevalue
; Bersihkan registry Run jika pengguna mengaktifkannya lewat menu aplikasi
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: none; ValueName: "SamOfficeAgent"; Flags: dontcreatekey uninsdeletevalue

[Run]
Filename: "{app}\SamTrayServer.exe"; Description: "Jalankan Sam Office Agent sekarang"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "cmd.exe"; Parameters: "/c taskkill /F /IM SamTrayServer.exe >nul 2>&1 || exit 0"; Flags: runhidden

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
