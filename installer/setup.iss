[Setup]
AppName=Office Tabs Workspace
AppVersion=1.0.0
DefaultDirName={localappdata}\OfficeTabs
DefaultGroupName=Office Tabs Workspace
UninstallDisplayIcon={app}\OfficeTabs.exe
Compression=lzma2
SolidCompression=yes
OutputDir=..\publish
OutputBaseFilename=Setup
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest

[Files]
Source: "..\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Office Tabs Workspace"; Filename: "{app}\OfficeTabs.exe"
Name: "{userstartup}\Office Tabs Workspace"; Filename: "{app}\OfficeTabs.exe"

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "OfficeTabs"; ValueData: """{app}\OfficeTabs.exe"""; Flags: uninsdeletevalue

[Run]
Filename: "{app}\OfficeTabs.exe"; Description: "הפעל את Office Tabs Workspace כעת"; Flags: nowait postinstall skipifsilent
