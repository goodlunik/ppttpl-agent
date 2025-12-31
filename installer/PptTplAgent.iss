; PptTplAgent.iss - Inno Setup script (per-user install, no admin required)
#define MyAppName "PptTplAgent"
#define MyAppPublisher "LUNIK"
#define MyAppURL "https://phpstack-1189557-6107953.cloudwaysapps.com/"
#define MyAppExeName "PptTplAgent.exe"

; Build defines:
; -DMyAppVersion=1.0.0
; -DMyPublishDir=..\src\PptTplAgent\bin\Release\net8.0-windows\win-x64\publish
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif
#ifndef MyPublishDir
  #define MyPublishDir "..\src\PptTplAgent\bin\Release\net8.0-windows\win-x64\publish"
#endif

[Setup]
AppId={{6B5AF29A-6D94-4C5C-8F22-0B1C4D05F5E2}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\{#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=PptTplAgent_Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"

[Files]
Source: "{#MyPublishDir}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Register ppttpl:// protocol (per-user, no admin)
Root: HKCU; Subkey: "Software\Classes\ppttpl"; ValueType: string; ValueName: ""; ValueData: "URL:PptTpl Protocol"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\ppttpl"; ValueType: string; ValueName: "URL Protocol"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Classes\ppttpl\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""; Flags: uninsdeletekey

[Icons]
Name: "{userprograms}\{#MyAppName}\{#MyAppName} (log)"; Filename: "{cmd}"; Parameters: "/c notepad ""{userprofile}\Desktop\PptTplAgent.log"""; WorkingDir: "{userprofile}"; IconFilename: "{app}\{#MyAppExeName}"
