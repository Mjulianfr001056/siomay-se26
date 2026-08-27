; SIOMAY Windows installer. Build after `flet build windows` succeeds.
; Requires Inno Setup 6 and build/windows/x64/runner/Release from Flet/Flutter.
; Keep AppId unchanged for every future SIOMAY upgrade.

#define MyAppName "SIOMAY"
#define MyAppVersion "2026.1.0.1"
#define MyAppPublisher "6304 - Muhammad Julian Firdaus, S.Tr.Stat."
#define MyAppURL "https://github.com/Mjulianfr001056/siomay-se26"
#define MyAppExeName "siomay.exe"
#define MyBuildDir "..\build\windows\x64\runner\Release"

[Setup]
AppId={{2F09AD98-2AC7-4CD4-BD85-856F232A1D64}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={autopf}\SIOMAY
DefaultGroupName=SIOMAY
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist
OutputBaseFilename=SIOMAY-Setup-2026.1.0.1
SetupIconFile=..\assets\icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=SIOMAY - Sistem Otomasi Massal dan Terpercaya
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}

; Uncomment and configure after a Windows code-signing certificate is available.
; SignTool=signtool sign /fd SHA256 /tr <timestamp-url> /td SHA256 /a $f

[Languages]
Name: "indonesian"; MessagesFile: "compiler:Languages\Indonesian.isl"

[Tasks]
Name: "desktopicon"; Description: "Buat ikon &desktop"; GroupDescription: "Ikon tambahan:"

[Files]
Source: "{#MyBuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\SIOMAY"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\SIOMAY"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Jalankan SIOMAY"; Flags: nowait postinstall skipifsilent