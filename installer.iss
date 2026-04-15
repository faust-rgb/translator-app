#define MyAppName "TranslatorApp"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Local"
#define MyAppExeName "TranslatorApp.exe"

[Setup]
AppId={{D7D16596-6B23-49A6-BD55-E9098DBA8102}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer-output
OutputBaseFilename=TranslatorApp-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "tessdata\*"; DestDir: "{app}\tessdata"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\TranslatorApp"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\TranslatorApp"; Filename: "{app}\{#MyAppExeName}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "启动 TranslatorApp"; Flags: nowait postinstall skipifsilent
