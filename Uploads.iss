; -- Inno Setup Script for AutoFlow Installer (Self-Contained .NET) --

[Setup]
AppName=AutoFlow
AppVersion=1.2
AppPublisher=Manchu Uday
DefaultDirName={localappdata}\Uploadings
OutputDir=C:\Users\Doctus.DBS\Fchanges\Converter\InstallerOutput
OutputBaseFilename=AutoFlowInstaller
Compression=lzma
SolidCompression=yes
DisableWelcomePage=no
PrivilegesRequired=lowest
CreateAppDir=yes

SetupIconFile=D:\PdfAutomationApp\PdfAutomationApp - Copy\publish\favi.ico
UninstallDisplayIcon={app}\favi.ico

[Files]
; ✅ Copy all published files
Source: "D:\PdfAutomationApp\PdfAutomationApp - Copy\publish\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

; ✅ Explicitly copy icon (optional)
Source: "D:\PdfAutomationApp\PdfAutomationApp - Copy\publish\favi.ico"; DestDir: "{app}"

[Icons]
Name: "{group}\AutoFlow"; Filename: "{app}\PdfAutomationApp.exe"; WorkingDir: "{app}"; IconFilename: "{app}\favi.ico"
Name: "{userdesktop}\AutoFlow"; Filename: "{app}\PdfAutomationApp.exe"; WorkingDir: "{app}"; IconFilename: "{app}\favi.ico"

[Run]
Filename: "{app}\PdfAutomationApp.exe"; Description: "Launch AutoFlow"; Flags: nowait postinstall skipifsilent

[Dirs]
; ✅ Create custom folders if required
Name: "{app}\Profile"

[UninstallDelete]
; ✅ Remove all files and folders from install directory
Type: filesandordirs; Name: "{app}"
; ✅ Also remove local app data folder used by this app
Type: filesandordirs; Name: "{localappdata}\Uploadings"
; ✅ Optionally clean cache or other app data
Type: filesandordirs; Name: "{userappdata}\AutoFlow"
