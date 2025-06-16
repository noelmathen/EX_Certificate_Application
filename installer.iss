; ─────────────────────────────────────────────────────────────
; installer.iss – build with Inno Setup (ISCC.exe)
; ─────────────────────────────────────────────────────────────

[Setup]
; ---- General ----
AppName=proEXy
AppVersion=1.0.0
AppPublisher=Noel Mathen
DefaultDirName={pf}\proEXy
DefaultGroupName=proEXy
OutputBaseFilename=proEXy-{#SetupSetting("AppVersion")}-Setup
SetupIconFile=assets\EX_logo.ico
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64
DisableProgramGroupPage=yes
WizardStyle=modern

[Files]
; 1️⃣  copy the entire frozen build
Source: "exe_build\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

; 2️⃣  add Ghostscript (optional but nice for lattice mode)
Source: "assets\gs10051w64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Icons]
Name: "{group}\proEXy";              Filename: "{app}\proEXy.exe"; \
      IconFilename: "{app}\assets\EX_logo.ico"
Name: "{commondesktop}\proEXy";      Filename: "{app}\proEXy.exe"; \
      Tasks: desktopicon; IconFilename: "{app}\assets\EX_logo.ico"

[Tasks]
Name: desktopicon; Description: "Create a &desktop shortcut"; Flags: unchecked

[Run]
; 3️⃣  install Ghostscript silently only if missing
Filename: "{tmp}\gs10051w64.exe"; Parameters: "/SILENT"; Check: not IsGhostscriptInstalled

[Code]
function IsGhostscriptInstalled: Boolean;
begin
  Result := RegKeyExists(HKLM, 'SOFTWARE\GPL Ghostscript\10.0');
end;
