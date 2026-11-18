#define AddInClsid "{{1B728F4B-5AF0-41FD-8589-86E44A7A5501}"
#define PaneClsid "{{7FA35AF7-2CF8-4B2A-8F5D-2A4BCC1E21A0}"
#define AddInClassKey "Software\Classes\CLSID\" + AddInClsid
#define PaneClassKey "Software\Classes\CLSID\" + PaneClsid
#define AddInServer AddInClassKey + "\InprocServer32"
#define ManagedVersionKey AddInServer + "\" + ManagedVersion
#define AddInProgId "Software\Classes\PptGcs.Connect"
#define PaneProgId "Software\Classes\PptGcs.ConstraintPaneHost"
#define OfficeAddInKey "Software\Microsoft\Office\PowerPoint\Addins\PptGcs.Connect"

[Setup]
AppId={{952A2DBE-4C8D-486A-AB9D-F82CC24C7031}
AppName=PPT GCS
AppVersion={#ProductVersion}
AppPublisher=Social-Mean
AppPublisherURL=https://github.com/Social-Mean/ppt-gcs
AppSupportURL=https://github.com/Social-Mean/ppt-gcs/issues
DefaultDirName={localappdata}\Programs\PPT GCS
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64os
ArchitecturesInstallIn64BitMode=x64os
MinVersion=10.0
CloseApplications=yes
RestartApplications=no
WizardStyle=modern
Compression=lzma2
SolidCompression=yes
VersionInfoVersion={#ProductVersion}.0
VersionInfoDescription=PPT GCS Installer

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
FinishedLabelNoIcons=Close and reopen PowerPoint. The PPT GCS tab will appear on the ribbon.

[Files]
Source: "{#PayloadDir}\PptGcs.AddIn.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PayloadDir}\PptGcs.Core.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PayloadDir}\PptGcs.PaneHost.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#PayloadDir}\MathNet.Numerics.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "MathNet.Numerics.LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKCU64; Subkey: "{#AddInClassKey}"; ValueType: string; ValueData: "PptGcs.AddIn.Connect"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "{#AddInClassKey}\ProgId"; ValueType: string; ValueData: "PptGcs.Connect"
Root: HKCU64; Subkey: "{#AddInServer}"; ValueType: string; ValueData: "mscoree.dll"
Root: HKCU64; Subkey: "{#AddInServer}"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU64; Subkey: "{#AddInServer}"; ValueType: string; ValueName: "Class"; ValueData: "PptGcs.AddIn.Connect"
Root: HKCU64; Subkey: "{#AddInServer}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#ManagedIdentity}"
Root: HKCU64; Subkey: "{#AddInServer}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU64; Subkey: "{#AddInServer}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:ManagedCodeBase}"
Root: HKCU64; Subkey: "{#ManagedVersionKey}"; ValueType: string; ValueName: "Class"; ValueData: "PptGcs.AddIn.Connect"
Root: HKCU64; Subkey: "{#ManagedVersionKey}"; ValueType: string; ValueName: "Assembly"; ValueData: "{#ManagedIdentity}"
Root: HKCU64; Subkey: "{#ManagedVersionKey}"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU64; Subkey: "{#ManagedVersionKey}"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:ManagedCodeBase}"
Root: HKCU64; Subkey: "{#AddInProgId}"; ValueType: string; ValueData: "PptGcs.AddIn.Connect"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "{#AddInProgId}\CLSID"; ValueType: string; ValueData: "{#AddInClsid}"
Root: HKCU64; Subkey: "{#PaneClassKey}"; ValueType: string; ValueData: "PPT GCS native task pane host"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "{#PaneClassKey}\InprocServer32"; ValueType: string; ValueData: "{app}\PptGcs.PaneHost.dll"
Root: HKCU64; Subkey: "{#PaneClassKey}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Apartment"
Root: HKCU64; Subkey: "{#PaneClassKey}\ProgId"; ValueType: string; ValueData: "PptGcs.ConstraintPaneHost"
Root: HKCU64; Subkey: "{#PaneClassKey}\Control"; ValueType: none
Root: HKCU64; Subkey: "{#PaneClassKey}\Implemented Categories\{{40FC6ED4-2438-11CF-A3DB-080036F12502}"; ValueType: none
Root: HKCU64; Subkey: "{#PaneProgId}"; ValueType: string; ValueData: "PPT GCS native task pane host"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "{#PaneProgId}\CLSID"; ValueType: string; ValueData: "{#PaneClsid}"
Root: HKCU64; Subkey: "{#OfficeAddInKey}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "PPT GCS"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "{#OfficeAddInKey}"; ValueType: string; ValueName: "Description"; ValueData: "Interactive layout constraints for PowerPoint shapes"
Root: HKCU64; Subkey: "{#OfficeAddInKey}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"

[Code]
const
  PowerPointAppPath = 'Software\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE';
  SCS_64BIT_BINARY = 6;

function GetBinaryType(FileName: String; var BinaryType: Cardinal): Integer;
  external 'GetBinaryTypeW@kernel32.dll stdcall';

function UrlCreateFromPath(Path, Url: String; var UrlLength: Cardinal; Flags: Cardinal): Integer;
  external 'UrlCreateFromPathW@shlwapi.dll stdcall setuponly';

function ManagedCodeBase(Param: String): String;
var
  Url: String;
  UrlLength: Cardinal;
begin
  UrlLength := 4096;
  SetLength(Url, UrlLength);
  if UrlCreateFromPath(ExpandConstant('{app}\PptGcs.AddIn.dll'), Url, UrlLength, 0) <> 0 then
    RaiseException('Could not convert the add-in path to a file URL.');
  SetLength(Url, UrlLength);
  Result := Url;
end;

function HasPowerPoint64(RegistryRoot: Integer): Boolean;
var
  ExePath: String;
  BinaryType: Cardinal;
begin
  Result := False;
  if not RegQueryStringValue(RegistryRoot, PowerPointAppPath, '', ExePath) then Exit;
  if GetBinaryType(RemoveQuotes(ExePath), BinaryType) <> 0 then
    Result := BinaryType = SCS_64BIT_BINARY;
end;

function InitializeSetup(): Boolean;
var
  NetRelease: Cardinal;
begin
  Result := False;
  if not RegQueryDWordValue(HKLM64, 'Software\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', NetRelease)
    or (NetRelease < 528040) then
  begin
    SuppressibleMsgBox('PPT GCS requires .NET Framework 4.8 or later.', mbError, MB_OK, IDOK);
    Exit;
  end;
  if not (HasPowerPoint64(HKCU64) or HasPowerPoint64(HKLM64) or HasPowerPoint64(HKLM32)) then
  begin
    SuppressibleMsgBox('PPT GCS requires 64-bit desktop Microsoft PowerPoint. Install it before running Setup.', mbError, MB_OK, IDOK);
    Exit;
  end;
  Result := True;
end;
