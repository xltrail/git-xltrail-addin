#define MyAppName "Git xltrail Excel Addin"
#define MyAppVersion "1.2.105"

#define PathToX86Binary "XltrailClient\bin\Release\XltrailClient-packed.xll"
#ifnexist PathToX86Binary
  #pragma error PathToX86Binary + " does not exist, please build it first."
#endif

#define PathToX64Binary "XltrailClient\bin\Release\XltrailClient64-packed.xll"
#ifnexist PathToX64Binary
  #pragma error PathToX64Binary + " does not exist, please build it first."
#endif

#define MyAppPublisher "Zoomer Analytics LLC"
#define MyAppURL "https://www.xltrail.com/git-xltrail"
#define MyAppFilePrefix "git-xltrail-addin"

[Setup]
PrivilegesRequired=lowest
AppCopyright=Zoomer Analytics LLC
AppId={{3567B627-1973-42A8-92BA-2C0E7F9587C8}
AppName={#MyAppName}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
AppVersion={#MyAppVersion}
ArchitecturesInstallIn64BitMode=x64
ChangesEnvironment=yes
Compression=lzma
DefaultDirName={localappdata}\xltrail
DirExistsWarning=no
DisableReadyPage=True
LicenseFile=LICENSE.md
OutputBaseFilename={#MyAppFilePrefix}-{#MyAppVersion}
OutputDir=bin
SetupIconFile=git-xltrail-logo.ico
SolidCompression=yes
;UninstallDisplayIcon={app}\git-xltrail.exe
UsePreviousAppDir=no
;VersionInfoVersion={#MyVersionInfoVersion}
;WizardImageFile=git-xltrail-wizard-image.bmp
;WizardSmallImageFile=git-xltrail-logo.bmp


;DefaultGroupName=Pathio
;UninstallFilesDir={localappdata}\Pathio\uninstall
;UninstallDisplayIcon={localappdata}\Pathio\bin\manager\{#AppVer}\icon.ico
;UninstallDisplayName=Pathio
;Compression=lzma2
;SolidCompression=yes
;OutputBaseFilename=Pathio


[Dirs]
Name: "{localappdata}\xltrail\bin"
Name: "{localappdata}\xltrail\config"
Name: "{localappdata}\xltrail\repositories"
Name: "{localappdata}\xltrail\staging"


[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Run]
; Uninstalls the old Git xltrail version that used a different installer in a different location:
;  If we don't do this, Git will prefer the old version as it is in the same directory as it.
Filename: "{code:GetExistingGitInstallation}\git-xltrail-uninstaller.exe"; Parameters: "/S"; Flags: skipifdoesntexist

[Files]
Source: {#PathToX86Binary}; DestDir: "{localappdata}"; Flags: ignoreversion; DestName: "git-xltrail.exe"; Check: not Is64BitInstallMode
Source: {#PathToX64Binary}; DestDir: "{localappdata}"; Flags: ignoreversion; DestName: "git-xltrail.exe"; Check: Is64BitInstallMode

[Files]
DestDir: {localappdata}\xltrail\bin; Source: XltrailClient\bin\Release\git-xltrail.xll; AfterInstall: RegisterAddin('{localappdata}\bin\git-xltrail.xll'); Check: not Is64BitInstallMode
DestDir: {localappdata}\xltrail\bin; Source: XltrailClient\bin\Release\git-xltrail64.xll; AfterInstall: RegisterAddin('{localappdata}\bin\git-xltrail64.xll'); Check: Is64BitInstallMode


[Code]
var
  WelcomePage: TOutputMsgWizardPage;
  HostPage : TInputQueryWizardPage;

procedure InitializeWizard;
begin
  WelcomePage := CreateOutputMsgPage(wpWelcome, 'Install Git xltrail Addin', 'Please read.',
  'This installs the Git xltrail Addin for Excel on your computer. ' +
    'This integrates Git for Excel into your workflow.'#13#13 +
    'Click Next to continue with the installation.');

end;

function HostPage_NextButtonClick(Page: TWizardPage): Boolean;
begin
  Result := True;
end;


[Code]
function GetDefaultDirName(Dummy: string): string;
begin
  if IsAdminLoggedOn then begin
    Result:=ExpandConstant('{pf}\{#MyAppName}');
  end else begin
    Result:=ExpandConstant('{userpf}\{#MyAppName}');
  end;
end;

// Uses cmd to parse and find the location of Git through the env vars.
// Currently only used to support running the uninstaller for the old Git xltrail version.
function GetExistingGitInstallation(Value: string): string;
var
  TmpFileName: String;
  ExecStdOut: AnsiString;
  ResultCode: integer;

begin
  TmpFileName := ExpandConstant('{tmp}') + '\git_location.txt';

  Exec(
    ExpandConstant('{cmd}'),
    '/C "for %i in (git.exe) do @echo. %~$PATH:i > "' + TmpFileName + '"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode
  );

  if LoadStringFromFile(TmpFileName, ExecStdOut) then begin
      if not (Pos('Git\cmd', ExtractFilePath(ExecStdOut)) = 0) then begin
        // Proxy Git path detected
        Result := ExpandConstant('{pf}');
      if Is64BitInstallMode then
        Result := Result + '\Git\mingw64\bin'
      else
        Result := Result + '\Git\mingw32\bin';
      end else begin
        Result := ExtractFilePath(ExecStdOut);
      end;

      DeleteFile(TmpFileName);
  end;
end;

// Checks to see if we need to add the dir to the env PATH variable.
function NeedsAddPath(Param: string): boolean;
var
  OrigPath: string;
  ParamExpanded: string;
begin
  //expand the setup constants like {app} from Param
  ParamExpanded := ExpandConstant(Param);
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE,
    'SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
    'Path', OrigPath)
  then begin
    Result := True;
    exit;
  end;
  // look for the path with leading and trailing semicolon and with or without \ ending
  // Pos() returns 0 if not found
  Result := Pos(';' + UpperCase(ParamExpanded) + ';', ';' + UpperCase(OrigPath) + ';') = 0;
  if Result = True then
    Result := Pos(';' + UpperCase(ParamExpanded) + '\;', ';' + UpperCase(OrigPath) + ';') = 0;
end;

// Runs the xltrail initialization.
procedure InstallGitxltrail();
var
  ResultCode: integer;
begin
  Exec(
    ExpandConstant('{cmd}'),
    ExpandConstant('/C ""{app}\git-xltrail.exe" install"'),
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode
  );
  if not ResultCode = 1 then
    MsgBox(
    'Git xltrail was not able to automatically initialize itself. ' +
    'Please run "git xltrail install" from the commandline.', mbInformation, MB_OK);
end;

// Event function automatically called when uninstalling:
function InitializeUninstall(): Boolean;
var
  ResultCode: integer;
begin
  Exec(
    ExpandConstant('{cmd}'),
    ExpandConstant('/C ""{app}\git-xltrail.exe" uninstall"'),
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode
  );
  Result := True;
end;


function GetExcelVersionNumberAsString(): String;
var
  CurVer: Cardinal;
  Version: string;
begin
  if RegQueryDWordValue(HKCR, 'Excel.Application\CurVer\','', CurVer) then begin
    Version := IntTOStr(CurVer)
	if Version = 'Excel.Application.16' then begin
	   Result := '16.0'
	end else if Version = 'Excel.Application.15' then begin
	   Result := '15.0'
	end else if Version = 'Excel.Application.14' then begin
	   Result := '14.0'
	end else if Version = 'Excel.Application.12' then begin
	   Result := '12.0'
	end else if Version = 'Excel.Application.11' then begin
	   Result := '11.0'
	end else if Version = 'Excel.Application.10' then begin
	   Result := '10.0'
	end else if Version = 'Excel.Application.9' then begin
	   Result := '9.0'
	end else begin
	   Result := '0.0';
	end;
  end;
end;



(*
    Importing a Windows API function for understanding if EXCEL.EXE is a 32-bit
    or 64-bit application. This function has a few named constants. There appears
    to be no other accurate option than looking at the exe's PE header.
*)
const
    SCS_32BIT_BINARY = 0;   // A 32-bit Windows-based application
    SCS_64BIT_BINARY = 6;   // A 64-bit Windows-based application
    SCS_DOS_BINARY   = 1;   // An MS-DOS – based application
    SCS_OS216_BINARY = 5;   // A 16-bit OS/2-based application
    SCS_PIF_BINARY   = 3;   // A PIF file that executes an MS-DOS – based application
    SCS_POSIX_BINARY = 4;   // A POSIX – based application
    SCS_WOW_BINARY   = 2;   // A 16-bit Windows-based application
 
function GetBinaryType(lpApplicationName: AnsiString; var lpBinaryType: Integer): Boolean;
    external 'GetBinaryTypeA@kernel32.dll stdcall';
 

function IsExcel64Bit(): Boolean;
(*
    64-bit versions of Office use the standard node of the Registry named
    'Software\Microsoft\' which is the same for 32-bit Excel on 32-bit Windows.
    The node named 'Software\Wow6432Node\Microsoft\' is only used for 32-bit
    Excel installed on 64-bit Windows so we can avoid checking.
 
    Makes sure you have these [Setup] directives correctly set otherwise it may
    not read from the correct branch in the Registry on 64-bit systems:
        ArchitecturesAllowed=x86 x64
        ArchitecturesInstallIn64BitMode=x64
*)
var
    InstallRoot: String;
    BinaryType: Integer;
begin
    Result := false;
 
    // Look for the InstallRoot of Excel. This is where Excel is installed.
    //
    if not RegQueryStringValue(HKEY_LOCAL_MACHINE, 'Software\Microsoft\Office\' + GetExcelVersionNumberAsString + '\Excel\InstallRoot', 'Path', InstallRoot) then exit;
 
    // Look what binary type 'EXCEL.EXE' is.
    //
    if GetBinaryType(AddBackslash(InstallRoot) + 'excel.exe', BinaryType) then begin
        Result := (BinaryType = SCS_64BIT_BINARY);
    end;
end;



procedure RegisterAddin(RegistryData: string);
var
    ExcelAddIn: Variant;
    OpenCounter: Integer;
    RegistryKey: String;
    RegistryValue: String;
begin
	// Close Excel. Must do so otherwise the next settings are voided.
	//
	//KillExcelApp;
 
	// To install an Add-In via the Registry, write a REG_SZ value called "OPENx"
	// to the "HKCU\Software\Microsoft\Office\\Excel\Options" key.
	//
	RegistryKey := 'Software\Microsoft\Office\' + GetExcelVersionNumberAsString + '\Excel\Options';
	OpenCounter := 0;
	while (OpenCounter >= 0) do begin
		// The "OPENx" value is "OPEN" for the first Add-In
		// and "OPEN1", "OPEN2", "OPEN..." for the next ones.
		//
		if (OpenCounter = 0) then begin
			RegistryValue := 'OPEN';
		end else begin
			RegistryValue := 'OPEN' + IntToStr(OpenCounter);
		end;
 
		// If the value exists then it pertains to another Add-In.
		// If missing, we add the value because this is our baby.
		//
		if not RegValueExists(HKEY_CURRENT_USER, RegistryKey, RegistryValue) then begin
			Log('Writing Registry entry: ' + RegistryKey + ', ' + RegistryValue + ', ' + RegistryData);
 
			RegWriteStringValue(HKEY_CURRENT_USER, RegistryKey, RegistryValue, RegistryData);
 
			// Stop here.
			//
			OpenCounter := -1;
		end else begin
			// Check next.
			//
			OpenCounter := OpenCounter + 1;
		end;
	end;
end;