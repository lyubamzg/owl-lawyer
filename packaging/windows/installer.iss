#ifndef AppVersion
  #define AppVersion "0.2.0"
#endif
#define ProjectRoot SourcePath + "..\.."

[Setup]
AppId={{32873ED1-22B2-48FA-9904-6856BE2D4A4D}
AppName=СОВА
AppVersion={#AppVersion}
AppPublisher=Owl Lawyer
AppPublisherURL=https://github.com/lyubamzg/owl-lawyer
DefaultDirName={localappdata}\Programs\OwlLawyer
DefaultGroupName=СОВА
PrivilegesRequired=lowest
ArchitecturesAllowed=x64os
ArchitecturesInstallIn64BitMode=x64os
MinVersion=10.0.17763
SourceDir={#ProjectRoot}
OutputDir={#ProjectRoot}\dist
OutputBaseFilename=SOVA-Setup-{#AppVersion}-x64
SetupIconFile={#ProjectRoot}\build\windows\owl-lawyer.ico
UninstallDisplayIcon={app}\owl-lawyer.ico
WizardStyle=modern
Compression=lzma2
SolidCompression=yes
DisableProgramGroupPage=yes
DisableWelcomePage=no
CloseApplications=yes
SetupLogging=yes
UninstallDisplayName=СОВА — рабочее место юриста

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "installmissing"; Description: "{cm:InstallMissing}"; GroupDescription: "{cm:Components}"
Name: "libreoffice"; Description: "{cm:InstallOffice}"; GroupDescription: "{cm:Components}"
Name: "desktopicon"; Description: "{cm:DesktopShortcut}"; GroupDescription: "{cm:Shortcuts}"

[CustomMessages]
russian.Components=Проверка компонентов:
russian.InstallMissing=Установить отсутствующие Python, Tesseract OCR и Poppler
russian.InstallOffice=Установить LibreOffice, если он отсутствует (редактирование и конвертация документов)
russian.Shortcuts=Ярлыки:
russian.DesktopShortcut=Создать ярлык СОВА на рабочем столе
russian.DependencyTitle=Проверка и установка компонентов
russian.DependencyInfo=Уже установленные программы будут использованы повторно. Библиотеки Python устанавливаются в отдельное окружение СОВА. Для загрузки недостающих компонентов нужен интернет; Windows может запросить права администратора.
russian.DependencyFailure=Не удалось подготовить компоненты. Исправьте причину и повторите установку.
russian.LogLocation=Журнал: %LOCALAPPDATA%\OwlLawyer\Logs\setup.log
russian.Launch=Запустить СОВА
english.Components=Component checks:
english.InstallMissing=Install missing Python, Tesseract OCR, and Poppler
english.InstallOffice=Install LibreOffice if missing (document editing and conversion)
english.Shortcuts=Shortcuts:
english.DesktopShortcut=Create a SOVA desktop shortcut
english.DependencyTitle=Checking and installing components
english.DependencyInfo=Existing programs will be reused. Python libraries are installed in a private SOVA environment. Missing components require an internet connection; Windows may request administrator approval.
english.DependencyFailure=Could not prepare the required components. Correct the problem and retry setup.
english.LogLocation=Log: %LOCALAPPDATA%\OwlLawyer\Logs\setup.log
english.Launch=Launch SOVA

[Dirs]
Name: "{app}\incoming"; Flags: uninsneveruninstall

[Files]
; Extract these before copying application files, so a dependency failure stops setup.
Source: "packaging\windows\setup-dependencies.ps1"; Flags: dontcopy
Source: "packaging\windows\requirements.txt"; Flags: dontcopy
Source: "app\*.py"; DestDir: "{app}\app"; Flags: ignoreversion
Source: "app\icons\*.png"; DestDir: "{app}\app\icons"; Flags: ignoreversion
Source: "app\input-workbook.xlsx"; DestDir: "{app}\app"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "app\demo-housing.xlsx"; DestDir: "{app}\app"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "app\demo-family-questionnaire.xlsx"; DestDir: "{app}\app"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "app\demo-consumer-questionnaire.xlsx"; DestDir: "{app}\app"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "templates\*.doc"; DestDir: "{app}\templates"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "incoming\README.txt"; DestDir: "{app}\incoming"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "incoming\demo\README.txt"; DestDir: "{app}\incoming\demo"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "incoming\demo\ledger-demo.xlsx"; DestDir: "{app}\incoming\demo"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "incoming\demo\f9-demo.pdf"; DestDir: "{app}\incoming\demo"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "build\windows\rus.traineddata"; DestDir: "{app}\tessdata"; Flags: ignoreversion
Source: "build\windows\owl-lawyer.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "packaging\windows\licenses\*.txt"; DestDir: "{app}\licenses"; Flags: ignoreversion
Source: "packaging\windows\launch.cmd"; DestDir: "{app}"; DestName: "OwlLawyer.cmd"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "windows-installation.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\СОВА"; Filename: "{app}\runtime\Scripts\pythonw.exe"; Parameters: """{app}\app\desktop_app.py"""; WorkingDir: "{app}\app"; IconFilename: "{app}\owl-lawyer.ico"
Name: "{autodesktop}\СОВА"; Filename: "{app}\runtime\Scripts\pythonw.exe"; Parameters: """{app}\app\desktop_app.py"""; WorkingDir: "{app}\app"; IconFilename: "{app}\owl-lawyer.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\runtime\Scripts\pythonw.exe"; Parameters: """{app}\app\desktop_app.py"""; WorkingDir: "{app}\app"; Description: "{cm:Launch}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; The private runtime is disposable. Never recursively remove app or incoming.
Type: filesandordirs; Name: "{app}\runtime"
Type: files; Name: "{app}\runtime-tools.json"

[Code]
procedure InitializeWizard;
var
  Page: TOutputMsgWizardPage;
begin
  Page := CreateOutputMsgPage(wpWelcome, CustomMessage('DependencyTitle'),
    '', CustomMessage('DependencyInfo'));
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Arguments, ErrorFile: String;
  Details: AnsiString;
  ExitCode: Integer;
begin
  Result := '';
  ExtractTemporaryFile('setup-dependencies.ps1');
  ExtractTemporaryFile('requirements.txt');
  ExtractTemporaryFile('rus.traineddata');
  ErrorFile := ExpandConstant('{tmp}\setup-error.txt');
  DeleteFile(ErrorFile);
  Arguments := '-NoProfile -ExecutionPolicy Bypass -File "' +
    ExpandConstant('{tmp}\setup-dependencies.ps1') + '" -InstallDir "' +
    ExpandConstant('{app}') + '" -RequirementsFile "' +
    ExpandConstant('{tmp}\requirements.txt') + '" -ReportFile "' + ErrorFile + '"';
  if WizardIsTaskSelected('installmissing') then
    Arguments := Arguments + ' -InstallMissing';
  if WizardIsTaskSelected('libreoffice') then
    Arguments := Arguments + ' -InstallLibreOffice';
  WizardForm.PreparingLabel.Caption := CustomMessage('DependencyTitle');
  if not Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
      Arguments, '', SW_SHOW, ewWaitUntilTerminated, ExitCode) then
    Result := CustomMessage('DependencyFailure')
  else if ExitCode <> 0 then
  begin
    if LoadStringFromFile(ErrorFile, Details) then
      Result := UTF8Decode(Details)
    else
      Result := CustomMessage('DependencyFailure');
  end;
  if Result <> '' then
    Result := Result + #13#10#13#10 + CustomMessage('LogLocation');
end;
