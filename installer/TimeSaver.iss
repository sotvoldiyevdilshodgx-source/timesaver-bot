; TimeSaver o'rnatuvchisi (Inno Setup 6). build.ps1 yig'adi: avval PyInstaller, keyin shu fayl.
#define AppName "TimeSaver"
#define AppVersion "1.0.0"
#define AppExe "TimeSaver.exe"
#define Src "..\build\dist\TimeSaver"
#define Brand "..\app\branding"

[Setup]
AppId={{6F1C2B7E-4A8D-4C2E-9B1F-7D3E5A9C1E50}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppName}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\release
OutputBaseFilename=TimeSaver-Setup-{#AppVersion}
SetupIconFile={#Brand}\timesaver.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
WizardStyle=modern
WizardImageFile={#Brand}\wizard.bmp,{#Brand}\wizard-2x.bmp
WizardSmallImageFile={#Brand}\wizard-small.bmp,{#Brand}\wizard-small-2x.bmp
Compression=lzma2/max
SolidCompression=yes
CloseApplications=force
RestartApplications=no
VersionInfoVersion={#AppVersion}
VersionInfoProductName={#AppName}

[Languages]
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "en"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
ru.TaskDesktop=Ярлык на рабочем столе
en.TaskDesktop=Desktop shortcut
ru.TaskStartup=Запускать вместе с Windows (бот работает постоянно)
en.TaskStartup=Start with Windows (the bot keeps running)
ru.RunNow=Запустить TimeSaver и открыть мастер настройки
en.RunNow=Start TimeSaver and open the setup wizard
ru.PanelLink=Панель TimeSaver
en.PanelLink=TimeSaver panel

[Tasks]
Name: "desktop"; Description: "{cm:TaskDesktop}"
Name: "startup"; Description: "{cm:TaskStartup}"

[Dirs]
; Ma'lumot papkasi: bot.db, config.ini, Excel — oddiy foydalanuvchi ham yoza oladi
Name: "{commonappdata}\TimeSaver"; Permissions: users-modify
Name: "{commonappdata}\TimeSaver\baza"; Permissions: users-modify
; C:\ProgramData\BazaReport (db.ini) ni dastur o'zi yaratadi — mavjud papkaning ruxsatlari o'zgartirilmaydi

[Files]
Source: "{#Src}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktop
Name: "{commonstartup}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: startup

[Run]
Filename: "{app}\{#AppExe}"; Description: "{cm:RunNow}"; Flags: postinstall nowait skipifsilent runasoriginaluser

[UninstallRun]
Filename: "{sys}\taskkill.exe"; Parameters: "/IM {#AppExe} /F"; Flags: runhidden; RunOnceId: "StopTimeSaver"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  Code: Integer;
begin
  { yangilashda eski nusxa ishlab tursa — to'xtatamiz (bot.db ochiq qolmasin) }
  if CurStep = ssInstall then
    Exec(ExpandConstant('{sys}\taskkill.exe'), '/IM {#AppExe} /F', '', SW_HIDE, ewWaitUntilTerminated, Code);
end;
