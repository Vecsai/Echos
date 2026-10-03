; ============================================================
;  Echos — Inno Setup Script
;  Собирает: Echos.exe + Echos_reader.exe + VC_redist + Ollama
;           + FFmpeg + Graphviz
; ============================================================

#define MyAppName "Echos"
#define MyAppVersion "7.0.0"
#define MyAppPublisher "Vecsai"
#define MyAppURL "https://github.com/Vecsai/Echos"
#define MyAppExeName "Echos.exe"

[Setup]
AppId={{8F3A2B1C-4D5E-4F6A-9B7C-1D2E3F4A5B6C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases

DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
AllowNoIcons=yes

OutputDir=Output
OutputBaseFilename=Echos_Setup_{#MyAppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes

SetupIconFile=Icon\Icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

; === КАРТИНКИ ВИЗАРДА (classic-стиль) ===
WizardImageFile=installer_assets\Background.bmp
WizardSmallImageFile=installer_assets\Corner.bmp
WizardImageStretch=yes

; === ЛИЦЕНЗИИ ===
; InfoBeforeFile — страница со сторонними лицензиями показывается ПОСЛЕ лицензии Echos,
; но ДО выбора папки установки.
InfoBeforeFile=installer_licenses.txt

ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

WizardStyle=classic
WizardResizable=no
DisableWelcomePage=no

MinVersion=10.0

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"; LicenseFile: "LICENSE.ru"
Name: "english"; MessagesFile: "compiler:Default.isl"; LicenseFile: "LICENSE"

[Messages]
; === Русский ===
russian.WizardLicense=Лицензионное соглашение Echos
russian.LicenseLabel=Пожалуйста, прочитайте лицензионное соглашение перед установкой Echos.
russian.WizardInfoBefore=Лицензии сторонних компонентов
russian.InfoBeforeLabel=Echos использует открытые библиотеки и программы (Ollama, Graphviz, FFmpeg, PySide6, torch, sentence-transformers и другие). Ниже приведены тексты их лицензий.

; === English ===
english.WizardLicense=End User License Agreement
english.LicenseLabel=Please read the license agreement before installing Echos.
english.WizardInfoBefore=Third-party licenses
english.InfoBeforeLabel=Echos uses open-source libraries and programs (Ollama, Graphviz, FFmpeg, PySide6, torch, sentence-transformers and others). Their license texts are shown below.

[Types]
Name: "full"; Description: "Полная установка (все компоненты)"
Name: "compact"; Description: "Только программа и runtime"
Name: "custom"; Description: "Выборочная установка"; Flags: iscustom

[Components]
Name: "main";     Description: "Echos (основная программа)"; Types: full compact custom; Flags: fixed
Name: "reader";   Description: "Echos Reader (консоль отладки)"; Types: full custom
Name: "vcredist"; Description: "Visual C++ Redistributable (требуется)"; Types: full compact custom; Flags: fixed
Name: "ollama";   Description: "Ollama (локальный сервер моделей)"; Types: full custom
Name: "ffmpeg";   Description: "FFmpeg (для транскрибации аудио/видео)"; Types: full custom
Name: "graphviz"; Description: "Graphviz (для схем и диаграмм)"; Types: full custom

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Дополнительные значки:"; Flags: checkedonce
Name: "quicklaunchicon"; Description: "Создать ярлык в меню «Пуск»"; GroupDescription: "Дополнительные значки:"; Flags: checkedonce

[Files]
Source: "dist\Echos\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: main
Source: "dist\Echos_reader.exe"; DestDir: "{app}"; Flags: ignoreversion; Components: reader
Source: "redist\VC_redist.x64.exe"; DestDir: "{app}\_redist"; Flags: ignoreversion; Components: vcredist; Check: VCRedistNeedsInstall
Source: "redist\OllamaSetup.exe"; DestDir: "{app}\_redist"; Flags: ignoreversion; Components: ollama; Check: not OllamaInstalled
Source: "redist\ffmpeg\bin\ffmpeg.exe";  DestDir: "{app}\ffmpeg"; Flags: ignoreversion; Components: ffmpeg
Source: "redist\ffmpeg\bin\ffprobe.exe"; DestDir: "{app}\ffmpeg"; Flags: ignoreversion; Components: ffmpeg
Source: "redist\graphviz-install.exe"; DestDir: "{app}\_redist"; Flags: ignoreversion; Components: graphviz; Check: not GraphvizInstalled

[Icons]
Name: "{group}\{#MyAppName}";        Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Удалить {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";  Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\_redist\VC_redist.x64.exe"; \
  Parameters: "/install /quiet /norestart"; \
  StatusMsg: "Установка Visual C++ Redistributable..."; \
  Flags: waituntilterminated; \
  Components: vcredist; \
  Check: VCRedistNeedsInstall

Filename: "{app}\_redist\OllamaSetup.exe"; \
  Parameters: "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; \
  StatusMsg: "Установка Ollama..."; \
  Flags: waituntilterminated; \
  Components: ollama; \
  Check: not OllamaInstalled

Filename: "{app}\_redist\graphviz-install.exe"; \
  Parameters: "/S /D=C:\Program Files\Graphviz"; \
  StatusMsg: "Установка Graphviz..."; \
  Flags: waituntilterminated; \
  Components: graphviz; \
  Check: not GraphvizInstalled

Filename: "{app}\{#MyAppExeName}"; \
  Description: "Запустить {#MyAppName}"; \
  Flags: nowait postinstall skipifsilent

[Registry]
Root: HKLM; Subkey: "SOFTWARE\Echos"; ValueType: string; ValueName: "FFmpegPath";   ValueData: "{app}\ffmpeg";         Flags: uninsdeletekey; Components: ffmpeg
Root: HKLM; Subkey: "SOFTWARE\Echos"; ValueType: string; ValueName: "GraphvizPath"; ValueData: "{pf}\Graphviz\bin";    Flags: uninsdeletekey; Components: graphviz

[UninstallDelete]
Type: filesandordirs; Name: "{app}\_redist"

[Code]
function VCRedistNeedsInstall: Boolean;
var
  Version: String;
begin
  if RegQueryStringValue(HKLM,
      'SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64',
      'Version', Version) then
  begin
    Log('VC++ Redist x64 найден, версия: ' + Version);
    Result := (CompareStr(Version, 'v14.30.0.0') < 0);
  end
  else
  begin
    Log('VC++ Redist x64 не найден');
    Result := True;
  end;
end;

function OllamaInstalled: Boolean;
var
  InstallPath: String;
begin
  Result := RegQueryStringValue(HKCU, 'SOFTWARE\Ollama', 'InstallLocation', InstallPath)
            and FileExists(InstallPath + '\ollama.exe');
  if Result then
    Log('Ollama найден: ' + InstallPath)
  else
    Log('Ollama не найден');
end;

function GraphvizInstalled: Boolean;
var
  InstallPath: String;
begin
  Result := RegQueryStringValue(HKLM, 'SOFTWARE\Graphviz', 'InstallPath', InstallPath)
            and FileExists(InstallPath + '\bin\dot.exe');
  if not Result then
    Result := FileExists(ExpandConstant('{pf}\Graphviz\bin\dot.exe'));
end;
