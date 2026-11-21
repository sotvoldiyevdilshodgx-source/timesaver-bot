# TimeSaver: setup.exe yig'ish.
#   1) logolarni yangilaydi  2) testlar  3) PyInstaller -> build\dist\TimeSaver  4) Inno Setup -> release\TimeSaver-Setup-<versiya>.exe
# Ishlatish (PowerShell):  .\build.ps1          (testsiz: .\build.ps1 -SkipTests)
param([switch]$SkipTests)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = Join-Path $root ".build-venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
  python -m venv (Join-Path $root ".build-venv")
  & $py -m pip install --disable-pip-version-check -q pyinstaller openpyxl pystray pillow pyodbc pywebview
}
# eski .build-venv da pywebview (panel oynasi) bo'lmasa — qo'shamiz
& $py -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('webview') else 1)"
if ($LASTEXITCODE -ne 0) { & $py -m pip install --disable-pip-version-check -q pywebview }
$iscc = @("$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe", "C:\Program Files (x86)\Inno Setup 6\ISCC.exe", "C:\Program Files\Inno Setup 6\ISCC.exe") |
  Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw "Inno Setup 6 topilmadi (https://jrsoftware.org/isdl.php)" }

Push-Location (Join-Path $root "app\branding"); & $py make_logo.py; Pop-Location

if (-not $SkipTests) {
  $t = Join-Path $env:TEMP ("ts_test_" + [guid]::NewGuid().ToString("N"))
  Copy-Item (Join-Path $root "app") $t -Recurse
  Set-Content (Join-Path $t "config.ini") "[bot]`r`ntoken =`r`nadmin_ids = 1`r`n[excel]`r`nfile =`r`n" -Encoding UTF8
  Push-Location $t
  $ErrorActionPreference = "Continue"   # testlar stderr ga yozadi — bu xato emas
  & $py -X utf8 -W ignore -m unittest discover -s tests -p "test_*.py" 2>&1 | ForEach-Object { "$_" } | Select-String -Pattern "^(Ran|OK|FAILED|FAIL:|ERROR:)"
  $code = $LASTEXITCODE
  $ErrorActionPreference = "Stop"
  Pop-Location
  Remove-Item $t -Recurse -Force -ErrorAction SilentlyContinue
  if ($code -ne 0) { Write-Warning "Testlarda xato bor (haqiqiy Excel faylni talab qiladigan test kutilgan)." }
}

$build = Join-Path $root "build"
& $py -m PyInstaller --noconfirm --clean --distpath (Join-Path $build "dist") --workpath (Join-Path $build "work") (Join-Path $root "installer\TimeSaver.spec")
if ($LASTEXITCODE -ne 0) { throw "PyInstaller xatosi" }
& $iscc /Q (Join-Path $root "installer\TimeSaver.iss")
if ($LASTEXITCODE -ne 0) { throw "Inno Setup xatosi" }
Get-ChildItem (Join-Path $root "release") -Filter *.exe | Select-Object Name, @{n="MB";e={[math]::Round($_.Length / 1MB, 1)}}, LastWriteTime
