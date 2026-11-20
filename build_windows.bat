@echo off
chcp 65001 >nul
echo === build_windows.bat  version 8 ===
set PYTHONUTF8=1
set "SRC=%~dp0"
set "WORK=%USERPROFILE%\ReceiptOCR_build"
set "LOG=%WORK%\build_log.txt"

REM Best: run on a normal 64-bit Windows 10/11 PC (Intel/AMD).
REM In a VM on a Mac (Parallels) the standard build crashes (PyInstaller loads
REM onnxruntime), so this script automatically retries in "compatibility mode".
REM The resulting .exe can NOT be started inside that VM - test it on a real PC.
REM OCR models are included in the "models" folder (no download needed).

set "PY="
call :findpy
if not defined PY (
  echo Python 3.11-3.13 not found. Trying winget ...
  winget install -e --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements >nul 2>&1
  call :findpy
)
if not defined PY (
  echo winget unavailable/failed. Downloading Python 3.12 from python.org ...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe' -OutFile '%TEMP%\python312_setup.exe'"
  if exist "%TEMP%\python312_setup.exe" (
    echo Installing Python 3.12 ^(per-user, about 1-2 minutes^) ...
    start /wait "" "%TEMP%\python312_setup.exe" /quiet InstallAllUsers=0 PrependPath=0 Include_test=0 Include_launcher=0 Include_doc=0
  )
  call :findpy
)
if not defined PY (
  echo.
  echo [!] Could not find/install Python 3.12 automatically ^(no internet?^).
  echo     Install it manually from https://www.python.org/downloads/release/python-3128/
  echo     then run this file again.
  pause
  exit /b 1
)
echo Using: %PY%

REM --- check Python can really start (0xc000007b / VCRUNTIME140.dll problem) ---
%PY% --version >nul 2>&1
if errorlevel 1 (
  echo Python was found but cannot start ^(bad VCRUNTIME140.dll^). Repairing ...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://aka.ms/vs/17/release/vc_redist.x64.exe' -OutFile '%TEMP%\vc_redist.x64.exe'"
  if exist "%TEMP%\vc_redist.x64.exe" start /wait "" "%TEMP%\vc_redist.x64.exe" /install /quiet /norestart
  for %%V in (312 313 311) do del /q "%LOCALAPPDATA%\Programs\Python\Python%%V\vcruntime140*.dll" >nul 2>&1
  %PY% --version >nul 2>&1
  if errorlevel 1 (
    echo.
    echo [X] Python still cannot start on this Windows.
    echo     This usually means Windows running inside a VM on a Mac ^(Parallels^) or an
    echo     unsupported CPU. Please run this script on a normal Intel/AMD Windows PC.
    pause
    exit /b 1
  )
)

echo [1/4] Copy files to %WORK%
if not exist "%WORK%" mkdir "%WORK%"
robocopy "%SRC%." "%WORK%" app.py receipt_core.py requirements.txt /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 goto :err
robocopy "%SRC%models" "%WORK%\models" /E /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 goto :err
if not exist "%WORK%\models\th_PP-OCRv5_rec_mobile.onnx" (
  echo [X] models folder is missing - copy the whole project folder including "models".
  goto :err
)
cd /d "%WORK%" || goto :err
echo build log > "%LOG%"

echo [2/4] Create virtual environment
if exist .venv rmdir /s /q .venv
%PY% -c "import sys,platform,struct;print('Python',sys.version);print('machine',platform.machine(),struct.calcsize('P')*8,'bit')" >>"%LOG%" 2>&1
%PY% -m venv .venv >>"%LOG%" 2>&1
if errorlevel 1 (
  echo      venv failed, trying again with --without-pip ...
  if exist .venv rmdir /s /q .venv
  %PY% -m venv --without-pip .venv >>"%LOG%" 2>&1 || goto :err
  echo      installing pip into the venv ...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%WORK%\get-pip.py'" >>"%LOG%" 2>&1
  .venv\Scripts\python.exe "%WORK%\get-pip.py" >>"%LOG%" 2>&1 || goto :err
)
set "VPY=%WORK%\.venv\Scripts\python.exe"
if not exist "%VPY%" (
  echo.
  echo [X] The virtual environment was NOT created ^(%VPY% is missing^).
  echo     Diagnostics - Python used: %PY%
  %PY% -c "import sys;print('exe:',sys.executable);print('version:',sys.version)"
  echo     Contents of %WORK%:
  dir /b "%WORK%"
  goto :err
)

echo [3/4] Install packages (first time takes a few minutes)
"%VPY%" -m pip install --upgrade pip >>"%LOG%" 2>&1
"%VPY%" -m pip install -r requirements.txt pyinstaller >>"%LOG%" 2>&1 || goto :err

echo [4/4] Build the .exe with PyInstaller (a few minutes)
"%VPY%" -m PyInstaller --noconfirm --clean --windowed --name ReceiptOCR --add-data "models;models" --collect-data rapidocr --collect-all tkinterdnd2 --collect-all pypdfium2_raw --collect-all pypdfium2 --exclude-module matplotlib --exclude-module scipy --exclude-module pandas --exclude-module IPython --exclude-module pytest --collect-binaries onnxruntime --hidden-import rapidocr.main --hidden-import rapidocr.inference_engine.onnxruntime --hidden-import rapidocr.utils.load_image app.py >>"%LOG%" 2>&1
if not errorlevel 1 goto :built
echo      Standard build failed ^(PyInstaller cannot load onnxruntime on this PC, e.g. a VM on a Mac^).
echo      Retrying in compatibility mode: onnxruntime is copied as plain files ...
"%VPY%" -m PyInstaller --noconfirm --clean --windowed --name ReceiptOCR --add-data "models;models" --collect-data rapidocr --collect-all tkinterdnd2 --collect-all pypdfium2_raw --collect-all pypdfium2 --exclude-module matplotlib --exclude-module scipy --exclude-module pandas --exclude-module IPython --exclude-module pytest --exclude-module onnxruntime --exclude-module rapidocr.inference_engine.onnxruntime --add-data ".venv\Lib\site-packages\onnxruntime;onnxruntime" --add-data ".venv\Lib\site-packages\rapidocr\inference_engine\onnxruntime;rapidocr\inference_engine\onnxruntime" --hidden-import rapidocr.main --hidden-import rapidocr.utils.load_image --hidden-import numpy --hidden-import psutil --hidden-import platform --hidden-import linecache --hidden-import ctypes --hidden-import argparse --hidden-import subprocess app.py >>"%LOG%" 2>&1 || goto :err
:built

if not exist "%WORK%\dist\ReceiptOCR\ReceiptOCR.exe" goto :err
robocopy "%WORK%\dist\ReceiptOCR" "%SRC%dist\ReceiptOCR" /E /NFL /NDL /NJH /NJS /NP >nul
echo Compressing to zip ...
if exist "%WORK%\ReceiptOCR_win64.zip" del "%WORK%\ReceiptOCR_win64.zip"
tar -a -c -f "%WORK%\ReceiptOCR_win64.zip" -C "%WORK%\dist" ReceiptOCR
if exist "%WORK%\ReceiptOCR_win64.zip" copy /y "%WORK%\ReceiptOCR_win64.zip" "%SRC%ReceiptOCR_win64.zip" >nul
echo.
echo ===== DONE =====
echo Program : %WORK%\dist\ReceiptOCR\ReceiptOCR.exe
echo Zip file: %WORK%\ReceiptOCR_win64.zip   (also copied next to this script)
echo To install on another PC: copy the zip, right-click Extract All,
echo then double-click ReceiptOCR.exe  (no Python needed on that PC).
echo Needs 64-bit Windows 10/11. If it fails to start, install the
echo "Microsoft Visual C++ Redistributable 2015-2022 x64" from microsoft.com.
pause
exit /b 0

:err
echo.
echo [X] Build failed. See %LOG%  (last lines below)
echo ------------------------------------------------------------
powershell -NoProfile -Command "Get-Content -Tail 25 '%LOG%'"
echo ------------------------------------------------------------
pause
exit /b 1

:findpy
for %%V in (312 313 311) do (
  if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" set PY="%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
  if not defined PY if exist "%ProgramFiles%\Python%%V\python.exe" set PY="%ProgramFiles%\Python%%V\python.exe"
)
for %%V in (3.12 3.13 3.11) do (
  if not defined PY (
    py -%%V --version >nul 2>&1 && set "PY=py -%%V"
  )
)
exit /b 0
