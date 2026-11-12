@echo off
setlocal
if "%~1"=="" (
  start "" powershell.exe -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Interface.ps1"
  exit /b
)
set "CONVERTER_PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%CONVERTER_PYTHON%" (
  echo The bundled Python runtime was not found. See Instructions.txt.
  pause
  exit /b 1
)
"%CONVERTER_PYTHON%" "%~dp0convert.py" %*
echo.
pause
