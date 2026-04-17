@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PYTHON_EXE=C:\Users\bobg6\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "SCRIPT_PATH=%SCRIPT_DIR%hospitalist_invoice_generator.py"

if exist "%PYTHON_EXE%" (
  echo Starting Hospitalist Invoice Generator...
  echo.
  echo Open http://127.0.0.1:8765 in your browser
  echo Press Ctrl+C in this window to stop the server.
  echo.
  "%PYTHON_EXE%" "%SCRIPT_PATH%" --serve
  goto :eof
)

echo Bundled Python runtime was not found.
echo.
echo You can still run the app by installing openpyxl into your normal Python:
echo   python -m pip install openpyxl
echo   python "%SCRIPT_PATH%" --serve
echo.
pause
