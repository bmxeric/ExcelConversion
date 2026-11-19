@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv-windows\Scripts\python.exe" (
    py -3.13 -m venv .venv-windows
    if errorlevel 1 goto failed
)

".venv-windows\Scripts\python.exe" -m pip install -r requirements-build.txt
if errorlevel 1 goto failed

".venv-windows\Scripts\python.exe" -m unittest test_excel_merge -v
if errorlevel 1 goto failed

".venv-windows\Scripts\python.exe" -m PyInstaller --clean --noconfirm ExcelConversion.spec
if errorlevel 1 goto failed
if not exist "dist\ExcelConversion.exe" goto failed

echo.
echo Build complete: "%~dp0dist\ExcelConversion.exe"
echo Distribute this single EXE. Python is not required on the target PC.
pause
exit /b 0

:failed
echo.
echo Build failed. Install Python 3.13 for Windows with the py launcher.
echo Check the error messages above, then run this script again.
pause
exit /b 1