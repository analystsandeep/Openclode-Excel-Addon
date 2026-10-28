@echo off
title Remove OpenCode from Excel
cd /d "%~dp0"

echo.
echo  Removing the OpenCode add-in from Excel...
echo  (Excel must be closed for this to work)
echo.

set "ADDIN_ID="
for /f "tokens=3 delims=<>" %%i in ('findstr /i "<Id>" manifest.xml') do set "ADDIN_ID=%%i"

if not defined ADDIN_ID (
  echo  Could not read the add-in id from manifest.xml.
  echo  Make sure this file is in the same folder as manifest.xml.
  echo.
  pause
  exit /b 1
)

reg delete "HKCU\Software\Microsoft\Office\16.0\Wef\Developer" /v "%ADDIN_ID%" /f >nul 2>&1
if errorlevel 1 (
  echo  The add-in was not registered, so there was nothing to remove.
) else (
  echo  Removed %ADDIN_ID% from Excel.
)

echo.
echo  Done. Close Excel and open it again.
pause
