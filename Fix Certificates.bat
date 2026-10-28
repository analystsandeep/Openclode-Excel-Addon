@echo off
title Fix OpenCode for Excel certificates
cd /d "%~dp0"
echo.
echo  Reinstalling the localhost security certificate used by Excel.
echo  If Windows asks "Do you want to install this certificate?", choose Yes.
echo.
npx --yes office-addin-dev-certs install
echo.
echo  Done. Close Excel, then run "Start OpenCode for Excel.bat" again.
pause
