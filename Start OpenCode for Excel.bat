@echo off
title OpenCode for Excel
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo.
  echo  Node.js is not installed yet.
  echo.
  echo  Opening the download page - install the LTS version, just click
  echo  Next through the installer, then double-click this file again.
  echo.
  start https://nodejs.org/en/download
  pause
  exit /b
)

node "tools\start.mjs"
echo.
echo  The add-in has stopped. You can close this window.
pause >nul
