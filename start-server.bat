@echo off
rem ============================================================
rem  data-cleaner : run the BUILT app as a local web server
rem
rem  HOW TO USE (Korean note is in CLAUDE.md of this folder)
rem    - Run "npm run build" first. This script serves dist\ .
rem    - Then double-click this file. A browser opens automatically.
rem    - Python must be installed.
rem    - Other port : start-server.bat 9001
rem    - Stop        : press Ctrl+C, or just close this window.
rem
rem  WHY dist\ ?
rem    This app imports the SheetJS package, so the source folder
rem    cannot be opened directly by a plain static server.
rem    "npm run build" bundles everything into dist\ .
rem    For development use "npm start" (vite, port 8081) instead.
rem
rem  NOTE: messages below are ASCII only on purpose.
rem        cmd.exe garbles Korean text in .bat files depending on
rem        the console code page, so do NOT add Korean here.
rem ============================================================

cd /d "%~dp0"

set PORT=%1
if "%PORT%"=="" set PORT=8081

if not exist "dist\index.html" (
  echo.
  echo  [ERROR] dist\index.html not found.
  echo          Run this first in this folder:
  echo.
  echo              npm install
  echo              npm run build
  echo.
  echo          Or use "npm start" for the dev server.
  echo.
  pause
  exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo  [ERROR] Python not found. Cannot start the local server.
  echo          Use "npm run preview" instead.
  echo.
  pause
  exit /b 1
)

echo.
echo  data-cleaner - local server (serving dist\)
echo.
echo    URL  : http://localhost:%PORT%
echo    Stop : Ctrl+C  (or close this window)
echo.

rem Pass "nobrowser" as the 2nd argument to skip opening the browser.
if /i "%2"=="nobrowser" goto serve
start "" "http://localhost:%PORT%"

:serve
python -m http.server %PORT% --directory dist
