@echo off
setlocal enableextensions

cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Node.js is not installed or not in PATH.
  echo Please install Node.js first: https://nodejs.org/
  pause
  exit /b 1
)

echo [991X] Building gaozhi index...
node scripts\build-gaozhi-index.js
if errorlevel 1 (
  echo [ERROR] Failed to build gaozhi index.
  pause
  exit /b 1
)

if /I "%~1"=="--build-only" (
  echo [991X] Build complete.
  exit /b 0
)

echo [991X] Starting local server: http://127.0.0.1:9910
if /I not "%~1"=="--no-open" (
  start "" "http://127.0.0.1:9910"
)
echo [991X] Press Ctrl+C to stop the server.
node server.js

set "exitCode=%errorlevel%"
if not "%exitCode%"=="0" (
  echo [991X] Server exited with code %exitCode%.
  pause
)
exit /b %exitCode%
