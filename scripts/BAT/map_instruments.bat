@echo off
setlocal EnableExtensions EnableDelayedExpansion
rem scripts/BAT/map_instruments.bat

set "HERE=%~dp0"
set "PS1=%HERE%..\PS1\run_map_instruments.ps1"

if not exist "%PS1%" (
  echo [ERROR] PS1 not found: "%PS1%"
  endlocal & exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
set "RC=%ERRORLEVEL%"
endlocal & exit /b %RC%
