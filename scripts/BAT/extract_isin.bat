@echo off
setlocal
cls
chcp 65001 >nul

REM Пробрасываем все аргументы дальше в PS1
pwsh -NoLogo -ExecutionPolicy Bypass -File "%~dp0..\PS1\run_extract_isin.ps1" %*
set "rc=%ERRORLEVEL%"

echo.
if %rc%==0 (
  echo ✅ extract_isin завершен успешно.
) else (
  echo ❌ extract_isin завершился с кодом %rc%.
)

pause
exit /b %rc%

