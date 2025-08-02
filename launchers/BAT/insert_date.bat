@echo off
cls
chcp 65001 >nul
pwsh.exe -NoExit -ExecutionPolicy Bypass -File "%~dp0run_insert_date.ps1"
