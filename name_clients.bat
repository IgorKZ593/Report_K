@echo off
cls
chcp 65001 >nul
pwsh -ExecutionPolicy Bypass -File "%~dp0run_name_clients.ps1"
