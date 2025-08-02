@echo off
cls
chcp 65001 >nul
pwsh -ExecutionPolicy Bypass -File "%~dp0..\PS1\run_clear_data_work.ps1"

