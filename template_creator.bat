@echo off
cls
chcp 65001 >nul
pwsh -ExecutionPolicy Bypass -File "%~dp0run_template_creator.ps1"
