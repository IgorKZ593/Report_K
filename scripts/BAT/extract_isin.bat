@echo off
chcp 65001 >nul
cd /d "F:\Python Projets\Report"
python extract_isin.py %*
pause
