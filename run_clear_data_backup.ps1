# run_clear_data_backup.ps1
Write-Host "`n🧹 Очистка папки Data_Backup..." -ForegroundColor Cyan
python "F:\Python Projets\Report\clear_data_backup.py"
Write-Host "`n✅ Завершено. Нажмите любую клавишу для выхода..." -ForegroundColor Green
[void][System.Console]::ReadKey($true)
