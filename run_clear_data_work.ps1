# run_clear_data_work.ps1
Write-Host "`n🧹 Очистка папки Data_work..." -ForegroundColor Cyan
python "F:\Python Projets\Report\clear_data_work.py"
Write-Host "`n✅ Завершено. Нажмите любую клавишу для выхода..." -ForegroundColor Green
[void][System.Console]::ReadKey($true)
