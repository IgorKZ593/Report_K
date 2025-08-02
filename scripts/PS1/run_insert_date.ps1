# run_insert_date.ps1
$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new()

Write-Host "`n🕓 Запуск insert_date.py..." -ForegroundColor Cyan
python "F:\Python Projets\Report\insert_date.py"
Write-Host "`n✅ insert_date.py завершён." -ForegroundColor Yellow
