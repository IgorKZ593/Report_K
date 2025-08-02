# run_name_clients.ps1
$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new()

Write-Host "`n🕓 Запуск name_clients.py..." -ForegroundColor Cyan
python "F:\Python Projets\Report\name_clients.py"
Write-Host "`n✅ name_clients.py завершён." -ForegroundColor Yellow
