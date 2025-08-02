# run_template_creator.ps1
$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new()

Write-Host "`n🕓 Запуск template_creator.py..." -ForegroundColor Cyan
python "F:\Python Projets\Report\template_creator.py"
Write-Host "`n✅ template_creator.py завершён." -ForegroundColor Yellow
