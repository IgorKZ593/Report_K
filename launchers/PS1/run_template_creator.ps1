chcp 65001 > $null
Write-Host "`n🔧 Запуск template_creator.py..." -ForegroundColor Cyan
python "F:\Python Projets\Report\template_creator.py"
Write-Host "`n✅ Скрипт завершён. Нажмите любую клавишу для выхода..." -ForegroundColor Green
[Console]::ReadKey($true) > $null
