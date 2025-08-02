chcp 65001 > $null
Write-Host "`n👤 Извлечение имени клиента запущено..." -ForegroundColor Cyan
python "F:\Python Projets\Report\name_clients.py"
Write-Host "`n✅ Готово" -ForegroundColor Yellow
[Console]::ReadKey($true) > $null
