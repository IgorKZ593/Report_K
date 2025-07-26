# run_template_creator.ps1
# Запускает скрипт создания Excel-шаблона с помощью Python

Write-Host "`n🔧 Запуск template_creator.py..." -ForegroundColor Cyan

# Указываем полный путь к Python-скрипту
$scriptPath = "F:\Python Projets\Report\template_creator.py"

# Запуск
python "$scriptPath"

Write-Host "`n✅ Скрипт завершён. Нажмите любую клавишу для выхода..." -ForegroundColor Green
[void][System.Console]::ReadKey($true)
