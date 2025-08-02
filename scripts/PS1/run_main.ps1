# main.ps1 — Главный управляющий скрипт

$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new()

Write-Host "`n📘 Запуск последовательности формирования отчёта..." -ForegroundColor Cyan

# Путь к папке со скриптами
$ps1Path = "$PSScriptRoot"

# 1️⃣ insert_date
Write-Host "`n1. ▶ insert_date.py" -ForegroundColor Green
& "$ps1Path\run_insert_date.ps1"

# 2️⃣ name_clients
Write-Host "`n2. ▶ name_clients.py" -ForegroundColor Green
& "$ps1Path\run_name_clients.ps1"

# 3️⃣ template_creator
Write-Host "`n3. ▶ template_creator.py" -ForegroundColor Green
& "$ps1Path\run_template_creator.ps1"

Write-Host "`n✅ Все модули успешно выполнены." -ForegroundColor Yellow
