# run_insert_date.ps1
$projectPath = "F:\Python Projets\Report"
$venvActivate = "F:\Python Projets\Tick\.venv\Scripts\Activate.ps1"

# Активируем виртуальное окружение
. $venvActivate

# Переходим в папку проекта и запускаем скрипт
Set-Location $projectPath
python insert_date.py
