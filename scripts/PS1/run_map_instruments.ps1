# robust runner for map_instruments.py (PowerShell 5+/7+)
$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new()

Write-Host "`n🧭 Запуск map_instruments..." -ForegroundColor Cyan

# Репозиторий: корень = два уровня вверх от scripts/PS1
$repoRoot  = Resolve-Path "$PSScriptRoot\..\.."
$python    = $null
$script    = Join-Path $repoRoot "map_instruments.py"

# Проверки наличия
if (-not (Test-Path $script)) {
  Write-Host "❌ Не найден файл: $script" -ForegroundColor Red
  exit 1
}

# Определяем Python интерпретатор
# 1) если активирован venv — достаточно 'python'
# 2) иначе пробуем py -3.10/3.11, затем просто py
function Test-Exe($cmd) { & $cmd --version *> $null; if ($LASTEXITCODE -eq 0) { return $true } return $false }

if (Test-Exe "python")       { $python = "python" }
elseif (Test-Exe "py -3.10") { $python = "py -3.10" }
elseif (Test-Exe "py -3.11") { $python = "py -3.11" }
elseif (Test-Exe "py")       { $python = "py" }
else {
  Write-Host "❌ Python не найден. Установите Python или активируйте venv." -ForegroundColor Red
  exit 1
}

# Пробрасываем все аргументы скрипту (если понадобятся)
Write-Host "▶ Интерпретатор: $python" -ForegroundColor DarkGray
Write-Host "▶ Скрипт:        $script" -ForegroundColor DarkGray
Write-Host "▶ Аргументы:     $args"   -ForegroundColor DarkGray

& $python $script @args
$code = $LASTEXITCODE

if ($code -eq 0) {
  Write-Host "`n✅ map_instruments завершён успешно." -ForegroundColor Green
} else {
  Write-Host "`n❌ map_instruments завершился с кодом: $code" -ForegroundColor Red
}
exit $code
