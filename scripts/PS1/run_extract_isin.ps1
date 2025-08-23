
# robust runner for extract_isin.py (PowerShell 5+/7+)
$OutputEncoding = [Console]::OutputEncoding = [Text.UTF8Encoding]::new()

Write-Host "`n🚀 Запуск extract_isin..." -ForegroundColor Cyan

# Репозиторий: корень = два уровня вверх от scripts/PS1
$repoRoot  = Resolve-Path "$PSScriptRoot\..\.."
$python    = $null
$script    = Join-Path $repoRoot "extract_isin.py"

# Проверки наличия
if (-not (Test-Path $script)) {
  Write-Host "❌ Не найден файл: $script" -ForegroundColor Red
  exit 1
}

# Определяем Python интерпретатор
# 1) если активирован venv — достаточно 'python'
# 2) иначе пробуем py -3.10, затем просто python
function Test-Exe($cmd) { & $cmd --version *> $null; if ($LASTEXITCODE -eq 0) { return $true } return $false }

if (Test-Exe "python")      { $python = "python" }
elseif (Test-Exe "py -3.10") { $python = "py -3.10" }
elseif (Test-Exe "py -3.11") { $python = "py -3.11" }
elseif (Test-Exe "py")       { $python = "py" }
else {
  Write-Host "❌ Python не найден. Установи Python или активируй venv." -ForegroundColor Red
  exit 1
}

# Пробрасываем все аргументы скрипту (например, --yes)
Write-Host "▶ Интерпретатор: $python" -ForegroundColor DarkGray
Write-Host "▶ Скрипт:        $script" -ForegroundColor DarkGray
Write-Host "▶ Аргументы:     $args"   -ForegroundColor DarkGray

& $python $script @args
$code = $LASTEXITCODE

if ($code -eq 0) {
  Write-Host "`n✅ extract_isin завершён успешно." -ForegroundColor Green
} else {
  Write-Host "`n❌ extract_isin завершился с кодом: $code" -ForegroundColor Red
}
exit $code
