# install_task.ps1 -- Регистрация ежедневной задачи «Мои чеки ФНС»
# Запускать от имени администратора (install_task.bat делает это автоматически).

$TaskName    = "MoiCheki - Синхронизация чеков ФНС"
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Definition
$RunnerBat   = Join-Path $ScriptDir "sync_runner.bat"
$RunAt       = "20:00"

Write-Host ""
Write-Host "============================================================"
Write-Host " Установка задачи Windows Task Scheduler"
Write-Host " Задача : $TaskName"
Write-Host " Скрипт : $RunnerBat"
Write-Host " Время  : ежедневно в $RunAt"
Write-Host "============================================================"
Write-Host ""

if (-not (Test-Path $RunnerBat)) {
    Write-Host "ОШИБКА: Файл не найден: $RunnerBat" -ForegroundColor Red
    exit 1
}

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Найдена существующая задача -- удаляю..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

$Action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument ("/c " + '"' + $RunnerBat + '"') `
    -WorkingDirectory $ScriptDir

$Trigger = New-ScheduledTaskTrigger -Daily -At $RunAt

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -DontStopOnIdleEnd `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -MultipleInstances IgnoreNew

try {
    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $Action `
        -Trigger $Trigger `
        -Settings $Settings `
        -RunLevel Highest `
        -Force | Out-Null

    Write-Host "Задача успешно зарегистрирована!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Следующий запуск: завтра в $RunAt"
    Write-Host "Лог: $env:USERPROFILE\.moi_cheki_exporter\sync.log"
    Write-Host ""
    Write-Host "Тестовый запуск прямо сейчас:"
    Write-Host "  Start-ScheduledTask -TaskName MoiCheki"
    Write-Host ""
} catch {
    Write-Host "ОШИБКА при регистрации задачи: $_" -ForegroundColor Red
    exit 1
}

$task = Get-ScheduledTask -TaskName $TaskName
Write-Host "Статус задачи: $($task.State)"
Write-Host "============================================================"
Write-Host ""
