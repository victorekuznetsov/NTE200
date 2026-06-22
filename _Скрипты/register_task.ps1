# NTE200 Daily Sync — регистрация задачи в Windows Task Scheduler
# Запускать один раз от имени Администратора:
#   powershell -ExecutionPolicy Bypass -File "C:\Projects\nte200\_Скрипты\register_task.ps1"

$PythonPath = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $PythonPath) {
    Write-Error "Python не найден. Установите Python и добавьте в PATH."
    exit 1
}

$ScriptPath = "C:\Projects\nte200\_Скрипты\daily_sync.py"
$WorkDir    = "C:\Projects\nte200"
$TaskName   = "NTE200 Daily Sync"

$action = New-ScheduledTaskAction `
    -Execute $PythonPath `
    -Argument "`"$ScriptPath`"" `
    -WorkingDirectory $WorkDir

$trigger = New-ScheduledTaskTrigger -Daily -At "12:00"

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -RunLevel Highest `
    -Force | Out-Null

Write-Host "Задача '$TaskName' зарегистрирована. Запуск ежедневно в 12:00."
Write-Host "Проверка: Get-ScheduledTask '$TaskName'"
Write-Host "Ручной запуск: Start-ScheduledTask '$TaskName'"
