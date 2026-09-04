# =============================================================================
#  Registra el worker Excel-COM como Tarea programada que ARRANCA AL INICIAR
#  SESION, en modo INTERACTIVO (imprescindible: Excel COM necesita escritorio).
#  Ejecutar UNA vez, como el usuario que hace el cierre:
#       powershell -ExecutionPolicy Bypass -File install_worker_task.ps1
#  Desinstalar:
#       Unregister-ScheduledTask -TaskName 'TancamentComWorker' -Confirm:$false
# =============================================================================
$ErrorActionPreference = 'Stop'
$TaskName = 'TancamentComWorker'
$bat = Join-Path $PSScriptRoot 'run_com_worker.bat'

if (-not (Test-Path $bat)) { throw "No encuentro $bat" }

$action  = New-ScheduledTaskAction -Execute $bat
$trigger = New-ScheduledTaskTrigger -AtLogOn
# ExecutionTimeLimit 0 = sin limite (el watch corre indefinidamente).
# Reinicia hasta 3 veces si el proceso cae.
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
# Interactive = corre en la sesion con escritorio del usuario -> Excel funciona.
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal `
    -Description 'Worker Excel-COM del cierre: refresca los .xlsx que caen en splitFiles\inbox (el ancla).' `
    -Force | Out-Null

Write-Host "OK -> tarea '$TaskName' registrada (arranca al iniciar sesion)."
Write-Host "  Arrancar ahora:  schtasks /run /tn $TaskName"
Write-Host "  Parar:           schtasks /end /tn $TaskName"
Write-Host "  Estado:          schtasks /query /tn $TaskName /v /fo LIST"
