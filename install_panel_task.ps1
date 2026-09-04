# =============================================================================
#  Registra el PANELL del tancament com a Tasca programada que ARRENCA AL INICIAR
#  SESSIÓ (interactiu). Corre `Panell.bat servei` (sense obrir navegador).
#  Executar UNA vegada, com l'usuari que fa el tancament:
#       powershell -ExecutionPolicy Bypass -File install_panel_task.ps1
#  Desinstal·lar:
#       Unregister-ScheduledTask -TaskName 'TancamentPanell' -Confirm:$false
#
#  Nota: Panell.bat escolta per defecte a tota la xarxa al port 5099
#  (http://NOM-EQUIP:5099). Per fer-lo només local, posa PANELL_HOST=127.0.0.1
#  a l'entorn abans d'arrencar-lo.
# =============================================================================
$ErrorActionPreference = 'Stop'
$TaskName = 'TancamentPanell'
$bat = Join-Path $PSScriptRoot 'Panell.bat'
if (-not (Test-Path $bat)) { throw "No trobo $bat" }

$action  = New-ScheduledTaskAction -Execute $bat -Argument 'servei'
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal `
    -Description 'Panell del tancament (Flask) — arrenca al iniciar sessió (Panell.bat servei).' `
    -Force | Out-Null

Write-Host "OK -> tasca '$TaskName' registrada (arrenca al iniciar sessió)."
Write-Host "  Arrencar ara:  schtasks /run /tn $TaskName"
Write-Host "  Parar:         schtasks /end /tn $TaskName"
Write-Host "  Estat:         schtasks /query /tn $TaskName /v /fo LIST"
