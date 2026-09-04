@echo off
REM ============================================================================
REM  Lanzador del worker Excel-COM (el ancla del cierre) en modo SERVICIO.
REM  Vigila splitFiles\inbox (donde el paso 'enqueue' deja los .xlsx ya hechos),
REM  los refresca y los DEVUELVE a splitFiles\pasteFiles ya recalculados, para
REM  que aguas abajo (envio a labs, etc.) no cambie nada.
REM  Debe correr en la SESION INTERACTIVA del usuario: Excel COM necesita
REM  escritorio. Lo arranca la Tarea programada 'TancamentComWorker' al iniciar
REM  sesion (ver install_worker_task.ps1).
REM ============================================================================
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
if not exist "splitFiles\inbox" mkdir "splitFiles\inbox"
if not exist "splitFiles\pasteFiles" mkdir "splitFiles\pasteFiles"
"C:\Users\meslava\AppData\Local\Programs\Python\Python313\python.exe" com_worker.py --watch "splitFiles\inbox" --done "splitFiles\pasteFiles" --interval 5
endlocal
