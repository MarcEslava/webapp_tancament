@echo off
setlocal EnableExtensions
REM ============================================================
REM  Prepara aquest equip com a SERVIDOR del panell de tancament.
REM
REM  Fa dues coses:
REM   1) Crea una tasca programada que arrenca el panell al iniciar
REM      sessio l'usuari actual.
REM   2) Obre el port al tallafocs nomes per a la xarxa local.
REM
REM  Us:  clic dret > "Executar com a administrador"
REM       InstalarServidor.bat elimina    (per desfer-ho tot)
REM
REM  IMPORTANT: la tasca s'executa DINS la sessio de l'usuari, no com a
REM  servei de sistema. Es a proposit: el Pas 2 automatitza Excel i Excel
REM  necessita una sessio interactiva de veritat, i la unitat Z: nomes
REM  existeix dins la sessio de l'usuari que la te mapada. Per tant
REM  l'equip servidor ha de quedar amb la sessio iniciada.
REM ============================================================
cd /d "%~dp0"

set "TASCA=Panell Tancament"
set "PORT=5099"
set "REGLA=Panell Tancament %PORT%"

REM -- Cal ser administrador per tocar el tallafocs
net session >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: cal executar aquest fitxer com a ADMINISTRADOR.
    echo        Clic dret sobre InstalarServidor.bat ^> "Executar com a administrador".
    echo.
    pause
    exit /b 1
)

if /i "%~1"=="elimina" goto :eliminar

REM -- 1) Preparar l'entorn abans de programar res: si aixo falla, millor
REM       saber-ho ara que no pas d'aqui a un reinici.
echo Preparant l'entorn virtual i les dependencies...
call "%~dp0Panell.bat" setup
if errorlevel 1 (
    echo.
    echo ERROR: no s'ha pogut preparar l'entorn. Revisa el missatge de dalt.
    pause
    exit /b 1
)

REM -- 2) Tasca programada a l'inici de sessio.
REM    /DELAY 1 minut: dona temps a que la xarxa i la unitat Z: estiguin
REM    llestes abans d'arrencar el panell.
echo Creant la tasca programada "%TASCA%"...
schtasks /Create /TN "%TASCA%" /TR "\"%~dp0Panell.bat\" servei" /SC ONLOGON /RU "%USERDOMAIN%\%USERNAME%" /DELAY 0001:00 /RL LIMITED /F
if errorlevel 1 (
    echo ERROR: no s'ha pogut crear la tasca programada.
    pause
    exit /b 1
)

REM -- 3) Tallafocs: nomes xarxa local, no perfil public.
echo Obrint el port %PORT% al tallafocs (nomes xarxa local)...
netsh advfirewall firewall delete rule name="%REGLA%" >nul 2>&1
netsh advfirewall firewall add rule name="%REGLA%" dir=in action=allow protocol=TCP localport=%PORT% profile=domain,private remoteip=localsubnet
if errorlevel 1 (
    echo AVIS: no s'ha pogut crear la regla del tallafocs. El panell funcionara
    echo       en aquest equip pero potser no des dels altres.
)

echo.
echo ============================================================
echo  Servidor preparat.
echo.
echo  El panell arrencara sol cada cop que aquest usuari inicii sessio.
echo  Per arrencar-lo ara mateix sense reiniciar:
echo      schtasks /Run /TN "%TASCA%"
echo.
echo  La resta d'usuaris hi entren des del navegador:
echo      http://%COMPUTERNAME%:%PORT%
echo ============================================================
echo.
pause
exit /b 0

:eliminar
echo Eliminant la tasca programada i la regla del tallafocs...
schtasks /Delete /TN "%TASCA%" /F
netsh advfirewall firewall delete rule name="%REGLA%"
echo.
echo Fet: aquest equip ja no arrenca el panell automaticament.
pause
exit /b 0
