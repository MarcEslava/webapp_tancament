@echo off
setlocal EnableExtensions
REM ============================================================
REM  Panell de tancament
REM  1) Comprova que hi ha Python (si no, l'instal.la amb winget)
REM  2) Crea l'entorn virtual la primera vegada (fora d'OneDrive)
REM  3) Instal.la les dependencies quan requirements.txt canvia
REM  4) Obre el navegador i arrenca el panell
REM  Us:  Panell.bat          (normal: arrenca i obre el navegador)
REM       Panell.bat servei   (arrenca sense obrir navegador -- autoarrencada)
REM       Panell.bat setup    (nomes prepara l'entorn, no arrenca)
REM
REM  El panell escolta a tota la xarxa: la resta d'usuaris no instal.len res,
REM  nomes obren http://NOM-EQUIP:5099 des del navegador.
REM ============================================================
cd /d "%~dp0"

REM El venv va a LOCALAPPDATA expressament: la carpeta del projecte es
REM dins d'OneDrive i no volem sincronitzar milers de fitxers del venv.
set "VENV_DIR=%LOCALAPPDATA%\WebAppTancament\venv"
set "VPY=%VENV_DIR%\Scripts\python.exe"
set "PYEXE="

REM -- 1) Buscar un Python que funcioni (launcher py primer, despres python)
py -3 --version >nul 2>&1 && set "PYEXE=py -3"
if not defined PYEXE python --version >nul 2>&1 && set "PYEXE=python"

if not defined PYEXE (
    echo Python no trobat. Provant d'instal.lar-lo amb winget...
    winget --version >nul 2>&1
    if errorlevel 1 (
        echo.
        echo ERROR: no hi ha ni Python ni winget en aquest equip.
        echo Instal.la Python manualment des de https://www.python.org/downloads/
        echo i torna a executar aquest panell.
        pause
        exit /b 1
    )
    winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
    REM El PATH d'aquesta finestra no es refresca sol: provem el launcher i la ruta tipica
    py -3 --version >nul 2>&1 && set "PYEXE=py -3"
    if not defined PYEXE if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
        set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    )
    if not defined PYEXE (
        echo.
        echo Python s'ha instal.lat pero aquesta finestra encara no el veu.
        echo Tanca aquesta finestra i torna a obrir el Panell.
        pause
        exit /b 1
    )
)

REM -- 2) Crear el venv si encara no existeix
if not exist "%VPY%" (
    echo Creant l'entorn virtual a "%VENV_DIR%"...
    %PYEXE% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo ERROR: no s'ha pogut crear l'entorn virtual.
        pause
        exit /b 1
    )
)

REM -- 3) Instal.lar dependencies nomes si requirements.txt ha canviat
fc /b requirements.txt "%VENV_DIR%\requirements.installed" >nul 2>&1
if errorlevel 1 (
    echo Instal.lant dependencies ^(nomes cal la primera vegada o si canvien^)...
    "%VPY%" -m pip install --upgrade pip
    "%VPY%" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: ha fallat la instal.lacio de dependencies. Revisa la connexio a internet.
        pause
        exit /b 1
    )
    copy /y requirements.txt "%VENV_DIR%\requirements.installed" >nul
)

if /i "%~1"=="setup" (
    echo Entorn preparat correctament.
    exit /b 0
)

REM -- 4) Arrencar el panell
REM En mode servei no obrim navegador ni deixem un "pause" que bloquejaria
REM l'autoarrencada esperant una tecla que ningu no prem mai.
if /i "%~1"=="servei" (
    "%VPY%" app.py
    exit /b %errorlevel%
)

start "" http://127.0.0.1:5099
"%VPY%" app.py
pause
