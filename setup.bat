@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
title Setup - Maintenance Costing Control Panel

echo.
echo  ╔══════════════════════════════════════════╗
echo  ║   INSTALLATION - Premiere utilisation    ║
echo  ╚══════════════════════════════════════════╝
echo.
echo  Ce script va verifier et installer tout ce qu'il faut.
echo  Il ne sera a executer qu'une seule fois.
echo.
pause

:: ── 1. Verifier Python 3.12 ──────────────────────────────────────
echo.
echo  [1/3] Verification de Python 3.12...

set PYTHON=
py -3.12 --version >nul 2>&1 && set PYTHON=py -3.12
if not defined PYTHON (
    python --version 2>&1 | findstr /C:"Python 3.12" >nul 2>&1 && set PYTHON=python
)

if not defined PYTHON (
    echo.
    echo  ╔══════════════════════════════════════════════════════════╗
    echo  ║  ATTENTION : Python 3.12 est requis                     ║
    echo  ║                                                          ║
    echo  ║  Cette application necessite Python 3.12 specifiquement.║
    echo  ║  Python 3.13 et 3.14 ne sont pas encore compatibles.    ║
    echo  ║                                                          ║
    echo  ║  Telecharge Python 3.12 ici :                           ║
    echo  ║  https://www.python.org/downloads/release/python-31211/ ║
    echo  ║                                                          ║
    echo  ║  IMPORTANT : coche "Add Python to PATH"                 ║
    echo  ║  puis relance ce script.                                ║
    echo  ╚══════════════════════════════════════════════════════════╝
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('%PYTHON% --version 2^>^&1') do echo  %%v detecte
echo  OK

:: ── 2. Verifier Git ───────────────────────────────────────────────
echo.
echo  [2/3] Verification de Git...
where git >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Git n'est pas installe.
    echo  Installe Git depuis : https://git-scm.com/download/win
    echo  (options par defaut, puis relance ce script)
    echo.
    pause
    exit /b 1
)
echo  OK

:: ── 3. Installer les dependances ──────────────────────────────────
echo.
echo  [3/3] Installation des dependances Python...
%PYTHON% -m pip install --upgrade pip -q --no-warn-script-location
%PYTHON% -m pip install -r requirements.txt -q --no-warn-script-location
if errorlevel 1 (
    echo.
    echo  [ERREUR] L'installation des dependances a echoue.
    echo  Verifie ta connexion internet et relance ce script.
    pause
    exit /b 1
)
echo  OK

:: ── Terminer ──────────────────────────────────────────────────────
echo.
echo  ════════════════════════════════════════════
echo  Installation terminee avec succes !
echo.
echo  Pour lancer l'application : double-clique sur launch.bat
echo  ════════════════════════════════════════════
echo.
pause
