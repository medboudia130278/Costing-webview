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

:: ── 1. Verifier Python ────────────────────────────────────────────
echo.
echo  [1/3] Verification de Python...
set PYTHON=
where py >nul 2>&1 && set PYTHON=py
if not defined PYTHON (
    where python >nul 2>&1 && set PYTHON=python
)

if not defined PYTHON (
    echo.
    echo  Python n'est pas installe.
    echo.
    echo  Installe Python 3.12 ou superieur depuis :
    echo  https://www.python.org/downloads/
    echo.
    echo  IMPORTANT : coche "Add Python to PATH" lors de l'installation,
    echo  puis relance ce script.
    echo.
    pause
    exit /b 1
)

%PYTHON% -c "import sys; v=sys.version_info; print(f'  Python {v.major}.{v.minor}.{v.micro} detecte')"
%PYTHON% -c "import sys; exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
if errorlevel 1 (
    echo  [ERREUR] Python 3.10 ou superieur requis. Mets Python a jour.
    pause
    exit /b 1
)
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
