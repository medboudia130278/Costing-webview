@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
title Maintenance Costing Control Panel

echo.
echo  ╔══════════════════════════════════════════╗
echo  ║   Maintenance Costing Control Panel      ║
echo  ║   Created by: Mohamed BOUDIA             ║
echo  ╚══════════════════════════════════════════╝
echo.

:: ── 1. Trouver Python ─────────────────────────────────────────────
set PYTHON=
where py >nul 2>&1 && set PYTHON=py
if not defined PYTHON (
    where python >nul 2>&1 && set PYTHON=python
)
if not defined PYTHON (
    echo  [ERREUR] Python n'est pas installe ou introuvable.
    echo.
    echo  Installe Python depuis : https://www.python.org/downloads/
    echo  Coche bien "Add Python to PATH" lors de l'installation.
    echo.
    pause
    exit /b 1
)

:: Verifier que c'est Python 3
%PYTHON% -c "import sys; exit(0 if sys.version_info >= (3,10) else 1)" >nul 2>&1
if errorlevel 1 (
    echo  [ERREUR] Python 3.10 ou superieur est requis.
    echo  Version actuelle trop ancienne. Mets Python a jour.
    pause
    exit /b 1
)

:: ── 2. Mise a jour depuis GitHub ──────────────────────────────────
where git >nul 2>&1
if errorlevel 1 (
    echo  [INFO] Git non installe - mise a jour ignoree.
) else (
    git rev-parse --is-inside-work-tree >nul 2>&1
    if errorlevel 1 (
        echo  [INFO] Dossier non suivi par Git - mise a jour ignoree.
    ) else (
        echo  Verification des mises a jour...
        git fetch --quiet 2>nul
        for /f %%i in ('git rev-list HEAD..origin/main --count 2^>nul') do set BEHIND=%%i
        if "!BEHIND!" == "" set BEHIND=0
        if !BEHIND! GTR 0 (
            echo  !BEHIND! mise(s) a jour disponible(s) - telechargement...
            git pull --quiet
            echo  Mis a jour avec succes.
        ) else (
            echo  Deja a jour.
        )
    )
)

:: ── 3. Verifier et installer les dependances ──────────────────────
echo.
echo  Verification des dependances...
%PYTHON% -m pip install -r requirements.txt -q --no-warn-script-location 2>nul
if errorlevel 1 (
    echo  [AVERTISSEMENT] Certaines dependances n'ont pas pu etre installees.
    echo  L'application va quand meme demarrer...
)

:: ── 4. Lancement ──────────────────────────────────────────────────
echo  Demarrage de l'application...
echo.
%PYTHON% main.py

if errorlevel 1 (
    echo.
    echo  [ERREUR] L'application s'est arretee avec une erreur.
    echo  Consulte le fichier last_error.log pour plus de details.
    pause
)
