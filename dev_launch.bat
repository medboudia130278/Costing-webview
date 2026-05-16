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

:: ── 1. Trouver Python 3.12 (seule version compatible) ─────────────
set PYTHON=
py -3.12 --version >nul 2>&1 && set PYTHON=py -3.12
if not defined PYTHON (
    python --version 2>&1 | findstr /C:"Python 3.12" >nul 2>&1 && set PYTHON=python
)

if not defined PYTHON (
    echo  ╔══════════════════════════════════════════════════════════╗
    echo  ║  ERREUR : Python 3.12 introuvable                       ║
    echo  ║                                                          ║
    echo  ║  Cette application necessite Python 3.12.               ║
    echo  ║  Python 3.13 et 3.14 ne sont pas compatibles.           ║
    echo  ║                                                          ║
    echo  ║  Telecharge Python 3.12 ici :                           ║
    echo  ║  https://www.python.org/downloads/release/python-31211/ ║
    echo  ║                                                          ║
    echo  ║  Coche "Add Python to PATH" lors de l'installation,     ║
    echo  ║  puis relance ce fichier.                               ║
    echo  ╚══════════════════════════════════════════════════════════╝
    echo.
    pause
    exit /b 1
)

echo  Python 3.12 detecte.

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

:: ── 3. Installer les dependances ──────────────────────────────────
echo.
echo  Verification des dependances...
%PYTHON% -m pip install -r requirements.txt -q --no-warn-script-location 2>nul
if errorlevel 1 (
    echo  [AVERTISSEMENT] Certaines dependances n'ont pas pu etre installees.
)

:: ── 4. Lancement ──────────────────────────────────────────────────
echo  Demarrage de l'application...
echo.
%PYTHON% main.py

if errorlevel 1 (
    echo.
    echo  [ERREUR] L'application s'est arretee avec une erreur.
    echo  Consulte le fichier last_error.log pour plus de details.
    echo.
    pause
)
