@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_comparar_bases.ps1"
if errorlevel 1 (
    echo.
    echo O comparador encerrou com erro.
    pause
    exit /b %errorlevel%
)
