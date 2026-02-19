@echo off
chcp 65001 >nul
title Convertitore DOCX in PDF
echo.
echo  =========================================
echo   Convertitore DOCX - PDF
echo  =========================================
echo.
echo  Avvio conversione di tutti i file .docx...
echo.

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%\converti-docx-in-pdf.ps1" "%SCRIPT_DIR%"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  Se lo script non funziona, assicurati di avere
    echo  Microsoft Word installato sul tuo PC.
    echo.
)
pause
