@echo off
chcp 65001 >nul
title Publication de SageBridge sur GitHub
cd /d "%~dp0"
rem Publie la version compilee (bin\) sur https://github.com/Gnaris/SageBridge/releases :
rem les SageBridge installes la detectent dans l'heure et se mettent a jour.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0outils\publier-version.ps1"
echo.
timeout /t 120
