@echo off
title Visor GIS Notificaciones - Subsector 1A-4
cd /d "%~dp0"
echo ========================================================
echo   INICIANDO VISOR GIS NOTIFICACIONES PROMUVI 1A-4
echo ========================================================
echo Abriendo servidor local en http://localhost:5174/ ...
start "" "http://localhost:5174/"
call npm run dev
pause
