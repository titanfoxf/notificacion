@echo off
title Publicar Visor GIS en GitHub Pages
cd /d "%~dp0"
echo ====================================================================
echo   ACTUALIZAR VISOR GIS EN GITHUB (GITHUB PAGES)
echo ====================================================================
echo.
git remote remove origin 2>nul
git remote add origin https://github.com/titanfoxf/notificacion.git
echo [*] Subiendo ultimos cambios a GitHub (rama main)...
git add .
git commit -m "update: actualizacion automatica del visor" 2>nul
git push -u origin main
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Hubo un problema al subir a GitHub.
    pause
    exit /b
)
echo.
echo ====================================================================
echo  EXITO: CAMBIOS SUBIDOS A GITHUB
echo  GitHub Actions compilara y desplegara la web automaticamente en:
echo  https://titanfoxf.github.io/notificacion/
echo ====================================================================
echo.
pause
