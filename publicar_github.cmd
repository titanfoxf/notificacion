@echo off
title Publicar Visor GIS en GitHub Pages
cd /d "%~dp0"
echo ====================================================================
echo   PUBLICACION DEL VISOR GIS EN GITHUB (GITHUB PAGES)
echo ====================================================================
echo.
echo Verificando remotos actuales...
git remote -v
echo.
set /p REPO_URL="Pegue la URL de su repositorio GitHub (ej: https://github.com/usuario/mi-visor-gis.git): "
if "%REPO_URL%"=="" (
    echo.
    echo No se ingreso ninguna URL. Operacion cancelada.
    pause
    exit /b
)
git remote remove origin 2>nul
git remote add origin %REPO_URL%
echo.
echo [*] Subiendo codigo a GitHub (rama main)...
git push -u origin main
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Hubo un problema al subir a GitHub. Verifique sus credenciales o permisos en el repositorio.
    pause
    exit /b
)
echo.
echo ====================================================================
echo  EXITO: CODIGO PUBLICADO EN GITHUB
echo ====================================================================
echo.
echo Para que se active como PAGINA WEB (GitHub Pages):
echo 1. Abra su repositorio en GitHub.
echo 2. Vaya a: Settings -> Pages
echo 3. En "Build and deployment" -> "Source", elija: "GitHub Actions"
echo 4. Su visor web estara en linea en segundos!
echo ====================================================================
echo.
pause
