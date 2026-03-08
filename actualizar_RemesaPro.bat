@echo off
setlocal enabledelayedexpansion
title Actualizador RemesaPro
color 0A

echo.
echo  ================================================
echo    Actualizador de RemesaPro
echo    Descarga la ultima version desde GitHub
echo  ================================================
echo.

REM --- Configuracion ---
set REPO=pollosev-sys/RemesaPro
set EXE_NAME=RemesaPro.exe
set TOKEN_FILE=%~dp0.github_token

REM --- Leer token desde archivo ---
if not exist "%TOKEN_FILE%" (
    echo  ERROR: No se encontro el archivo de token.
    echo.
    echo  Crea un archivo llamado .github_token en esta carpeta
    echo  con tu Personal Access Token de GitHub ^(solo lectura^).
    echo.
    echo  Instrucciones:
    echo    1. Ve a https://github.com/settings/tokens/new
    echo    2. Marca el scope: repo ^(read:release es suficiente^)
    echo    3. Genera el token y copialo
    echo    4. Ejecuta este comando en CMD:
    echo       echo ghp_TUTOKEN ^> "%TOKEN_FILE%"
    echo.
    pause
    exit /b 1
)

set /p GITHUB_TOKEN=<"%TOKEN_FILE%"
set GITHUB_TOKEN=%GITHUB_TOKEN: =%

if "%GITHUB_TOKEN%"=="" (
    echo  ERROR: El archivo .github_token esta vacio.
    pause
    exit /b 1
)

REM --- Buscar y descargar ultima version con PowerShell ---
echo  Buscando ultima version disponible...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference = 'Stop';" ^
    "try {" ^
    "  $headers = @{ Authorization = 'token %GITHUB_TOKEN%'; 'User-Agent' = 'RemesaUpdater' };" ^
    "  $release = Invoke-RestMethod -Uri 'https://api.github.com/repos/%REPO%/releases/latest' -Headers $headers;" ^
    "  $asset = $release.assets | Where-Object { $_.name -eq '%EXE_NAME%' };" ^
    "  if (-not $asset) { throw 'No se encontro RemesaPro.exe en el Release.' }" ^
    "  $sizeMB = [math]::Round($asset.size / 1MB, 1);" ^
    "  Write-Host ('  Version : ' + $release.name) -ForegroundColor Cyan;" ^
    "  Write-Host ('  Tamano  : ' + $sizeMB + ' MB');" ^
    "  Write-Host ('  Fecha   : ' + $release.published_at.Substring(0,10));" ^
    "  Write-Host '';" ^
    "  Write-Host '  Descargando...' -NoNewline;" ^
    "  $dlHeaders = @{ Authorization = 'token %GITHUB_TOKEN%'; Accept = 'application/octet-stream'; 'User-Agent' = 'RemesaUpdater' };" ^
    "  $outPath = '%~dp0%EXE_NAME%';" ^
    "  Invoke-WebRequest -Uri $asset.url -Headers $dlHeaders -OutFile $outPath;" ^
    "  Write-Host ' HECHO' -ForegroundColor Green;" ^
    "} catch {" ^
    "  Write-Host '';" ^
    "  Write-Host ('  ERROR: ' + $_.Exception.Message) -ForegroundColor Red;" ^
    "  exit 1;" ^
    "}"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo  ================================================
    echo    RemesaPro.exe actualizado correctamente^^!
    echo  ================================================
    echo.
    echo  Puedes ejecutar RemesaPro.exe ahora.
) else (
    echo.
    echo  No se pudo descargar la actualizacion.
    echo  Comprueba tu token y conexion a internet.
)

echo.
pause
