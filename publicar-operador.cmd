@echo off
REM ============================================================================
REM  Entrega a operador - Remuneracion UAESP
REM  Genera una carpeta autocontenida (win-x64), sin .pdb, con version en el
REM  nombre y comprimida en Entrega\. Doble clic y listo.
REM
REM  Uso:  publicar-operador.cmd [version]
REM        publicar-operador.cmd 1.0.1
REM ============================================================================
setlocal

set "VERSION=%~1"
if "%VERSION%"=="" set "VERSION=1.0.0"

pushd "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0publicar-operador.ps1" -Version "%VERSION%"
set "EXITCODE=%ERRORLEVEL%"
popd

if not "%EXITCODE%"=="0" (
  echo.
  echo [ERROR] La publicacion fallo. Revisa el mensaje de arriba.
)
echo.
pause
exit /b %EXITCODE%
