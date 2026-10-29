<#
    Entrega a operador - Remuneracion UAESP
    ---------------------------------------
    Publica la app autocontenida (win-x64), la deja sin simbolos de depuracion
    y la comprime con la version en el nombre.

    Salida:  Entrega\Remuneracion-<version>-win-x64.zip
             Entrega\Remuneracion-<version>-win-x64\   (carpeta lista para copiar)

    El operador descomprime en una carpeta ESCRIBIBLE y FUERA de OneDrive
    (los logs se escriben junto al ejecutable).
#>
[CmdletBinding()]
param(
    [string]$Version = "1.0.0"
)

$ErrorActionPreference = "Stop"
$raiz = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $raiz

$proyecto = Join-Path $raiz "Remuneracion.WinForms\Remuneracion.WinForms.csproj"
$perfil = "Operador-win-x64"
$salidaPublicacion = Join-Path $raiz "Remuneracion.WinForms\bin\Release\net10.0-windows\publish\win-x64"
$entrega = Join-Path $raiz "Entrega"
$carpetaEntrega = Join-Path $entrega "Remuneracion-$Version-win-x64"
$zip = Join-Path $entrega "Remuneracion-$Version-win-x64.zip"

Write-Host "=== Publicando Remuneracion $Version (win-x64, autocontenida) ===" -ForegroundColor Cyan
& dotnet publish $proyecto -p:PublishProfile=$perfil -p:Version=$Version -p:FileVersion=$Version.0 -v:m
if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish fallo con codigo $LASTEXITCODE"
}

if (-not (Test-Path $salidaPublicacion)) {
    throw "No se encontro la carpeta publicada: $salidaPublicacion"
}

Write-Host "=== Limpiando simbolos y restos de pruebas ===" -ForegroundColor Cyan
Get-ChildItem $salidaPublicacion -Recurse -Include *.pdb -File | Remove-Item -Force
Get-ChildItem $salidaPublicacion -Filter "remuneracion_log_*.txt" -File | Remove-Item -Force

$exe = Join-Path $salidaPublicacion "Remuneracion.WinForms.exe"
if (-not (Test-Path $exe)) { throw "No se encontro el ejecutable en $salidaPublicacion" }

$archivos = Get-ChildItem $salidaPublicacion -Recurse -File
$mb = [math]::Round((($archivos | Measure-Object Length -Sum).Sum / 1MB), 1)
Write-Host ("    {0} archivos | {1} MB | version {2}" -f $archivos.Count, $mb, (Get-Item $exe).VersionInfo.FileVersion)

Write-Host "=== Preparando carpeta de entrega ===" -ForegroundColor Cyan
if (Test-Path $carpetaEntrega) { Remove-Item $carpetaEntrega -Recurse -Force }
New-Item -ItemType Directory -Path $carpetaEntrega -Force | Out-Null
Copy-Item (Join-Path $salidaPublicacion "*") $carpetaEntrega -Recurse -Force

$leeme = @"
Remuneracion UAESP - version $Version
=====================================

COMO USARLA
  1. Copie esta carpeta completa a una ruta ESCRIBIBLE y FUERA de OneDrive.
     Ejemplo: C:\FN\Remuneracion
  2. Abra Remuneracion.WinForms.exe
  3. Cierre los archivos de Excel antes de ejecutar el proceso.

NO SEPARE EL EJECUTABLE DE LA CARPETA
  El .exe necesita los archivos que estan a su lado (runtime de .NET incluido).

REQUISITOS
  Windows x64. No hace falta instalar .NET ni nada adicional.

DONDE QUEDAN LAS COSAS
  - El Excel de salida y el reporte de novedades: donde usted elija en la app.
  - Los registros de la corrida: remuneracion_log_<fecha>.txt, junto al .exe.
"@
$leeme | Set-Content (Join-Path $carpetaEntrega "LEEME.txt") -Encoding UTF8

Write-Host "=== Comprimiendo ===" -ForegroundColor Cyan
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $carpetaEntrega "*") -DestinationPath $zip -CompressionLevel Optimal
$zipMb = [math]::Round((Get-Item $zip).Length / 1MB, 1)

Write-Host ""
Write-Host "LISTO" -ForegroundColor Green
Write-Host ("  Carpeta : {0}" -f $carpetaEntrega)
Write-Host ("  ZIP     : {0}  ({1} MB)" -f $zip, $zipMb)
Write-Host ""
Write-Host "Entregue el ZIP. El operador lo descomprime en una carpeta escribible fuera de OneDrive." -ForegroundColor Yellow
