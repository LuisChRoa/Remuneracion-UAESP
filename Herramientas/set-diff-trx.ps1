# Set-diff de rojos entre dos trx (Plan 62 H5). Uso: .\Herramientas\set-diff-trx.ps1 <baseline.trx> <nuevo.trx>
param([Parameter(Mandatory=$true)][string]$Baseline, [Parameter(Mandatory=$true)][string]$Nuevo)

function Leer($ruta) {
    [xml]$x = Get-Content $ruta -Raw
    $ns = New-Object System.Xml.XmlNamespaceManager($x.NameTable)
    $ns.AddNamespace('s', 'http://microsoft.com/schemas/VisualStudio/TeamTest/2010')
    $todos = $x.SelectNodes('//s:UnitTestResult', $ns)
    $rojos = $todos | Where-Object { $_.outcome -ne 'Passed' } | ForEach-Object { $_.testName }
    [pscustomobject]@{ Total = $todos.Count; Verdes = ($todos | Where-Object { $_.outcome -eq 'Passed' }).Count; Rojos = $rojos }
}

$b = Leer $Baseline
$n = Leer $Nuevo
Write-Output "baseline : $($b.Total) tests / $($b.Verdes) verdes / $($b.Rojos.Count) rojos"
Write-Output "nuevo    : $($n.Total) tests / $($n.Verdes) verdes / $($n.Rojos.Count) rojos"

$altas = $n.Rojos | Where-Object { $b.Rojos -notcontains $_ }
$bajas = $b.Rojos | Where-Object { $n.Rojos -notcontains $_ }
Write-Output "ALTAS (rojo nuevo) : $($altas.Count)"
$altas | ForEach-Object { "  + $_" }
Write-Output "BAJAS (rojo cerrado): $($bajas.Count)"
$bajas | ForEach-Object { "  - $_" }
