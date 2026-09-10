using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-10 (2.4, plan §2.7): Golden Capa A extendida a 2.4 con insumos Q1 reales.
/// A1: celdas BCE escritas en la SALIDA (D3:E7 en valores) vs mismas celdas leaf del golden ±0.5.
/// A2: visibles de DOMINIO por ASE (BCE = fuente con asignación T0; F=D+E; H≈F; sumas fila 11)
///     vs caché golden de filas 3–7 + fila 11.
/// A3: F3:F7/H3:H7/I3:I7 + filas 9/10/11/12/13 + bloque 18–24 + CONSOLIDADO J/K/M + refs
///     DetRetri/DetValiRetri siguen siendo fórmula en la salida.
/// A4: SHA256 plantilla origen igual antes/después.
/// A5: PROHIBIDO comparar caché de la SALIDA vs golden; PROHIBIDO usar CONSOLIDADO J/K/M como
///     oráculo de celdas D/E de BCE.
/// </summary>
public sealed class GoldenBalanceScTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;
    private const string HojaBce = WorkbookLeafCellMapBalanceSc.HojaBce;

    /// <summary>
    /// Golden Q1 filas 3–7 (caché template, T0-0.5): D=Contribución(+), E=Subsidio(−), F=D+E, H.
    /// </summary>
    private static readonly (int Ase, decimal D, decimal E, decimal F, decimal H)[] Golden =
    [
        (1, 3256235169.97m, -1223871491.13m, 2032363678.84m, 2032363679m),
        (2, 1946491828.39m, -6958007481.98m, -5011515653.59m, -5011515654m),
        (3, 1936109874.45m, -3132932285.76m, -1196822411.31m, -1196822411m),
        (4, 474156195.46m, -438200379.42m, 35955816.04m, 35955816m),
        (5, 1266769543.80m, -2477195172.11m, -1210425628.31m, -1210425628m)
    ];

    [Fact]
    public void ProcesadorPeriodo_CapaA2_4_BceContraGolden()
    {
        var insumos = Insumos.Plantilla;
        Assert.True(File.Exists(insumos), $"Falta plantilla: {insumos}");
        var origenHashAntes = Sha256(insumos);
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-hu10-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator());

        var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.Periodo(),
            CarpetaPeriodo = Insumos.CarpetaPeriodo,
            RutaPlantilla = insumos,
            RutaSalida = salida
        });

        Assert.True(File.Exists(salida));

        // Los 5 leafs traen su fila de balance SC (asignación T0; H = D2(b) null).
        foreach (var leaf in resultado.Leafs)
        {
            Assert.NotNull(leaf.BalanceSc);
            Assert.Single(leaf.BalanceSc!.Ases);
            Assert.Equal(leaf.Ase.Id, leaf.BalanceSc.Ases[0].Ase.Id);
            Assert.Null(leaf.BalanceSc.Ases[0].Sistema); // D2(b)
            Assert.InRange(
                leaf.BalanceSc.Ases[0].TotalBsc - leaf.BalanceSc.Ases[0].TotalFuente,
                -Tolerancia, Tolerancia); // BCE = fuente (gate D5-i)
        }

        // ---- A1: celdas BCE escritas en la SALIDA (D/E en valores) vs mismas celdas leaf del golden ----
        foreach (var (aseId, d, e, _, _) in Golden)
        {
            var celdaD = $"D{aseId + 2}";
            var celdaE = $"E{aseId + 2}";
            var escritoD = LeerCeldaNumerica(salida, HojaBce, celdaD);
            var escritoE = LeerCeldaNumerica(salida, HojaBce, celdaE);
            var goldenD = LeerCeldaNumerica(insumos, HojaBce, celdaD);
            var goldenE = LeerCeldaNumerica(insumos, HojaBce, celdaE);

            Assert.InRange(escritoD - goldenD, -Tolerancia, Tolerancia);
            Assert.InRange(escritoE - goldenE, -Tolerancia, Tolerancia);
            Assert.InRange(escritoD - d, -Tolerancia, Tolerancia);
            Assert.InRange(escritoE - e, -Tolerancia, Tolerancia);
            Assert.False(CeldaEsFormula(salida, HojaBce, celdaD), $"{HojaBce}!{celdaD} debió ser valor en la salida.");
            Assert.False(CeldaEsFormula(salida, HojaBce, celdaE), $"{HojaBce}!{celdaE} debió ser valor en la salida.");
            Assert.False(CeldaEsFormula(insumos, HojaBce, celdaD), $"Golden {HojaBce}!{celdaD} debía ser valor.");
        }

        // Columna C (id ASE) ya venía en el template y se conserva (valor 1..5).
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            Assert.InRange(LeerCeldaNumerica(salida, HojaBce, $"C{aseId + 2}") - aseId, -Tolerancia, Tolerancia);
        }

        // ---- A2: visibles de DOMINIO vs caché golden (filas 3–7 + fila 11) ----
        foreach (var (aseId, d, e, f, h) in Golden)
        {
            var fila = resultado.Leafs.Single(l => l.Ase.Id == aseId).BalanceSc!.Ases[0];

            // BCE = fuente con asignación T0 (Contribucion→D, Subsidio→E).
            Assert.InRange(fila.Contribucion - d, -Tolerancia, Tolerancia);
            Assert.InRange(fila.Subsidio - e, -Tolerancia, Tolerancia);
            // F = D+E (aritmética de dominio) == caché golden F.
            Assert.InRange(fila.TotalBsc - f, -Tolerancia, Tolerancia);
            // H≈F con regla T0-0.4 (redondeo a entero) == caché golden H.
            Assert.InRange(Math.Round(fila.TotalBsc, MidpointRounding.AwayFromZero) - h, -Tolerancia, Tolerancia);
        }

        // Sumas fila 11 en dominio vs caché golden (D11/E11/F11).
        var sumaD = Golden.Sum(g => g.D);
        var sumaE = Golden.Sum(g => g.E);
        var sumaF = Golden.Sum(g => g.F);
        Assert.InRange(sumaD - LeerCeldaNumerica(insumos, HojaBce, "D11"), -Tolerancia, Tolerancia);
        Assert.InRange(sumaE - LeerCeldaNumerica(insumos, HojaBce, "E11"), -Tolerancia, Tolerancia);
        Assert.InRange(sumaF - LeerCeldaNumerica(insumos, HojaBce, "F11"), -Tolerancia, Tolerancia);

        // ---- A3: fórmulas protegidas 2.4 siguen siendo fórmula en la salida ----
        foreach (var (hoja, celda, _) in WorkbookLeafCellMapBalanceSc.Protegidas)
        {
            Assert.True(CeldaEsFormula(salida, hoja, celda), $"{hoja}!{celda} debió seguir siendo fórmula.");
        }

        // ---- A4: plantilla origen intacta ----
        Assert.Equal(origenHashAntes, Sha256(insumos));

        // ---- A5: este test NO compara caché de la SALIDA contra golden (OpenXML no recalcula).
        // Tampoco usa CONSOLIDADO J9 como oráculo de D/E: J9 = BCE!F3 (total) — difiere de D3/E3.
        var j9 = LeerCeldaNumerica(insumos, "CONSOLIDADO_TOTAL RECAUDO", "J9");
        Assert.InRange(j9 - LeerCeldaNumerica(insumos, HojaBce, "F3"), -Tolerancia, Tolerancia);
        Assert.NotEqual(j9, LeerCeldaNumerica(insumos, HojaBce, "D3"));
        Assert.NotEqual(j9, LeerCeldaNumerica(insumos, HojaBce, "E3"));
    }

    private static string Sha256(string ruta)
    {
        using var stream = File.OpenRead(ruta);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static bool CeldaEsFormula(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var workbookXml = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook null");
        var sheet = workbookXml.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var sheetId = sheet.Id?.Value ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Id.");
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheetId)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var cell = ws.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        return cell?.CellFormula is not null;
    }

    private static decimal LeerCeldaNumerica(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var workbookXml = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook null");
        var sheet = workbookXml.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var sheetId = sheet.Id?.Value ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Id.");
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheetId)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var cell = ws.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda}.");

        if (cell.CellValue is null)
        {
            return 0m;
        }

        return decimal.TryParse(cell.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0m;
    }
}
