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
/// HU-07 Req 7 / plan §2.7: Golden Capa A × 5 ASE con insumos Q1 reales.
/// Matriz: A1 leaf vs golden, A2 visibles de dominio vs golden + CONSOLIDADO, A3 fórmulas
/// protegidas, A4 hash plantilla, A5 prohibiciones, A6 D104:D108/D109 fórmula + aritmética.
/// </summary>
public sealed class GoldenMultiAseTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    [Fact]
    public void ProcesadorPeriodo_CapaAX5_LeafVisiblesFormulasHashYSumas()
    {
        var insumos = ObtenerInsumos();
        var origenHashAntes = Sha256(insumos);
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-hu07-" + Guid.NewGuid().ToString("N"));
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

        // ---- A1: leaf escritos en la SALIDA vs leaf del golden (celdas del mapa por ASE) ----
        foreach (var leaf in resultado.Leafs)
        {
            var editables = WorkbookLeafCellMapPorAse.ObtenerEditables(leaf.Ase.Id);
            foreach (var (hoja, celda, _) in editables)
            {
                var escrito = LeerCeldaNumerica(salida, hoja, celda);
                var golden = LeerCeldaNumerica(insumos, hoja, celda);
                Assert.InRange(escrito - golden, -Tolerancia, Tolerancia);
            }
        }

        // ---- A2: visibles de DOMINIO por ASE vs cache golden del bloque + CONSOLIDADO ----
        // Cache golden §2.1: (D TOT_OPT, D R2, D EXTEMP, D R4) por ASE.
        var visiblesGolden = new[]
        {
            (1, 16704332434.57m, 54216385.68m, 11673020m, -12054255.65m),
            (2, 12157780441.19m, 79400801.26m, 0m, -9889189.72m),
            (3, 10101988514.82m, 31111803.75m, 10198723.07m, -16103442.89m),
            (4, 10552409503.83m, 17236000.33m, 5419780m, -21288908.57m),
            (5, 8519310329.28m, 30774154.23m, 0m, -6421274.68m)
        };

        foreach (var leaf in resultado.Leafs.OrderBy(l => l.Ase.Id))
        {
            var (_, totOpt, r2, extemp, r4) = visiblesGolden[leaf.Ase.Id - 1];
            Assert.InRange(leaf.R1.TotalOportunoEsperadoPorAse - totOpt, -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R2.TotalOportunoEsperado - r2, -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R1.ExtemporaneoEsperadoPorAse - extemp, -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R4.TotalReversionEsperada - r4, -Tolerancia, Tolerancia);

            // Visibles de dominio vs CONSOLIDADO golden (filas por ASE).
            Assert.InRange(leaf.R1.TotalOportunoEsperadoPorAse - LeerCeldaNumerica(insumos, WorkbookLeafCellMap.HojaConsolidado, $"D{8 + leaf.Ase.Id}"), -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R2.TotalOportunoEsperado - LeerCeldaNumerica(insumos, WorkbookLeafCellMap.HojaConsolidado, $"D{27 + leaf.Ase.Id}"), -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R1.ExtemporaneoEsperadoPorAse - LeerCeldaNumerica(insumos, WorkbookLeafCellMap.HojaConsolidado, $"D{46 + leaf.Ase.Id}"), -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R4.TotalReversionEsperada - LeerCeldaNumerica(insumos, WorkbookLeafCellMap.HojaConsolidado, $"D{65 + leaf.Ase.Id}"), -Tolerancia, Tolerancia);
        }

        // ---- A3: fórmulas protegidas (mapa ampliado) siguen siendo fórmula en la salida ----
        foreach (var aseId in Enumerable.Range(1, 5))
        {
            foreach (var (hoja, celda, _) in WorkbookLeafCellMapPorAse.ProtectedFormulasPorAse[aseId])
            {
                Assert.True(CeldaEsFormula(salida, hoja, celda), $"{hoja}!{celda} (ASE {aseId}) debió seguir siendo fórmula.");
            }
        }

        // D106 es shared follower (texto vacío + SharedIndex) — debe seguir siendo fórmula.
        Assert.True(CeldaEsFormula(salida, WorkbookLeafCellMap.HojaConsolidado, "D106"), "D106 debió seguir siendo fórmula (shared follower).");

        // ---- Riesgo F178 (plan §7 / T0-0.2): R1!F178 (ASE2) y R1!F521 (ASE5) son celdas
        // VALOR 0 sin fórmula en el template y deben conservarse intactas en la salida ----
        foreach (var celda in new[] { "F178", "F521" })
        {
            Assert.Equal(0m, LeerCeldaNumerica(salida, WorkbookLeafCellMap.HojaR1, celda));
            Assert.False(CeldaEsFormula(salida, WorkbookLeafCellMap.HojaR1, celda),
                $"{WorkbookLeafCellMap.HojaR1}!{celda} debió conservarse como celda de valor (sin fórmula).");
        }

        // ---- A4: plantilla origen intacta ----
        Assert.Equal(origenHashAntes, Sha256(insumos));

        // ---- A5: este test NO compara cache de fórmula DE LA SALIDA contra golden;
        // prohibido TotOpt HU-02 vs visibles R1 (el dominio usa los visibles por bloque).
        Assert.NotEqual(19556118465.99m, resultado.Leafs[0].R1.TotalOportunoEsperadoPorAse);

        // ---- A6: D104:D108 son fórmulas de suma por fila y D109 = SUM(D104:D108);
        // aritmética de dominio Σ por ASE vs golden D104:D108 y D109 ±0.5 ----
        for (var i = 0; i < 5; i++)
        {
            var leaf = resultado.Leafs.First(l => l.Ase.Id == i + 1);
            var sumaDominio = leaf.R1.TotalOportunoEsperadoPorAse
                + leaf.R2.TotalOportunoEsperado
                + leaf.R1.ExtemporaneoEsperadoPorAse
                + leaf.R4.TotalReversionEsperada; // AjustesSfT = 0 en Q1
            var goldenFila = LeerCeldaNumerica(insumos, WorkbookLeafCellMap.HojaConsolidado, $"D{104 + i}");
            Assert.InRange(sumaDominio - goldenFila, -Tolerancia, Tolerancia);

            Assert.True(CeldaEsFormula(salida, WorkbookLeafCellMap.HojaConsolidado, $"D{104 + i}"), $"D{104 + i} debió seguir siendo fórmula.");
        }

        // A6 (honesto): D104:D108 = Σ visibles leaf por ASE (aritmética de dominio por bloque),
        // NO agregados HU-02 (prohibido A5). D109 = SUM(D104:D108) por estructura.
        var granTotalPostExcel = resultado.Leafs.Sum(l =>
            l.R1.TotalOportunoEsperadoPorAse
            + l.R2.TotalOportunoEsperado
            + l.R1.ExtemporaneoEsperadoPorAse
            + l.R4.TotalReversionEsperada);
        var goldenGranTotal = LeerCeldaNumerica(insumos, WorkbookLeafCellMap.HojaConsolidado, "D109");
        Assert.InRange(granTotalPostExcel - goldenGranTotal, -Tolerancia, Tolerancia);
        Assert.True(CeldaEsFormula(salida, WorkbookLeafCellMap.HojaConsolidado, "D109"), "D109 debió seguir siendo fórmula.");
    }

    private static string ObtenerInsumos()
    {
        Assert.True(File.Exists(Insumos.Plantilla), $"Falta plantilla: {Insumos.Plantilla}");
        return Insumos.Plantilla;
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

        if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString)
        {
            var shared = workbookPart.SharedStringTablePart?.SharedStringTable
                ?? throw new InvalidOperationException("SharedStringTable ausente.");
            var index = int.Parse(cell.CellValue.InnerText, CultureInfo.InvariantCulture);
            var texto = shared.Elements<SharedStringItem>().ElementAt(index).InnerText;
            return decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed) ? parsed : 0m;
        }

        return decimal.TryParse(cell.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0m;
    }
}
