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
/// HU-13 (2.7, Plan 13 §2.7 — Unidad 4): Golden Capa A de VALIDACIONES en ambos períodos.
/// A2: gates de dominio por ASE × validación vs CACHÉ GOLDEN de las hojas de validación
/// (Q1 = Remuneracion 202607-1 Total.xlsx; Q2 valores = Remuneracion 202607-2 Total.xlsx).
/// A3: las hojas de validación siguen siendo fórmula en la SALIDA (mapa protegido extendido).
/// A4: canónicos no mutados (hash). A5/A7: prohibido comparar caché de salida vs golden
/// (OpenXML no recalcula) y la regresión sin snapshot = HU-12 puro (suite 128/128).
/// </summary>
public sealed class GoldenValidacionesCruzadasTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    [Fact]
    public void CapaA_Q1_GatesVsGoldenCache_SinErrores_Y_ValidacionesFormulaEnLaSalida()
    {
        var (resultado, leafs, salida) = EjecutarQ1();
        var hashQ1 = Sha256(Insumos.Plantilla);

        // A2-Q1: gates 2.7 de dominio contra el caché golden del canónico Q1 (= plantilla).
        var snapshotsGolden = new ValidacionOracleReader().LeerSnapshots(Insumos.Plantilla, Insumos.Periodo());
        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshotsGolden);
        Assert.Empty(errores);

        // A3-Q1: las hojas de validación siguen siendo fórmula en la salida (mapa 2.7).
        foreach (var (hoja, celda, _) in WorkbookLeafCellMapValidaciones.ProtegidasValidacionesParaPeriodo(1))
        {
            Assert.True(CeldaEsFormula(salida, hoja, celda), $"{hoja}!{celda} debió seguir siendo fórmula en la salida Q1.");
        }

        // A4: el canónico Q1 no mutó.
        Assert.Equal(hashQ1, Sha256(Insumos.Plantilla));
    }

    [Fact]
    public void CapaA_Q2_GatesVsGoldenCache_SinErrores_Y_ValidacionesFormulaEnLaSalida()
    {
        var (resultado, leafs, salida) = EjecutarWriterQ2();
        var hashPlantillaQ2 = Sha256(Insumos.PlantillaQ2);
        var hashGoldenQ2 = Sha256(Insumos.GoldenQ2);

        // A2-Q2: gates 2.7 de dominio contra el CACHÉ GOLDEN Q2 (valores reales post-Excel).
        var snapshotsGolden = new ValidacionOracleReader().LeerSnapshots(Insumos.GoldenQ2, Insumos.PeriodoQ2());
        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshotsGolden);
        Assert.Empty(errores);

        // A3-Q2: validaciones siguen siendo fórmula en la salida Q2 (extensión 2.7; el mapa Q2
        // DetRetriProtected ya cubierto por la suite HU-12).
        foreach (var (hoja, celda, _) in WorkbookLeafCellMapValidaciones.ProtegidasValidacionesParaPeriodo(2))
        {
            Assert.True(CeldaEsFormula(salida, hoja, celda), $"{hoja}!{celda} debió seguir siendo fórmula en la salida Q2.");
        }

        // A4: canónicos no mutados.
        Assert.Equal(hashPlantillaQ2, Sha256(Insumos.PlantillaQ2));
        Assert.Equal(hashGoldenQ2, Sha256(Insumos.GoldenQ2));
    }

    [Fact]
    public void CapaA_Q2_SalidaCacheCero_NoEsOracle_ProhibidoCompararVsGolden()
    {
        // A5 (honestidad): OpenXML NO recalcula → el caché de VALIDACION de la salida Q2 queda en
        // ceros (plantilla "8 agos"); compararlo contra el golden "cerraría en falso". El oráculo
        // honesto es el golden cache / dominio, no la caché de la salida.
        var (_, _, salida) = EjecutarWriterQ2();
        var cacheSalida = LeerCeldaNumerica(salida, "VALIDACION_ENEL", "H3");
        var golden = LeerCeldaNumerica(Insumos.GoldenQ2, "VALIDACION_ENEL", "H3");
        Assert.NotEqual(golden, cacheSalida);
    }

    // ── Helpers ───────────────────────────────────────────────────────────────────────────────

    private static (ResultadoRemuneracion Resultado, List<WorkbookLeafInputs> Leafs, string Salida) EjecutarQ1()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-golden-val-q1-" + Guid.NewGuid().ToString("N"));
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
            RutaPlantilla = Insumos.Plantilla,
            RutaSalida = salida
        });

        return (resultado.Resultado, resultado.Leafs.ToList(), salida);
    }

    private static (ResultadoRemuneracion Resultado, List<WorkbookLeafInputs> Leafs, string Salida) EjecutarWriterQ2()
    {
        var (datos, _, leafs) = AjustesSfTTests.LeerDatosQ2ConAjustes();

        foreach (var leaf in leafs)
        {
            leaf.DetRetriQ2 = new DetRetriQ2Inputs
            {
                Ase = leaf.Ase,
                TotalD104 = leaf.R1.TotalOportunoEsperadoPorAse
                    + leaf.R2.TotalOportunoEsperado
                    + leaf.R1.ExtemporaneoEsperadoPorAse
                    + leaf.R4.TotalReversionEsperada
                    + (leaf.AjustesSfT?.TotalAjustes ?? 0m)
            };
        }

        var resultado = new CalculoRemuneracion().CalcularConsolidado(Insumos.PeriodoQ2(), datos);
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-golden-val-q2-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.PeriodoQ2().NombreArchivo);

        new OpenXmlPlantillaWriter().GenerarWorkbook(Insumos.PlantillaQ2, salida, resultado, leafs);

        return (resultado, leafs, salida);
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
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var cell = ws.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        return cell?.CellFormula is not null;
    }

    private static decimal LeerCeldaNumerica(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var cell = ws.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda}.");
        if (cell.CellValue is null)
        {
            return 0m;
        }

        return decimal.TryParse(cell.CellValue.InnerText, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var value)
            ? value
            : 0m;
    }
}
