using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-12 (2.6 ampliada, plan §2.7 — Unidad 4/PR5): matriz Capa A Q2 COMPLETA contra el golden
/// canónico (no se re-fija) y el caché de valores <c>Remuneracion 202607-2 Total.xlsx</c>.
///
/// Incluye A6 (TotalAse/GranTotal Q2 vs caché D104:D109 ±0.5) que CIERRA el pendiente HU-11
/// (recorte T0-0.6 levantado por el dispatch Q2 del reader + variante ASE5 V0.3) y A8/M1
/// (path Q2 del writer ejercitado con <c>esQuincena2 = true</c> + <c>ProtegidasBceParaPeriodo(true)</c>
/// contra hojas/celdas reales del canónico, incl. DetRetri2026072/DetValiRetri2026072).
///
/// Honestidad HU-06..HU-11 (A5): NUNCA se compara caché de fórmula de la salida vs golden
/// (OpenXML no recalcula); se comparan leafs/dominio contra el caché golden. Capa B (manual
/// Excel) queda como acción del usuario (§5.3).
/// </summary>
public sealed class GoldenDetRetriQ2Tests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    [Fact]
    public void CapaA_Q2_WriterEjercitado_A1CeldasEscritasVsGoldenCache()
    {
        // A8 (M1) + A1: el path Q2 del writer se ejercita por primera vez (esQuincena2 = true,
        // sufijo real 2026072 verificado contra hojas 36/37 del canónico) y las celdas escritas
        // (R1/R2/R4-Q2 por ASE + DetRetri-D) coinciden con las mismas celdas leaf del caché golden.
        var (resultado, leafs, salida) = EjecutarWriterQ2();

        Assert.True(File.Exists(salida), "El writer Q2 debe producir salida certificada.");

        foreach (var leaf in leafs)
        {
            foreach (var (celda, _) in WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(leaf.Ase.Id))
            {
                var escrita = LeerCeldaNumerica(salida, WorkbookLeafCellMapQ2.HojaR1, celda);
                var golden = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaR1, celda);
                Assert.InRange(escrita - golden, -Tolerancia, Tolerancia);
            }

            var r2 = WorkbookLeafCellMapQ2.ObtenerR2Q2Editables(leaf.Ase.Id);
            foreach (var celda in new[] { r2.Componente, r2.SubsCont, r2.Especiales })
            {
                Assert.InRange(
                    LeerCeldaNumerica(salida, WorkbookLeafCellMapQ2.HojaR2, celda)
                        - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaR2, celda),
                    -Tolerancia, Tolerancia);
            }

            var r4 = WorkbookLeafCellMapQ2.ObtenerR4Q2Editables(leaf.Ase.Id);
            foreach (var celda in new[] { r4.Total, r4.P })
            {
                Assert.InRange(
                    LeerCeldaNumerica(salida, WorkbookLeafCellMapQ2.HojaR4, celda)
                        - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaR4, celda),
                    -Tolerancia, Tolerancia);
            }
        }

        // A1b: DetRetri-D escrito (enteros V0.4) == golden DetRetri2026072 D9:D13 + D14.
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var celda = WorkbookLeafCellMapQ2.ObtenerDetRetriDestino(aseId);
            Assert.InRange(
                LeerCeldaNumerica(salida, WorkbookLeafCellMapQ2.HojaDetRetri, celda)
                    - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaDetRetri, celda),
                -Tolerancia, Tolerancia);
        }

        Assert.InRange(
            LeerCeldaNumerica(salida, WorkbookLeafCellMapQ2.HojaDetRetri, WorkbookLeafCellMapQ2.DetRetriTotal)
                - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaDetRetri, WorkbookLeafCellMapQ2.DetRetriTotal),
            -Tolerancia, Tolerancia);

        // A8: BCE parametrizado al período (2026072) matchea celdas reales del canónico — el
        // writer YA lo validó (ValidarFormulasProtegidasMultiAseQ2); se re-asegura explícito.
        foreach (var (hoja, celda, _) in ProtegidasBceParaPeriodo(true))
        {
            Assert.True(CeldaEsFormula(Insumos.PlantillaQ2, hoja, celda), $"{hoja}!{celda} (BCE Q2) debió ser fórmula en el canónico.");
        }
    }

    [Fact]
    public void CapaA_Q2_A2DetalleDominioVsGoldenCacheDetRetri()
    {
        // A2: Detalle de dominio (ROUND(D104:D108,0) vía DetRetriRounder) == caché golden
        // DetRetri2026072 D9:D13 ±0.5 en los 5 ASE (matcheo estricto por Id).
        var (_, leafs, _) = EjecutarWriterQ2();
        foreach (var leaf in leafs)
        {
            var detalle = leaf.DetRetriQ2 ?? throw new InvalidOperationException($"ASE {leaf.Ase.Id}: falta DetRetriQ2.");
            var golden = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaDetRetri, WorkbookLeafCellMapQ2.ObtenerDetRetriDestino(leaf.Ase.Id));
            Assert.InRange(detalle.Detalle - golden, -Tolerancia, Tolerancia);
        }
    }

    [Fact]
    public void CapaA_Q2_A3EstructuraFormulasProtegidasSiguenEnLaSalida()
    {
        // A3 (estructural): DetRetri/DetValiRetri Q2 (D23:D28/D32:D36 y D16:D21/D24:D29),
        // VALIDACION_*, GERENTES_*, INTERVENTORIA y ANT EXT-REV siguen siendo FÓRMULA en la salida
        // (el writer no las toca; Requirement 5).
        var (_, _, salida) = EjecutarWriterQ2();

        foreach (var (hoja, celda, _) in WorkbookLeafCellMapQ2.DetRetriProtected)
        {
            Assert.True(CeldaEsFormula(salida, hoja, celda), $"{hoja}!{celda} debió seguir siendo fórmula en la salida Q2.");
        }

        foreach (var (hoja, celda, _) in WorkbookLeafCellMapQ2.ProtegidasAdicionalesQ2)
        {
            Assert.True(CeldaEsFormula(salida, hoja, celda), $"{hoja}!{celda} debió seguir siendo fórmula en la salida Q2.");
        }
    }

    [Fact]
    public void CapaA_Q2_A4CanonicoNoMutadoY_A6TotalAseGranTotalVsGoldenCache()
    {
        // A4: el canónico no muta por el proceso (hash antes/después).
        var hashAntes = Sha256(Insumos.PlantillaQ2);
        var (_, leafs, _) = EjecutarWriterQ2();
        Assert.Equal(hashAntes, Sha256(Insumos.PlantillaQ2));

        // A6 (cierra el pendiente HU-11): TotalAse de DOMINIO por ASE (Σ visibles leaf Q2 +
        // AJUSTES-SF-T — la composición exacta de D104:D108) == caché golden D104:D108 ±0.5, y
        // GranTotal == D109 ±0.5. NUNCA agregados HU-02 como oráculo (A5: el TotOpt de la fuente
        // es otra cantidad que la visible del workbook; no se mezclan).
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var leaf = leafs.Single(l => l.Ase.Id == aseId);
            var totalD104 = leaf.R1.TotalOportunoEsperadoPorAse
                + leaf.R2.TotalOportunoEsperado
                + leaf.R1.ExtemporaneoEsperadoPorAse
                + leaf.R4.TotalReversionEsperada
                + (leaf.AjustesSfT?.TotalAjustes ?? 0m);

            var goldenD104 = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, $"D{103 + aseId}");
            Assert.InRange(totalD104 - goldenD104, -Tolerancia, Tolerancia);
        }

        var granTotal = leafs.Sum(l =>
            l.R1.TotalOportunoEsperadoPorAse
            + l.R2.TotalOportunoEsperado
            + l.R1.ExtemporaneoEsperadoPorAse
            + l.R4.TotalReversionEsperada
            + (l.AjustesSfT?.TotalAjustes ?? 0m));
        var goldenGranTotal = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, "D109");
        Assert.InRange(granTotal - goldenGranTotal, -Tolerancia, Tolerancia);
    }

    [Fact]
    public void CapaA_Q2_A5ProhibidoCacheDeFormulaDeLaSalidaVsGolden()
    {
        // A5 (honestidad): OpenXML NO recalcula → el caché D104 de la SALIDA queda en 0 (plantilla)
        // aunque el valor post-Excel esperado sea el golden. Compararlo contra el golden
        // "cerraría en falso"; está PROHIBIDO. El oráculo honesto es el dominio (leaf), que SÍ cierra.
        var (_, leafs, salida) = EjecutarWriterQ2();

        var cacheSalida = LeerCeldaNumerica(salida, WorkbookLeafCellMap.HojaConsolidado, "D104");
        var golden = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, "D104");
        Assert.NotEqual(golden, cacheSalida); // cache ≠ golden → usar cache como oráculo sería falso

        var dominio = leafs.Single(l => l.Ase.Id == 1);
        var totalD104 = dominio.R1.TotalOportunoEsperadoPorAse
            + dominio.R2.TotalOportunoEsperado
            + dominio.R1.ExtemporaneoEsperadoPorAse
            + dominio.R4.TotalReversionEsperada
            + (dominio.AjustesSfT?.TotalAjustes ?? 0m);
        Assert.InRange(totalD104 - golden, -Tolerancia, Tolerancia);
    }

    [Fact]
    public void CapaA_Q2_A7Q1Intacto_D104GoldenQ1SigueCejando()
    {
        // A7: Q1 intacto — el golden Q1 (D104 = 58210094820.50) sigue cerrando con los visibles
        // leaf Q1 (regresión ciega complementaria a la suite 105/105).
        var lector = new ExcelDataReaderRecaudoReader();
        var leafReader = new ExcelDataReaderWorkbookLeafInputReader();
        var periodo = Insumos.Periodo();
        var leafs = new List<WorkbookLeafInputs>();
        for (var i = 1; i <= 5; i++)
        {
            leafs.Add(leafReader.LeerLeafInputs(Insumos.Ase(i), periodo, Insumos.R1(i), Insumos.R2(i), Insumos.R4(i)));
        }

        var granTotalQ1 = leafs.Sum(l =>
            l.R1.TotalOportunoEsperadoPorAse
            + l.R2.TotalOportunoEsperado
            + l.R1.ExtemporaneoEsperadoPorAse
            + l.R4.TotalReversionEsperada);
        Assert.InRange(granTotalQ1 - 58210094820.50m, -Tolerancia, Tolerancia);
    }

    /// <summary>
    /// Ejecuta el writer Q2 end-to-end sobre el canónico (copia temporal) con los leafs Q2 reales
    /// de los 5 ASE (dispatch Q2 del reader + variante ASE5 V0.3 + DetRetriQ2 V0.4). Devuelve el
    /// resultado, los leafs y la ruta de salida. Es el corazón de A8/M1: el path Q2 del writer se
    /// ejercita por primera vez en CI (HU-11 lo dejó como recorte T0-0.6).
    /// </summary>
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
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-golden-q2-" + Guid.NewGuid().ToString("N"));
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

    /// <summary>
    /// Espejo de <c>OpenXmlPlantillaWriter.ProtegidasBceParaPeriodo</c> (no toca el mapa HU-10).
    /// </summary>
    private static IEnumerable<(string Hoja, string Celda, string[] Fragmentos)> ProtegidasBceParaPeriodo(bool esQuincena2)
    {
        if (!esQuincena2)
        {
            return WorkbookLeafCellMapBalanceSc.Protegidas;
        }

        return WorkbookLeafCellMapBalanceSc.Protegidas
            .Select(p => (
                p.Hoja.Replace("2026071", "2026072", StringComparison.Ordinal),
                p.Celda,
                p.Fragmentos.Select(f => f.Replace("2026071", "2026072", StringComparison.Ordinal)).ToArray()));
    }
}