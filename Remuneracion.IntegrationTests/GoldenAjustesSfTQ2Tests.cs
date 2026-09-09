using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-11 (2.5, plan §2.7 Capa A Q2): matriz golden contra el golden canónico Q2.
///
/// Golden canónico (T0-0.1): <c>Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx</c>
/// (coincide con la referencia en los valores cacheados de validación; SHA256 95825422…).
/// Golden de VALORES (caché): <c>Remuneracion 202607-2 Total.xlsx</c> (D85:D89, D104:D109,
/// AJUSTES-SF-T D47:D51). La otra plantilla = control, NUNCA oráculo (G6/D8).
///
/// RECORTE HONESTO T0-0.6 (Riesgo 5): el path completo del PROCESADOR Q2 no es certificable
/// end-to-end porque el R1 de ASE5-Q2 DIVERGE del Q1 (solo 2 filas Mes/Total vs 3 que exige el
/// reader leaf HU-07; F519/F521 son VALORES en el template Q2, no la fórmula F513+F498+F478
/// del mapa HU-07). El plan §0.2 prohíbe reescribir HU-07. Por eso la matriz Capa A se ejecuta
/// construyendo los leafs con los readers 2.5 directamente (A1/A2/A3/A4 en los 5 ASE) y el
/// TotalAse (A6) se certifica para ASE1-4 (R1/R2/R4 Q2 legibles); ASE5 queda declarado no
/// certificable en CI (Capa B manual residual) — NUNCA se inventa su valor.
///
/// NOTA HU-12 (2.6 ampliada): este recorte quedó LEVANTADO — el dispatch Q2 del reader
/// (mapa <see cref="WorkbookLeafCellMapQ2"/> con variante ASE5 de 2 filas V0.3) hace certificable
/// el procesador Q2 5/5 y la matriz Capa A COMPLETA (incl. A6 y M1) vive en
/// <see cref="GoldenDetRetriQ2Tests"/>. Este archivo conserva la certificación de la cadena 2.5.
///
/// Honestidad HU-06..HU-10 (A5): nunca se compara caché de fórmula de la salida vs golden
/// (OpenXML no recalcula); se comparan leafs/dominio contra el caché golden.
/// </summary>
public sealed class GoldenAjustesSfTQ2Tests
{
    private const decimal Tolerancia = 0.5m;

    /// <summary>
    /// Golden Q2 (T0-0.3, probado): TotalAjustes por ASE = D85:D89 = AJUSTES-SF-T D47:D51.
    /// </summary>
    private static readonly decimal[] GoldenAjustes = [973693.46m, 216025.77m, 104231.83m, 35954.44m, 0m];

    /// <summary>
    /// Golden Q2 CONSOLIDADO D104:D109 (caché de la referencia completada). Se documenta como
    /// referencia de Capa B: el TotalAse/GranTotal post-Excel NO se certifica en CI por el
    /// recorte T0-0.6 (R1-Q2 del template diverge del mapa HU-07).
    /// </summary>
    private static readonly decimal[] GoldenTotales = [17450228673.35m, 20516143969.65m, 15221896467.45m, 7179595395.67m, 12073344661.59m, 72441209167.71m];

[Fact]
    public void CapaA_Q2_Celdas2_5VsGolden_TotalAjustesDominio()
    {
        var (datos, _, leafs) = AjustesSfTTests.LeerDatosQ2ConAjustes();
        var resultado = new CalculoRemuneracion().CalcularConsolidado(Insumos.PeriodoQ2(), datos);

        // Gate previo: coherencia ajustes-vs-consolidado para los leafs construidos (ASE1-4).
        // El validador multi-ASE completo exige 5 leafs y ASE5 no es construible por el reader
        // leaf HU-07 (recorte T0-0.6) → aquí se valida la coherencia por ASE con matcheo
        // estricto (misma regla §2.5 punto 6) sin el gate de cantidad.
        foreach (var leaf in leafs)
        {
            var consolidado = resultado.Consolidados.Single(c => c.Ase.Id == leaf.Ase.Id);
            Assert.InRange(leaf.AjustesSfT!.TotalAjustes - consolidado.AjustesSfT, -Tolerancia, Tolerancia);
        }

        // A1 — operandos 2.5 leídos (SaldosNotas/RetribucionNegativa) vs caché golden ±0.5.
        // Los readers 2.5 SÍ certifican los 5 ASE (independiente del reader leaf HU-07).
        var ajustesLos5 = AjustesSfTTests.LeerAjustesQ2Los5();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var ajustes = ajustesLos5.Single(a => a.Ase.Id == aseId);
            foreach (var (celda, valor) in ajustes.SaldosNotas.Celdas)
            {
                var golden = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapAjustesSfT.HojaSaldosNotas, celda);
                Assert.InRange(valor - golden, -Tolerancia, Tolerancia);
            }

            foreach (var (celda, valor) in ajustes.RetribucionNegativa.Celdas)
            {
                var golden = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapAjustesSfT.HojaRetribucionNegativa, celda);
                Assert.InRange(valor - golden, -Tolerancia, Tolerancia);
            }
        }

        // A2 — TotalAjustes de DOMINIO (composición T0-0.3) vs caché golden D85:D89 y
        // AJUSTES-SF-T D47:D51 ±0.5 en los 5 ASE (nunca caché de fórmula de la salida).
        for (var i = 0; i < 5; i++)
        {
            var aseId = i + 1;
            var totalAjustes = ajustesLos5.Single(a => a.Ase.Id == aseId).TotalAjustes;

            var dConsolidado = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, $"D{85 + i}");
            var dAjustes = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapAjustesSfT.HojaAjustesSfT, $"D{47 + i}");

            Assert.InRange(totalAjustes - dConsolidado, -Tolerancia, Tolerancia);
            Assert.InRange(totalAjustes - dAjustes, -Tolerancia, Tolerancia);
            Assert.InRange(totalAjustes - GoldenAjustes[i], -Tolerancia, Tolerancia);
        }

        // A6 (RECORTE HONESTO T0-0.6 de HU-11 — LEVANTADO por HU-12): TotalAse/GranTotal Q2 se
        // certifican en <see cref="GoldenDetRetriQ2Tests.CapaA_Q2_A4CanonicoNoMutadoY_A6TotalAseGranTotalVsGoldenCache"/>
        // contra el caché golden D104:D109 (dispatch Q2 + variante ASE5). Aquí solo se re-afirma
        // que la cadena 2.5 (A1/A2) queda certificada.
    }

    [Fact]
    public void CapaA_Q2_EstructuraCadenaAjustesFormulaYHashPlantillaIntacto()
    {
        // A3 (estructural) + A4 sobre la PLANTILLA canónica Q2: la cadena AJUSTES-SF-T (mapa
        // 2.5) está en FÓRMULA en el template Q2 y el hash de la plantilla no muta por lectura.
        // RECORTE T0-0.6: el writer multi-ASE Q2 NO se certifica en CI porque valida el mapa
        // HU-07 (F519/F521 de R1-ASE5 como fórmulas) y el template Q2 los trae como VALORES
        // (layout R1-ASE5 divergente vs Q1). La escritura de la cadena 2.5 queda a Capa B manual.
        foreach (var (hoja, celda, _) in WorkbookLeafCellMapAjustesSfT.Protegidas)
        {
            Assert.True(CeldaEsFormula(Insumos.PlantillaQ2, hoja, celda), $"{hoja}!{celda} debió ser fórmula en la plantilla canónica Q2.");
        }

        // A4 — lectura no muta la plantilla canónica.
        var hashAntes = Sha256(Insumos.PlantillaQ2);
        var ajustes = AjustesSfTTests.LeerAjustesQ2Los5();
        Assert.Equal(5, ajustes.Count);
        Assert.Equal(hashAntes, Sha256(Insumos.PlantillaQ2));

        // A7 — Q1 intacto: golden Q1 D85:D89 = 0 y AjustesSfT = 0 (regresión ciega).
        for (var i = 0; i < 5; i++)
        {
            Assert.InRange(
                LeerCeldaNumerica(Insumos.Plantilla, WorkbookLeafCellMap.HojaConsolidado, $"D{85 + i}") - 0m,
                -Tolerancia, Tolerancia);
        }
    }

    [Fact]
    public void CapaA_Q2_LaSegundaPlantillaNuncaEsOraculo()
    {
        // Requirement 4 / A5: PROHIBIDO usar la segunda plantilla 202607-2 como oráculo de merge.
        // El test lo demuestra: la plantilla control tiene valores cacheados de validación
        // distintos (1/-1 vs 0) que NO corresponden al golden de valores.
        var d62Control = LeerCeldaNumerica(Insumos.PlantillaQ2Control, "REPORTE RECAUDO x BANCO", "D62");
        var d62Canonica = LeerCeldaNumerica(Insumos.PlantillaQ2, "REPORTE RECAUDO x BANCO", "D62");
        var d62Golden = LeerCeldaNumerica(Insumos.GoldenQ2, "REPORTE RECAUDO x BANCO", "D62");

        Assert.Equal(1m, d62Golden); // la referencia completada tiene D62 = 1 (C17 == C62)
        Assert.Equal(1m, d62Canonica);
        Assert.NotEqual(d62Canonica, d62Control); // la control no coincide → nunca oráculo
    }

    [Fact]
    public void CapaA_Q2_ProhibidoCompararCacheDeFormulaDeLaSalidaVsGolden()
    {
        // Requirement 4 / A5: OpenXML NO recalcula → el caché D85:D89 de la SALIDA quedaría en 0
        // (plantilla) aunque el valor post-Excel esperado sea el golden. Compararlo contra el
        // golden "cerraría en falso" solo por no recalcular; está PROHIBIDO. El test demuestra
        // que el caché de la plantilla (lo que OpenXML conservaría) NO es el oráculo: difiere
        // del golden, y el dominio (leaf 2.5) SÍ cierra contra el golden.
        var cachePlantilla = LeerCeldaNumerica(Insumos.PlantillaQ2, WorkbookLeafCellMap.HojaConsolidado, "D85");
        var golden = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, "D85");
        Assert.NotEqual(golden, cachePlantilla); // cache ≠ golden → usar cache como oráculo sería falso

        // El oráculo honesto es el dominio (reader 2.5), que SÍ cierra contra el golden.
        var dominio = AjustesSfTTests.LeerAjustesQ2Los5().Single(a => a.Ase.Id == 1).TotalAjustes;
        Assert.InRange(dominio - golden, -Tolerancia, Tolerancia);
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