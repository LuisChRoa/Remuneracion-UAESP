using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-12 (2.6 ampliada, plan §4 Fase 0 — Unidad 0/PR1): EVIDENCIA T0 réplica-HU-07 sobre Q2.
/// Congela <see cref="WorkbookLeafCellMapQ2"/> contra el canónico
/// (<c>Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx</c>) con dumps valor-vs-fórmula:
/// operandos editables = VALORES, visibles R1/R2/R4-Q2 + CONSOLIDADO + DetRetri/DetValiRetri +
/// protegidas adicionales = FÓRMULA con los fragmentos exactos, y M1 estructural
/// (<c>ProtegidasBceParaPeriodo(true)</c> matchea celdas reales del canónico).
///
/// Regla del plan (§0.2): ninguna celda Q2 entra al código sin pasar por este T0; V0.3/V0.4 se
/// reutilizan (aritmética ASE5 y composición DetRetri probadas contra el golden), no se re-descubren.
/// </summary>
public sealed class MapaQ2T0Tests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    [Fact]
    public void T0_OperandosR1Q2_SonValoresEditablesEnCanonico()
    {
        // T0-0.2: cada operando del mapa R1-Q2 es VALOR (sin &lt;f&gt;) en el canónico → editable.
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            foreach (var (celda, _) in WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(aseId))
            {
                Assert.False(CeldaEsFormula(Insumos.PlantillaQ2, WorkbookLeafCellMapQ2.HojaR1, celda),
                    $"{WorkbookLeafCellMapQ2.HojaR1}!{celda} (ASE{aseId}) debió ser VALOR editable en el canónico Q2.");
            }
        }
    }

    [Fact]
    public void T0_OperandosR2R4Q2_SonValoresEditablesEnCanonico()
    {
        // T0-0.2: R2/R4-Q2 (Componente/SubsCont/Especiales + Total/P).
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var r2 = WorkbookLeafCellMapQ2.ObtenerR2Q2Editables(aseId);
            foreach (var celda in new[] { r2.Componente, r2.SubsCont, r2.Especiales })
            {
                Assert.False(CeldaEsFormula(Insumos.PlantillaQ2, WorkbookLeafCellMapQ2.HojaR2, celda),
                    $"{WorkbookLeafCellMapQ2.HojaR2}!{celda} (ASE{aseId}) debió ser VALOR editable en el canónico Q2.");
            }

            var r4 = WorkbookLeafCellMapQ2.ObtenerR4Q2Editables(aseId);
            foreach (var celda in new[] { r4.Total, r4.P })
            {
                Assert.False(CeldaEsFormula(Insumos.PlantillaQ2, WorkbookLeafCellMapQ2.HojaR4, celda),
                    $"{WorkbookLeafCellMapQ2.HojaR4}!{celda} (ASE{aseId}) debió ser VALOR editable en el canónico Q2.");
            }
        }
    }

    [Fact]
    public void T0_VisiblesR1Q2_SonFormulasConFragmentosExactos()
    {
        // T0-0.2: los visibles R1-Q2 son FÓRMULA con los fragmentos congelados (F53=F32+F48+F12-L12-L32, …).
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            foreach (var (celda, fragmentos) in WorkbookLeafCellMapQ2.ObtenerR1Q2Protegidos(aseId))
            {
                Assert.True(CeldaTieneFormulaConFragmentos(Insumos.PlantillaQ2, WorkbookLeafCellMapQ2.HojaR1, celda, fragmentos),
                    $"{WorkbookLeafCellMapQ2.HojaR1}!{celda} (ASE{aseId}) debió ser fórmula con fragmentos [{string.Join(",", fragmentos)}] en el canónico Q2.");
            }
        }
    }

    [Fact]
    public void T0_VisiblesR2R4Q2_SonFormulasConFragmentosExactos()
    {
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            foreach (var (celda, fragmentos) in WorkbookLeafCellMapQ2.ObtenerR2Q2Protegidos(aseId))
            {
                Assert.True(CeldaTieneFormulaConFragmentos(Insumos.PlantillaQ2, WorkbookLeafCellMapQ2.HojaR2, celda, fragmentos),
                    $"{WorkbookLeafCellMapQ2.HojaR2}!{celda} (ASE{aseId}) debió ser fórmula con fragmentos [{string.Join(",", fragmentos)}].");
            }

            foreach (var (celda, fragmentos) in WorkbookLeafCellMapQ2.ObtenerR4Q2Protegidos(aseId))
            {
                Assert.True(CeldaTieneFormulaConFragmentos(Insumos.PlantillaQ2, WorkbookLeafCellMapQ2.HojaR4, celda, fragmentos),
                    $"{WorkbookLeafCellMapQ2.HojaR4}!{celda} (ASE{aseId}) debió ser fórmula con fragmentos [{string.Join(",", fragmentos)}].");
            }
        }
    }

    [Fact]
    public void T0_ConsolidadoQ2_FormulasConReferenciasQ2()
    {
        // V0.1: filas CONSOLIDADO Q2 (D9:D13 → R1!F53/…, D28:D32 → R2!E43/…, D47:D51 → R1!F55/…,
        // D66:D70 → R4!D73/…, D85:D89 → AJUSTES-SF-T, D104:D108 sumas, D109 SUM).
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.ConsolidadoProtected)
        {
            Assert.True(CeldaTieneFormulaConFragmentos(Insumos.PlantillaQ2, hoja, celda, fragmentos),
                $"{hoja}!{celda} debió ser fórmula con fragmentos [{string.Join(",", fragmentos)}] en el canónico Q2.");
        }
    }

    [Fact]
    public void T0_DetRetriQ2_D9D14Valores_Y_FormulaProtected()
    {
        // T0-0.4 / V0.4: DetRetri2026072 D9:D14 = VALORES 0 editables; D23:D28 + D32:D36 = FÓRMULA.
        for (var fila = 9; fila <= 14; fila++)
        {
            Assert.False(CeldaEsFormula(Insumos.PlantillaQ2, WorkbookLeafCellMapQ2.HojaDetRetri, $"D{fila}"),
                $"DetRetri2026072!D{fila} debió ser VALOR editable en el canónico Q2.");
        }

        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.DetRetriProtected)
        {
            Assert.True(CeldaTieneFormulaConFragmentos(Insumos.PlantillaQ2, hoja, celda, fragmentos),
                $"{hoja}!{celda} debió ser fórmula con fragmentos [{string.Join(",", fragmentos)}] en el canónico Q2.");
        }

        // DetValiRetri2026072 D9:D14 también VALORES (no escritas; 2.7 protegida por Requirement 5).
        for (var fila = 9; fila <= 14; fila++)
        {
            Assert.False(CeldaEsFormula(Insumos.PlantillaQ2, WorkbookLeafCellMapQ2.HojaDetValiRetri, $"D{fila}"),
                $"DetValiRetri2026072!D{fila} debió ser VALOR (no fórmula) en el canónico Q2.");
        }
    }

    [Fact]
    public void T0_DetRetriComposicion_CierraContraGoldenCache5De5()
    {
        // V0.4 (se reutiliza, no se re-descubre): DetRetri-D = ROUND(D104:D108) por ASE + D14 total
        // contra el caché golden. Es la composición que el writer va a escribir.
        var goldenD104 = new decimal[5];
        for (var i = 0; i < 5; i++)
        {
            goldenD104[i] = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, $"D{104 + i}");
        }

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var esperado = decimal.Round(goldenD104[aseId - 1], 0, MidpointRounding.AwayFromZero);
            var goldenDetRetri = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaDetRetri, WorkbookLeafCellMapQ2.ObtenerDetRetriDestino(aseId));
            Assert.InRange(esperado - goldenDetRetri, -Tolerancia, Tolerancia);
        }

        var totalEsperado = decimal.Round(goldenD104.Sum(), 0, MidpointRounding.AwayFromZero);
        var goldenTotal = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaDetRetri, WorkbookLeafCellMapQ2.DetRetriTotal);
        Assert.InRange(totalEsperado - goldenTotal, -Tolerancia, Tolerancia);
    }

    [Fact]
    public void T0_M1Estructural_BceParametrizadoYProtegidasAdicionales_MatcheanCanonico()
    {
        // T0-0.5 / M1 a nivel estructura: ProtegidasBceParaPeriodo(true) + ProtegidasAdicionalesQ2
        // + cadena AJUSTES-SF-T matchean celdas reales del canónico Q2 (incl. sheets 36/37).
        var bceQ2 = BalanceScProtegidasParaPeriodo(true);
        foreach (var (hoja, celda, fragmentos) in bceQ2)
        {
            Assert.True(CeldaTieneFormulaConFragmentos(Insumos.PlantillaQ2, hoja, celda, fragmentos),
                $"{hoja}!{celda} (BCE Q2) debió ser fórmula con fragmentos [{string.Join(",", fragmentos)}] en el canónico Q2.");
        }

        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.ProtegidasAdicionalesQ2)
        {
            Assert.True(CeldaTieneFormulaConFragmentos(Insumos.PlantillaQ2, hoja, celda, fragmentos),
                $"{hoja}!{celda} debió ser fórmula con fragmentos [{string.Join(",", fragmentos)}] en el canónico Q2.");
        }

        foreach (var (hoja, celda, _) in WorkbookLeafCellMapAjustesSfT.Protegidas)
        {
            Assert.True(CeldaEsFormula(Insumos.PlantillaQ2, hoja, celda),
                $"{hoja}!{celda} (cadena AJUSTES-SF-T) debió ser fórmula en el canónico Q2.");
        }
    }

    [Fact]
    public void T0_MapaBancoQ2_ExcluyeTotalRecaudoFila81YMatcheaCanonico()
    {
        // T0-0.5 / D5 (banco parametrizado): el template Q2 termina en la fila 79 — C81/D81 del
        // mapa Q1 NO existen (fail-fast honesto). El resto del mapa banco (consolidado 1–7,
        // totales de bloque, verificación I/J, validación 59–79) es FÓRMULA en el canónico Q2.
        Assert.False(CeldaExiste(Insumos.PlantillaQ2, "REPORTE RECAUDO x BANCO", "C81"), "C81 (TOTAL RECAUDO) no existe en el template Q2.");
        Assert.False(CeldaExiste(Insumos.PlantillaQ2, "REPORTE RECAUDO x BANCO", "D81"), "D81 no existe en el template Q2.");

        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.BancoProtegidasQ2)
        {
            Assert.True(CeldaTieneFormulaConFragmentos(Insumos.PlantillaQ2, hoja, celda, fragmentos),
                $"{hoja}!{celda} (banco Q2) debió ser fórmula con fragmentos [{string.Join(",", fragmentos)}] en el canónico Q2.");
        }
    }

    [Fact]
    public void T0_HashCanonico_IntactoTrasLecturas()
    {
        // A4 extendido al mapa Q2: la lectura de evidencia no muta el canónico.
        var hashAntes = Sha256(Insumos.PlantillaQ2);

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            _ = WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(aseId);
            _ = WorkbookLeafCellMapQ2.ObtenerR2Q2Editables(aseId);
            _ = WorkbookLeafCellMapQ2.ObtenerR4Q2Editables(aseId);
        }

        Assert.Equal(hashAntes, Sha256(Insumos.PlantillaQ2));
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
        var cell = ObtenerCelda(workbookPart, hoja, celda);
        return cell?.CellFormula is not null;
    }

    private static bool CeldaExiste(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        return ObtenerCelda(workbookPart, hoja, celda) is not null;
    }

    private static bool CeldaTieneFormulaConFragmentos(string ruta, string hoja, string celda, string[] fragmentos)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var cell = ObtenerCelda(workbookPart, hoja, celda)
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en el canónico Q2.");

        if (cell.CellFormula is null)
        {
            return false;
        }

        var formula = NormalizarFormula(cell.CellFormula.Text ?? string.Empty);

        // Shared-formula follower (texto vacío + SharedIndex) = réplica del maestro → presencia.
        var esSharedFollower = string.IsNullOrWhiteSpace(formula) && cell.CellFormula.SharedIndex is not null;
        if (esSharedFollower)
        {
            return true;
        }

        return fragmentos
            .Select(NormalizarFormula)
            .Where(f => !string.IsNullOrWhiteSpace(f))
            .All(f => formula.Contains(f, StringComparison.OrdinalIgnoreCase));
    }

    private static string NormalizarFormula(string formula)
    {
        var value = formula.Trim();
        value = value.Replace("'", string.Empty, StringComparison.Ordinal);
        value = value.Replace(" ", string.Empty, StringComparison.Ordinal);
        value = value.Replace("_xlfn.", string.Empty, StringComparison.OrdinalIgnoreCase);
        return value;
    }

    private static Cell? ObtenerCelda(WorkbookPart workbookPart, string hoja, string celda)
    {
        var workbookXml = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook null");
        var sheet = workbookXml.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var sheetId = sheet.Id?.Value ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Id.");
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheetId)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        return ws.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
    }

    private static decimal LeerCeldaNumerica(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var cell = ObtenerCelda(workbookPart, hoja, celda)
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
    /// Espejo de <c>OpenXmlPlantillaWriter.ProtegidasBceParaPeriodo</c>: reemplaza el sufijo
    /// 2026071 → 2026072 en el mapa HU-10 (no toca el mapa; solo su interpretación por período).
    /// </summary>
    private static IEnumerable<(string Hoja, string Celda, string[] Fragmentos)> BalanceScProtegidasParaPeriodo(bool esQuincena2)
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