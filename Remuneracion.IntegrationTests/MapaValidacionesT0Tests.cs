using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-13 (2.7, Plan 13 §4 Fase 0 — Unidad 0): EVIDENCIA T0 contra ambos canónicos.
/// Congela la semántica del oráculo de validaciones cruzadas (nada entra al código sin esta
/// tabla): filas-ASE de VALIDACION_* (3..7) con O/P fórmulas, DetValiRetri D16:D20 ≤ ±0.5 y
/// D24:D28/D29 booleanos TRUE, D21 (Total) divergencia documentada (D6), D9:D14 ≠ ROUND
/// documentado, Valida -* vistas y GERENTES_* fórmulas puras. Fe de erratas: sheets 36/37 por
/// NOMBRE con sufijo 2026071/2026072 (nunca 93/94, A8).
/// </summary>
public sealed class MapaValidacionesT0Tests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    [Fact]
    public void T0_VALIDACION_PorEmpresa_Filas3A7Ase_OPFormula_EnQ1()
    {
        // Q1 golden (= plantilla canónica): O/P filas 3..7 (ASE 1..5) son FÓRMULA en las 5 hojas.
        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var fila = WorkbookLeafCellMapValidaciones.FilaValidacionPorAse(aseId);
                Assert.True(CeldaEsFormula(Insumos.Plantilla, empresa.HojaValidacion, $"O{fila}"),
                    $"{empresa.HojaValidacion}!O{fila} debió ser fórmula en Q1.");
                Assert.True(CeldaEsFormula(Insumos.Plantilla, empresa.HojaValidacion, $"P{fila}"),
                    $"{empresa.HojaValidacion}!P{fila} debió ser fórmula en Q1.");
            }
        }
    }

    [Fact]
    public void T0_VALIDACION_PorEmpresa_OPCeroTrue_EnQ1YQ2()
    {
        // Caché golden: O = 0 y P = TRUE en toda fila-ASE de las 5 empresas, en Q1 y Q2 (ceros
        // legítimos donde la empresa no tiene actividad en un ASE).
        foreach (var (ruta, etiqueta) in new[] { (Insumos.Plantilla, "Q1"), (Insumos.GoldenQ2, "Q2") })
        {
            foreach (var empresa in EmpresaFacturacion.Catalogo)
            {
                for (var aseId = 1; aseId <= 5; aseId++)
                {
                    var fila = WorkbookLeafCellMapValidaciones.FilaValidacionPorAse(aseId);
                    Assert.InRange(LeerCeldaNumerica(ruta, empresa.HojaValidacion, $"O{fila}"), -Tolerancia, Tolerancia);
                    Assert.True(LeerCeldaBooleana(ruta, empresa.HojaValidacion, $"P{fila}"),
                        $"{empresa.HojaValidacion}!P{fila} debió ser TRUE en {etiqueta}.");
                }
            }
        }
    }

    [Fact]
    public void T0_VALIDACION_TOTAL_OPFilas3A9_FormulaYCeroTrue_EnAmbos()
    {
        foreach (var (ruta, etiqueta) in new[] { (Insumos.Plantilla, "Q1"), (Insumos.GoldenQ2, "Q2") })
        {
            for (var fila = 3; fila <= 9; fila++)
            {
                Assert.True(CeldaEsFormula(ruta, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, $"O{fila}"),
                    $"VALIDACION_TOTAL!O{fila} debió ser fórmula en {etiqueta}.");
                Assert.True(CeldaEsFormula(ruta, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, $"P{fila}"),
                    $"VALIDACION_TOTAL!P{fila} debió ser fórmula en {etiqueta}.");
                Assert.InRange(LeerCeldaNumerica(ruta, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, $"O{fila}"), -Tolerancia, Tolerancia);
                Assert.True(LeerCeldaBooleana(ruta, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, $"P{fila}"),
                    $"VALIDACION_TOTAL!P{fila} debió ser TRUE en {etiqueta}.");
            }
        }
    }

    [Fact]
    public void T0_DetValiRetri_DiferenciasAse16a20_DentroTolerancia_EnAmbos()
    {
        foreach (var (ruta, quincena, etiqueta) in new[]
                 { (Insumos.Plantilla, 1, "Q1"), (Insumos.GoldenQ2, 2, "Q2") })
        {
            var hoja = WorkbookLeafCellMapValidaciones.HojaDetValiRetri(quincena);
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var fila = WorkbookLeafCellMapValidaciones.FilaDiferenciaDetValiRetri(aseId);
                var valor = LeerCeldaNumerica(ruta, hoja, $"D{fila}");
                Assert.InRange(valor, -Tolerancia, Tolerancia);
            }
        }
    }

    [Fact]
    public void T0_DetValiRetri_D21Total_DivergenciaDocumentada_EnAmbos()
    {
        // D6: la fila Total D21 = D14 − CONSOLIDADO U109 NO cierra ±0.5 en NINGÚN canónico
        // (acumulado de ruido float de las 5 filas). Queda documentada y EXCLUIDA del gate.
        var q1 = LeerCeldaNumerica(Insumos.Plantilla, WorkbookLeafCellMapValidaciones.HojaDetValiRetri(1), "D21");
        var q2 = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapValidaciones.HojaDetValiRetri(2), "D21");
        Assert.True(Math.Abs(q1) > Tolerancia, $"Q1 D21 = {q1}; se esperaba divergencia > ±0.5 (documentada).");
        Assert.True(Math.Abs(q2) > Tolerancia, $"Q2 D21 = {q2}; se esperaba divergencia > ±0.5 (documentada).");
    }

    [Fact]
    public void T0_DetValiRetri_Verificaciones24a29_TRUE_EnAmbos()
    {
        foreach (var (ruta, quincena, etiqueta) in new[]
                 { (Insumos.Plantilla, 1, "Q1"), (Insumos.GoldenQ2, 2, "Q2") })
        {
            var hoja = WorkbookLeafCellMapValidaciones.HojaDetValiRetri(quincena);
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var fila = WorkbookLeafCellMapValidaciones.FilaVerificacionDetValiRetri(aseId);
                Assert.True(LeerCeldaBooleana(ruta, hoja, $"D{fila}"),
                    $"{hoja}!D{fila} debió ser TRUE en {etiqueta}.");
            }

            Assert.True(LeerCeldaBooleana(ruta, hoja, "D29"), $"{hoja}!D29 debió ser TRUE en {etiqueta}.");
        }
    }

    [Fact]
    public void T0_D9D14_ValoresFuente_NoExigenRound_EnAmbos()
    {
        // Hallazgo HU-12 (D6): DetRetri/DetValiRetri D14/J14 del golden son VALORES fuente
        // (enteros escritos) que NO equivalen a ROUND(Σ float) en la fila Total. Documentado:
        // DetValiRetri D14 ≠ ROUND(CONSOLIDADO U109) en ambos canónicos → el gate NUNCA exige
        // ROUND sobre D9:D14 (prohibido "cerrarlos").
        var u109Q1 = LeerCeldaNumerica(Insumos.Plantilla, WorkbookLeafCellMap.HojaConsolidado, "U109");
        var d14Q1 = LeerCeldaNumerica(Insumos.Plantilla, WorkbookLeafCellMapValidaciones.HojaDetValiRetri(1), "D14");
        Assert.NotEqual(decimal.Round(u109Q1, 0, MidpointRounding.AwayFromZero), d14Q1);

        var u109Q2 = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, "U109");
        var d14Q2 = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapValidaciones.HojaDetValiRetri(2), "D14");
        Assert.NotEqual(decimal.Round(u109Q2, 0, MidpointRounding.AwayFromZero), d14Q2);
    }

    [Fact]
    public void T0_ValidaYGerentes_RepresentantesFormula_EnAmbos()
    {
        foreach (var (ruta, etiqueta) in new[] { (Insumos.Plantilla, "Q1"), (Insumos.GoldenQ2, "Q2") })
        {
            // Valida -*: vistas lado-a-lado con celdas fórmula (no gate numérico; T0-0.5).
            Assert.True(CeldaEsFormula(ruta, "Valida -Remunera", "D9"), $"Valida -Remunera!D9 debió ser fórmula en {etiqueta}.");
            Assert.True(CeldaEsFormula(ruta, "Valida - Anticipos", "D6"), $"Valida - Anticipos!D6 debió ser fórmula en {etiqueta}.");

            // GERENTES_*: fórmulas puras (T0-0.6, cierra el pendiente Plan 08).
            foreach (var empresa in EmpresaFacturacion.Catalogo)
            {
                Assert.True(CeldaEsFormula(ruta, empresa.HojaGerentes, "D9"),
                    $"{empresa.HojaGerentes}!D9 debió ser fórmula en {etiqueta}.");
                Assert.True(CeldaEsFormula(ruta, empresa.HojaGerentes, "D14"),
                    $"{empresa.HojaGerentes}!D14 debió ser fórmula (SUM) en {etiqueta}.");
            }
        }
    }

    [Fact]
    public void T0_FeDeErratas_HojasDetRetri36_37_PorNombre_EnAmbos()
    {
        // G7/A8: las hojas DetRetri/DetValiRetri existen por NOMBRE con sufijo de período
        // (2026071 en Q1, 2026072 en Q2); nunca números 93/94.
        Assert.True(HojaExiste(Insumos.Plantilla, "DetRetri2026071"), "DetRetri2026071 debió existir en Q1.");
        Assert.True(HojaExiste(Insumos.Plantilla, "DetValiRetri2026071"), "DetValiRetri2026071 debió existir en Q1.");
        Assert.True(HojaExiste(Insumos.GoldenQ2, "DetRetri2026072"), "DetRetri2026072 debió existir en Q2.");
        Assert.True(HojaExiste(Insumos.GoldenQ2, "DetValiRetri2026072"), "DetValiRetri2026072 debió existir en Q2.");
    }

    [Fact]
    public void T0_CeldasOracle_SinExternalLinks_EnAmbos()
    {
        // §0.1 punto 7: ninguna celda-oráculo depende de externalLink roto.
        var celdas = new List<(string Hoja, string Celda)>();
        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var fila = WorkbookLeafCellMapValidaciones.FilaValidacionPorAse(aseId);
                celdas.Add((empresa.HojaValidacion, $"O{fila}"));
                celdas.Add((empresa.HojaValidacion, $"P{fila}"));
            }
        }

        celdas.Add((WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "O9"));
        celdas.Add((WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "P9"));

        foreach (var (ruta, etiqueta) in new[] { (Insumos.Plantilla, "Q1"), (Insumos.GoldenQ2, "Q2") })
        {
            foreach (var (hoja, celda) in celdas)
            {
                var formula = LeerFormula(ruta, hoja, celda);
                Assert.DoesNotContain("[", formula, StringComparison.Ordinal);
            }
        }
    }

    // ── Helpers OpenXML (read-only; copia local de los helpers de MapaQ2T0Tests) ───────────────

    private static bool HojaExiste(string ruta, string hoja)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        return workbookPart.Workbook!.Descendants<Sheet>()
            .Any(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
    }

    private static bool CeldaEsFormula(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var cell = ObtenerCelda(workbookPart, hoja, celda);
        return cell?.CellFormula is not null;
    }

    private static string LeerFormula(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var cell = ObtenerCelda(workbookPart, hoja, celda);
        return cell?.CellFormula?.Text ?? string.Empty;
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

    private static bool LeerCeldaBooleana(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var cell = ObtenerCelda(workbookPart, hoja, celda)
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda}.");
        if (cell.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
        {
            return false;
        }

        var texto = cell.CellValue.InnerText.Trim();
        if (cell.DataType is not null && cell.DataType.Value == CellValues.Boolean)
        {
            return texto != "0";
        }

        return decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor) && valor != 0m;
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
}
