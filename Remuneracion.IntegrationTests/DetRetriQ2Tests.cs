using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-12 (2.6 ampliada, plan §4 Fase 1/2/4.1 — Unidad 1/PR2): dominio + lectura Q2 reales SIN
/// salida. Cubre: variante R1-Q2-ASE5 (V0.3: 2 filas Mes/Total, TOT_OPT = 12033011685.71 = D13),
/// leafs R1/R2/R4-Q2 contra el caché golden (A1 leaf), visibles por ASE vs CONSOLIDADO golden,
/// composición DetRetri V0.4 exacta (ROUND(D104:D108,0) 5/5), origen-equivocado prohibido
/// (agregado HU-02 ≠ DetRetri) y fail-fast que nombra ASE + reporte ante slot ausente.
/// </summary>
public sealed class DetRetriQ2Tests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    [Fact]
    public void R1Q2_Ase5_Variante2Filas_SlotsV03YTotOptGolden()
    {
        // V0.3: la fuente ASE5-Q2 trae SOLO 2 filas Mes/Total; el reader despacha a la rama
        // ASE5-Q2 y los slots F531/F552/L531 se pueblan; TOT_OPT = F558 = 12033011685.71 = D13.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var leaf = reader.LeerLeafInputs(Insumos.Ase(5), Insumos.PeriodoQ2(), Insumos.R1Q2(5), Insumos.R2Q2(5), Insumos.R4Q2(5));

        Assert.InRange(leaf.R1.CeldasPorAse["F531"] - 13084233937.44m, -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R1.CeldasPorAse["F552"] - (-1036320522.40m), -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R1.CeldasPorAse["L531"] - 14901729.33m, -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R1.TotalOportunoEsperadoPorAse - 12033011685.71m, -Tolerancia, Tolerancia);

        // EXTEMP ASE5-Q2: F560 = F538+F512-L512 = 441849 = D51 golden.
        Assert.InRange(leaf.R1.ExtemporaneoEsperadoPorAse - 441849m, -Tolerancia, Tolerancia);

        // Coherencia contra la fuente (gate HU-02): F25 = primera fila Mes/Total = Extemporaneo.
        Assert.InRange(leaf.R1.F25 - 13084233937.44m, -Tolerancia, Tolerancia);
    }

    [Fact]
    public void R1Q2_Los5_SlotsVsGoldenCache_5De5()
    {
        // A1 leaf: cada celda operando R1-Q2 leída == misma celda del caché golden ±0.5.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var leaf = reader.LeerLeafInputs(Insumos.Ase(aseId), Insumos.PeriodoQ2(), Insumos.R1Q2(aseId), Insumos.R2Q2(aseId), Insumos.R4Q2(aseId));
            foreach (var (celda, _) in WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(aseId))
            {
                var golden = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaR1, celda);
                Assert.InRange(leaf.R1.CeldasPorAse[celda] - golden, -Tolerancia, Tolerancia);
            }
        }
    }

    [Fact]
    public void R2R4Q2_Los5_SlotsVsGoldenCache_5De5()
    {
        // A1 leaf: R2 (Componente/SubsCont/Especiales) y R4 (Total/P) vs caché golden.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var leaf = reader.LeerLeafInputs(Insumos.Ase(aseId), Insumos.PeriodoQ2(), Insumos.R1Q2(aseId), Insumos.R2Q2(aseId), Insumos.R4Q2(aseId));

            var r2 = WorkbookLeafCellMapQ2.ObtenerR2Q2Editables(aseId);
            Assert.InRange(leaf.R2.CeldasPorAse[r2.Componente] - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaR2, r2.Componente), -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R2.CeldasPorAse[r2.SubsCont] - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaR2, r2.SubsCont), -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R2.CeldasPorAse[r2.Especiales] - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaR2, r2.Especiales), -Tolerancia, Tolerancia);

            var r4 = WorkbookLeafCellMapQ2.ObtenerR4Q2Editables(aseId);
            Assert.InRange(leaf.R4.CeldasPorAse[r4.Total] - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaR4, r4.Total), -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R4.CeldasPorAse[r4.P] - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaR4, r4.P), -Tolerancia, Tolerancia);
        }
    }

    [Fact]
    public void VisiblesQ2_Los5_CoincidenConConsolidadoGolden()
    {
        // V0.1 + T0-0.2: visibles leaf Q2 == CONSOLIDADO golden D9:D13 (TOT_OPT), D28:D32 (R2),
        // D47:D51 (EXTEMP), D66:D70 (R4).
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var leaf = reader.LeerLeafInputs(Insumos.Ase(aseId), Insumos.PeriodoQ2(), Insumos.R1Q2(aseId), Insumos.R2Q2(aseId), Insumos.R4Q2(aseId));

            Assert.InRange(leaf.R1.TotalOportunoEsperadoPorAse - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, $"D{8 + aseId}"), -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R2.TotalOportunoEsperado - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, $"D{27 + aseId}"), -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R1.ExtemporaneoEsperadoPorAse - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, $"D{46 + aseId}"), -Tolerancia, Tolerancia);
            Assert.InRange(leaf.R4.TotalReversionEsperada - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, $"D{65 + aseId}"), -Tolerancia, Tolerancia);
        }
    }

    [Fact]
    public void DetRetriQ2_Composicion_TotalD104YDetalleVsGoldenCache_5De5()
    {
        // V0.4 (congelada): DetRetri_D(ase) = ROUND(D104:D108,0). TotalD104 = Σ visibles leaf
        // (+ AJUSTES-SF-T) == golden D104:D108; Detalle == golden DetRetri2026072 D9:D14.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var leafReader = new ExcelDataReaderWorkbookLeafInputReader();

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var leaf = leafReader.LeerLeafInputs(Insumos.Ase(aseId), Insumos.PeriodoQ2(), Insumos.R1Q2(aseId), Insumos.R2Q2(aseId), Insumos.R4Q2(aseId));
            leaf.AjustesSfT = new AjustesSfTInputs
            {
                Ase = leaf.Ase,
                SaldosNotas = reader.LeerSaldosNotas(leaf.Ase, Insumos.SaldosNotas(aseId)),
                RetribucionNegativa = reader.LeerRetribucionNegativa(leaf.Ase, Insumos.RetribucionNegativa(aseId))
            };

            var detretri = new DetRetriQ2Inputs
            {
                Ase = leaf.Ase,
                TotalD104 = leaf.R1.TotalOportunoEsperadoPorAse
                    + leaf.R2.TotalOportunoEsperado
                    + leaf.R1.ExtemporaneoEsperadoPorAse
                    + leaf.R4.TotalReversionEsperada
                    + leaf.AjustesSfT.TotalAjustes
            };

            var goldenD104 = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, $"D{103 + aseId}");
            var goldenDetalle = LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaDetRetri, WorkbookLeafCellMapQ2.ObtenerDetRetriDestino(aseId));

            Assert.InRange(detretri.TotalD104 - goldenD104, -Tolerancia, Tolerancia);
            Assert.InRange(detretri.Detalle - goldenDetalle, -Tolerancia, Tolerancia);
            Assert.Equal(DetRetriRounder.Round(detretri.TotalD104), detretri.Detalle); // única regla de redondeo
        }

        // D14 = total: ROUND(Σ D104:D108) == golden DetRetri D14.
        var totalDetalle = DetRetriRounder.Round(SumarGoldenD104());
        Assert.InRange(totalDetalle - LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMapQ2.HojaDetRetri, WorkbookLeafCellMapQ2.DetRetriTotal), -Tolerancia, Tolerancia);
    }

    [Fact]
    public void DetRetriQ2_ProhibidoOrigenEquivocado_TotOptHU02NoEsDetRetri()
    {
        // §5.2 / A5: PROHIBIDO usar el agregado HU-02 (TotOpt de la fuente) como DetRetri-D.
        // El test demuestra la confusión fallando: TotOpt de ASE1 (fila Componente/Total de la
        // fuente) ≠ D104:D108 leaf-based.
        var lector = new ExcelDataReaderRecaudoReader();
        var r1 = lector.LeerR1(Insumos.R1Q2(1));
        var leafReader = new ExcelDataReaderWorkbookLeafInputReader();
        var leaf = leafReader.LeerLeafInputs(Insumos.Ase(1), Insumos.PeriodoQ2(), Insumos.R1Q2(1), Insumos.R2Q2(1), Insumos.R4Q2(1));

        var totalD104 = leaf.R1.TotalOportunoEsperadoPorAse
            + leaf.R2.TotalOportunoEsperado
            + leaf.R1.ExtemporaneoEsperadoPorAse
            + leaf.R4.TotalReversionEsperada;

        Assert.NotEqual(r1.TotalOportuno, totalD104);
        Assert.NotEqual(DetRetriRounder.Round(r1.TotalOportuno), DetRetriRounder.Round(totalD104));
    }

    [Fact]
    public void R1Q2_SlotAusente_FallaNombrandoAseYReporte()
    {
        // §5.2 / Riesgo 6: si falta un slot del mapa T0 (rol Mes2 para ASE1), fail-fast nombra
        // ASE + reporte; nunca 0 silencioso.
        var ruta = CrearFuenteR1SinMes2();
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            reader.LeerLeafInputs(Insumos.Ase(1), Insumos.PeriodoQ2(), ruta, Insumos.R2Q2(1), Insumos.R4Q2(1)));

        Assert.Contains("ASE 1", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("R1-Q2", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static decimal SumarGoldenD104()
    {
        decimal suma = 0m;
        for (var i = 0; i < 5; i++)
        {
            suma += LeerCeldaNumerica(Insumos.GoldenQ2, WorkbookLeafCellMap.HojaConsolidado, $"D{104 + i}");
        }

        return suma;
    }

    /// <summary>
    /// Fuente sintética R1 con SOLO 2 filas Mes/Total (ASE1-Q2 exige 3 → rol Mes2 ausente).
    /// </summary>
    private static string CrearFuenteR1SinMes2()
    {
        var ruta = Path.Combine(Path.GetTempPath(), "r1-sin-mes2-" + Guid.NewGuid().ToString("N") + ".xlsx");
        using (var doc = SpreadsheetDocument.Create(ruta, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var sheetPart = workbookPart.AddNewPart<WorksheetPart>();

            Row FilaMes(int fila)
            {
                return new Row(
                    new Cell { CellReference = $"B{fila}", DataType = CellValues.String, CellValue = new CellValue("Mes") },
                    new Cell { CellReference = $"C{fila}", DataType = CellValues.String, CellValue = new CellValue("Total") },
                    new Cell { CellReference = $"F{fila}", CellValue = new CellValue("100") });
            }

            sheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(new Cell
                {
                    CellReference = "A1",
                    DataType = CellValues.String,
                    CellValue = new CellValue("Recaudo Desde: 16/07/2026 Hasta: 31/07/2026")
                }),
                // Header con "Componente TDF" + F=Total (requisitos del agregado HU-02).
                new Row(
                    new Cell { CellReference = "F3", DataType = CellValues.String, CellValue = new CellValue("Total") },
                    new Cell { CellReference = "G3", DataType = CellValues.String, CellValue = new CellValue("Componente TDF") }),
                new Row(
                    new Cell { CellReference = "A4", DataType = CellValues.String, CellValue = new CellValue("Componente") },
                    new Cell { CellReference = "B4", DataType = CellValues.String, CellValue = new CellValue("Total") },
                    new Cell { CellReference = "F4", CellValue = new CellValue("100") }),
                FilaMes(11),
                FilaMes(31)));

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(sheetPart), SheetId = 1, Name = "Sheet1" });
            workbookPart.Workbook.Save();
        }

        return ruta;
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
