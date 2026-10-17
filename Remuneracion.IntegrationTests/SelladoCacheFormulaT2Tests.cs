using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 40 (T2, R-S-1/R-S-2/R-S-3 — D-E): el writer SELA los cachés computables por read-back
/// (cadena <c>C3:C8→C9→C14→C15</c>) tras la escritura de negocio y antes del saneamiento calcChain.
/// Obligatorio: el <c>&lt;v&gt;</c> sellado COINCIDE con la evaluación independiente del gate
/// (sellado == evaluación) y el <c>&lt;f&gt;</c> queda byte-idéntica a la de la base.
/// </summary>
public sealed class SelladoCacheFormulaT2Tests
{
    private const string Hoja = "VALIDACION_TOTAL";

    [Fact]
    public void T2_SelladoCoincideConEvaluacion_YFormulaIntacta()
    {
        var (salida, baseXlsx) = EjecutarFlujoQ2();

        // 1) El sellado dejó el caché de la cadena correcto (antes: C9/C14/C15 heredaban 0).
        Assert.Equal("1", Valor(salida, Hoja, "C15"));
        Assert.Equal(Valor(salida, Hoja, "C14"), Valor(salida, Hoja, "C9"));

        // 2) Sellado == evaluación independiente del gate (S6).
        var evidencia = new ValidacionOracleReader().LeerSnapshots(salida, Insumos.PeriodoQ2())[0]
            .SubBloquesValidacionTotalEvidencia["C15"];
        Assert.True(evidencia.Computable);
        Assert.Equal(decimal.Parse(Valor(salida, Hoja, "C14")!, CultureInfo.InvariantCulture), evidencia.OperandoIzquierdo);
        Assert.Equal(decimal.Parse(Valor(salida, Hoja, "C9")!, CultureInfo.InvariantCulture), evidencia.OperandoDerecho);
        Assert.True(FormasValidacionTotal.Igualdad(evidencia.OperandoIzquierdo, evidencia.OperandoDerecho));

        // 3) El <f> de las celdas selladas es byte-idéntico a la base (nunca se sobrescribe).
        foreach (var celda in new[] { "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C14", "C15" })
        {
            Assert.Equal(FormulaOuterXml(baseXlsx, Hoja, celda), FormulaOuterXml(salida, Hoja, celda));
        }
    }

    [Fact]
    public void T2_Sellado_SoloAllowList_SubBloquesSelladosConElValorEvaluado()
    {
        // Sana (Plan 46: el espejo de remuneración CIERRA): los 5 sub-bloques evalúan TRUE → sellan "1".
        var (salidaSana, baseXlsx) = EjecutarFlujoQ2();
        foreach (var celda in new[] { "C15", "D25", "O25", "D34", "F34" })
        {
            Assert.Equal("1", Valor(salidaSana, Hoja, celda));
        }

        // Sintética (FIXTURE Plan 46 H2: Total de recaudo +1 vs celdas ⇒ C15 FALSE con Δ = 1):
        // el sellado (read-back) escribe "0" en C15 y "1" en el resto — nunca inventa un TRUE (D-E).
        var salidaSintetica = EjecutarFlujoQ2Sintetico();
        Assert.Equal("0", Valor(salidaSintetica, Hoja, "C15"));
        foreach (var celda in new[] { "D25", "O25", "D34", "F34" })
        {
            Assert.Equal("1", Valor(salidaSintetica, Hoja, celda));
        }

        // Los operandos intermedios (O24/D33/F33) NO están en la allow-list del writer: su caché
        // queda igual que en la base (no se sellan).
        foreach (var celda in new[] { "O24", "D33", "F33" })
        {
            Assert.Equal(Valor(baseXlsx, Hoja, celda), Valor(salidaSana, Hoja, celda));
        }
    }

    // ── Helpers ─────────────────────────────────────────────────────────────────────────────────

    private static (string Salida, string Base) EjecutarFlujoQ2()
    {
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "t2-sellado-" + Guid.NewGuid().ToString("N"))).FullName;
        var salida = Path.Combine(dir, "salida.xlsx");
        var procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader());
        procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.PeriodoQ2(),
            CarpetaPeriodo = Insumos.CarpetaInsumosJulioQ2,
            RutaPlantilla = Insumos.PlantillaQ2,
            RutaSalida = salida
        });
        return (salida, Insumos.PlantillaQ2);
    }

    /// <summary>
    /// Flujo Q2 con el FIXTURE SINTÉTICO (Plan 46 H2): sin oráculo ni reporte, para observar SOLO el
    /// sellado. El hecho sintético (<c>Recaudo *</c> Total +1) hace que <c>C15</c> evalúe FALSE.
    /// </summary>
    private static string EjecutarFlujoQ2Sintetico()
    {
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "t2-sellado-sint-" + Guid.NewGuid().ToString("N"))).FullName;
        var salida = Path.Combine(dir, "salida.xlsx");
        var procesador = FixtureDivergenciaSintetica.CrearProcesador(conOraculo: false, conReporte: false);
        procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.PeriodoQ2(),
            CarpetaPeriodo = Insumos.CarpetaInsumosJulioQ2,
            RutaPlantilla = Insumos.PlantillaQ2,
            RutaSalida = salida
        });
        return salida;
    }

    private static string? Valor(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        return Celda(workbook, hoja, celda).CellValue?.InnerText;
    }

    private static string? FormulaOuterXml(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        return Celda(workbook, hoja, celda).CellFormula?.OuterXml;
    }

    private static Cell Celda(SpreadsheetDocument workbook, string hoja, string celda)
    {
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        return worksheet.Descendants<Cell>().First(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
    }
}
