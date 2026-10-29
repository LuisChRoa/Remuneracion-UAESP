using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-13 (2.7, deuda W1-guarda D8): la escritura DetRetri-Q2 (D9:D14) lleva guarda de capacidad.
/// Si una plantilla futura cambia la capacidad (más/menos filas ASE), el writer FALLA nombrando
/// la hoja — nunca truncado silencioso, nunca insert/delete de filas.
/// </summary>
public sealed class DetRetriCapacidadTests
{
    [Fact]
    public void PlantillaSinCeldaDetRetriD10_FallaNombrandoLaHoja()
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

        // Plan 62 (T2.2/T2.5, reclassificado): el asiento DetRetri (D9:D14) ya NO lo escribe el
        // dominio; lo PEGA el bloque del R10 del periodo. La guarda de capacidad que este fixture
        // ejercita vive ahora en el plan-then-write del pegado: sin el bloque no habría nada que
        // escribir sobre D10 y el test no probaría nada.
        var lectorR10 = new ExcelDataReaderDetRetriR10Reader();
        resultado.BloqueDetRetriR10 = lectorR10.LeerBloqueDetRetri(Insumos.PeriodoQ2(), Insumos.R10Q2);
        resultado.BloqueValiRetriR10 = lectorR10.LeerBloqueDetValiRetri(Insumos.PeriodoQ2(), Insumos.R10Q2);

        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-capacidad-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var plantillaFutura = Path.Combine(salidaDir, "plantilla-sin-D10.xlsx");
        File.Copy(Insumos.PlantillaQ2, plantillaFutura, overwrite: true);

        var hojaDetRetri = WorkbookLeafCellMapQ2.HojaDetRetri(Insumos.PeriodoQ2());
        QuitarCelda(plantillaFutura, hojaDetRetri, "D10");

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            new OpenXmlPlantillaWriter().GenerarWorkbook(plantillaFutura, Path.Combine(salidaDir, "salida.xlsx"), resultado, leafs));

        Assert.Contains(hojaDetRetri, ex.Message);
        Assert.Contains("D10", ex.Message);
        Assert.False(File.Exists(Path.Combine(salidaDir, "salida.xlsx")), "No debe quedar salida parcial ante fallo (patrón existente).");
    }

    private static void QuitarCelda(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var cell = ws.Descendants<Cell>().First(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        cell.Remove();
        ws.Save();
    }
}
