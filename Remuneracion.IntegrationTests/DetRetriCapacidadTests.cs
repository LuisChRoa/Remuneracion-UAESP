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

        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-capacidad-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var plantillaFutura = Path.Combine(salidaDir, "plantilla-sin-D10.xlsx");
        File.Copy(Insumos.PlantillaQ2, plantillaFutura, overwrite: true);

        QuitarCelda(plantillaFutura, "DetRetri2026072", "D10");

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            new OpenXmlPlantillaWriter().GenerarWorkbook(plantillaFutura, Path.Combine(salidaDir, "salida.xlsx"), resultado, leafs));

        Assert.Contains("DetRetri2026072", ex.Message);
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
