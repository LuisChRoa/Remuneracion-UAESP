using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

public sealed class WorkbookLeafInputsTests
{
    private const decimal Tolerancia = 0.5m;

    [Fact]
    public void LeerLeafInputs_PromoambientalQ1_ProduceVisiblesEsperadosDelWorkbook()
    {
        var insumos = ResolverInsumos();
        var leaf = new ExcelDataReaderWorkbookLeafInputReader()
            .LeerLeafInputs(CrearAse(), CrearPeriodo(), insumos.R1, insumos.R2, insumos.R4);

        Assert.True(leaf.R1.F25 != 0m);
        Assert.True(leaf.R2.E15 != 0m);
        Assert.True(leaf.R4.D9 != 0m);
        Assert.InRange(leaf.R1.TotalOportunoEsperado - 16704332434.57m, -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R1.ExtemporaneoEsperado - 11673020m, -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R2.TotalOportunoEsperado - 54216385.68m, -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R4.TotalReversionEsperada - (-12054255.65m), -Tolerancia, Tolerancia);
    }

    [Fact]
    public void GenerarWorkbook_EscribeSoloLeafYPreservaFormulas()
    {
        var insumos = ResolverInsumos();
        var ase = CrearAse();
        var periodo = CrearPeriodo();
        var reader = new ExcelDataReaderRecaudoReader();
        var r1 = reader.LeerR1(insumos.R1);
        var r2 = reader.LeerR2(insumos.R2);
        var r4 = reader.LeerR4(insumos.R4);
        var resultado = new CalculoRemuneracion().CalcularConsolidado(periodo, [(ase, r1, r2, r4)]);
        var leaf = new ExcelDataReaderWorkbookLeafInputReader().LeerLeafInputs(ase, periodo, insumos.R1, insumos.R2, insumos.R4);
        var salida = Path.Combine(Path.GetTempPath(), "remuneracion-it-" + Guid.NewGuid().ToString("N"), "salida.xlsx");

        new OpenXmlPlantillaWriter().GenerarWorkbook(insumos.Plantilla, salida, resultado, leaf);

        Assert.True(File.Exists(salida));
        Assert.True(CeldaEsFormula(salida, "Reporte Componentes R1", "F46"));
        Assert.True(CeldaEsFormula(salida, "CONSOLIDADO_TOTAL RECAUDO", "D9"));
        Assert.True(CeldaEsFormula(salida, "CONSOLIDADO_TOTAL RECAUDO", "D109"));
    }

    [Fact]
    public void GenerarWorkbook_MismatchLeafVsAgregado_NoCertificaArchivo()
    {
        var insumos = ResolverInsumos();
        var ase = CrearAse();
        var periodo = CrearPeriodo();
        var reader = new ExcelDataReaderRecaudoReader();
        var r1 = reader.LeerR1(insumos.R1);
        var r2 = reader.LeerR2(insumos.R2);
        var r4 = reader.LeerR4(insumos.R4);
        var resultado = new CalculoRemuneracion().CalcularConsolidado(periodo, [(ase, r1, r2, r4)]);
        var leaf = new ExcelDataReaderWorkbookLeafInputReader().LeerLeafInputs(ase, periodo, insumos.R1, insumos.R2, insumos.R4);
        leaf.R2.E15 += 10m;
        var salida = Path.Combine(Path.GetTempPath(), "remuneracion-it-fail-" + Guid.NewGuid().ToString("N"), "salida.xlsx");

        Assert.Throws<CalculoInvalidoException>(() =>
            new OpenXmlPlantillaWriter().GenerarWorkbook(insumos.Plantilla, salida, resultado, leaf));
        Assert.False(File.Exists(salida));
    }

    private static Ase CrearAse() => AseFactory.DesdeId(1); // HU-14 (S-3): factoría única

    private static Periodo CrearPeriodo() =>
        new() { CodigoAAAAMM = "202607", NumeroQuincena = 1 };

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

    private static (string Plantilla, string R1, string R2, string R4) ResolverInsumos()
    {
        var raiz = BuscarRaizRepo()
            ?? throw new DirectoryNotFoundException("No se encontró la raíz del repositorio con Docs/Insumos.");

        var plantilla = Path.Combine(raiz, "Docs", "Insumos", "Remuneracion 202607-1 Total.xlsx");
        var carpeta = Path.Combine(raiz, "Docs", "Insumos", "REMUNERACION 2026071", "1-Promoambiental");
        var r1 = Path.Combine(carpeta, "Recaudoporcomponente_to_date01072026ddMMyyyy_to_date15072026ddMMyyyy___20267161653925.xlsx");
        var r2 = Path.Combine(carpeta, "RerpoteDetalleSaldosaFavor_to_date01072026ddMMyyyy_to_date15072026ddMMyyyy___202671616325453.xlsx");
        var r4 = Path.Combine(carpeta, "ReversiónPorComponente_to_date01072026ddMMyyyy_to_date15072026ddMMyyyy___2026716163221489.xlsx");

        Assert.True(File.Exists(plantilla), $"Falta plantilla: {plantilla}");
        Assert.True(File.Exists(r1), $"Falta R1: {r1}");
        Assert.True(File.Exists(r2), $"Falta R2: {r2}");
        Assert.True(File.Exists(r4), $"Falta R4: {r4}");
        return (plantilla, r1, r2, r4);
    }

    private static string? BuscarRaizRepo()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")) && Directory.Exists(Path.Combine(dir.FullName, "Docs", "Insumos")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        return null;
    }
}
