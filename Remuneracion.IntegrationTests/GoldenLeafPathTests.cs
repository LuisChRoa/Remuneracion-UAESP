using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

public sealed class GoldenLeafPathTests
{
    private const decimal Tolerancia = 0.5m;

    [Fact]
    public void Procesador_CapaA_LeafVisiblesFormulasYHashSinUsarCacheDeSalida()
    {
        var insumos = ResolverInsumos();
        var origenHashAntes = Sha256(insumos.Plantilla);
        var ase = CrearAse();
        var periodo = CrearPeriodo();
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-hu06-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, periodo.NombreArchivo);

        var procesador = new ProcesadorRemuneracion(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter());

        var resultado = procesador.Ejecutar(new SolicitudProcesoAse
        {
            Ase = ase,
            Periodo = periodo,
            RutaR1 = insumos.R1,
            RutaR2 = insumos.R2,
            RutaR4 = insumos.R4,
            RutaPlantilla = insumos.Plantilla,
            RutaSalida = salida
        });

        Assert.True(File.Exists(salida));

        // A1 — leaf escritos en la SALIDA vs leaf del golden (no celdas fórmula).
        foreach (var (hoja, celda, _) in WorkbookLeafCellMap.EditableLeafCells)
        {
            var escrito = LeerCeldaNumerica(salida, hoja, celda);
            var golden = LeerCeldaNumerica(insumos.Plantilla, hoja, celda);
            Assert.InRange(escrito - golden, -Tolerancia, Tolerancia);
        }

        // A2 — visibles de DOMINIO vs CACHE del golden (nunca cache de fórmulas de la salida).
        var leaf = resultado.Leaf;
        Assert.InRange(leaf.R1.TotalOportunoEsperado - LeerCeldaNumerica(insumos.Plantilla, WorkbookLeafCellMap.HojaR1, "F46"), -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R1.ExtemporaneoEsperado - LeerCeldaNumerica(insumos.Plantilla, WorkbookLeafCellMap.HojaR1, "F48"), -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R2.TotalOportunoEsperado - LeerCeldaNumerica(insumos.Plantilla, WorkbookLeafCellMap.HojaR2, "E41"), -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R4.TotalReversionEsperada - LeerCeldaNumerica(insumos.Plantilla, WorkbookLeafCellMap.HojaR4, "D67"), -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R1.TotalOportunoEsperado - LeerCeldaNumerica(insumos.Plantilla, WorkbookLeafCellMap.HojaConsolidado, "D9"), -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R1.ExtemporaneoEsperado - LeerCeldaNumerica(insumos.Plantilla, WorkbookLeafCellMap.HojaConsolidado, "D47"), -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R2.TotalOportunoEsperado - LeerCeldaNumerica(insumos.Plantilla, WorkbookLeafCellMap.HojaConsolidado, "D28"), -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R4.TotalReversionEsperada - LeerCeldaNumerica(insumos.Plantilla, WorkbookLeafCellMap.HojaConsolidado, "D66"), -Tolerancia, Tolerancia);

        // A3 — fórmulas protegidas siguen siendo fórmula en la salida.
        foreach (var (hoja, celda, _) in WorkbookLeafCellMap.ProtectedFormulas)
        {
            Assert.True(CeldaEsFormula(salida, hoja, celda), $"{hoja}!{celda} debió seguir siendo fórmula.");
        }

        // A4 — plantilla origen intacta.
        Assert.Equal(origenHashAntes, Sha256(insumos.Plantilla));

        // A5 — este test no compara cache de D9/F46/etc. DE LA SALIDA contra golden.
        Assert.NotEqual(19556118465.99m, leaf.R1.TotalOportunoEsperado);
    }

    private static Ase CrearAse() => AseFactory.DesdeId(1); // HU-14 (S-3): factoría única

    private static Periodo CrearPeriodo() =>
        new() { CodigoAAAAMM = "202607", NumeroQuincena = 1 };

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
