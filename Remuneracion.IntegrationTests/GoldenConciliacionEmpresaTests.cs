using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-08 (2.2) plan §2.7: Golden Capa A extendida a 2.2 con insumos Q1 reales.
/// Matriz: A1 leaf 2.2 escritos en la SALIDA vs golden, A2 visibles de dominio por empresa vs
/// caché golden de <c>REMUNERACION_*</c>, A3 fórmulas protegidas 2.2, A4 hash plantilla,
/// A5 prohibido comparar caché de la SALIDA (OpenXML no recalcula), A6 Recaudo * en valores.
/// </summary>
public sealed class GoldenConciliacionEmpresaTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    [Fact]
    public void ProcesadorPeriodo_CapaA2_2_EmpresasYRecaudosContraGolden()
    {
        var insumos = ObtenerInsumos();
        var origenHashAntes = Sha256(insumos);
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-hu08-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator());

        var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.Periodo(),
            CarpetaPeriodo = Insumos.CarpetaPeriodo,
            RutaPlantilla = insumos,
            RutaSalida = salida
        });

        Assert.True(File.Exists(salida));

        // La lectura 2.2 llena las 5 empresas por ASE + las 5 hojas Recaudo *.
        foreach (var leaf in resultado.Leafs)
        {
            Assert.Equal(5, leaf.Conciliacion.Count);
            Assert.Equal(5, leaf.Recaudos.Count);
        }

        // ---- A1: celdas 2.2 escritas en la SALIDA vs mismas celdas leaf del golden ----
        foreach (var leaf in resultado.Leafs)
        {
            foreach (var conc in leaf.Conciliacion)
            {
                VerificarCeldas(conc.CeldasR1, WorkbookLeafCellMap.HojaR1, salida, insumos);
                VerificarCeldas(conc.CeldasR2, WorkbookLeafCellMap.HojaR2, salida, insumos);
                VerificarCeldas(conc.CeldasR4, WorkbookLeafCellMap.HojaR4, salida, insumos);
            }
        }

        // ---- A2: visibles de DOMINIO por empresa vs caché golden de REMUNERACION_* ----
        // REMUNERACION_{empresa}: D9:D13 = R1 OPORTUNO, D28:D32 = R2, D66:D70 = R4.
        foreach (var leaf in resultado.Leafs)
        {
            foreach (var conc in leaf.Conciliacion)
            {
                Assert.InRange(
                    conc.VisibleR1 - LeerCeldaNumerica(insumos, conc.Empresa.HojaRemuneracion, $"D{8 + leaf.Ase.Id}"),
                    -Tolerancia, Tolerancia);
                Assert.InRange(
                    conc.VisibleR2 - LeerCeldaNumerica(insumos, conc.Empresa.HojaRemuneracion, $"D{27 + leaf.Ase.Id}"),
                    -Tolerancia, Tolerancia);
                Assert.InRange(
                    conc.VisibleR4 - LeerCeldaNumerica(insumos, conc.Empresa.HojaRemuneracion, $"D{65 + leaf.Ase.Id}"),
                    -Tolerancia, Tolerancia);
            }
        }

        // ---- A1/A6: Recaudo * escritos en la SALIDA vs golden (valores, no fórmulas) ----
        foreach (var recaudo in resultado.Leafs[0].Recaudos)
        {
            foreach (var (celda, _) in recaudo.Celdas)
            {
                var escrito = LeerCeldaNumerica(salida, recaudo.HojaRecaudo, celda);
                var golden = LeerCeldaNumerica(insumos, recaudo.HojaRecaudo, celda);
                Assert.InRange(escrito - golden, -Tolerancia, Tolerancia);
                Assert.False(CeldaEsFormula(salida, recaudo.HojaRecaudo, celda),
                    $"{recaudo.HojaRecaudo}!{celda} debe ser valor (no fórmula) en la salida.");
            }

            // Referencias Q1 (Conciliaciones): ENEL OPORTUNO ASE1 = 16.020.970.408, TOTAL = 55.775.447.693.
            if (recaudo.Empresa.Id == 2)
            {
                Assert.InRange(recaudo.TotalOportuno - 55775447693m, -Tolerancia, Tolerancia);
            }
        }

        // ---- A3: fórmulas protegidas 2.2 siguen siendo fórmula en la salida ----
        foreach (var (hoja, celda, _) in WorkbookLeafCellMapPorEmpresa.ProtectedFormulasPorEmpresa)
        {
            Assert.True(CeldaEsFormula(salida, hoja, celda), $"{hoja}!{celda} debió seguir siendo fórmula.");
        }

        // REMUNERACION_ENEL D9 = 'Reporte Componentes R1'!F56 (verificado T0) — sigue fórmula.
        Assert.True(CeldaEsFormula(salida, "REMUNERACION_ENEL", "D9"));

        // ---- A4: plantilla origen intacta ----
        Assert.Equal(origenHashAntes, Sha256(insumos));

        // ---- A5: este test NO compara caché de fórmula DE LA SALIDA contra golden ----
        // (OpenXML no recalcula; A5 del plan lo prohíbe). El Recaudo ENEL D3 golden = 16020970408
        // es valor de insumo, no resultado de recálculo.
        Assert.Equal(16020970408m, LeerCeldaNumerica(insumos, "Recaudo ENEL", "D3"));
    }

    private static void VerificarCeldas(
        IReadOnlyDictionary<string, decimal> celdas,
        string hoja,
        string salida,
        string insumos)
    {
        foreach (var (celda, _) in celdas)
        {
            var escrito = LeerCeldaNumerica(salida, hoja, celda);
            var golden = LeerCeldaNumerica(insumos, hoja, celda);
            Assert.InRange(escrito - golden, -Tolerancia, Tolerancia);
            Assert.False(CeldaEsFormula(salida, hoja, celda), $"{hoja}!{celda} debió ser valor en la salida.");
        }
    }

    private static string ObtenerInsumos()
    {
        Assert.True(File.Exists(Insumos.Plantilla), $"Falta plantilla: {Insumos.Plantilla}");
        return Insumos.Plantilla;
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