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
/// HU-09 (2.3, plan §2.7): Golden Capa A extendida a 2.3 con insumos Q1 reales.
/// A1: celdas banco escritas en la SALIDA vs mismas celdas leaf del golden ±0.5.
/// A2: visibles de dominio (bloque = fuente y Σ consolidado = Σ bloques) vs caché golden
///     de bloques y filas 1–7.
/// A3: filas 1–7, Total de bloque, validación 59–80 y TOTAL siguen siendo fórmula en la salida.
/// A4: hash plantilla intacta. A5: prohibido comparar caché de la SALIDA; prohibido usar
/// TotOpt HU-02 / Recaudo * HU-08 como oráculo de valores banco.
/// </summary>
public sealed class GoldenReporteBancoTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;
    private const string HojaBanco = WorkbookLeafCellMapReporteBanco.HojaBanco;

    [Fact]
    public void ProcesadorPeriodo_CapaA2_3_BloquesYConsolidadoContraGolden()
    {
        var insumos = Insumos.Plantilla;
        Assert.True(File.Exists(insumos), $"Falta plantilla: {insumos}");
        var origenHashAntes = Sha256(insumos);
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-hu09-" + Guid.NewGuid().ToString("N"));
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

        // Los 5 leafs traen su bloque banco (empresas del mapa T0-0.7).
        foreach (var leaf in resultado.Leafs)
        {
            Assert.NotNull(leaf.ReporteBanco);
            Assert.Single(leaf.ReporteBanco!.Ases);
            Assert.True(leaf.ReporteBanco.Ases[0].Empresas.Count >= 2);
            Assert.Equal(1, leaf.ReporteBanco.Quincena);
            Assert.Null(leaf.ReporteBanco.Consolidado); // D2(b)
        }

        // ---- A1: celdas banco escritas en la SALIDA vs mismas celdas leaf del golden ----
        // Referencias golden de las celdas de concepto (T0-0.1): ENEL C13/C23/C32/C41/C51,
        // NUEVO ESQUEMA F23/F41, ENERBIT G51, OCCIDENTE H13/H23/H32/H41/H51 (concepto 1).
        var esperadosGolden = new (int Ase, string Empresa, string Celda, decimal Esperado)[]
        {
            (1, "ENEL", "C13", 16194517449m),
            (1, "OCCIDENTE", "H13", 523782724.81m),
            (2, "ENEL", "C23", 11785775189m),
            (2, "NUEVO ESQUEMA", "F23", 13570060m),
            (2, "OCCIDENTE", "H23", 390172374m),
            (3, "ENEL", "C32", 9664176949m),
            (3, "OCCIDENTE", "H32", 466628267.89m),
            (4, "ENEL", "C41", 10340898687m),
            (4, "NUEVO ESQUEMA", "F41", 7264980m),
            (4, "OCCIDENTE", "H41", 200757639.13m),
            (5, "ENEL", "C51", 8240568549m),
            (5, "ENERBIT", "G51", 126372915m),
            (5, "OCCIDENTE", "H51", 178634316.15m)
        };

        foreach (var (aseId, empresa, celda, esperado) in esperadosGolden)
        {
            var escrito = LeerCeldaNumerica(salida, HojaBanco, celda);
            var golden = LeerCeldaNumerica(insumos, HojaBanco, celda);
            Assert.InRange(escrito - golden, -Tolerancia, Tolerancia);
            Assert.InRange(escrito - esperado, -Tolerancia, Tolerancia);
            Assert.False(CeldaEsFormula(salida, HojaBanco, celda), $"{HojaBanco}!{celda} debió ser valor en la salida.");
        }

        // C59 = quincena (valor) en la salida (Requirement 4).
        Assert.InRange(LeerCeldaNumerica(salida, HojaBanco, "C59") - 1m, -Tolerancia, Tolerancia);
        Assert.False(CeldaEsFormula(salida, HojaBanco, "C59"), "C59 debió ser valor (quincena).");

        // ---- A2: visibles de DOMINIO vs caché golden ----
        // (i) Bloque = fuente: Total bloque por empresa (Σ conceptos) == caché golden del
        // Total de bloque (C17/H17/F27/H27/...) ±0.5 (V5).
        var totalesGolden = new (string Celda, decimal Esperado)[]
        {
            ("C17", 16244803518m), ("H17", 540475692m),
            ("C27", 11833210912m), ("F27", 29345550m), ("H27", 422221285m),
            ("C36", 9685219461m), ("H36", 472519065m),
            ("C45", 10376694382m), ("F45", 15240310m), ("H45", 239142386m),
            ("C55", 8278977659m), ("G55", 129536656m), ("H55", 181733562m)
        };

        foreach (var leaf in resultado.Leafs)
        {
            foreach (var empresa in leaf.ReporteBanco!.Ases[0].Empresas)
            {
                // El Total de bloque del golden se obtiene del caché de la fórmula (C17 etc.).
                var celdaTotal = ObtenerCeldaTotalGolden(leaf.Ase.Id, empresa.Empresa);
                var goldenTotal = LeerCeldaNumerica(insumos, HojaBanco, celdaTotal);
                Assert.InRange(empresa.Total - goldenTotal, -Tolerancia, Tolerancia);
            }
        }

        // (ii) Σ consolidado 1–7 por empresa (aritmética de dominio) == caché golden filas 1–7
        // (C2..H7 e I6) ±0.5 (G3: verificación Σ; nunca caché de la SALIDA).
        var sumaConsolidado = SumaConsolidadoPorEmpresa(resultado.Leafs);
        var consolidadoGolden = new (string Celda, decimal Esperado)[]
        {
            ("C2", 56225936823m),   // ENEL 1-APLICADOS
            ("F2", 20835040m),      // NUEVO ESQUEMA 1-APLICADOS
            ("G2", 126372915m),     // ENERBIT 1-APLICADOS
            ("H2", 1759975321.98m), // OCCIDENTE 1-APLICADOS
            ("I2", 58133120099.98m),
            ("C3", 192969109m),     // ENEL 2-SALDOS
            ("F3", 23750820m),      // NUEVO ESQUEMA 2-SALDOS
            ("G3", 3163741m),       // ENERBIT 2-SALDOS
            ("H3", 61262182.79m),   // OCCIDENTE 2-SALDOS
            ("I3", 281145852.79m),
            ("H4", 34854485.23m),   // OCCIDENTE 3-FINANCIACIONES
            ("I4", 34854485.23m),
            ("C6", 56418905932m),   // ENEL Total
            ("F6", 44585860m),      // NUEVO ESQUEMA Total
            ("G6", 129536656m),     // ENERBIT Total
            ("H6", 1856091990m),    // OCCIDENTE Total
            ("I6", 58449120438.00m) // Gran Total consolidado
        };

        foreach (var (celda, esperado) in consolidadoGolden)
        {
            var golden = LeerCeldaNumerica(insumos, HojaBanco, celda);
            Assert.InRange(golden - esperado, -Tolerancia, Tolerancia);
        }

        // Σ por empresa (dominio) == caché golden C6 (Total ENEL), F6, G6, H6.
        Assert.InRange(sumaConsolidado["ENEL"] - LeerCeldaNumerica(insumos, HojaBanco, "C6"), -Tolerancia, Tolerancia);
        Assert.InRange(sumaConsolidado["NUEVO ESQUEMA"] - LeerCeldaNumerica(insumos, HojaBanco, "F6"), -Tolerancia, Tolerancia);
        Assert.InRange(sumaConsolidado["ENERBIT"] - LeerCeldaNumerica(insumos, HojaBanco, "G6"), -Tolerancia, Tolerancia);
        Assert.InRange(sumaConsolidado["OCCIDENTE"] - LeerCeldaNumerica(insumos, HojaBanco, "H6"), -Tolerancia, Tolerancia);

        // ---- A3: fórmulas protegidas 2.3 siguen siendo fórmula en la salida ----
        foreach (var (hoja, celda, _) in WorkbookLeafCellMapReporteBanco.Protegidas)
        {
            Assert.True(CeldaEsFormula(salida, hoja, celda), $"{hoja}!{celda} debió seguir siendo fórmula.");
        }

        // ---- A4: plantilla origen intacta ----
        Assert.Equal(origenHashAntes, Sha256(insumos));

        // ---- A5: este test NO compara caché de la SALIDA contra golden ----
        // (OpenXML no recalcula). Tampoco usa TotOpt HU-02 como oráculo: el banco ENEL ASE1
        // (16.244.803.518) ≠ TotOpt HU-02 del consolidado (~19.556.118.465.99).
        var bancoEnelAse1 = resultado.Leafs[0].ReporteBanco!.Ases[0].Empresas.Single(e => e.Empresa == "ENEL").Total;
        var totOptAse1 = resultado.Resultado.Consolidados.Single(c => c.Ase.Id == 1).TotOpt;
        Assert.NotEqual(totOptAse1, bancoEnelAse1);
    }

    private static string ObtenerCeldaTotalGolden(int aseId, string empresa) => (aseId, empresa) switch
    {
        (1, "ENEL") => "C17",
        (1, "OCCIDENTE") => "H17",
        (2, "ENEL") => "C27",
        (2, "NUEVO ESQUEMA") => "F27",
        (2, "OCCIDENTE") => "H27",
        (3, "ENEL") => "C36",
        (3, "OCCIDENTE") => "H36",
        (4, "ENEL") => "C45",
        (4, "NUEVO ESQUEMA") => "F45",
        (4, "OCCIDENTE") => "H45",
        (5, "ENEL") => "C55",
        (5, "ENERBIT") => "G55",
        (5, "OCCIDENTE") => "H55",
        _ => throw new ArgumentOutOfRangeException(nameof(aseId), $"Sin celda total golden para ASE {aseId} · {empresa}.")
    };

    private static IReadOnlyDictionary<string, decimal> SumaConsolidadoPorEmpresa(IReadOnlyList<WorkbookLeafInputs> leafs) =>
        leafs
            .Where(l => l.ReporteBanco is not null)
            .SelectMany(l => l.ReporteBanco!.Ases)
            .SelectMany(b => b.Empresas)
            .GroupBy(e => e.Empresa, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Sum(e => e.Total), StringComparer.OrdinalIgnoreCase);

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