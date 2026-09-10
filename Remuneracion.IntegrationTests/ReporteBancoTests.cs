using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-09 (2.3, plan §4.1/§5.2): gates D5 del reporte por banco in-memory + lectura del
/// "Resumen Recaudo Aplicado Por Servicio" con fuentes Q1 reales (sin salida). Cubre:
/// Σ bloques == TotalFuente, variante "FINANCIACIONES" sin NUEVAS matchea, C59 != quincena
/// falla, mismatch nombra ASE+empresa-columna, TotOpt HU-02 NO es oráculo de valores banco.
/// </summary>
public sealed class ReporteBancoTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    [Fact]
    public void LeerReporteBanco_Ase1Q1_ValoresExactosVsGolden()
    {
        // Plan §2.1: bloque ASE1 golden EXACTO: ENEL 16194517449/50286069/0/0;
        // OCCIDENTE 523782724.81/12466967.19/4226000/0.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(Insumos.Ase(1), Insumos.Periodo(), Insumos.ReporteBanco(1));

        var enel = resultado.Ases[0].Empresas.Single(e => e.Empresa == "ENEL");
        Assert.InRange(enel.AplicadosFacturacion - 16194517449m, -Tolerancia, Tolerancia);
        Assert.InRange(enel.SaldosFavorGenerados - 50286069m, -Tolerancia, Tolerancia);
        Assert.Equal(0m, enel.FinanciacionesNuevas);
        Assert.Equal(0m, enel.RecibosServEspeciales);
        Assert.InRange(enel.Total - 16244803518m, -Tolerancia, Tolerancia);
        Assert.InRange(enel.TotalFuente - 16244803518m, -Tolerancia, Tolerancia);

        var occ = resultado.Ases[0].Empresas.Single(e => e.Empresa == "OCCIDENTE");
        Assert.InRange(occ.AplicadosFacturacion - 523782724.81m, -Tolerancia, Tolerancia);
        Assert.InRange(occ.SaldosFavorGenerados - 12466967.19m, -Tolerancia, Tolerancia);
        Assert.InRange(occ.FinanciacionesNuevas - 4226000m, -Tolerancia, Tolerancia);
        Assert.Equal(0m, occ.RecibosServEspeciales);
        Assert.InRange(occ.Total - 540475692m, -Tolerancia, Tolerancia);

        Assert.Equal(1, resultado.Quincena);
        Assert.Null(resultado.Consolidado); // D2(b): filas 1–7 son fórmulas (T0-0.1).
    }

    [Fact]
    public void LeerReporteBanco_Ase2Q1_TresEmpresasSinFinanciaciones()
    {
        // ASE2 LIME: ENEL + NUEVO ESQUEMA + OCCIDENTE; sin fila de FINANCIACIONES (0 demostrable).
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(Insumos.Ase(2), Insumos.Periodo(), Insumos.ReporteBanco(2));

        Assert.Equal(3, resultado.Ases[0].Empresas.Count);
        var enel = resultado.Ases[0].Empresas.Single(e => e.Empresa == "ENEL");
        var nuevo = resultado.Ases[0].Empresas.Single(e => e.Empresa == "NUEVO ESQUEMA");
        var occ = resultado.Ases[0].Empresas.Single(e => e.Empresa == "OCCIDENTE");

        Assert.InRange(enel.Total - 11833210912m, -Tolerancia, Tolerancia);
        Assert.InRange(nuevo.Total - 29345550m, -Tolerancia, Tolerancia);
        Assert.InRange(occ.Total - 422221285m, -Tolerancia, Tolerancia);
        Assert.Equal(0m, enel.FinanciacionesNuevas); // fila ausente = 0 demostrable
        Assert.Equal(0m, nuevo.RecibosServEspeciales);
    }

    [Fact]
    public void LeerReporteBanco_Ase4Q1_OccidenteExisteConValores()
    {
        // T0-0.3: en ASE4 (Bogotá Limpia) OCCIDENTE EXISTE (columna c3) con valores.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(Insumos.Ase(4), Insumos.Periodo(), Insumos.ReporteBanco(4));

        var occ = resultado.Ases[0].Empresas.Single(e => e.Empresa == "OCCIDENTE");
        Assert.InRange(occ.Total - 239142386m, -Tolerancia, Tolerancia);
        Assert.InRange(occ.FinanciacionesNuevas - 26456711.27m, -Tolerancia, Tolerancia);
    }

    [Fact]
    public void LeerReporteBanco_Ase5Q1_EnerbitPresente()
    {
        // T0-0.3: ENERBIT solo en ASE5 (columna c2).
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(Insumos.Ase(5), Insumos.Periodo(), Insumos.ReporteBanco(5));

        var enerbit = resultado.Ases[0].Empresas.Single(e => e.Empresa == "ENERBIT");
        Assert.InRange(enerbit.Total - 129536656m, -Tolerancia, Tolerancia);
    }

    [Fact]
    public void LeerReporteBanco_VarianteFinanciacionesSinNuevas_Matchea()
    {
        // Plan §5.2 / T0-0.4: la variante "3-APLICADOS A FINANCIACIONES" (sin NUEVAS) debe
        // matchear por prefijo normalizado. Fixture sintético con OpenXML (in-memory, sin salida).
        var ruta = CrearReporteSinteticoSinNuevas();
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(Insumos.Ase(1), Insumos.Periodo(), ruta);

        var enel = resultado.Ases[0].Empresas.Single(e => e.Empresa == "ENEL");
        var occ = resultado.Ases[0].Empresas.Single(e => e.Empresa == "OCCIDENTE");
        Assert.InRange(enel.AplicadosFacturacion - 100m, -Tolerancia, Tolerancia);
        Assert.InRange(enel.SaldosFavorGenerados - 10m, -Tolerancia, Tolerancia);
        Assert.Equal(0m, enel.FinanciacionesNuevas);
        Assert.InRange(enel.Total - 110m, -Tolerancia, Tolerancia);
        Assert.InRange(occ.FinanciacionesNuevas - 20m, -Tolerancia, Tolerancia);
        Assert.InRange(occ.Total - 75m, -Tolerancia, Tolerancia);
    }

    [Fact]
    public void LeerReporteBanco_SinEtiquetaResumen_FallaNombrandoAse()
    {
        // Requirement 1: si falta la etiqueta del Resumen → fallo que nombra el ASE.
        var ruta = CrearReporteSinteticoSinEtiqueta();
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            reader.LeerReporteBanco(Insumos.Ase(3), Insumos.Periodo(), ruta));

        Assert.Contains("ASE 3", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Resumen Recaudo Aplicado Por Servicio", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Validar_CasoValidoConReporteBanco_SinErrores()
    {
        var (resultado, leafs) = CrearCasoValido();
        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs));
    }

    [Fact]
    public void Validar_C59NoCoincideConQuincena_Falla()
    {
        // Requirement 4: C59 == Periodo.NumeroQuincena; si no, error.
        var (resultado, leafs) = CrearCasoValido();
        leafs[2].ReporteBanco!.Quincena = 2; // período es Q1

        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("C59", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(errores, e => e.Contains("ASE 3", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_MismatchBloqueVsFuente_NombraAseYEmpresa()
    {
        // Gate D5-i: bloque banco != resumen fuente → error nombra ASE y empresa.
        var (resultado, leafs) = CrearCasoValido();
        var occAse3 = leafs[2].ReporteBanco!.Ases[0].Empresas.Single(e => e.Empresa == "OCCIDENTE");
        occAse3.TotalFuente += 1000m;

        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("ASE 3", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(errores, e => e.Contains("OCCIDENTE", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_TotOptHu02NoEsOraculoDeValoresBanco()
    {
        // Plan §2.5 regla 5 / §4.1: prohibido usar TotOpt HU-02 como valor banco. El test
        // demuestra que son cantidades DISTINTAS y que el validador NO mezcla oráculos
        // (el caso válido pasa aunque TotOpt ≠ Total banco).
        var (resultado, leafs) = CrearCasoValido();

        var bancoEnelAse1 = leafs[0].ReporteBanco!.Ases[0].Empresas.Single(e => e.Empresa == "ENEL").Total;
        var totOptAse1 = resultado.Consolidados.Single(c => c.Ase.Id == 1).TotOpt;
        Assert.NotEqual(totOptAse1, bancoEnelAse1); // son magnitudes distintas (V9: la hoja banco resume TODO el recaudo)

        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs));
    }

    [Fact]
    public void Validar_ReporteBancoNull_ComportamientoHu08Intacto()
    {
        // Plan §8 / G4: lista 2.3 vacía (ReporteBanco == null) = comportamiento HU-08 puro.
        var (resultado, leafs) = CrearCasoValido();
        foreach (var leaf in leafs)
        {
            leaf.ReporteBanco = null;
        }

        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs));
    }

    /// <summary>
    /// Caso Q1 válido: leafs HU-07/HU-08 (reutiliza el helper de ConciliacionEmpresaTests) +
    /// ReporteBanco leído de las fuentes Q1 reales (el reader ya verifica Σ conceptos == TotalFuente).
    /// </summary>
    internal static (ResultadoRemuneracion Resultado, List<WorkbookLeafInputs> Leafs) CrearCasoValido()
    {
        var (resultado, leafs) = ConciliacionEmpresaTests.CrearCasoValido();
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        for (var i = 0; i < leafs.Count; i++)
        {
            leafs[i].ReporteBanco = reader.LeerReporteBanco(
                Insumos.Ase(i + 1),
                Insumos.Periodo(),
                Insumos.ReporteBanco(i + 1));
        }

        return (resultado, leafs);
    }

    /// <summary>
    /// Fixture sintético: Resumen con "3-APLICADOS A FINANCIACIONES" (sin NUEVAS) que debe
    /// matchear por prefijo normalizado (T0-0.4).
    /// </summary>
    private static string CrearReporteSinteticoSinNuevas()
    {
        var ruta = Path.Combine(Path.GetTempPath(), "reporte-sin-nuevas-" + Guid.NewGuid().ToString("N") + ".xlsx");
        using (var doc = SpreadsheetDocument.Create(ruta, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var sheetPart = workbookPart.AddNewPart<WorksheetPart>();
            sheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("Resumen Recaudo Aplicado Por Servicio") }),
                new Row(
                    new Cell { CellReference = "A2", DataType = CellValues.String, CellValue = new CellValue("CONCEPTO") },
                    new Cell { CellReference = "B2", DataType = CellValues.String, CellValue = new CellValue("ENEL") },
                    new Cell { CellReference = "C2", DataType = CellValues.String, CellValue = new CellValue("OCCIDENTE") },
                    new Cell { CellReference = "D2", DataType = CellValues.String, CellValue = new CellValue("Total") }),
                new Row(
                    new Cell { CellReference = "A3", DataType = CellValues.String, CellValue = new CellValue("1-APLICADOS A FACTURACION") },
                    new Cell { CellReference = "B3", CellValue = new CellValue("100") },
                    new Cell { CellReference = "C3", CellValue = new CellValue("50") }),
                new Row(
                    new Cell { CellReference = "A4", DataType = CellValues.String, CellValue = new CellValue("2-SALDOS A FAVOR GENERADOS") },
                    new Cell { CellReference = "B4", CellValue = new CellValue("10") },
                    new Cell { CellReference = "C4", CellValue = new CellValue("5") }),
                new Row(
                    new Cell { CellReference = "A5", DataType = CellValues.String, CellValue = new CellValue("3-APLICADOS A FINANCIACIONES") },
                    new Cell { CellReference = "B5", CellValue = new CellValue("0") },
                    new Cell { CellReference = "C5", CellValue = new CellValue("20") }),
                new Row(
                    new Cell { CellReference = "A6", DataType = CellValues.String, CellValue = new CellValue("Total") },
                    new Cell { CellReference = "B6", CellValue = new CellValue("110") },
                    new Cell { CellReference = "C6", CellValue = new CellValue("75") })));
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(sheetPart), SheetId = 1, Name = "Sheet1" });
            workbookPart.Workbook.Save();
        }

        return ruta;
    }

    /// <summary>
    /// Fixture sintético: Resumen sin la etiqueta (para el fail-fast del Requirement 1).
    /// </summary>
    private static string CrearReporteSinteticoSinEtiqueta()
    {
        var ruta = Path.Combine(Path.GetTempPath(), "reporte-sin-etiqueta-" + Guid.NewGuid().ToString("N") + ".xlsx");
        using (var doc = SpreadsheetDocument.Create(ruta, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var sheetPart = workbookPart.AddNewPart<WorksheetPart>();
            sheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("OTRA COSA") })));
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(sheetPart), SheetId = 1, Name = "Sheet1" });
            workbookPart.Workbook.Save();
        }

        return ruta;
    }
}
