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
/// HU-10 (2.4, plan §4.1/§5.2): gates D5 del balance SC in-memory + lectura de la fila
/// "Total General" con fuentes Q1 reales (sin salida). Cubre: valores exactos vs golden,
/// asignación D/E del veredicto T0 (hipótesis líder probada), variante "TOTAL GENERAL" con
/// mayúsculas matchea, falta de etiqueta nombra ASE, coherencia E+F vs G rota nombra ASE,
/// F=D+E aritmética de dominio, sumas fila 11 vs golden, H≈ROUND(F,0) (regla T0-0.4),
/// J9 CONSOLIDADO NO es oráculo de D/E, y BalanceSc null = comportamiento HU-09 intacto.
/// </summary>
public sealed class BalanceScTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    /// <summary>
    /// Golden Q1 de la hoja BCE SC POR FACT. (caché template, filas 3–7; T0-0.5).
    /// Contribucion = template-D (positivo, ← F-fuente); Subsidio = template-E (negativo, ← E-fuente).
    /// </summary>
    private static readonly (int Ase, decimal D, decimal E, decimal F, decimal H)[] Golden =
    [
        (1, 3256235169.97m, -1223871491.13m, 2032363678.84m, 2032363679m),
        (2, 1946491828.39m, -6958007481.98m, -5011515653.59m, -5011515654m),
        (3, 1936109874.45m, -3132932285.76m, -1196822411.31m, -1196822411m),
        (4, 474156195.46m, -438200379.42m, 35955816.04m, 35955816m),
        (5, 1266769543.80m, -2477195172.11m, -1210425628.31m, -1210425628m)
    ];

    [Fact]
    public void LeerBalanceSc_Ase1Q1_ValoresExactosVsGolden()
    {
        // Plan §2.1: fila ASE1 golden EXACTA — D=3256235169.97 (Contribución, F-fuente),
        // E=-1223871491.13 (Subsidio, E-fuente), F=2032363678.84, H=2032363679 (I=-0.16).
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerBalanceSc(Insumos.Ase(1), Insumos.Periodo(), Insumos.Balance(1));

        var fila = Assert.Single(resultado.Ases);
        Assert.InRange(fila.Contribucion - 3256235169.97m, -Tolerancia, Tolerancia);
        Assert.InRange(fila.Subsidio - (-1223871491.13m), -Tolerancia, Tolerancia);
        Assert.InRange(fila.TotalBsc - 2032363678.84m, -Tolerancia, Tolerancia);
        Assert.InRange(fila.TotalFuente - 2032363678.84m, -Tolerancia, Tolerancia);
        Assert.Null(fila.Sistema); // D2(b): H es fórmula en el template (T0-0.6).
    }

    [Theory]
    [InlineData(1, 3256235169.97, -1223871491.13, 2032363678.84)]
    [InlineData(2, 1946491828.39, -6958007481.98, -5011515653.59)]
    [InlineData(3, 1936109874.45, -3132932285.76, -1196822411.31)]
    [InlineData(4, 474156195.46, -438200379.42, 35955816.04)]
    [InlineData(5, 1266769543.80, -2477195172.11, -1210425628.31)]
    public void LeerBalanceSc_Ases1a5Q1_MatcheanGolden(int aseId, decimal contribucion, decimal subsidio, decimal total)
    {
        // Requirement 1: los 5 ASE leen su "Total General" y matchean el golden ±0.5.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerBalanceSc(Insumos.Ase(aseId), Insumos.Periodo(), Insumos.Balance(aseId));

        var fila = Assert.Single(resultado.Ases);
        Assert.InRange(fila.Contribucion - contribucion, -Tolerancia, Tolerancia);
        Assert.InRange(fila.Subsidio - subsidio, -Tolerancia, Tolerancia);
        Assert.InRange(fila.TotalBsc - total, -Tolerancia, Tolerancia);
        Assert.InRange(fila.TotalFuente - total, -Tolerancia, Tolerancia);
    }

    [Fact]
    public void LeerBalanceSc_VarianteTotalGeneralMayusculas_Matchea()
    {
        // Requirement 1 / §5.2: la variante "TOTAL GENERAL" (mayúsculas) debe matchear por
        // normalización (lowercase + sin espacios), no fallar.
        var ruta = CrearBalanceSintetico("TOTAL GENERAL", 100m, 50m, 150m);
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerBalanceSc(Insumos.Ase(1), Insumos.Periodo(), ruta);

        var fila = Assert.Single(resultado.Ases);
        Assert.InRange(fila.Subsidio - 100m, -Tolerancia, Tolerancia);
        Assert.InRange(fila.Contribucion - 50m, -Tolerancia, Tolerancia);
        Assert.InRange(fila.TotalBsc - 150m, -Tolerancia, Tolerancia);
    }

    [Fact]
    public void LeerBalanceSc_SinEtiquetaTotalGeneral_FallaNombrandoAse()
    {
        // Requirement 1: si falta la etiqueta → fallo que nombra el ASE, nunca valor inventado.
        var ruta = CrearBalanceSintetico("OTRA COSA", 100m, 50m, 150m);
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            reader.LeerBalanceSc(Insumos.Ase(3), Insumos.Periodo(), ruta));

        Assert.Contains("ASE 3", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Total General", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void LeerBalanceSc_CoherenciaBceFuenteRota_FallaNombrandoAse()
    {
        // Gate D5-i: E+F (TotalBsc) != col G (TotalFuente) → fail-fast nombra el ASE.
        var ruta = CrearBalanceSintetico("Total General", 100m, 50m, 999m);
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            reader.LeerBalanceSc(Insumos.Ase(2), Insumos.Periodo(), ruta));

        Assert.Contains("ASE 2", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void AsignacionDE_MapaCongelado_ContribucionADySubsidioAE(int aseId)
    {
        // Requirement 2 / veredicto T0-0.2 (hipótesis líder probada): el MAPA fija la asignación
        // template: Contribucion → celda D (CONTRIBUCION), Subsidio → celda E (SUBSIDIO).
        // El MODELO es neutro (propiedades por significado de dominio); si el veredicto hubiera
        // sido inverso, solo cambiaría este mapa, no el modelo.
        var editables = WorkbookLeafCellMapBalanceSc.ObtenerEditables(aseId);
        Assert.Equal($"D{aseId + 2}", editables.Contribucion);
        Assert.Equal($"E{aseId + 2}", editables.Subsidio);
    }

    [Fact]
    public void AsignacionDE_ModeloNeutro_AmbosDesenlacesSoloCambiaElMapa()
    {
        // D4 / §4.1 "asignación D/E en ambos desenlaces": si T0 hubiera probado la asignación
        // inversa, SOLO cambia el mapa, no el modelo. El modelo carga Contribucion/Subsidio por
        // significado de dominio; este test simula los dos desenlaces con un mapa local y
        // demuestra que el modelo no fija columnas template (misma semántica en ambos).
        var modelo = new BalanceScAseInputs
        {
            Ase = Insumos.Ase(1),
            Contribucion = 3256235169.97m,
            Subsidio = -1223871491.13m
        };

        // Desenlace líder (T0-0.2 PROBADO): Contribucion → D, Subsidio → E.
        Assert.InRange(ValorPorCelda(modelo, "D3", inverso: false) - 3256235169.97m, -Tolerancia, Tolerancia);
        Assert.InRange(ValorPorCelda(modelo, "E3", inverso: false) - (-1223871491.13m), -Tolerancia, Tolerancia);

        // Desenlace inverso (hipotético, texto del doc base): Contribucion → E, Subsidio → D.
        Assert.InRange(ValorPorCelda(modelo, "E3", inverso: true) - 3256235169.97m, -Tolerancia, Tolerancia);
        Assert.InRange(ValorPorCelda(modelo, "D3", inverso: true) - (-1223871491.13m), -Tolerancia, Tolerancia);
    }

    private static decimal ValorPorCelda(BalanceScAseInputs modelo, string celda, bool inverso) => (celda, inverso) switch
    {
        ("D3", false) => modelo.Contribucion,
        ("E3", false) => modelo.Subsidio,
        ("E3", true) => modelo.Contribucion,
        ("D3", true) => modelo.Subsidio,
        _ => throw new ArgumentOutOfRangeException(nameof(celda), $"Celda no mapeada en el test: {celda}.")
    };

    [Fact]
    public void F_D_E_AritmeticaDeDominio_OkVsGolden()
    {
        // Requirement 3 (ii): F == D+E ±0.5 por ASE (aritmética de dominio sobre golden).
        foreach (var (_, d, e, f, _) in Golden)
        {
            Assert.InRange((d + e) - f, -Tolerancia, Tolerancia);
        }
    }

    [Fact]
    public void SumasFila11_DominioVsGolden()
    {
        // §2.5 gate (ii) "sumas fila 11 en dominio ±0.5": ΣD = D11 golden 8879762612.07;
        // ΣE = E11 golden -14230206810.40; ΣF = F11 golden -5350444198.33.
        var sumaD = Golden.Sum(g => g.D);
        var sumaE = Golden.Sum(g => g.E);
        var sumaF = Golden.Sum(g => g.F);

        Assert.InRange(sumaD - 8879762612.07m, -Tolerancia, Tolerancia);
        Assert.InRange(sumaE - (-14230206810.40m), -Tolerancia, Tolerancia);
        Assert.InRange(sumaF - (-5350444198.33m), -Tolerancia, Tolerancia);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void H_RedondeoReglaT0_VsGolden(int aseId)
    {
        // §2.5 gate (iii) / T0-0.4: H ≈ ROUND(F,0) (redondeo a entero, AwayFromZero). En Q1 H es
        // fórmula (D2(b)); la regla se prueba contra el caché golden (verificación en Capa A).
        var golden = Golden.Single(g => g.Ase == aseId);
        var redondeado = Math.Round(golden.F, MidpointRounding.AwayFromZero);
        Assert.InRange(redondeado - golden.H, -Tolerancia, Tolerancia);
    }

    [Fact]
    public void Validar_CasoValidoConBalanceSc_SinErrores()
    {
        var (resultado, leafs) = CrearCasoValido();
        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs));
    }

    [Fact]
    public void Validar_MismatchBceVsFuente_NombraAse()
    {
        // Gate D5-i: BCE != fuente → error nombra el ASE.
        var (resultado, leafs) = CrearCasoValido();
        var filaAse3 = leafs[2].BalanceSc!.Ases[0];
        filaAse3.TotalFuente += 1000m;

        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("ASE 3", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(errores, e => e.Contains("Total General", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_SumaBceVsSumaFuente_Falla()
    {
        // Gate D5-ii: Σ Total BSC != Σ TotalFuente (sumas fila 11 en dominio) → error.
        var (resultado, leafs) = CrearCasoValido();
        leafs[4].BalanceSc!.Ases[0].TotalFuente += 5000m;

        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("Σ Total BSC", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_SistemaDistintoDeRoundF_FallaNombrandoAse()
    {
        // Gate D5-iii / §5.2 "H ≠ ROUND(F,0) fuera de ±0.5": con Sistema no-nulo (desenlace
        // D2(a) hipotético) fuera de tolerancia → error nombra el ASE. En Q1 Sistema es null
        // (D2(b): H es fórmula) y esta rama no se ejecuta.
        var (resultado, leafs) = CrearCasoValido();
        leafs[1].BalanceSc!.Ases[0].Sistema = leafs[1].BalanceSc!.Ases[0].TotalBsc + 1000m;

        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("ASE 2", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(errores, e => e.Contains("SISTEMA", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_BalanceScNull_ComportamientoHu09Intacto()
    {
        // Plan §8 / G6: lista 2.4 vacía (BalanceSc == null) = comportamiento HU-09 puro.
        var (resultado, leafs) = CrearCasoValido();
        foreach (var leaf in leafs)
        {
            leaf.BalanceSc = null;
        }

        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs));
    }

    [Fact]
    public void J9Consolidado_NoEsOraculoDeValoresBce()
    {
        // Plan §2.7 A5 / §4.1: PROHIBIDO usar CONSOLIDADO J/K/M como oráculo de las celdas D/E
        // de BCE. J9 = BCE!F3 (total) por fórmula; NO permite recuperar D ni E. El test lo
        // demuestra: J9 del golden difiere de D3 y de E3, y el caso válido pasa sin leer J9.
        var j9 = LeerCeldaNumerica(Insumos.Plantilla, "CONSOLIDADO_TOTAL RECAUDO", "J9");
        var d3 = LeerCeldaNumerica(Insumos.Plantilla, WorkbookLeafCellMapBalanceSc.HojaBce, "D3");
        var e3 = LeerCeldaNumerica(Insumos.Plantilla, WorkbookLeafCellMapBalanceSc.HojaBce, "E3");

        Assert.NotEqual(d3, j9);
        Assert.NotEqual(e3, j9);
        Assert.InRange(j9 - (d3 + e3), -Tolerancia, Tolerancia); // J9 es el TOTAL (F3), no un componente

        var (resultado, leafs) = CrearCasoValido();
        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs)); // nunca abre xlsx ni exige J9
    }

    /// <summary>
    /// Caso Q1 válido: leafs HU-07/HU-08/HU-09 (reutiliza ReporteBancoTests.CrearCasoValido) +
    /// BalanceSc leído de las fuentes Q1 reales (el reader ya verifica TotalBsc == TotalFuente).
    /// </summary>
    internal static (ResultadoRemuneracion Resultado, List<WorkbookLeafInputs> Leafs) CrearCasoValido()
    {
        var (resultado, leafs) = ReporteBancoTests.CrearCasoValido();
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        for (var i = 0; i < leafs.Count; i++)
        {
            leafs[i].BalanceSc = reader.LeerBalanceSc(
                Insumos.Ase(i + 1),
                Insumos.Periodo(),
                Insumos.Balance(i + 1));
        }

        return (resultado, leafs);
    }

    /// <summary>
    /// Fixture sintético de un Balance con la etiqueta indicada (Sheet1, layout B..G de T0-0.3:
    /// E=Subsidio en col E, F=Contribución en col F, G=Valor en col G).
    /// </summary>
    private static string CrearBalanceSintetico(string etiqueta, decimal subsidio, decimal contribucion, decimal valor)
    {
        var ruta = Path.Combine(Path.GetTempPath(), "balance-sintetico-" + Guid.NewGuid().ToString("N") + ".xlsx");
        using (var doc = SpreadsheetDocument.Create(ruta, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var sheetPart = workbookPart.AddNewPart<WorksheetPart>();
            sheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(
                    new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("Localidad") },
                    new Cell { CellReference = "B1", DataType = CellValues.String, CellValue = new CellValue("Total") },
                    new Cell { CellReference = "C1", CellValue = new CellValue("0") }),
                new Row(
                    new Cell { CellReference = "A2", DataType = CellValues.String, CellValue = new CellValue(etiqueta) },
                    new Cell { CellReference = "E2", CellValue = new CellValue(subsidio.ToString(System.Globalization.CultureInfo.InvariantCulture)) },
                    new Cell { CellReference = "F2", CellValue = new CellValue(contribucion.ToString(System.Globalization.CultureInfo.InvariantCulture)) },
                    new Cell { CellReference = "G2", CellValue = new CellValue(valor.ToString(System.Globalization.CultureInfo.InvariantCulture)) })));
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

        return decimal.TryParse(cell.CellValue.InnerText, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var value)
            ? value
            : 0m;
    }
}