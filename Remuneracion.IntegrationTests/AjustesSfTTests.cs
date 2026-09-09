using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-11 (2.5, plan §4.1/§5.2): tests de dominio + lectura de las fuentes Q2 reales
/// (SALDOS POR NOTA / RETRIBUCION NEGATIVA) sin salida. Cubre: composición T0-0.3 probada
/// contra el golden, Especiales ausente = 0, header obligatorio ausente = fallo ASE+reporte,
/// fuente vacía = 0 legítimo, locator agnóstico a rango/diacríticos, Balance-Optimizado Q2
/// (fallback obligatorio), TotOpt-vs-ajuste prohibido y gates de validador Q2.
/// </summary>
public sealed class AjustesSfTTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    /// <summary>
    /// Golden Q2 (T0-0.3, probado contra <c>Remuneracion 202607-2 Total.xlsx</c>): TotalAjustes
    /// por ASE = D85:D89 del CONSOLIDADO = AJUSTES-SF-T D47:D51.
    /// </summary>
    private static readonly decimal[] GoldenAjustes = [973693.46m, 216025.77m, 104231.83m, 35954.44m, 0m];

    [Theory]
    [InlineData(1, 973693.46)]
    [InlineData(2, 216025.77)]
    [InlineData(3, 104231.83)]
    [InlineData(4, 35954.44)]
    public void LeerSaldosNotas_AseConDatos_TotalMatcheaGolden(int aseId, decimal totalEsperado)
    {
        // Requirement 1/2: lectura header-driven de la fuente real; Total (col C fila Total)
        // = golden ±0.5; Especiales ausente → TieneColumnaEspeciales=false, ServEspK=0.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var saldos = reader.LeerSaldosNotas(Insumos.Ase(aseId), Insumos.SaldosNotas(aseId));

        Assert.False(saldos.TieneColumnaEspeciales);
        Assert.Equal(0m, saldos.ServEspK);
        Assert.InRange(saldos.Total - totalEsperado, -Tolerancia, Tolerancia);
        Assert.InRange(saldos.TotalSaldosNotas - totalEsperado, -Tolerancia, Tolerancia);
    }

    [Fact]
    public void LeerSaldosNotas_Ase5FuenteVacia_TotalCeroLegitimo()
    {
        // T0-0.5: ASE5-Q2 trae SaldosaFavorAplicadosPorNotas SOLO con la fila 1 (rango de fechas)
        // → 0 legítimo (el golden D89 = 0 lo confirma). No es un fallo de lectura.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var saldos = reader.LeerSaldosNotas(Insumos.Ase(5), Insumos.SaldosNotas(5));

        Assert.Equal(0m, saldos.Total);
        Assert.Equal(0m, saldos.TotalSaldosNotas);
        Assert.Empty(saldos.Celdas);
    }

    [Fact]
    public void LeerSaldosNotas_Ase2TraeDebCred_YSeMapeaAColumnaO()
    {
        // T0-0.5: ASE2 es la única fuente con columna "Deb/Cred" → el reader la mapea a la
        // columna O del template (opcional; ausente = 0). El Total no depende de ella.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var saldos = reader.LeerSaldosNotas(Insumos.Ase(2), Insumos.SaldosNotas(2));

        // Fila template 15 (Vlr Servicio) del bloque ASE2: O15 = Deb/Cred de la fuente (N4=-0.37).
        Assert.True(saldos.Celdas.ContainsKey("O15"));
        Assert.InRange(saldos.Celdas["O15"] - (-0.37m), -Tolerancia, Tolerancia);
    }

    [Fact]
    public void LeerRetribucionNegativa_Las5FuentesVacias_TotalCeroLegitimo()
    {
        // T0-0.4/0.5: las 5 fuentes RETRIBUCION-NEGATIVA Q2 traen solo la fila 1 (rango) → 0
        // legítimo, confirmado por el golden (D28:D32 = 0). Fuente vacía ≠ header ausente.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        for (var i = 1; i <= 5; i++)
        {
            var retribucion = reader.LeerRetribucionNegativa(Insumos.Ase(i), Insumos.RetribucionNegativa(i));
            Assert.Equal(0m, retribucion.Total);
            Assert.Equal(0m, retribucion.TotalRetribucionNegativa);
            Assert.Empty(retribucion.Celdas);
        }
    }

    [Fact]
    public void LeerSaldosNotas_HeaderObligatorioAusente_FallaNombrandoAse()
    {
        // Requirement 2 / §5.2: si falta un header obligatorio del mapa T0-0.5, fail-fast nombra
        // ASE + reporte; NUNCA valor inventado. Se simula con una fuente sintética SIN el header.
        var ruta = CrearFuenteSaldosSinHeader("Total");
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            reader.LeerSaldosNotas(Insumos.Ase(3), ruta));

        Assert.Contains("ASE 3", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("SALDOS POR NOTA", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ComponerAjustesSfT_PorAse_TotalAjustesGolden()
    {
        // Requirement 3 / T0-0.3: TotalAjustes = TotalSaldosNotas + TotalRetribucionNegativa
        // con las fuentes Q2 reales → golden D85:D89 ±0.5 en los 5 ASE.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        for (var i = 1; i <= 5; i++)
        {
            var ajustes = new AjustesSfTInputs
            {
                Ase = Insumos.Ase(i),
                SaldosNotas = reader.LeerSaldosNotas(Insumos.Ase(i), Insumos.SaldosNotas(i)),
                RetribucionNegativa = reader.LeerRetribucionNegativa(Insumos.Ase(i), Insumos.RetribucionNegativa(i))
            };

            Assert.InRange(ajustes.TotalAjustes - GoldenAjustes[i - 1], -Tolerancia, Tolerancia);
        }
    }

    [Fact]
    public void CalcularConsolidado_Q2ConAjustes_CuadraTotalesYGranTotal()
    {
        // Requirement 1: tuplas Q2 CON ajustes → 5 consolidados con AjustesSfT = fuente ±0.5
        // y GranTotal = Σ TotalAse.
        var (datos, _, _) = LeerDatosQ2ConAjustes();
        var periodo = Insumos.PeriodoQ2();
        var resultado = new CalculoRemuneracion().CalcularConsolidado(periodo, datos);

        Assert.Equal(5, resultado.Consolidados.Count);
        for (var i = 0; i < 5; i++)
        {
            Assert.InRange(resultado.Consolidados[i].AjustesSfT - GoldenAjustes[i], -Tolerancia, Tolerancia);
        }

        Assert.InRange(resultado.GranTotal - resultado.Consolidados.Sum(c => c.TotalAse), -Tolerancia, Tolerancia);
    }

    [Fact]
    public void CalcularConsolidado_Q2ConOverloadViejo_LanzaFailFast()
    {
        // Requirement 1 (red de seguridad): el overload viejo (sin ajustes) sigue lanzando para
        // Q2 — garantía contra cálculo silencioso incompleto.
        var (_, datosViejos, _) = LeerDatosQ2ConAjustes();
        var periodo = Insumos.PeriodoQ2();

        Assert.Throws<CalculoInvalidoException>(() =>
            new CalculoRemuneracion().CalcularConsolidado(periodo, datosViejos));
    }

    [Fact]
    public void Validar_Q2_AjustesCoincidenConFuente_SinErrores()
    {
        // §2.5 gates Q2: AjustesSfT == TotalAjustes ±0.5 por ASE (matcheo estricto por Id).
        // Leafs sintéticos in-memory (patrón ReporteBancoTests.CrearCasoValido): el gate de
        // ajustes es lógica pura de dominio; el leaf ASE5-Q2 completo no es construible con el
        // reader HU-07 (R1 divergente, recorte T0-0.6) — no afecta la prueba del gate.
        var (resultado, leafs) = CrearCasoQ2Sintetico();
        var errores = new ValidadorBasico().Validar(resultado, leafs);

        Assert.DoesNotContain(errores, e => e.Contains("AjustesSfT", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_Q2_AjustesDifierenDeFuente_NombraAse()
    {
        // §2.5 gate Q2 negativo: si el consolidado trae un ajuste distinto de la fuente, el
        // error nombra el ASE.
        var (resultado, leafs) = CrearCasoQ2Sintetico();
        resultado.Consolidados.Single(c => c.Ase.Id == 3).AjustesSfT += 1000m;

        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("ASE 3", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_Q2_LeafSinAjustes_NombraAse()
    {
        // §2.5 regla 2: en Q2 un leaf sin AjustesSfT (null) es error que nombra el ASE — nunca
        // 0 silencioso.
        var (resultado, leafs) = CrearCasoQ2Sintetico();
        leafs.Single(l => l.Ase.Id == 4).AjustesSfT = null;

        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("ASE 4", StringComparison.OrdinalIgnoreCase)
            && e.Contains("AjustesSfT", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Caso Q2 sintético (5 leafs in-memory con AjustesSfT) para probar los gates del validador
    /// sin depender del reader leaf Q2 (recorte T0-0.6). Los consolidados se calculan con el
    /// overload Q2; los totales de R1/R2/R4 son arbitrarios pero coherentes (el gate de ajustes
    /// no depende de ellos).
    /// </summary>
    private static (ResultadoRemuneracion Resultado, List<WorkbookLeafInputs> Leafs) CrearCasoQ2Sintetico()
    {
        var periodo = Insumos.PeriodoQ2();
        var datos = new List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4, decimal ajustesSfT)>();
        var leafs = new List<WorkbookLeafInputs>();

        for (var i = 1; i <= 5; i++)
        {
            var ase = Insumos.Ase(i);
            var r1 = new RecaudoComponenteR1 { TotalOportuno = 1000m + i, Extemporaneo = 100m + i };
            var r2 = new SaldosFavorR2 { GrandTotal = 2000m + i, ServEspK = 0m };
            var r4 = new ReversionR4 { TotalReversiones = -(300m + i) };

            var ajustes = new AjustesSfTInputs
            {
                Ase = ase,
                SaldosNotas = new SaldosNotasAseInputs { Ase = ase, Total = 500m + i, ServEspK = 0m },
                RetribucionNegativa = new RetribucionNegativaAseInputs { Ase = ase, Total = 50m + i, ServEspK = 0m }
            };

            var leaf = new WorkbookLeafInputs
            {
                Ase = ase,
                Periodo = periodo,
                R1 = new WorkbookLeafInputsR1 { F25 = 100m + i, F41 = 200m + i, L25 = 0m, F30 = 0m, F10 = 0m, L10 = 0m },
                R2 = new WorkbookLeafInputsR2 { E15 = 2000m + i, E26 = 0m, K15 = 0m },
                R4 = new WorkbookLeafInputsR4 { D9 = -(300m + i), P9 = 0m },
                AjustesSfT = ajustes,
                DetRetriQ2 = new DetRetriQ2Inputs
                {
                    Ase = ase,
                    TotalD104 = (2000m + i) + (-(300m + i)) + ajustes.TotalAjustes // = R2.TotalOportunoEsperado + R4.TotalReversionEsperada + TotalAjustes
                }
            };

            datos.Add((ase, r1, r2, r4, ajustes.TotalAjustes));
            leafs.Add(leaf);
        }

        var resultado = new CalculoRemuneracion().CalcularConsolidado(periodo, datos);
        return (resultado, leafs);
    }

    [Fact]
    public void Q1_Regresion_AjustesSfTCeroEnLos5_Intacto()
    {
        // Requirement 1 / V7: Q1 sigue exigiendo AjustesSfT == 0 (regresión ciega 78/78).
        var (resultado, leafs) = ReporteBancoTests.CrearCasoValido();
        Assert.All(resultado.Consolidados, c => Assert.Equal(0m, c.AjustesSfT));
        Assert.All(leafs, l => Assert.Null(l.AjustesSfT));
        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs));
    }

    [Fact]
    public void Locator_BuscarSaldosNotas_ResuelveLos5Q2()
    {
        // Requirement 6: los 5 ASE resuelven por prefijo sin fechas (rango ASE4 ≠ resto).
        var locator = new ArchivoFuenteLocator();
        for (var i = 1; i <= 5; i++)
        {
            var ruta = locator.BuscarSaldosNotas(Insumos.CarpetasAseQ2[i - 1]);
            Assert.NotNull(ruta);
            Assert.StartsWith("SaldosaFavorAplicadosPorNotas", Path.GetFileNameWithoutExtension(ruta!), StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void Locator_BuscarRetribucionNegativa_DiacriticosYRangoAse4Resuelven()
    {
        // Requirement 6 / V9: el nombre en disco lleva tilde (RetribuciónNegativa) y ASE4 trae
        // rango 16072026–31072026 → el matcher normalizado (sin acentos, sin fechas) resuelve.
        var locator = new ArchivoFuenteLocator();
        for (var i = 1; i <= 5; i++)
        {
            var ruta = locator.BuscarRetribucionNegativa(Insumos.CarpetasAseQ2[i - 1]);
            Assert.NotNull(ruta);
            Assert.StartsWith("Retribuci", Path.GetFileNameWithoutExtension(ruta!), StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void BalanceOptimizado_Q2_ResuelvePorSegundoPrefijo()
    {
        // Requirement 5 (OBLIGATORIO, V4): ASE5-Q2 solo trae la variante -Optimizado → el
        // fallback de BuscarBalance pasa de opcional a ejercitado. Sin base ni Optimizado →
        // fail-fast que nombra el ASE.
        var locator = new ArchivoFuenteLocator();
        var ruta = locator.BuscarBalance(Insumos.CarpetasAseQ2[4]);
        Assert.NotNull(ruta);
        Assert.StartsWith("R4-BalanceSubsidioyContribuciones-Optimizado", Path.GetFileNameWithoutExtension(ruta!), StringComparison.OrdinalIgnoreCase);

        // Negativa: carpeta sin Balance (base ni Optimizado) → null (el procesador fail-fast nombra ASE).
        var carpetaVacia = Path.Combine(Path.GetTempPath(), "remuneracion-sin-balance-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(carpetaVacia);
        Assert.Null(locator.BuscarBalance(carpetaVacia));
    }

    [Fact]
    public void Prohibido_TotOptComoValorDeAjuste_NoCuadra()
    {
        // Requirement A5 / §5.2: PROHIBIDO usar TotOpt (HU-02) como valor de ajuste. El test
        // demuestra la confusión fallando: TotOpt de ASE1 (agregado HU-02) ≠ AjustesSfT golden.
        var lector = new ExcelDataReaderRecaudoReader();
        var r1 = lector.LeerR1(Insumos.R1(1));
        var r2 = lector.LeerR2(Insumos.R2(1));
        var r4 = lector.LeerR4(Insumos.R4(1));

        var totOpt = r1.TotalOportuno;
        Assert.NotEqual(GoldenAjustes[0], totOpt);

        var consolidado = new CalculoRemuneracion().Calcular(Insumos.Ase(1), r1, r2, r4, totOpt);
        Assert.NotEqual(GoldenAjustes[0], consolidado.AjustesSfT);
    }

    [Fact]
    public void WriterQ2_AjustesSfTNull_ComportamientoHu10Intacto()
    {
        // Plan §8 / G3: AjustesSfT == null = comportamiento HU-10 puro (Q1). El procesador Q1
        // real genera la salida con leafs completos (CeldasPorAse) sin tocar la cadena 2.5.
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-writer-q1-" + Guid.NewGuid().ToString("N"));
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
            RutaPlantilla = Insumos.Plantilla,
            RutaSalida = salida
        });

        Assert.True(File.Exists(salida));
        Assert.All(resultado.Leafs, l => Assert.Null(l.AjustesSfT));
    }

    /// <summary>
    /// Lee R1/R2/R4 + 2.5 de las fuentes Q2 reales y arma leafs completos (sin escritura).
    /// HU-12 (2.6 ampliada): el dispatch Q2 del reader leaf (mapa <see cref="WorkbookLeafCellMapQ2"/>
    /// con variante ASE5 de 2 filas Mes/Total, V0.3) levanta el recorte T0-0.6 de HU-11 → los 5
    /// leafs se construyen (R1/R2/R4 Q2 legibles). La conciliación por empresa (HU-08) NO se
    /// ejecuta en Q2 (recorte 1 HU-11, V0.6: layout R4 por empresa divergente; el mapa HU-08 está
    /// congelado para Q1).
    /// </summary>
    internal static (List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4, decimal ajustesSfT)> Datos,
        List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4)> DatosViejos,
        List<WorkbookLeafInputs> Leafs) LeerDatosQ2ConAjustes()
    {
        var lector = new ExcelDataReaderRecaudoReader();
        var leafReader = new ExcelDataReaderWorkbookLeafInputReader();
        var periodo = Insumos.PeriodoQ2();

        var datos = new List<(Ase, RecaudoComponenteR1, SaldosFavorR2, ReversionR4, decimal)>();
        var datosViejos = new List<(Ase, RecaudoComponenteR1, SaldosFavorR2, ReversionR4)>();
        var leafs = new List<WorkbookLeafInputs>();

        for (var i = 1; i <= 5; i++)
        {
            var ase = Insumos.Ase(i);
            var r1 = lector.LeerR1(Insumos.R1Q2(i));
            var r2 = lector.LeerR2(Insumos.R2Q2(i));
            var r4 = lector.LeerR4(Insumos.R4Q2(i));

            var leaf = leafReader.LeerLeafInputs(ase, periodo, Insumos.R1Q2(i), Insumos.R2Q2(i), Insumos.R4Q2(i));
            leaf.ReporteBanco = leafReader.LeerReporteBanco(ase, periodo, Insumos.ReporteBancoQ2(i));
            leaf.BalanceSc = leafReader.LeerBalanceSc(ase, periodo, Insumos.BalanceQ2(i));
            leaf.AjustesSfT = new AjustesSfTInputs
            {
                Ase = ase,
                SaldosNotas = leafReader.LeerSaldosNotas(ase, Insumos.SaldosNotas(i)),
                RetribucionNegativa = leafReader.LeerRetribucionNegativa(ase, Insumos.RetribucionNegativa(i))
            };

            datos.Add((ase, r1, r2, r4, leaf.AjustesSfT.TotalAjustes));
            datosViejos.Add((ase, r1, r2, r4));
            leafs.Add(leaf);
        }

        return (datos, datosViejos, leafs);
    }

    /// <summary>
    /// TotalAjustes de los 5 ASE-Q2 con los readers 2.5 (independiente del reader leaf HU-07;
    /// certificable en los 5 pese al recorte T0-0.6).
    /// </summary>
    internal static IReadOnlyList<AjustesSfTInputs> LeerAjustesQ2Los5()
    {
        var leafReader = new ExcelDataReaderWorkbookLeafInputReader();
        var ajustes = new List<AjustesSfTInputs>();
        for (var i = 1; i <= 5; i++)
        {
            ajustes.Add(new AjustesSfTInputs
            {
                Ase = Insumos.Ase(i),
                SaldosNotas = leafReader.LeerSaldosNotas(Insumos.Ase(i), Insumos.SaldosNotas(i)),
                RetribucionNegativa = leafReader.LeerRetribucionNegativa(Insumos.Ase(i), Insumos.RetribucionNegativa(i))
            });
        }

        return ajustes;
    }

    /// <summary>
    /// Fuente sintética SALDOS-NOTAS sin el header indicado (para fail-fast de header ausente).
    /// </summary>
    private static string CrearFuenteSaldosSinHeader(string headerOmitido)
    {
        var ruta = Path.Combine(Path.GetTempPath(), "saldos-sintetico-" + Guid.NewGuid().ToString("N") + ".xlsx");
        using (var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(ruta, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            var sheetPart = workbookPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorksheetPart>();
            var headers = new[] { "Total", "Componente TDF", "Componente TTL", "Componente TVIAT", "Aprovechamiento", "CCSA Prest.Aprov.", "Componente TCS", "Componente TLU", "Componente TBL", "Componente TRT", "CCSA Prest. No Aprov." };
            var celdasHeaders = new List<DocumentFormat.OpenXml.Spreadsheet.Cell>();
            for (var i = 0; i < headers.Length; i++)
            {
                if (string.Equals(headers[i], headerOmitido, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                celdasHeaders.Add(new DocumentFormat.OpenXml.Spreadsheet.Cell
                {
                    CellReference = $"{(char)('C' + i)}{3}",
                    DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                    CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(headers[i])
                });
            }

            sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new DocumentFormat.OpenXml.Spreadsheet.SheetData(
                new DocumentFormat.OpenXml.Spreadsheet.Row(new DocumentFormat.OpenXml.Spreadsheet.Cell
                {
                    CellReference = "A1",
                    DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String,
                    CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Recaudo Desde: 01/07/2026 Hasta: 31/07/2026")
                }),
                new DocumentFormat.OpenXml.Spreadsheet.Row(celdasHeaders),
                new DocumentFormat.OpenXml.Spreadsheet.Row(
                    new DocumentFormat.OpenXml.Spreadsheet.Cell { CellReference = "B4", DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String, CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("Vlr Servicio") },
                    new DocumentFormat.OpenXml.Spreadsheet.Cell { CellReference = "C4", CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("100") })));

            var sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
            sheets.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheet { Id = workbookPart.GetIdOfPart(sheetPart), SheetId = 1, Name = "Sheet1" });
            workbookPart.Workbook.Save();
        }

        return ruta;
    }
}