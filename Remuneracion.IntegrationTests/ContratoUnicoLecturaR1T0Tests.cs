using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 54 (H0, S1.1-S1.6): evidencia T0 del contrato ÚNICO de lectura R1 (Q1/Q2), leída por BYTES
/// con BCL (ZIP + XML crudo). Sin OpenXML SDK para leer: el oráculo son las etiquetas <c>&lt;f&gt;</c>
/// y los literales <c>&lt;v&gt;</c> del paquete, nunca el caché de una celda con fórmula.
///
/// Reglas del plan que este fixture hace cumplir:
/// - CA0: todo número de criterio viaja con su derivación (fuente × <c>&lt;f&gt;</c> + hoja/celda).
/// - S1.1 (T0.1): las 10 <c>&lt;f&gt;</c> TOT_OPT/EXTEMP de la v2 y los slots editables literales.
/// - S1.2 (T0.2): censo octubre-1 por ASE (filas <c>Mes/Total</c>, <c>Aplicacion</c>, cierre <c>Total</c>).
/// - S1.3 (T0.3): derivación al centavo de D9/D28/D47/D66/D104 y paridad con el R10 oficial.
/// - S1.4 (T0.4): caracterización R2/R4 Q1-vs-Q2 sobre fixtures reales (julio-1 + julio-2) y destinos
///   literales escribibles por bytes.
/// - S1.5 (T0.5): geometría v1 congelada como legacy + forma-fuente R1 por período (dato, no supuesto).
///
/// Rutas EXPLÍCITAS (familias de carpeta + sufijo de archivo congelados). Prohibido <c>Insumos.Raiz()</c>:
/// este fixture resuelve la raíz por su cuenta.
/// </summary>
public sealed class ContratoUnicoLecturaR1T0Tests
{
    // ── G-SHA (CA6 / D8): plantillas oficiales INTOCABLES (valor medido en disco, 64 hex) ────────
    private const string ShaPlantillaV2 = "0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB";
    private const string ShaPlantillaV1 = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";

    private const string HojaR1 = "Reporte Componentes R1";
    private const string HojaR2 = "Rem. Anticipos R2";
    private const string HojaR4 = "Reversion Pagos R4";

    // ── Rutas explícitas ────────────────────────────────────────────────────────────────────
    private static string PlantillaV2 => Ruta("Docs", "Plantilla estandarizada", "v2", "Formato Remuneracion 2026xx-x Total vf.xlsx");

    private static string PlantillaV1 => Ruta("Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string R10Octubre1 => Ruta("Docs", "Prueba Octubre-1", "Insumos", "R10_Remuneracion_2026101.xlsx");

    /// <summary>Octubre-1 (caso rojo del plan): carpeta + sufijo de archivo por ASE.</summary>
    private static readonly (int Ase, string Carpeta, string R1, string R2, string R4)[] Octubre1 =
    [
        (1, "1-Promoambiental",
            "Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174133882.xlsx",
            "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174159251.xlsx",
            "ReversiónPorComponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174217316.xlsx"),
        (2, "2-Lime",
            "Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174350151.xlsx",
            "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___2026101617441360.xlsx",
            "ReversiónPorComponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174430392.xlsx"),
        (3, "3-Ciudad Limpia",
            "Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174753997.xlsx",
            "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174813931.xlsx",
            "ReversiónPorComponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174833303.xlsx"),
        (4, "4-Bogota Limpia",
            "Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174948502.xlsx",
            "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___202610161750610.xlsx",
            "ReversiónPorComponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016175022764.xlsx"),
        (5, "5-Área Limpia",
            "Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016175148433.xlsx",
            "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___202610161752844.xlsx",
            "ReversiónPorComponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___202610161752242.xlsx")
    ];

    private static string CarpetaAse(int aseId) => aseId switch
    {
        1 => "1-Promoambiental",
        2 => "2-Lime",
        3 => "3-Ciudad Limpia",
        4 => "4-Bogota Limpia",
        5 => "5-Área Limpia",
        _ => throw new ArgumentOutOfRangeException(nameof(aseId))
    };

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.1 / G-SHA — plantillas intactas (CA6, D8)
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void T0_1_PlantillasOficiales_ShaIntacto()
    {
        Assert.Equal(ShaPlantillaV2, Sha256(PlantillaV2));
        Assert.Equal(ShaPlantillaV1, Sha256(PlantillaV1));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.1 (S1.1) — geometría v2 R1: las 10 <f> TOT_OPT/EXTEMP y los slots literales
    // ════════════════════════════════════════════════════════════════════════════════════════

    public static TheoryData<int, string, string> FormulasTotOptExtemp => new()
    {
        { 1, "F75", "F47+F70+F16-L16-L47" },
        { 1, "F77", "F59+F32-L32" },
        { 2, "F254", "F171+F208-L171+F120-L120" },
        { 2, "F256", "F190+F146-L146" },
        { 3, "F415", "F330+F353+F299-L299-L330" },
        { 3, "F417", "F342+F315-L315" },
        { 4, "F568", "F511+F548+F460-L460-L511" },
        { 4, "F570", "F530+F486-L486" },
        { 5, "F680", "F674+F648+F613-L613-L648" },
        { 5, "F682", "F660+F629-L629" }
    };

    [Theory]
    [MemberData(nameof(FormulasTotOptExtemp))]
    public void T0_1_GeometriaV2_FormulasTotOptYExtemp_PorBytes(int aseId, string celda, string formulaEsperada)
    {
        var cruda = CeldaCrudaDe(PlantillaV2, HojaR1, celda);

        Assert.True(cruda.EsFormula, $"{HojaR1}!{celda} (ASE{aseId}) debe ser FÓRMULA en la v2.");
        Assert.Equal(formulaEsperada, NormalizarFormula(cruda.Formula));
    }

    /// <summary>
    /// Slots editables R1 v2 por ASE (T0.1): LITERALES sin <c>&lt;f&gt;</c>, exactamente los que declara
    /// el mapa productivo (<see cref="WorkbookLeafCellMapQ2.R1Q2EditablesPorAse"/>), incluidos los
    /// <c>L{Aplic0}</c> (L32/L146/L315/L486/L629). ASE5 no declara slot <c>Mes2</c> (D6).
    /// </summary>
    public static TheoryData<int, string[]> SlotsR1V2 => new()
    {
        { 1, ["F16", "L16", "F32", "L32", "F47", "L47", "F59", "F70"] },
        { 2, ["F120", "L120", "F146", "L146", "F171", "L171", "F190", "F208"] },
        { 3, ["F299", "L299", "F315", "L315", "F330", "L330", "F342", "F353"] },
        { 4, ["F460", "L460", "F486", "L486", "F511", "L511", "F530", "F548"] },
        { 5, ["F613", "L613", "F629", "L629", "F648", "F660"] }
    };

    [Theory]
    [MemberData(nameof(SlotsR1V2))]
    public void T0_1_SlotsR1V2_SonLiteralesSinFormula(int aseId, string[] celdas)
    {        foreach (var celda in celdas)
        {
            var cruda = CeldaCrudaDe(PlantillaV2, HojaR1, celda);
            Assert.False(cruda.EsFormula, $"{HojaR1}!{celda} (ASE{aseId}) debió ser VALOR editable (sin <f>) en la v2.");
            Assert.True(cruda.Valor is not null, $"Falta {HojaR1}!{celda} (ASE{aseId}) en la v2.");
            Assert.Equal(0m, NumeroDe(cruda.Valor));
        }

        // Simetría mapa ↔ bytes: el mapa productivo declara exactamente estas celdas.
        var declaradas = WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(aseId).Select(p => p.Celda).ToArray();
        Assert.Equal(celdas.OrderBy(c => c, StringComparer.Ordinal), declaradas.OrderBy(c => c, StringComparer.Ordinal));
    }

    /// <summary>
    /// D6: el slot <c>Mes2</c> de ASE5 (<c>F674</c>) NO está en el mapa (forma-fuente de 2 filas) y en
    /// el template vale literal <c>0</c>; su visible <c>F680</c> lo consume como término nulo.
    /// </summary>
    [Fact]
    public void T0_1_Ase5_F674_LiteralCeroFueraDelMapa()
    {
        var cruda = CeldaCrudaDe(PlantillaV2, HojaR1, "F674");

        Assert.False(cruda.EsFormula);
        Assert.Equal(0m, NumeroDe(cruda.Valor));
        Assert.DoesNotContain(WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(5), p => p.Celda == "F674");
        Assert.Contains("F674", NormalizarFormula(CeldaCrudaDe(PlantillaV2, HojaR1, "F680").Formula), StringComparison.Ordinal);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.2 (S1.2) — censo octubre-1 por bytes
    // ════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Censo congelado (§0.2-F5 + T0 ejecutado 2026-10-22). <c>FilasMes</c>/<c>FilasAplicacion</c> son
    /// números de fila 1-based del XML; <c>Cierre</c> es el valor de la columna <c>Total</c> de la fila
    /// <c>A='Total' &amp; B</c> vacío.
    /// </summary>
    public static TheoryData<int, int[], int[], int, decimal> CensoOctubre1 => new()
    {
        { 1, [8, 33, 52], [18, 41], 54, 14835602537.78m },
        { 2, [11, 38, 57], [23, 46], 59, 11238640614.90m },
        { 3, [8, 24, 36], [], 38, 8830777878.00m },
        { 4, [11, 46, 70], [31, 59], 72, 10504168126.00m },
        { 5, [23, 38], [], 40, 7556633218.00m }
    };

    [Theory]
    [MemberData(nameof(CensoOctubre1))]
    public void T0_2_CensoOctubre1_PorBytes(int aseId, int[] filasMes, int[] filasAplicacion, int filaCierre, decimal cierreEsperado)
    {
        var filas = LeerFilas(RutaOctubre1(aseId, R1: true));

        Assert.Equal(filasMes, filas.Where(f => EsMesTotal(f.Celdas)).Select(f => f.Numero).ToArray());
        Assert.Equal(filasAplicacion, filas.Where(f => EsAplicacionTotal(f.Celdas)).Select(f => f.Numero).ToArray());

        var filaTotal = filas.FirstOrDefault(f => Texto(f.Celdas, 0).Equals("Total", StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrWhiteSpace(Texto(f.Celdas, 1)));
        Assert.NotNull(filaTotal);
        Assert.Equal(filaCierre, filaTotal!.Numero);
        Assert.Equal(cierreEsperado, Numero(filaTotal.Celdas[5]));

        // La fila 'Total' de la fuente NO es un visible: es Σ Mes + Σ Aplic (no resta Especiales).
        var suma = filas.Where(f => EsMesTotal(f.Celdas) || EsAplicacionTotal(f.Celdas)).Sum(f => Numero(f.Celdas[5]));
        Assert.Equal(cierreEsperado, suma);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.3 (S1.3, CA0/CA2/CA3/CA7) — derivación al centavo + paridad con el R10 oficial
    // ════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Derivación (CA0) por ASE: <c>D9</c> (TOT_OPT = Σ F(Mes) − Σ Esp(Mes sin la última)),
    /// <c>D28</c> (R2: E(Componente)+E(Subs/Cont)−K(Especiales)), <c>D47</c> (EXTEMP =
    /// Σ F(Aplic) − Esp(Aplic[0])), <c>D66</c> (R4: D(Total)), <c>D104</c> y el <c>RECAUDO TOTAL</c>
    /// del R10 oficial (hoja <c>DetRetri2026101</c>, D9:D13).
    /// </summary>
    public static TheoryData<int, decimal, decimal, decimal, decimal, decimal, long> DerivacionR10 => new()
    {
        { 1, 14815021974.75m, 30588071.02m, 648870.00m, -3833780.00m, 14842425135.77m, 14842425136L },
        { 2, 11219101305.09m, 95634974.32m, 9704974.00m, -21803648.45m, 11302637604.96m, 11302637605L },
        { 3, 8811472160.34m, 14574412.03m, 0m, -2259272.00m, 8823787300.37m, 8823787300L },
        { 4, 10473967813.42m, 29206357.07m, 21557177.00m, -41558585.40m, 10483172762.09m, 10483172762L },
        { 5, 7544253279.94m, 10379542.05m, 0m, -3964157.00m, 7550668664.99m, 7550668665L }
    };

    [Theory]
    [MemberData(nameof(DerivacionR10))]
    public void T0_3_DerivacionD104_ContraR10Oficial(
        int aseId, decimal d9Esperado, decimal d28Esperado, decimal d47Esperado, decimal d66Esperado,
        decimal d104Esperado, long r10Esperado)
    {
        var r1 = LeerFilas(RutaOctubre1(aseId, R1: true));
        var r2 = LeerFilas(RutaOctubre1(aseId, R1: false, R2: true));
        var r4 = LeerFilas(RutaOctubre1(aseId, R1: false, R4: true));

        var d9 = TotalOportunoR1(r1);
        var d28 = TotalOportunoR2(r2);
        var d47 = ExtemporaneoR1(r1);
        var d66 = TotalR4(r4);

        Assert.Equal(d9Esperado, d9);
        Assert.Equal(d28Esperado, d28);
        Assert.Equal(d47Esperado, d47);
        Assert.Equal(d66Esperado, d66);

        var d104 = d9 + d28 + d47 + d66; // D85 (AJUSTES-SF-T) = 0 en Q1: regla de negocio.
        Assert.Equal(d104Esperado, d104);

        // CA2/CA3: el oráculo es el R10 OFICIAL del período (nunca el log de una corrida).
        var r10 = MonedaAEntero(CeldaCrudaDe(R10Octubre1, "DetRetri2026101", $"D{8 + aseId}").Valor);
        Assert.Equal(r10Esperado, r10);
        Assert.Equal(r10Esperado, Redondear(d104));
    }

    /// <summary>
    /// CA7: el EXTEMP v2 es <c>Σ Aplic − Esp(Aplic[0])</c>; el término v1 <c>subsidio[0] + Aplic[0]</c>
    /// (F12) queda FUERA. Octubre-1 ASE1: Aplic fila 18 = 636.047,07 y fila 41 = 12.822,93;
    /// subsidio[0] fila 35 = −32.099,64.
    /// </summary>
    [Fact]
    public void T0_3_CA7_ExtempDeAse1_ExcluyeElSubsidioV1()
    {
        var r1 = LeerFilas(RutaOctubre1(1, R1: true));

        var aplicaciones = r1.Where(f => EsAplicacionTotal(f.Celdas)).ToList();
        Assert.Equal(2, aplicaciones.Count);
        Assert.Equal(636047.07m, Numero(aplicaciones[0].Celdas[5]));
        Assert.Equal(12822.93m, Numero(aplicaciones[1].Celdas[5]));

        var subsidio = r1.First(f => EsSubsidio(f.Celdas));
        Assert.Equal(-32099.64m, Numero(subsidio.Celdas[5]));

        Assert.Equal(648870.00m, ExtemporaneoR1(r1));

        // El agregado v1 (F30 + F10 − L10 = subsidio[0] + Aplic[0]) NO es el EXTEMP de la v2.
        const decimal extempV1 = -32099.64m + 636047.07m; // 603.947,43
        Assert.NotEqual(extempV1, ExtemporaneoR1(r1));
    }

    /// <summary>
    /// CA3: ASE3 y ASE5 traen CERO filas <c>Aplicacion</c> ⇒ EXTEMP = 0,00 (sus slots Aplic están en
    /// el 0 del template). No es un valor "faltante": es la forma de la fuente.
    /// </summary>
    [Theory]
    [InlineData(3)]
    [InlineData(5)]
    public void T0_3_CeroAplicaciones_ExtempEsCero(int aseId)
    {
        var r1 = LeerFilas(RutaOctubre1(aseId, R1: true));

        Assert.DoesNotContain(r1, f => EsAplicacionTotal(f.Celdas));
        Assert.Equal(0m, ExtemporaneoR1(r1));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.4 (S1.4) — caracterización R2/R4: Q1 ≡ Q2 sobre las MISMAS fuentes reales
    // ════════════════════════════════════════════════════════════════════════════════════════

    public static TheoryData<int, string> FixturesCaracterizacion => new()
    {
        { 1, "julio-1" }, { 2, "julio-1" }, { 3, "julio-1" }, { 4, "julio-1" }, { 5, "julio-1" },
        { 1, "julio-2" }, { 2, "julio-2" }, { 3, "julio-2" }, { 4, "julio-2" }, { 5, "julio-2" }
    };

    /// <summary>
    /// Combos en los que una de las dos ramas NO puede ni leer el R1 de la fuente (aborto por ordinal
    /// congelado, ajeno a R2/R4). Se declaran aquí como DATO: la comparación de R2/R4 se hace donde
    /// ambas ramas viven, y el conjunto observado de abortos debe ser SUBCONJUNTO de esta lista
    /// (tras la unificación la lista debe quedar vacía sin editar el test).
    /// </summary>
    private static readonly HashSet<(int Ase, string Fixture, string Rama)> AbortosDeLecturaConocidos =
    [
        // Q1 exigia el rol core Mes2 y julio-1 ASE1 trae 2 filas Mes (post H1: Mes2 es opcional).
        (1, "julio-1", "Q2"),
        // Q1 (MapearR1 case 5) exige 3 filas Mes y julio-2 ASE5 trae 2.
        (5, "julio-2", "Q1")
    ];

    /// <summary>
    /// T0.4 (R-TABLA): la rama R2 unica y la rama R4 unica deben
    /// producir VALORES idénticos sobre el mismo archivo. Se resuelven por la API pública con dos
    /// <see cref="Periodo"/> distintos sobre las MISMAS rutas: la única variable es la quincena.
    /// </summary>
    [Theory]
    [MemberData(nameof(FixturesCaracterizacion))]
    public void T0_4_R2R4_Q1IdenticoAQ2(int aseId, string fixture)
    {
        var (r1, r2, r4) = RutasFixture(aseId, fixture);
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var ase = AseFactory.DesdeId(aseId);

        WorkbookLeafInputs? q1 = null;
        WorkbookLeafInputs? q2 = null;

        try
        {
            q1 = reader.LeerLeafInputs(ase, Periodo.Parse("2026071"), r1, r2, r4);
        }
        catch (CalculoInvalidoException ex)
        {
            Assert.True(AbortosDeLecturaConocidos.Contains((aseId, fixture, "Q1")),
                $"Aborto Q1 NO declarado en el fixture ({aseId}, {fixture}): {ex.Message}");
        }

        try
        {
            q2 = reader.LeerLeafInputs(ase, Periodo.Parse("2026072"), r1, r2, r4);
        }
        catch (CalculoInvalidoException ex)
        {
            Assert.True(AbortosDeLecturaConocidos.Contains((aseId, fixture, "Q2")),
                $"Aborto Q2 NO declarado en el fixture ({aseId}, {fixture}): {ex.Message}");
        }

        Assert.True(q1 is not null || q2 is not null,
            $"La fuente ({aseId}, {fixture}) no es legible por NINGUNA de las dos ramas.");

        if (q1 is null || q2 is null)
        {
            return; // combo con rama no viable: se compara donde ambas viven (subconjunto declarado)
        }

        Assert.Equal(q2.R2.E15, q1.R2.E15);
        Assert.Equal(q2.R2.E26, q1.R2.E26);
        Assert.Equal(q2.R2.K15, q1.R2.K15);
        AssertCeldasIguales(q1.R2.CeldasPorAse, q2.R2.CeldasPorAse, aseId, "R2");

        Assert.Equal(q2.R4.D9, q1.R4.D9);
        Assert.Equal(q2.R4.P9, q1.R4.P9);
        AssertCeldasIguales(q1.R4.CeldasPorAse, q2.R4.CeldasPorAse, aseId, "R4");
    }

    /// <summary>
    /// T0.4 (segunda mitad): las celdas DESTINO de R2/R4 en la v2 son literales escribibles (sin
    /// <c>&lt;f&gt;</c>). Si alguna fuera fórmula sería un STOP de clase <c>R-MAPA-V2-R2-ASE4</c>.
    /// </summary>
    [Fact]
    public void T0_4_DestinosR2R4V2_SonLiteralesSinFormula()
    {
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var r2 = WorkbookLeafCellMapQ2.ObtenerR2Q2Editables(aseId);
            foreach (var celda in new[] { r2.Componente, r2.SubsCont, r2.Especiales })
            {
                Assert.False(CeldaCrudaDe(PlantillaV2, HojaR2, celda).EsFormula,
                    $"{HojaR2}!{celda} (ASE{aseId}) debió ser VALOR editable en la v2.");
            }

            var r4 = WorkbookLeafCellMapQ2.ObtenerR4Q2Editables(aseId);
            foreach (var celda in new[] { r4.Total, r4.P })
            {
                Assert.False(CeldaCrudaDe(PlantillaV2, HojaR4, celda).EsFormula,
                    $"{HojaR4}!{celda} (ASE{aseId}) debió ser VALOR editable en la v2.");
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.5 (S1.5) — geometría v1 aislada como legacy + FORMA-FUENTE R1 por período (dato duro)
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void T0_5_MapaV1_EsGeometriaLegacy_NoEsLaProductiva()
    {
        // v1: F25/F41/L25/F30/F10/L10 (ASE1) — CELDAS DISTINTAS de la geometría v2 (F16/F32/F47/…).
        var v1 = WorkbookLeafCellMapPorAse.EditableLeafCellsPorAse[1].Select(c => c.Celda).ToArray();
        Assert.Contains("F25", v1);
        Assert.Contains("F41", v1);
        Assert.Contains("L25", v1);
        Assert.Contains("F30", v1);
        Assert.Contains("F10", v1);
        Assert.Contains("L10", v1);

        var v2 = WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(1).Select(p => p.Celda).ToArray();
        Assert.DoesNotContain(v1, celda => v2.Contains(celda, StringComparer.Ordinal));

        // En la v2 esas celdas v1 NO son los operandos del visible: F75 lee F47/F70/F16/L16/L47.
        var formula = NormalizarFormula(CeldaCrudaDe(PlantillaV2, HojaR1, "F75").Formula);
        Assert.DoesNotContain("F25", formula, StringComparison.Ordinal);
        Assert.DoesNotContain("F41", formula, StringComparison.Ordinal);
    }

    /// <summary>
    /// T0.5: forma de la fuente R1 (filas <c>Mes/Total</c> y <c>Aplicacion/Total</c>) por período y ASE.
    /// Es un DATO de la fuente (no del código): la variabilidad de forma es la que el lector único debe
    /// absorber. Casos relevantes: ASE1-julio-1 trae 2 filas Mes (el mapa declara 3 slots) y ASE5 trae
    /// 2 en julio-2/octubre-1 y 3 en julio-1/agosto-2 (el mapa declara 2 slots).
    /// </summary>
    public static TheoryData<string, int, int, int> FormaFuenteR1 => new()
    {
        { "julio-1", 1, 2, 2 }, { "julio-1", 2, 3, 0 }, { "julio-1", 3, 3, 2 }, { "julio-1", 4, 3, 2 }, { "julio-1", 5, 3, 0 },
        { "julio-2", 1, 3, 2 }, { "julio-2", 2, 3, 2 }, { "julio-2", 3, 3, 2 }, { "julio-2", 4, 3, 2 }, { "julio-2", 5, 2, 2 },
        { "octubre-1", 1, 3, 2 }, { "octubre-1", 2, 3, 2 }, { "octubre-1", 3, 3, 0 }, { "octubre-1", 4, 3, 2 }, { "octubre-1", 5, 2, 0 }
    };

    [Theory]
    [MemberData(nameof(FormaFuenteR1))]
    public void T0_5_FormaFuenteR1_PorPeriodo(string periodo, int aseId, int mesEsperadas, int aplicacionEsperadas)
    {
        var filas = LeerFilas(RutaR1(aseId, periodo));

        Assert.Equal(mesEsperadas, filas.Count(f => EsMesTotal(f.Celdas)));
        Assert.Equal(aplicacionEsperadas, filas.Count(f => EsAplicacionTotal(f.Celdas)));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.6 — Δ-1 (GAP de H0): ASE1-julio-1 trae 2 filas Mes y el mapa v2 declara Mes2 (core)
    // ════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Evidencia del GAP que bloquea la puerta H0→H1. Julio-1 ASE1 trae <b>2</b> filas
    /// <c>Mes/Total</c> (no existe el bloque <c>AFaseo</c>) y la geometría v2 le declara 3 slots, uno
    /// de ellos <c>Mes2</c> → <c>F70</c>. La semántica Σ de la ruta única reproduce su R10 oficial al
    /// peso con <c>F70 = 0</c>; lo único que impide leerla es el fail-fast del rol core <c>Mes2</c>.
    /// </summary>
    [Fact]
    public void T0_6_Ase1Julio1_DosFilasMes_ReproduceSuR10ConMes2EnCero()
    {
        var r1 = LeerFilas(RutaR1(1, "julio-1"));
        Assert.Equal(2, r1.Count(f => EsMesTotal(f.Celdas)));

        // El mapa v2 SÍ declara el slot Mes2 para ASE1 (a diferencia de ASE5).
        Assert.Contains(WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(1),
            p => p.Celda == "F70" && p.Fuente == WorkbookLeafCellMapQ2.R1Q2Fuente.Mes2);

        var r2 = LeerFilas(RutasFixture(1, "julio-1").R2);
        var r4 = LeerFilas(RutasFixture(1, "julio-1").R4);

        var d9 = TotalOportunoR1(r1);        // Σ F(Mes) − Σ Esp(Mes sin la última) — no depende de slots
        var d28 = TotalOportunoR2(r2);
        var d47 = ExtemporaneoR1(r1);
        var d66 = TotalR4(r4);

        Assert.Equal(16704332434.57m, d9);
        Assert.Equal(54216385.68m, d28);
        Assert.Equal(11673020.00m, d47);
        Assert.Equal(-12054255.65m, d66);

        var d104 = d9 + d28 + d47 + d66; // Q1: D85 = 0
        Assert.Equal(16758167584.60m, d104);

        var r10 = MonedaAEntero(CeldaCrudaDe(
            Ruta("Docs", "Prueba Julio-1", "Insumos", "R10_Remuneracion_2026071.xlsx"),
            "DetRetri2026071",
            "D9").Valor);
        Assert.Equal(16758167585L, r10);
        Assert.Equal(r10, Redondear(d104));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // Derivaciones (semántica única de la ruta R1, Plan 25 D-C/D-F — NO se re-arbitra)
    // ════════════════════════════════════════════════════════════════════════════════════════

    private static decimal TotalOportunoR1(IReadOnlyList<FilaCruda> r1)
    {
        var mes = r1.Where(f => EsMesTotal(f.Celdas)).ToList();
        var indiceEsp = ExcelWorksheetNavigator.IndiceColumnaPorEncabezado(r1.Select(f => f.Celdas).ToList(), "SERVICIO ESPECIALES");
        var suma = mes.Sum(f => Numero(f.Celdas[5]));
        var resta = mes.Take(mes.Count - 1).Sum(f => indiceEsp >= 0 ? Numero(f.Celdas[indiceEsp]) : 0m);
        return suma - resta;
    }

    private static decimal ExtemporaneoR1(IReadOnlyList<FilaCruda> r1)
    {
        var indiceEsp = ExcelWorksheetNavigator.IndiceColumnaPorEncabezado(r1.Select(f => f.Celdas).ToList(), "SERVICIO ESPECIALES");
        var aplicaciones = r1.Where(f => EsAplicacionTotal(f.Celdas)).ToList();
        var suma = aplicaciones.Sum(f => Numero(f.Celdas[5]));
        var esp = aplicaciones.Count > 0 && indiceEsp >= 0 ? Numero(aplicaciones[0].Celdas[indiceEsp]) : 0m;
        return suma - esp;
    }

    private static decimal TotalOportunoR2(IReadOnlyList<FilaCruda> r2)
    {
        var celdas = r2.Select(f => f.Celdas).ToList();
        var componente = celdas.First(f => Texto(f, 0).Equals("Componente", StringComparison.OrdinalIgnoreCase)
            && Texto(f, 1).Equals("Total", StringComparison.OrdinalIgnoreCase));
        var subsCont = celdas.First(f => Texto(f, 0).Equals("Subs/Cont", StringComparison.OrdinalIgnoreCase)
            && Texto(f, 1).Equals("Total", StringComparison.OrdinalIgnoreCase));
        var indiceEsp = ExcelWorksheetNavigator.IndiceColumnaPorEncabezado(celdas, "Especiales");
        var esp = indiceEsp >= 0 ? Numero(componente[indiceEsp]) : 0m;

        return Numero(componente[4]) + Numero(subsCont[4]) - esp;
    }

    private static decimal TotalR4(IReadOnlyList<FilaCruda> r4)
    {
        var total = r4.Select(f => f.Celdas).First(f => Texto(f, 0).Equals("Total", StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrWhiteSpace(Texto(f, 1)));
        return Numero(total[3]); // D = 0-based 3; P (col 15) no tiene análogo en la fuente → 0
    }

    private static bool EsMesTotal(object?[] fila) =>
        Texto(fila, 1).Equals("Mes", StringComparison.OrdinalIgnoreCase)
        && Texto(fila, 2).Equals("Total", StringComparison.OrdinalIgnoreCase);

    private static bool EsAplicacionTotal(object?[] fila) =>
        Texto(fila, 1).Contains("Aplicacion nuevos x reversion", StringComparison.OrdinalIgnoreCase)
        && Texto(fila, 2).Equals("Total", StringComparison.OrdinalIgnoreCase);

    private static bool EsSubsidio(object?[] fila) =>
        Texto(fila, 4).Contains("Subsidio(-)/Contribucion(+)", StringComparison.OrdinalIgnoreCase);

    private static void AssertCeldasIguales(
        IReadOnlyDictionary<string, decimal> esperado, IReadOnlyDictionary<string, decimal> real, int aseId, string reporte)
    {
        Assert.Equal(esperado.Count, real.Count);
        foreach (var (celda, valor) in esperado)
        {
            Assert.True(real.TryGetValue(celda, out var otro), $"{reporte} ASE{aseId}: falta la celda {celda} en la rama Q2.");
            Assert.Equal(valor, otro);
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // Infraestructura BCL: ZIP + XML crudo
    // ════════════════════════════════════════════════════════════════════════════════════════

    private static (string R1, string R2, string R4) RutasFixture(int aseId, string fixture)
    {
        var carpeta = fixture switch
        {
            "julio-1" => Ruta("Docs", "Prueba Julio-1", "Insumos", CarpetaAse(aseId)),
            "julio-2" => Ruta("Docs", "Prueba Julio-2", "Insumos", CarpetaAse(aseId)),
            "octubre-1" => Ruta("Docs", "Prueba Octubre-1", "Insumos", CarpetaAse(aseId)),
            _ => throw new ArgumentOutOfRangeException(nameof(fixture))
        };

        return (
            Unico(carpeta, "Recaudoporcomponente"),
            Unico(carpeta, "RerpoteDetalleSaldosaFavor"),
            Unico(carpeta, "ReversiónPorComponente"));
    }

    private static string RutaR1(int aseId, string periodo) => RutasFixture(aseId, periodo).R1;

    private static string Unico(string carpeta, string prefijo)
    {
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .SingleOrDefault(f => Path.GetFileName(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta '{prefijo}*' en {carpeta}.");
        return archivo!;
    }

    private static string RutaOctubre1(int aseId, bool R1 = false, bool R2 = false, bool R4 = false)
    {
        var fila = Octubre1.Single(f => f.Ase == aseId);
        var archivo = R1 ? fila.R1 : R2 ? fila.R2 : R4 ? fila.R4 : throw new ArgumentException("Indicá el reporte.");
        return Ruta("Docs", "Prueba Octubre-1", "Insumos", fila.Carpeta, archivo);
    }

    /// <summary>Raíz del repositorio POR CUENTA PROPIA (prohibido <c>Insumos.Raiz()</c> en este fixture).</summary>
    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }

    private static string Ruta(params string[] partes) => Path.Combine(partes.Prepend(Raiz()).ToArray());

    private static string Sha256(string ruta)
    {
        using var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private sealed record CeldaCruda(string? Formula, string? Valor)
    {
        public bool EsFormula => Formula is not null;
    }

    /// <summary>Fila cruda con su NÚMERO real (atributo <c>r</c>), no su posición en la lista.</summary>
    private sealed record FilaCruda(int Numero, object?[] Celdas);

    /// <summary>Lee una celda por referencia (p. ej. <c>F75</c>) del XML crudo de la hoja por NOMBRE.</summary>
    private static CeldaCruda CeldaCrudaDe(string xlsx, string hoja, string referencia)
    {
        var (doc, compartidos) = CargarHoja(xlsx, hoja);
        var celda = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "c"
            && string.Equals((string?)e.Attribute("r"), referencia, StringComparison.OrdinalIgnoreCase));

        if (celda is null)
        {
            return new CeldaCruda(null, null);
        }

        var formula = celda.Elements().FirstOrDefault(e => e.Name.LocalName == "f")?.Value;
        var valor = celda.Elements().FirstOrDefault(e => e.Name.LocalName == "v")?.Value;

        var tipo = (string?)celda.Attribute("t");
        if (tipo == "s" && valor is not null)
        {
            valor = compartidos[(int)decimal.Parse(valor, CultureInfo.InvariantCulture)];
        }
        else if (tipo == "inlineStr")
        {
            valor = string.Concat(celda.Elements().Where(e => e.Name.LocalName == "is")
                .SelectMany(e => e.Descendants()).Where(e => e.Name.LocalName == "t").Select(e => e.Value));
        }

        return new CeldaCruda(formula, valor);
    }

    /// <summary>
    /// Filas de <c>Sheet1</c> como <c>object?[]</c> (índice 0 = columna A), con la MISMA conversión
    /// numérica del lector: <c>double → decimal</c> (15 dígitos significativos). Los XML de origen
    /// guardan artefactos binarios (p. ej. <c>636047.06999999995</c>) que esa ruta normaliza, así el
    /// fixture deriva exactamente lo mismo que producción.
    /// </summary>
    private static List<FilaCruda> LeerFilas(string xlsx)
    {
        var (doc, compartidos) = CargarHoja(xlsx, "Sheet1");
        var filas = new List<FilaCruda>();

        foreach (var nodo in doc.Descendants().Where(e => e.Name.LocalName == "row"))
        {
            var porColumna = new Dictionary<int, object?>();

            foreach (var celda in nodo.Elements().Where(e => e.Name.LocalName == "c"))
            {
                var referencia = (string?)celda.Attribute("r") ?? string.Empty;
                var columna = IndiceColumna(referencia);
                var tipo = (string?)celda.Attribute("t");
                var valor = celda.Elements().FirstOrDefault(e => e.Name.LocalName == "v")?.Value;

                object? contenido = tipo == "s" && valor is not null
                    ? compartidos[(int)decimal.Parse(valor, CultureInfo.InvariantCulture)]
                    : tipo == "inlineStr"
                        ? string.Concat(celda.Elements().Where(e => e.Name.LocalName == "is")
                            .SelectMany(e => e.Descendants()).Where(e => e.Name.LocalName == "t").Select(e => e.Value))
                        : valor is null
                            ? null
                            : Convert.ToDecimal(double.Parse(valor, NumberStyles.Any, CultureInfo.InvariantCulture));

                porColumna[columna] = contenido;
            }

            var ancho = porColumna.Count == 0 ? 0 : porColumna.Keys.Max() + 1;
            var fila = new object?[ancho];
            for (var i = 0; i < ancho; i++)
            {
                fila[i] = porColumna.GetValueOrDefault(i);
            }

            filas.Add(new FilaCruda(int.Parse((string?)nodo.Attribute("r") ?? "0", CultureInfo.InvariantCulture), fila));
        }

        return filas;
    }

    private static (XDocument Documento, List<string> Compartidos) CargarHoja(string xlsx, string hoja)
    {
        using var zip = ZipFile.OpenRead(xlsx);

        var compartidos = new List<string>();
        var entradaCompartidos = zip.Entries.FirstOrDefault(e => e.FullName.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase));
        if (entradaCompartidos is not null)
        {
            using var streamCompartidos = entradaCompartidos.Open();
            var doc = XDocument.Load(streamCompartidos);
            foreach (var si in doc.Descendants().Where(e => e.Name.LocalName == "si"))
            {
                compartidos.Add(string.Concat(si.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value)));
            }
        }

        var rutaHoja = RutaHoja(zip, hoja);
        var entrada = zip.Entries.First(e => e.FullName.Equals(rutaHoja, StringComparison.OrdinalIgnoreCase));
        using var stream = entrada.Open();
        var copia = new MemoryStream();
        stream.CopyTo(copia);
        copia.Position = 0;
        return (XDocument.Load(copia), compartidos);
    }

    private static string RutaHoja(ZipArchive zip, string hoja)
    {
        var entradaWorkbook = zip.Entries.First(e => e.FullName.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase));
        XDocument workbook;
        using (var stream = entradaWorkbook.Open())
        {
            workbook = XDocument.Load(stream);
        }

        var nodo = workbook.Descendants().First(e => e.Name.LocalName == "sheet"
            && string.Equals((string?)e.Attribute("name"), hoja, StringComparison.OrdinalIgnoreCase));
        var relId = nodo.Attributes().First(a => a.Name.LocalName == "id").Value;

        var entradaRels = zip.Entries.First(e => e.FullName.Equals("xl/_rels/workbook.xml.rels", StringComparison.OrdinalIgnoreCase));
        XDocument rels;
        using (var streamRels = entradaRels.Open())
        {
            rels = XDocument.Load(streamRels);
        }

        var destino = rels.Descendants().First(e => e.Name.LocalName == "Relationship"
            && string.Equals((string?)e.Attribute("Id"), relId, StringComparison.Ordinal)).Attribute("Target")!.Value;

        return destino.StartsWith("/", StringComparison.Ordinal)
            ? destino.TrimStart('/')
            : $"xl/{destino.TrimStart('.', '/')}";
    }

    private static int IndiceColumna(string referencia)
    {
        var letras = new string(referencia.TakeWhile(char.IsLetter).ToArray()).ToUpperInvariant();
        var indice = 0;
        foreach (var letra in letras)
        {
            indice = (indice * 26) + (letra - 'A' + 1);
        }

        return indice - 1;
    }

    private static string Texto(object?[] fila, int indice) =>
        indice < fila.Length ? ExcelWorksheetNavigator.CeldaTexto(fila[indice]) : string.Empty;

    private static decimal Numero(object? valor) => ExcelWorksheetNavigator.CeldaNumero(valor);

    private static decimal NumeroDe(string? valor) =>
        string.IsNullOrWhiteSpace(valor) ? 0m : Convert.ToDecimal(double.Parse(valor, NumberStyles.Any, CultureInfo.InvariantCulture));

    private static long Redondear(decimal valor) => (long)Math.Round(valor, 0, MidpointRounding.AwayFromZero);

    /// <summary>
    /// El R10 oficial guarda sus importes como TEXTO con formato de moneda colombiana
    /// (<c>$14.842.425.136</c>): punto = miles, coma = decimales.
    /// </summary>
    private static long MonedaAEntero(string? valor)
    {
        Assert.False(string.IsNullOrWhiteSpace(valor), "El oráculo R10 no trae valor en la celda esperada.");

        var limpio = valor!.Replace("$", string.Empty, StringComparison.Ordinal)
            .Replace(" ", string.Empty, StringComparison.Ordinal)
            .Replace("\u00A0", string.Empty, StringComparison.Ordinal)
            .Replace(".", string.Empty, StringComparison.Ordinal)
            .Replace(',', '.');

        return Redondear(decimal.Parse(limpio, NumberStyles.Any, CultureInfo.InvariantCulture));
    }

    private static string NormalizarFormula(string? formula)
    {
        var texto = (formula ?? string.Empty).Trim();
        texto = texto.Replace("'", string.Empty, StringComparison.Ordinal);
        texto = texto.Replace(" ", string.Empty, StringComparison.Ordinal);
        return texto.Replace("_xlfn.", string.Empty, StringComparison.OrdinalIgnoreCase);
    }
}
