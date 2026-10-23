using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 55 (H0, S1.1-S1.6): evidencia T0 de la CONCILIACIÓN POR EMPRESA vs visible de bloque en la
/// corrida 5-ASE de <c>2026101</c> (las 5 deltas del gate <c>ValidarSigmaEmpresas</c>).
///
/// Método (D1 del plan, forma normativa): BCL puro (<see cref="ZipArchive"/> + <see cref="XDocument"/>),
/// sin Excel/COM, sin OpenXml SDK y sin ExcelDataReader para LEER; rutas EXPLÍCITAS (prohibido
/// <c>Insumos.Raiz()</c>: este fixture resuelve la raíz por su cuenta). Cero escritura sobre fuentes
/// y plantillas: todo el fixture es read-only (H0 no toca código de producción).
///
/// Invariantes del plan que este fixture hace cumplir:
/// - R-ORACULO: el byte de una celda con <c>&lt;f&gt;</c> se usa SOLO por el TEXTO de la fórmula; nunca
///   su <c>&lt;v&gt;</c> cacheado como oráculo.
/// - CA0: todo número de criterio viaja con su derivación (fuente × <c>&lt;f&gt;</c> de la v2 × mapa).
/// - G-SHA: plantillas v1/v2 intactas (valor medido en disco, 64 hex).
///
/// Trazabilidad: T0.1 (S1.1) destinos por bytes · T0.2 (S1.2) censo octubre-1 · T0.3 (S1.3) deltas y
/// atribución · T0.4 (S1.4) H-ASE4-R1 / H-ASE1-R1 · T0.5 (S1.5) veredicto R-RECORTE-Q2 ·
/// T0.6 (S1.6) cobertura del gate · T0.7 (G-SHA/G-FORMULA).
/// Evidencia narrativa: <c>Docs/T0-Plan55-Conciliacion-Por-Empresa.md</c>.
/// </summary>
public sealed class ConciliacionEmpresaCierreBloqueT0Tests
{
    // ── G-SHA (CA4 / S0): plantillas oficiales INTOCABLES ───────────────────────────────────────
    private const string ShaPlantillaV2 = "0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB";
    private const string ShaPlantillaV1 = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";

    private const string HojaR1 = "Reporte Componentes R1";
    private const string HojaR2 = "Rem. Anticipos R2";
    private const string HojaR4 = "Reversion Pagos R4";

    /// <summary>Tolerancia del gate (<c>WorkbookLeafCoherence.Tolerancia</c>): no se toca, solo se cita.</summary>
    private const decimal ToleranciaGate = 0.5m;

    // ── Rutas EXPLÍCITAS (sin Insumos.Raiz) ─────────────────────────────────────────────────────
    private static string PlantillaV2 => Ruta("Docs", "Plantilla estandarizada", "v2", "Formato Remuneracion 2026xx-x Total vf.xlsx");

    private static string PlantillaV1 => Ruta("Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string TrxBaseP53 => Ruta("Remuneracion.IntegrationTests", "TestResults", "p53-h2-suite.trx");

    /// <summary>
    /// Octubre-1 (caso rojo del plan): carpeta + nombre completo de archivo por ASE. Solo los tres ASE
    /// con delta medida (1, 2 y 4); los demás no participan de la atribución de este T0.
    /// </summary>
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
        (4, "4-Bogota Limpia",
            "Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174948502.xlsx",
            "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___202610161750610.xlsx",
            "ReversiónPorComponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016175022764.xlsx")
    ];

    private static string FuenteR1(int aseId) => Ruta("Docs", "Prueba Octubre-1", "Insumos", Archivo(aseId).Carpeta, Archivo(aseId).R1);

    private static string FuenteR2(int aseId) => Ruta("Docs", "Prueba Octubre-1", "Insumos", Archivo(aseId).Carpeta, Archivo(aseId).R2);

    private static string FuenteR4(int aseId) => Ruta("Docs", "Prueba Octubre-1", "Insumos", Archivo(aseId).Carpeta, Archivo(aseId).R4);

    private static (int Ase, string Carpeta, string R1, string R2, string R4) Archivo(int aseId) =>
        Octubre1.Single(o => o.Ase == aseId);

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T0.7 (S0 G-SHA / G-FORMULA) — plantillas intactas y las 10 <f> de R1 intactas
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void T0_7_PlantillasOficiales_ShaIntacto()
    {
        Assert.Equal(ShaPlantillaV2, Sha256(PlantillaV2));
        Assert.Equal(ShaPlantillaV1, Sha256(PlantillaV1));
    }

    /// <summary>
    /// Las 10 <c>&lt;f&gt;</c> de <c>Reporte Componentes R1</c> (TOT_OPT + EXTEMP por ASE) siguen
    /// byte-idénticas en la v2 (misma lista del T0 del Plan 54 §1: CA4 «las 10 &lt;f&gt; de R1
    /// intactas»).
    /// </summary>
    public static TheoryData<int, string, string> FormulasR1 =>
        new()
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
    [MemberData(nameof(FormulasR1))]
    public void T0_7_FormulasR1_ByteIdenticas(int aseId, string celda, string formulaEsperada)
    {
        var cruda = CeldaCrudaDe(PlantillaV2, HojaR1, celda);

        Assert.True(cruda.EsFormula, $"{HojaR1}!{celda} (ASE{aseId}) debe seguir siendo FÓRMULA en la v2.");
        Assert.Equal(formulaEsperada, NormalizarFormula(cruda.Formula));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T0.1 (S1.1 / CA4) — destinos por bytes: celda, naturaleza (<f> vs literal) y texto de la <f>
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Tabla congelada por bytes de las <b>112</b> celdas destino del mapa por empresa
    /// (<see cref="WorkbookLeafCellMapPorEmpresa"/>: 56 R1 + 36 R2 + 20 R4). <c>Formula == null</c> =
    /// LITERAL sin <c>&lt;f&gt;</c>; <c>Formula != null</c> = FÓRMULA (texto exacto de la <c>&lt;f&gt;</c>;
    /// vacío cuando es miembro de una fórmula compartida <c>t="shared"</c>, que no lleva texto propio).
    /// </summary>
    private static readonly (string Reporte, int Ase, int Empresa, string Celda, string? Formula, bool Compartida)[] DestinosV2 =
    [
        // R1 — ASE1
        ("R1", 1, 2, "F33", null, false), ("R1", 1, 2, "F14", null, false), ("R1", 1, 2, "L14", null, false),
        ("R1", 1, 4, "F40", null, false), ("R1", 1, 4, "F24", null, false), ("R1", 1, 4, "L24", null, false),
        ("R1", 1, 4, "F29", null, false), ("R1", 1, 4, "F9", null, false), ("R1", 1, 4, "L9", null, false),
        // R1 — ASE2
        ("R1", 2, 1, "F102", null, false), ("R1", 2, 1, "F122", null, false), ("R1", 2, 1, "F86", "F62", false),
        ("R1", 2, 2, "F117", null, false), ("R1", 2, 2, "F95", "F69+F46+F15-L15-L46", false),
        ("R1", 2, 2, "F83", null, false), ("R1", 2, 2, "L95", string.Empty, true),
        ("R1", 2, 4, "F129", null, false), ("R1", 2, 4, "F112", null, false), ("R1", 2, 4, "F89", null, false),
        ("R1", 2, 4, "L112", null, false),
        // R1 — ASE3
        ("R1", 3, 2, "F246", null, false), ("R1", 3, 2, "F227", null, false), ("R1", 3, 2, "F213", null, false),
        ("R1", 3, 2, "L227", null, false),
        ("R1", 3, 4, "F253", null, false), ("R1", 3, 4, "F237", null, false), ("R1", 3, 4, "F216", null, false),
        ("R1", 3, 4, "L237", null, false), ("R1", 3, 4, "F242", null, false), ("R1", 3, 4, "F222", null, false),
        // R1 — ASE4
        ("R1", 4, 1, "F409", null, false), ("R1", 4, 1, "F380", null, false), ("R1", 4, 1, "L380", null, false),
        ("R1", 4, 2, "F404", null, false), ("R1", 4, 2, "F373", null, false), ("R1", 4, 2, "F353", null, false),
        ("R1", 4, 2, "L373", null, false), ("R1", 4, 2, "F395", null, false), ("R1", 4, 2, "F362", null, false),
        ("R1", 4, 4, "F416", "F353", false), ("R1", 4, 4, "F390", null, false), ("R1", 4, 4, "F356", null, false),
        ("R1", 4, 4, "L390", null, false), ("R1", 4, 4, "F400", null, false), ("R1", 4, 4, "F368", null, false),
        // R1 — ASE5
        ("R1", 5, 2, "F502", null, false), ("R1", 5, 2, "F483", null, false), ("R1", 5, 2, "F474", null, false),
        ("R1", 5, 2, "L483", null, false),
        ("R1", 5, 3, "F505", null, false), ("R1", 5, 3, "F487", null, false), ("R1", 5, 3, "L487", null, false),
        ("R1", 5, 4, "F512", null, false), ("R1", 5, 4, "F497", null, false), ("R1", 5, 4, "F477", null, false),
        ("R1", 5, 4, "L497", null, false),

        // R2 — 12 pares empresa×ASE (empresas 3 y 5 no declaran R2 en ningún ASE)
        ("R2", 1, 2, "E24", null, false), ("R2", 1, 2, "E10", null, false), ("R2", 1, 2, "K10", null, false),
        ("R2", 1, 4, "E20", null, false), ("R2", 1, 4, "E31", null, false), ("R2", 1, 4, "K20", null, false),
        ("R2", 2, 1, "E104", null, false), ("R2", 2, 1, "E83", null, false), ("R2", 2, 1, "K83", null, false),
        ("R2", 2, 2, "E97", null, false), ("R2", 2, 2, "E74", null, false), ("R2", 2, 2, "K74", null, false),
        ("R2", 2, 4, "E93", null, false), ("R2", 2, 4, "E111", null, false), ("R2", 2, 4, "K93", null, false),
        ("R2", 3, 2, "E185", null, false), ("R2", 3, 2, "E171", null, false), ("R2", 3, 2, "K171", null, false),
        ("R2", 3, 4, "E181", null, false), ("R2", 3, 4, "E192", null, false), ("R2", 3, 4, "K181", null, false),
        ("R2", 4, 1, "E319", null, false), ("R2", 4, 1, "E285", "E192+E203-K192", false),
        ("R2", 4, 1, "K285", string.Empty, true),
        ("R2", 4, 2, "E312", null, false), ("R2", 4, 2, "E289", null, false), ("R2", 4, 2, "K289", null, false),
        ("R2", 4, 4, "E308", null, false), ("R2", 4, 4, "E326", null, false), ("R2", 4, 4, "K308", null, false),
        ("R2", 5, 2, "E407", null, false), ("R2", 5, 2, "E389", null, false), ("R2", 5, 2, "K389", null, false),
        ("R2", 5, 4, "E403", null, false), ("R2", 5, 4, "E417", null, false), ("R2", 5, 4, "K403", null, false),

        // R4 — 10 pares empresa×ASE (todos literales)
        ("R4", 1, 4, "D14", null, false), ("R4", 1, 4, "P14", null, false),
        ("R4", 2, 1, "D104", null, false), ("R4", 2, 1, "P104", null, false),
        ("R4", 2, 2, "D98", null, false), ("R4", 2, 2, "P98", null, false),
        ("R4", 3, 2, "D199", null, false), ("R4", 3, 2, "P199", null, false),
        ("R4", 3, 4, "D205", null, false), ("R4", 3, 4, "P205", null, false),
        ("R4", 4, 1, "D242", null, false), ("R4", 4, 1, "P242", null, false),
        ("R4", 4, 2, "D236", null, false), ("R4", 4, 2, "P236", null, false),
        ("R4", 4, 4, "D248", null, false), ("R4", 4, 4, "P248", null, false),
        ("R4", 5, 2, "D351", null, false), ("R4", 5, 2, "P351", null, false),
        ("R4", 5, 4, "D357", null, false), ("R4", 5, 4, "P357", null, false)
    ];

    /// <summary>
    /// S1.1: para CADA celda destino del mapa por empresa, la v2 trae exactamente la naturaleza
    /// congelada (literal vs <c>&lt;f&gt;</c>) y el texto exacto de la <c>&lt;f&gt;</c>. El conjunto de
    /// celdas de la tabla es IDÉNTICO al que declara el mapa de producción: si alguien añade/borra una
    /// celda del mapa sin re-congelarla aquí, el fixture enrojece.
    /// </summary>
    [Fact]
    public void T0_1_DestinosPorBytes_CoincidenConElMapaYSuNaturaleza()
    {
        Assert.Equal(56, DestinosV2.Count(d => d.Reporte == "R1"));
        Assert.Equal(36, DestinosV2.Count(d => d.Reporte == "R2"));
        Assert.Equal(20, DestinosV2.Count(d => d.Reporte == "R4"));

        // Simetría mapa ↔ tabla (por reporte): mismo multiconjunto de (ASE, empresa, celda).
        foreach (var reporte in new[] { "R1", "R2", "R4" })
        {
            var delMapa = DestinosDelMapa(reporte).OrderBy(x => x).ToArray();
            var deLaTabla = DestinosV2
                .Where(d => d.Reporte == reporte)
                .Select(d => (d.Ase, d.Empresa, d.Celda))
                .OrderBy(x => x)
                .ToArray();

            Assert.Equal(deLaTabla, delMapa);
        }

        foreach (var (reporte, ase, empresa, celda, formula, compartida) in DestinosV2)
        {
            var hoja = HojaDe(reporte);
            var cruda = CeldaCrudaDe(PlantillaV2, hoja, celda);

            Assert.True(cruda.Existe, $"{hoja}!{celda} (ASE{ase}, empresa {empresa}) no existe en la v2.");
            Assert.Equal(formula is not null, cruda.EsFormula);
            Assert.Equal(compartida, cruda.Compartida);
            if (formula is not null)
            {
                Assert.Equal(NormalizarFormula(formula), NormalizarFormula(cruda.Formula));
            }
        }
    }

    /// <summary>
    /// Lista CERRADA de celdas destino que son FÓRMULA en la v2 (clase <c>R-MAPA-V2-R2-ASE4</c>):
    /// escribir cualquiera de ellas aborta con <c>ERR-PLANTILLA</c> (guard anti-fórmula de
    /// <c>EscribirValorNumerico</c>). Son 6: 4 de R1 y 2 de R2.
    /// </summary>
    [Fact]
    public void T0_1_DestinosFormula_ListaCerradaSeis()
    {
        var formulas = DestinosV2
            .Where(d => d.Formula is not null)
            .Select(d => $"{d.Reporte}!{d.Celda}")
            .OrderBy(x => x, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(new[] { "R1!F416", "R1!F86", "R1!F95", "R1!L95", "R2!E285", "R2!K285" }, formulas);
    }

    /// <summary>
    /// Fórmula del VISIBLE DE BLOQUE por (reporte × ASE) en la v2 — la <c>&lt;f&gt;</c> que el bloque
    /// usa para su total. Se usa SOLO su texto (nunca su <c>&lt;v&gt;</c>).
    /// </summary>
    public static TheoryData<string, int, string, string> VisibleBloqueV2 =>
        new()
        {
            { "R1", 1, "F75", "F47+F70+F16-L16-L47" },
            { "R1", 2, "F254", "F171+F208-L171+F120-L120" },
            { "R1", 3, "F415", "F330+F353+F299-L299-L330" },
            { "R1", 4, "F568", "F511+F548+F460-L460-L511" },
            { "R1", 5, "F680", "F674+F648+F613-L613-L648" },
            { "R2", 1, "E47", "E21+E32-K21" },
            { "R2", 2, "E155", "E101+E123-K101" },
            { "R2", 3, "E273", "E193+E204-K193" },
            { "R2", 4, "E384", "E327+E349-K327" },
            { "R2", 5, "E468", "E440+E426-K426" },
            { "R4", 1, "D73", "D15-P15" },
            { "R4", 2, "D174", "D111-P111" },
            { "R4", 3, "D211", "D206-P206" },
            { "R4", 4, "D326", "D249-P249" },
            { "R4", 5, "D361", "D358-P358" }
        };

    /// <summary>
    /// S1.1: NINGUNA de las 112 celdas destino del mapa por empresa es término de la <c>&lt;f&gt;</c>
    /// del visible de bloque del ASE que la declara: el operando viaja a una celda que el total del
    /// bloque no recorre (hallazgo estructural, coherente con T0-Plan53 §6: 34/56 celdas R1 caen en el
    /// bloque de OTRO ASE).
    /// </summary>
    [Theory]
    [MemberData(nameof(VisibleBloqueV2))]
    public void T0_1_NingunDestinoEsTerminoDelVisibleDeBloque(string reporte, int ase, string celdaVisible, string formulaEsperada)
    {
        var cruda = CeldaCrudaDe(PlantillaV2, HojaDe(reporte), celdaVisible);
        Assert.True(cruda.EsFormula, $"{HojaDe(reporte)}!{celdaVisible} debe ser FÓRMULA.");
        Assert.Equal(formulaEsperada, NormalizarFormula(cruda.Formula));

        var terminos = Terminos(cruda.Formula!);
        var destinos = DestinosV2
            .Where(d => d.Reporte == reporte && d.Ase == ase)
            .Select(d => d.Celda)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var interseccion = terminos.Intersect(destinos, StringComparer.OrdinalIgnoreCase).ToArray();
        Assert.True(
            interseccion.Length == 0,
            $"{reporte} ASE{ase}: los destinos [{string.Join(", ", interseccion)}] SÍ son términos del visible " +
            $"'{celdaVisible}={formulaEsperada}' (el T0 congeló lo contrario).");
    }

    /// <summary>
    /// Contaminación cruzada medida: <c>F353</c> es destino declarado de <c>(ENEL, ASE4)</c> y a la vez
    /// término del TOT_OPT de <b>ASE3</b> (<c>F415</c>): escribir la empresa de ASE4 mueve el visible de
    /// ASE3. Se congela como hecho (no se corrige en H0).
    /// </summary>
    [Fact]
    public void T0_1_ContaminacionCruzada_F353_EsDestinoDeAse4YTerminoDeAse3()
    {
        Assert.Contains(DestinosV2, d => d.Reporte == "R1" && d.Ase == 4 && d.Celda == "F353");

        var terminosAse3 = Terminos(CeldaCrudaDe(PlantillaV2, HojaR1, "F415").Formula!);
        Assert.Contains("F353", terminosAse3, StringComparer.OrdinalIgnoreCase);

        var terminosAse4 = Terminos(CeldaCrudaDe(PlantillaV2, HojaR1, "F568").Formula!);
        Assert.DoesNotContain("F353", terminosAse4, StringComparer.OrdinalIgnoreCase);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T0.2 (S1.2) — censo octubre-1 por bytes de los 5 bloques que participan de las deltas
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Censo congelado por bytes. R1: filas <c>Mes/Total</c> (B='Mes' ∧ C='Total'), filas
    /// <c>Aplicacion</c>, y fila de cierre (<c>A='Total'</c> ∧ B vacío). R2: filas
    /// <c>Componente/Total</c> y <c>Subs/Cont/Total</c>. R4: fila <c>Total</c> con B vacío.
    /// </summary>
    public static TheoryData<int, string, int[], int[], int, decimal, decimal> CensoBloques =>
        new()
        {
            // ase, reporte, filas etiqueta A, filas etiqueta B, fila cierre, Σ col F/E, especiales (sin la última)
            { 1, "R1", new[] { 8, 33, 52 }, new[] { 18, 41 }, 54, 14834953667.78m, 19931693.03m },
            { 4, "R1", new[] { 11, 46, 70 }, new[] { 31, 59 }, 72, 10482610949.00m, 8643135.58m },
            { 2, "R2", new[] { 33 }, new[] { 56 }, 57, 116410420.67m, 70736.95m },
            { 1, "R4", Array.Empty<int>(), Array.Empty<int>(), 17, -3833780.00m, 0m },
            { 2, "R4", Array.Empty<int>(), Array.Empty<int>(), 17, -21803648.45m, 0m }
        };

    [Theory]
    [MemberData(nameof(CensoBloques))]
    public void T0_2_CensoOctubre1_PorBytes(
        int aseId, string reporte, int[] filasPrimarias, int[] filasSecundarias, int filaCierre, decimal suma, decimal especiales)
    {
        switch (reporte)
        {
            case "R1":
            {
                var filas = LeerFilas(FuenteR1(aseId));

                var mes = filas.Where(f => EsMesTotal(f.Celdas)).ToArray();
                Assert.Equal(filasPrimarias, mes.Select(f => f.Numero).ToArray());
                Assert.Equal(filasSecundarias, filas.Where(f => EsAplicacionTotal(f.Celdas)).Select(f => f.Numero).ToArray());
                Assert.Equal(suma, mes.Sum(f => Numero(f.Celdas[5])));
                Assert.Equal(especiales, mes.Take(mes.Length - 1).Sum(f => Numero(f.Celdas[11])));

                Assert.Equal(filaCierre, filas.Single(f => EsTotalFinal(f.Celdas)).Numero);
                break;
            }

            case "R2":
            {
                var filas = LeerFilas(FuenteR2(aseId));

                var componente = filas.Single(f => Texto(f.Celdas, 0).Equals("Componente", StringComparison.OrdinalIgnoreCase)
                    && Texto(f.Celdas, 1).Equals("Total", StringComparison.OrdinalIgnoreCase));
                var subsCont = filas.Single(f => Texto(f.Celdas, 0).Equals("Subs/Cont", StringComparison.OrdinalIgnoreCase)
                    && Texto(f.Celdas, 1).Equals("Total", StringComparison.OrdinalIgnoreCase));

                Assert.Equal(filasPrimarias, new[] { componente.Numero });
                Assert.Equal(filasSecundarias, new[] { subsCont.Numero });
                Assert.Equal(suma, Numero(componente.Celdas[4]));
                Assert.Equal(especiales, Numero(componente.Celdas[10]));

                Assert.Equal(filaCierre, filas.Single(f => EsTotalFinal(f.Celdas)).Numero);
                break;
            }

            default:
            {
                var filas = LeerFilas(FuenteR4(aseId));
                var cierre = filas.Single(f => EsTotalFinal(f.Celdas));

                Assert.Equal(filaCierre, cierre.Numero);
                Assert.Equal(suma, Numero(cierre.Celdas[3]));
                break;
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T0.3 (S1.3 / CA0) — derivación operando por operando y las 5 deltas
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Valor que el lector de producción escribe en cada celda destino (derivación
    /// <c>fuente × mapa</c>): R1 por sección declarada (<see cref="ResolvedorOperandosEmpresaR1"/>),
    /// R2 por magnitud (Total = mayor |E|, SubsCont = menor |E|, Especiales = K de la fila Total),
    /// R4 por fila <c>Total</c> del label con <c>P = 0</c>. Congelado por bytes.
    /// </summary>
    public static TheoryData<string, int, int, string, decimal> OperandosPorEmpresa =>
        new()
        {
            // ASE1 R1
            { "R1", 1, 2, "F14", 17734947270.61m },
            { "R1", 1, 2, "F33", -3253607640.54m },
            { "R1", 1, 2, "L14", 9937924.10m },
            { "R1", 1, 4, "F24", 290244120.78m },
            { "R1", 1, 4, "F40", 63371932.00m },
            { "R1", 1, 4, "L24", 9993768.93m },
            { "R1", 1, 4, "F9", 89987.43m },
            { "R1", 1, 4, "F29", 44922.57m },
            { "R1", 1, 4, "L9", 0m },
            // ASE4 R1
            { "R1", 4, 1, "F380", 356584.67m },
            { "R1", 4, 1, "F409", -47234.67m },
            { "R1", 4, 1, "L380", 0m },
            { "R1", 4, 2, "F373", 10331368139.96m },
            { "R1", 4, 2, "F404", -440994165.67m },
            { "R1", 4, 2, "L373", 7972372.83m },
            { "R1", 4, 2, "F353", -10060.29m },
            { "R1", 4, 2, "F395", -1560876.97m },
            { "R1", 4, 2, "F362", 21957966.97m },
            { "R1", 4, 4, "F390", 445264511.25m },
            { "R1", 4, 4, "F416", 147002644.40m },
            { "R1", 4, 4, "L390", 670762.75m },
            { "R1", 4, 4, "F356", -20120.65m },
            { "R1", 4, 4, "F400", 133706.30m },
            { "R1", 4, 4, "F368", 717030.70m },
            // ASE2 R2
            { "R2", 2, 1, "E83", 16015076.93m },
            { "R2", 2, 1, "E104", -4383219.66m },
            { "R2", 2, 1, "K83", 0m },
            { "R2", 2, 2, "E74", 58544042.33m },
            { "R2", 2, 2, "E97", -19130403.64m },
            { "R2", 2, 2, "K74", 70736.95m },
            { "R2", 2, 4, "E93", 41702163.54m },
            { "R2", 2, 4, "E111", 2827527.77m },
            { "R2", 2, 4, "K93", 0m },
            // ASE1 R4 / ASE2 R4
            { "R4", 1, 4, "D14", -134910.00m },
            { "R4", 1, 4, "P14", 0m },
            { "R4", 2, 1, "D104", 0m },
            { "R4", 2, 2, "D98", -20858585.45m },
            { "R4", 2, 2, "P98", 0m }
        };

    [Theory]
    [MemberData(nameof(OperandosPorEmpresa))]
    public void T0_3_OperandoPorEmpresa_DerivacionFuenteMapa(string reporte, int aseId, int empresaId, string celda, decimal esperado)
    {
        var celdas = ExtraerCeldasDelMapa(reporte, aseId, empresaId);

        Assert.True(celdas.ContainsKey(celda), $"{reporte} (ASE{aseId}, empresa {empresaId}) no deriva la celda {celda}.");
        Assert.Equal(esperado, celdas[celda]);
    }

    /// <summary>
    /// Las 5 deltas medidas por el gate en la corrida real <c>2026101</c>, derivadas aquí por bytes:
    /// <c>Σ visibles-empresa (dominio)</c> vs <c>visible de bloque (dominio)</c>, tolerancia ±0,5.
    /// Ningún número se ajusta: la delta es la que el gate reportó (verbatim del §0.2 del plan).
    /// </summary>
    public static TheoryData<int, string, decimal, decimal, decimal, string> Deltas =>
        new()
        {
            // ase, reporte, Σ empresas, visible de bloque, delta, causa propuesta (i/ii/iii)
            { 1, "R1", 14815023989.82m, 14815021974.75m, 2015.07m, "i" },
            { 4, "R1", 10474277163.42m, 10473967813.42m, 309350.00m, "i" },
            { 2, "R2", 95504450.32m, 95634974.32m, 130524.00m, "ii" },
            { 1, "R4", -134910.00m, -3833780.00m, 3698870.00m, "i" },
            { 2, "R4", -20858585.45m, -21803648.45m, 945063.00m, "i" }
        };

    [Theory]
    [MemberData(nameof(Deltas))]
    public void T0_3_DeltaMedida_SeReproducePorBytes(
        int aseId, string reporte, decimal sumaEsperada, decimal bloqueEsperado, decimal deltaEsperada, string causa)
    {
        Assert.Contains(causa, new[] { "i", "ii", "iii" });

        var (sumaEmpresas, visibleBloque) = DerivarGate(reporte, aseId);

        Assert.Equal(sumaEsperada, sumaEmpresas);
        Assert.Equal(bloqueEsperado, visibleBloque);
        Assert.Equal(deltaEsperada, Math.Abs(sumaEmpresas - visibleBloque));
        Assert.True(
            Math.Abs(sumaEmpresas - visibleBloque) > ToleranciaGate,
            "La delta debe superar la tolerancia del gate (si no, no habría ERR-VALIDACION).");
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T0.4 (S1.4) — los dos delatores: H-ASE4-R1 y H-ASE1-R1
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// H-ASE4-R1: Δ = <c>309.350,00</c> = <c>ReciprocidadEaab</c> exacto.
    /// CONFIRMADA la identidad numérica y REFUTADA la lectura «valor huérfano o duplicado»: en
    /// octubre-1 el label <c>NUEVO ESQUEMA</c> de ASE4 **no trae ninguna fila en las secciones
    /// Oportuno** (solo <c>Aplicacion#1</c>), así que la resolución <c>UltimaOportuna</c> del operando
    /// cae en la única sección presente = Aplicación. El TOT_OPT del bloque suma SOLO las filas
    /// <c>Mes/Total</c>, que no recorren Aplicación ⇒ el dinero de RECIP entra en Σ empresas y no en el
    /// bloque. No hay duplicado: 356.584,67 + (−47.234,67) = 309.350,00 es la MISMA sección vista por
    /// sus dos caras (componente y subsidio).
    /// </summary>
    [Fact]
    public void T0_4_Hipotesis_Ase4R1_ReciprocidadEsAplicacionYNoOportuno()
    {
        var ventana = Ventana(LeerFilas(FuenteR1(4)));
        var espejo = EspejoDe(ventana);

        // El label de fuente de RECIP en ASE4 es 'NUEVO ESQUEMA' (ObtenerLabelFuente).
        var totales = FilasEmpresa(ventana, "NUEVO ESQUEMA", antesDeComponente: true);
        var subs = FilasEmpresa(ventana, "NUEVO ESQUEMA", antesDeComponente: false);

        // Una sola sección propia en cada bloque (Totales y Subs): ambas Aplicacion#1.
        Assert.Single(totales);
        Assert.Single(subs);
        Assert.Equal(TipoSeccionEspejoR1.Aplicacion, Seccion(espejo, totales[0].Indice)!.Value.Tipo);

        Assert.Equal(356584.67m, Numero(ventana[totales[0].Indice].Celdas[5]));
        Assert.Equal(-47234.67m, Numero(ventana[subs[0].Indice].Celdas[5]));
        Assert.Equal(309350.00m, 356584.67m - 47234.67m);

        // El bloque (Σ Mes) no recorre la sección Aplicación donde vive RECIP.
        var (sumaEmpresas, visibleBloque) = DerivarGate("R1", 4);
        Assert.Equal(309350.00m, sumaEmpresas - visibleBloque);

        // La sección Aplicacion#1 y su total por bytes (no se usa su <v>: es un literal de la fuente).
        var aplicacion = ventana.First(f => EsAplicacionTotal(f.Celdas));
        Assert.Equal(31, aplicacion.Numero);
        Assert.Equal(23031582.34m, Numero(aplicacion.Celdas[5]));
    }

    /// <summary>
    /// H-ASE1-R1: Δ = <c>2.015,07</c> = |primera fila <c>Mes/Total</c>|.
    /// CONFIRMADA la identidad y REFUTADA la dirección que la hipótesis enunciaba como primera
    /// alternativa: la sección «primera» (Oportuno#1, cerrada por la fila <c>Mes/Total</c> 8) <b>SÍ</b>
    /// entra en el visible de bloque (es el término <c>F16</c> de <c>F75</c> y su importe es −2.015,07)
    /// y <b>NO</b> entra en el agregado por empresa: la cadena de ASE1 es de 3 términos
    /// (<c>totalMain + subsMain − especialesMain</c>) y el mapa <c>(2,1)</c> no declara
    /// <c>TotalMenor</c>. La fuente cuadra al peso (la fila de ENEL de esa sección vale −2.015,07).
    /// </summary>
    [Fact]
    public void T0_4_Hipotesis_Ase1R1_PrimeraSeccionEntraEnBloqueYNoEnEmpresa()
    {
        var filas = LeerFilas(FuenteR1(1));
        var ventana = Ventana(filas);
        var espejo = EspejoDe(ventana);

        var mes = ventana.Where(f => EsMesTotal(f.Celdas)).ToArray();
        Assert.Equal(new[] { 8, 33, 52 }, mes.Select(f => f.Numero).ToArray());
        Assert.Equal(-2015.07m, Numero(mes[0].Celdas[5]));

        // ENEL tiene 3 secciones propias en Totales: Oportuno#1 (−2.015,07), Aplicacion#1 y Oportuno#2.
        var enel = FilasEmpresa(ventana, "ENEL", antesDeComponente: true);
        Assert.Equal(3, enel.Count);
        Assert.Equal(TipoSeccionEspejoR1.Oportuno, Seccion(espejo, enel[0].Indice)!.Value.Tipo);
        Assert.Equal(-2015.07m, Numero(ventana[enel[0].Indice].Celdas[5]));

        // El operando por empresa NO declara esa sección para ENEL (mapa (2,1) sin TotalMenor).
        Assert.DoesNotContain(
            WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa[(2, 1)],
            e => e.Fuente == WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor);

        var (sumaEmpresas, visibleBloque) = DerivarGate("R1", 1);
        Assert.Equal(2015.07m, sumaEmpresas - visibleBloque);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T0.5 (S1.5 / D6) — veredicto R-RECORTE-Q2: NO se levanta
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T4 (D6): el recorte Q2 (<c>ProcesadorPeriodo.cs:198-217</c>) sigue vigente porque el destino
    /// <c>(RECIP, ASE4)</c> de R2 es FÓRMULA en la v2 (<c>E285=E192+E203-K192</c> y <c>K285</c>
    /// compartida): escribirlo aborta con <c>ERR-PLANTILLA</c>. El T0 NO respalda el levantamiento.
    /// </summary>
    [Fact]
    public void T0_5_RecorteQ2_NoSeLevanta_DestinoE285EsFormula()
    {
        var e285 = CeldaCrudaDe(PlantillaV2, HojaR2, "E285");
        var k285 = CeldaCrudaDe(PlantillaV2, HojaR2, "K285");

        Assert.True(e285.EsFormula);
        Assert.Equal("E192+E203-K192", NormalizarFormula(e285.Formula));
        Assert.True(k285.EsFormula);
        Assert.True(k285.Compartida);

        // El mapa por empresa del par (RECIP, ASE4) es exactamente el que cae en esas dos fórmulas.
        Assert.Contains(WorkbookLeafCellMapPorEmpresa.EditablesR2PorEmpresa[(1, 4)], e => e.Celda == "E285");
        Assert.Contains(WorkbookLeafCellMapPorEmpresa.EditablesR2PorEmpresa[(1, 4)], e => e.Celda == "K285");
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T0.6 (S1.6 / F5) — el gate NO tiene cobertura en la línea base
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// F5 verificado por bytes sobre la línea base ratificada: 0 ocurrencias del nombre del gate y de
    /// su mensaje en <c>p53-h2-suite.trx</c> (1289/1222/67/0). Si alguien añade cobertura, este
    /// assertion se re-congela en el plan que la introduzca (H2 del Plan 55).
    /// </summary>
    [Fact]
    public void T0_6_GateSigmaEmpresas_SinCoberturaEnLaLineaBase()
    {
        Assert.True(File.Exists(TrxBaseP53), $"No se encontró la línea base: {TrxBaseP53}");
        var texto = File.ReadAllText(TrxBaseP53);

        Assert.DoesNotContain("ValidarSigmaEmpresas", texto, StringComparison.Ordinal);
        Assert.DoesNotContain("visibles de empresas", texto, StringComparison.Ordinal);
        Assert.Contains("total=\"1289\"", texto, StringComparison.Ordinal);
        Assert.Contains("passed=\"1222\"", texto, StringComparison.Ordinal);
        Assert.Contains("failed=\"67\"", texto, StringComparison.Ordinal);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Derivación — espejo EXACTO del lector de producción, alimentado por bytes BCL
    // ════════════════════════════════════════════════════════════════════════════════════════════

    private static (decimal SumaEmpresas, decimal VisibleBloque) DerivarGate(string reporte, int aseId) =>
        reporte switch
        {
            "R1" => DerivarGateR1(aseId),
            "R2" => DerivarGateR2(aseId),
            "R4" => DerivarGateR4(aseId),
            _ => throw new ArgumentOutOfRangeException(nameof(reporte))
        };

    /// <summary>
    /// Espejo de <c>LeerConciliacionEmpresas</c> + <c>CalcularVisibleR1</c> + el TOT_OPT de
    /// <c>MapearR1</c>. Filas-empresa tipadas por
    /// <see cref="ContratoSeccionPeriodFreeR1.SeccionPorCierre"/> y resueltas por
    /// <see cref="ResolvedorOperandosEmpresaR1"/> (una sola semántica, la del contrato del Plan 53).
    /// </summary>
    private static (decimal SumaEmpresas, decimal VisibleBloque) DerivarGateR1(int aseId)
    {
        var ventana = Ventana(LeerFilas(FuenteR1(aseId)));
        var suma = 0m;

        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            suma += VisibleR1(aseId, empresa.Id, ExtraerCeldasR1(aseId, empresa, ventana));
        }

        var mes = ventana.Where(f => EsMesTotal(f.Celdas)).ToArray();
        var totOpt = mes.Sum(f => Numero(f.Celdas[5]))
            - (mes.Length > 1 ? mes.Take(mes.Length - 1).Sum(f => Numero(f.Celdas[11])) : 0m);

        return (suma, totOpt);
    }

    /// <summary>
    /// Espejo de <c>LeerConciliacionEmpresas</c> + <c>CalcularVisibleR2</c> + el visible de
    /// <c>MapearR2</c> (<c>E(Componente/Total) + E(Subs/Cont/Total) − K(Componente/Total)</c>).
    /// </summary>
    private static (decimal SumaEmpresas, decimal VisibleBloque) DerivarGateR2(int aseId)
    {
        var filas = LeerFilas(FuenteR2(aseId));
        var suma = 0m;

        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            var celdas = ExtraerCeldasR2(aseId, empresa, filas);
            suma += celdas["#T"] + celdas["#S"] - celdas["#E"];
        }

        var componente = filas.Single(f => Texto(f.Celdas, 0).Equals("Componente", StringComparison.OrdinalIgnoreCase)
            && Texto(f.Celdas, 1).Equals("Total", StringComparison.OrdinalIgnoreCase));
        var subsCont = filas.Single(f => Texto(f.Celdas, 0).Equals("Subs/Cont", StringComparison.OrdinalIgnoreCase)
            && Texto(f.Celdas, 1).Equals("Total", StringComparison.OrdinalIgnoreCase));

        return (suma, Numero(componente.Celdas[4]) + Numero(subsCont.Celdas[4]) - Numero(componente.Celdas[10]));
    }

    /// <summary>
    /// Espejo de <c>LeerConciliacionEmpresas</c> + <c>CalcularVisibleR4</c> + el visible de
    /// <c>MapearR4</c> (<c>D(Total con B vacío) − P</c>, con <c>P = 0</c>).
    /// </summary>
    private static (decimal SumaEmpresas, decimal VisibleBloque) DerivarGateR4(int aseId)
    {
        var filas = LeerFilas(FuenteR4(aseId));
        var suma = EmpresaFacturacion.Catalogo
            .Sum(empresa => ExtraerCeldasR4(aseId, empresa, filas)["#T"]);

        var cierre = filas.Single(f => EsTotalFinal(f.Celdas));
        return (suma, Numero(cierre.Celdas[3]));
    }

    private static IReadOnlyDictionary<string, decimal> ExtraerCeldasDelMapa(string reporte, int aseId, int empresaId)
    {
        var empresa = EmpresaFacturacion.Obtener(empresaId);
        return reporte switch
        {
            "R1" => ExtraerCeldasR1(aseId, empresa, Ventana(LeerFilas(FuenteR1(aseId)))),
            "R2" => ExtraerCeldasR2(aseId, empresa, LeerFilas(FuenteR2(aseId))),
            "R4" => ExtraerCeldasR4(aseId, empresa, LeerFilas(FuenteR4(aseId))),
            _ => throw new ArgumentOutOfRangeException(nameof(reporte))
        };
    }

    /// <summary>
    /// Operandos R1 por empresa sobre la ventana del espejo. Espejo de <c>ExtraerCeldasR1</c>:
    /// bloque de Totales antes de <c>Componente</c>, bloque de Subs después; fila-empresa por
    /// <c>C=label ∧ D=Total</c>; sección por <see cref="ContratoSeccionPeriodFreeR1.SeccionPorCierre"/>;
    /// operando resuelto por <see cref="ResolvedorOperandosEmpresaR1"/> (ausente → 0).
    /// </summary>
    private static IReadOnlyDictionary<string, decimal> ExtraerCeldasR1(int aseId, EmpresaFacturacion empresa, List<FilaCruda> ventana)
    {
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        if (!WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.TryGetValue((empresa.Id, aseId), out var mapa))
        {
            return celdas; // empresa sin detalle no-estático en este ASE (visible 0)
        }

        var label = LabelFuente(empresa.Id, aseId);
        var espejo = EspejoDe(ventana);
        var indiceComponente = ventana.FindIndex(f => Texto(f.Celdas, 0).Equals("Componente", StringComparison.OrdinalIgnoreCase));
        Assert.True(indiceComponente >= 0, $"ASE {aseId}: la fuente R1 no trae la fila 'Componente'.");

        List<FilaSeccionadaR1> Bloque(bool subs)
        {
            var bloque = new List<FilaSeccionadaR1>();
            for (var i = 0; i < ventana.Count; i++)
            {
                if ((subs ? i <= indiceComponente : i >= indiceComponente)
                    || !Texto(ventana[i].Celdas, 2).Equals(label, StringComparison.OrdinalIgnoreCase)
                    || !Texto(ventana[i].Celdas, 3).Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var seccion = Seccion(espejo, i);
                if (seccion is null)
                {
                    continue; // firma AFaseo: no participa de la alineación por sección
                }

                bloque.Add(new FilaSeccionadaR1(i, seccion.Value));
            }

            return bloque;
        }

        var filasTotal = Bloque(subs: false);
        var filasSubs = Bloque(subs: true);

        foreach (var (celda, fuente, seccionEsperada) in mapa)
        {
            var esDeSubs = fuente is WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMain
                or WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMenor;
            var resultado = ResolvedorOperandosEmpresaR1.Resolver(esDeSubs ? filasSubs : filasTotal, seccionEsperada);

            if (resultado.EsAusente)
            {
                celdas[celda] = 0m;
                continue;
            }

            Assert.False(resultado.EsAmbigua, $"ASE {aseId}: empate de sección para {empresa.Nombre} ({celda}).");
            var indice = resultado.Fila!.Value;
            var esEspeciales = fuente is WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMain
                or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMenor
                or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMedio;

            celdas[celda] = esEspeciales ? Numero(ventana[indice].Celdas[11]) : Numero(ventana[indice].Celdas[5]);
        }

        return celdas;
    }

    /// <summary>
    /// Operandos R2 por empresa: filas <c>B=label ∧ C=Total</c> ordenadas por |E| desc;
    /// Total = [0], SubsCont = [última], Especiales = K de la fila Total (0 si no hay columna).
    /// Devuelve además los marcadores <c>#T</c>/<c>#S</c>/<c>#E</c> para el visible.
    /// </summary>
    private static IReadOnlyDictionary<string, decimal> ExtraerCeldasR2(int aseId, EmpresaFacturacion empresa, List<FilaCruda> filas)
    {
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        if (!WorkbookLeafCellMapPorEmpresa.EditablesR2PorEmpresa.TryGetValue((empresa.Id, aseId), out var mapa))
        {
            celdas["#T"] = 0m;
            celdas["#S"] = 0m;
            celdas["#E"] = 0m;
            return celdas;
        }

        var label = LabelFuente(empresa.Id, aseId);
        var filasEmpresa = filas
            .Where(f => Texto(f.Celdas, 1).Equals(label, StringComparison.OrdinalIgnoreCase)
                && Texto(f.Celdas, 2).Equals("Total", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(f => Math.Abs(Numero(f.Celdas[4])))
            .ToArray();

        var total = filasEmpresa.Length == 0 ? 0m : Numero(filasEmpresa[0].Celdas[4]);
        var subsCont = filasEmpresa.Length == 0 ? 0m : Numero(filasEmpresa[^1].Celdas[4]);
        var especiales = filasEmpresa.Length == 0 ? 0m : Numero(filasEmpresa[0].Celdas[10]);

        foreach (var (celda, fuente) in mapa)
        {
            celdas[celda] = fuente switch
            {
                WorkbookLeafCellMapPorEmpresa.FuenteR2.Total => total,
                WorkbookLeafCellMapPorEmpresa.FuenteR2.SubsCont => subsCont,
                WorkbookLeafCellMapPorEmpresa.FuenteR2.Especiales => especiales,
                _ => 0m
            };
        }

        celdas["#T"] = total;
        celdas["#S"] = subsCont;
        celdas["#E"] = especiales;
        return celdas;
    }

    /// <summary>
    /// Operandos R4 por empresa: fila <c>A=label ∧ B=Total</c>, valor en D (negativo), <c>P = 0</c>.
    /// Devuelve además el marcador <c>#T</c> (Total del bloque de la empresa) para el visible.
    /// </summary>
    private static IReadOnlyDictionary<string, decimal> ExtraerCeldasR4(int aseId, EmpresaFacturacion empresa, List<FilaCruda> filas)
    {
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        if (!WorkbookLeafCellMapPorEmpresa.EditablesR4PorEmpresa.TryGetValue((empresa.Id, aseId), out var mapa))
        {
            celdas["#T"] = 0m;
            return celdas;
        }

        var label = LabelFuente(empresa.Id, aseId);
        var filaTotal = filas.FirstOrDefault(f =>
            Texto(f.Celdas, 0).Equals(label, StringComparison.OrdinalIgnoreCase)
            && Texto(f.Celdas, 1).Equals("Total", StringComparison.OrdinalIgnoreCase));

        var total = filaTotal is null ? 0m : Numero(filaTotal.Celdas[3]);
        foreach (var (celda, fuente) in mapa)
        {
            celdas[celda] = fuente switch
            {
                WorkbookLeafCellMapPorEmpresa.FuenteR4.Total => total,
                WorkbookLeafCellMapPorEmpresa.FuenteR4.P => 0m,
                _ => 0m
            };
        }

        celdas["#T"] = total;
        return celdas;
    }

    /// <summary>Cadena T0 por ASE del visible R1 (misma que <c>CalcularVisibleR1</c>).</summary>
    private static decimal VisibleR1(int aseId, int empresaId, IReadOnlyDictionary<string, decimal> celdas)
    {
        if (celdas.Count == 0)
        {
            return 0m;
        }

        decimal totalMain = 0m, especialesMain = 0m, subsMain = 0m, totalMenor = 0m, block5 = 0m;
        foreach (var (celda, fuente, _) in WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa[(empresaId, aseId)])
        {
            var valor = celdas.GetValueOrDefault(celda);
            switch (fuente)
            {
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMain: totalMain = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMain: especialesMain = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMain: subsMain = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor: totalMenor = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.Block5Total: block5 = valor; break;
            }
        }

        return aseId switch
        {
            1 => totalMain + subsMain - especialesMain,
            2 when empresaId == 1 => totalMain + subsMain + block5,
            2 => totalMain + subsMain + totalMenor - especialesMain,
            3 => totalMain + subsMain + totalMenor - especialesMain,
            4 when empresaId == 1 => totalMain + subsMain - especialesMain,
            4 => totalMain + subsMain + totalMenor - especialesMain,
            5 => totalMain + subsMain + totalMenor - especialesMain,
            _ => 0m
        };
    }

    /// <summary>
    /// Ventana del espejo R1 (misma que <c>ConstruirSecuenciaEspejoR1</c>): desde la primera fila
    /// <c>E='Vlr Servicio'</c> hasta la fila invariante <c>A='Total'</c> con <c>B</c> vacío, ambas
    /// inclusive, descartando filas fantasma (sin texto A–E y sin valor en ninguna columna de la fila).
    /// El índice devuelto es el índice de la VENTANA (el contrato del reader).
    /// </summary>
    private static List<FilaCruda> Ventana(List<FilaCruda> filas)
    {
        var desde = filas.FindIndex(f => Texto(f.Celdas, 4).Equals("Vlr Servicio", StringComparison.OrdinalIgnoreCase));
        Assert.True(desde >= 0, "La fuente R1 no trae la fila 'Vlr Servicio' que delimita el bloque espejo.");

        var hasta = -1;
        for (var i = desde; i < filas.Count; i++)
        {
            if (EsTotalFinal(filas[i].Celdas))
            {
                hasta = i;
                break;
            }
        }

        Assert.True(hasta >= 0, "La fuente R1 no trae la fila invariante 'Total' (A='Total', B vacío).");

        var ventana = new List<FilaCruda>();
        for (var i = desde; i <= hasta; i++)
        {
            var celdas = filas[i].Celdas;
            var hayTexto = Enumerable.Range(0, 5).Any(j => Texto(celdas, j).Length > 0);
            var hayValor = celdas.Any(v => v is not null);
            if (!hayTexto && !hayValor)
            {
                continue; // fila fantasma: no forma parte de la secuencia observada
            }

            ventana.Add(filas[i]);
        }

        return ventana;
    }

    private static List<FilaEspejoR1> EspejoDe(List<FilaCruda> ventana) =>
        ventana.Select(f => new FilaEspejoR1
        {
            A = Texto(f.Celdas, 0),
            B = Texto(f.Celdas, 1),
            C = Texto(f.Celdas, 2),
            D = Texto(f.Celdas, 3),
            E = Texto(f.Celdas, 4),
            ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase)
            {
                ["Total"] = f.Celdas.Length > 5 ? Numero(f.Celdas[5]) : null,
                ["SERVICIO ESPECIALES"] = f.Celdas.Length > 11 ? Numero(f.Celdas[11]) : null
            }
        }).ToList();

    /// <summary>
    /// Sección de una fila por el predicado vigente del Plan 46
    /// (<see cref="ContratoSeccionPeriodFreeR1.SeccionPorCierre"/>). <c>null</c> = firma AFaseo.
    /// </summary>
    private static SeccionFronteraArriba? Seccion(List<FilaEspejoR1> espejo, int indice) =>
        ContratoSeccionPeriodFreeR1.SeccionPorCierre(espejo, indice);

    /// <summary>Filas-empresa del label por firma <c>C=label ∧ D=Total</c>, antes/después de <c>Componente</c>.</summary>
    private static List<(int Indice, FilaCruda Fila)> FilasEmpresa(List<FilaCruda> ventana, string label, bool antesDeComponente)
    {
        var indiceComponente = ventana.FindIndex(f => Texto(f.Celdas, 0).Equals("Componente", StringComparison.OrdinalIgnoreCase));
        Assert.True(indiceComponente >= 0, "La fuente R1 no trae la fila 'Componente'.");

        var resultado = new List<(int, FilaCruda)>();
        for (var i = 0; i < ventana.Count; i++)
        {
            if (antesDeComponente ? i >= indiceComponente : i <= indiceComponente)
            {
                continue;
            }

            if (Texto(ventana[i].Celdas, 2).Equals(label, StringComparison.OrdinalIgnoreCase)
                && Texto(ventana[i].Celdas, 3).Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                resultado.Add((i, ventana[i]));
            }
        }

        return resultado;
    }

    private static string LabelFuente(int empresaId, int aseId) =>
        empresaId == 1 && aseId is 2 or 4 ? "NUEVO ESQUEMA" : EmpresaFacturacion.Obtener(empresaId).LabelTemplate;

    private static string HojaDe(string reporte) => reporte switch
    {
        "R1" => HojaR1,
        "R2" => HojaR2,
        "R4" => HojaR4,
        _ => throw new ArgumentOutOfRangeException(nameof(reporte))
    };

    private static IEnumerable<(int Ase, int Empresa, string Celda)> DestinosDelMapa(string reporte) =>
        reporte switch
        {
            "R1" => WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa
                .SelectMany(kv => kv.Value.Select(v => (kv.Key.Item2, kv.Key.Item1, v.Celda))),
            "R2" => WorkbookLeafCellMapPorEmpresa.EditablesR2PorEmpresa
                .SelectMany(kv => kv.Value.Select(v => (kv.Key.Item2, kv.Key.Item1, v.Celda))),
            "R4" => WorkbookLeafCellMapPorEmpresa.EditablesR4PorEmpresa
                .SelectMany(kv => kv.Value.Select(v => (kv.Key.Item2, kv.Key.Item1, v.Celda))),
            _ => throw new ArgumentOutOfRangeException(nameof(reporte))
        };

    /// <summary>Términos de una <c>&lt;f&gt;</c> (referencias A1): <c>F47+F70-L16</c> → {F47,F70,L16}.</summary>
    private static HashSet<string> Terminos(string formula) =>
        Regex.Matches(NormalizarFormula(formula), "[A-Z]{1,2}\\d{1,4}")
            .Select(m => m.Value)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Lectura BCL (ZIP + XML crudo) — sin Excel, sin OpenXml SDK, sin ExcelDataReader
    // ════════════════════════════════════════════════════════════════════════════════════════════

    private sealed record CeldaCruda(string? Formula, string? Valor, bool Compartida, bool Existe)
    {
        public bool EsFormula => Formula is not null;
    }

    private sealed record FilaCruda(int Numero, object?[] Celdas);

    private static CeldaCruda CeldaCrudaDe(string xlsx, string hoja, string referencia)
    {
        var (doc, compartidos) = CargarHoja(xlsx, hoja);
        var celda = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "c"
            && string.Equals((string?)e.Attribute("r"), referencia, StringComparison.OrdinalIgnoreCase));

        if (celda is null)
        {
            return new CeldaCruda(null, null, false, false);
        }

        var nodoFormula = celda.Elements().FirstOrDefault(e => e.Name.LocalName == "f");
        var formula = nodoFormula?.Value;
        var compartida = nodoFormula is not null
            && string.Equals((string?)nodoFormula.Attribute("t"), "shared", StringComparison.OrdinalIgnoreCase);
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

        return new CeldaCruda(formula, valor, compartida, true);
    }

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

    private static bool EsMesTotal(object?[] fila) =>
        Texto(fila, 1).Equals("Mes", StringComparison.OrdinalIgnoreCase)
        && Texto(fila, 2).Equals("Total", StringComparison.OrdinalIgnoreCase);

    private static bool EsAplicacionTotal(object?[] fila) =>
        Texto(fila, 1).Contains("Aplicacion nuevos x reversion", StringComparison.OrdinalIgnoreCase)
        && Texto(fila, 2).Equals("Total", StringComparison.OrdinalIgnoreCase);

    private static bool EsTotalFinal(object?[] fila) =>
        Texto(fila, 0).Equals("Total", StringComparison.OrdinalIgnoreCase)
        && string.IsNullOrWhiteSpace(Texto(fila, 1));

    private static string NormalizarFormula(string? formula) =>
        (formula ?? string.Empty).Trim()
            .Replace("'", string.Empty, StringComparison.Ordinal)
            .Replace(" ", string.Empty, StringComparison.Ordinal);

    private static string Sha256(string ruta)
    {
        using var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Convert.ToHexString(SHA256.HashData(stream));
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
}
