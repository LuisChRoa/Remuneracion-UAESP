using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (H0/T0, S1.1-S1.5 — §2-D0): evidencia MEDIDA POR BYTES que habilita (o bloquea) H1.
/// CERO código de producción: este fixture y sus asserts SON la evidencia; nunca escribe nada.
///
/// <para>
/// Lectura SOLO con BCL (<see cref="ZipArchive"/> + <see cref="XDocument"/>): sin Excel/COM, sin
/// DocumentFormat.OpenXml, sin ExcelDataReader, sin <c>Insumos.Raiz()</c>. Rutas EXPLÍCITAS a la
/// plantilla v3, a las históricas v1/v2 (solo para SHA/superconjunto) y a las fuentes reales de
/// octubre-1, julio-1.
/// </para>
///
/// <para>
/// <b>Los cuatro frentes.</b> T0.1 re-mide §0.4 en disco (SHA, filas, cabeceras de bloque,
/// superconjunto de etiquetas, cero <c>&lt;f&gt;</c> en las zonas nuevas). T0.2 fija el VOCABULARIO
/// de cierres por hoja (grupo/subsección/sección/gran total + primera fila de datos) y confirma la
/// semántica <b>trailer</b> de §0.4-d con aritmética de la fuente. T0.3 dictamina, célula por célula,
/// si los destinos CONSOLIDADO/DetRetri son <b>derivables por etiqueta</b>. T0.4 mapea el grupo
/// <c>canal A / empresa EAB</c> del 2.º cuadro de R2 en los 5 ASE y confirma que <c>R2!E285</c> dejó
/// de ser fórmula en la v3.
/// </para>
///
/// <para>
/// <b>Por qué los números van asertados.</b> El propósito del fixture es que una deriva futura de la
/// plantilla (o del disco) ROMPA el test en vez de habilitar H1 sobre evidencia falsa. Cada assert
/// lleva su derivación citada (hoja/fila/columna o archivo).
/// </para>
///
/// <para>
/// <b>Estatus de la v3 (2026-10-27).</b> La v3 quedó <b>HISTÓRICA</b> cuando la v4 pasó a ser la
/// plantilla vigente; este fixture se conserva como <b>evidencia congelada del H0</b> y por eso sigue
/// abriendo el archivo v3 (motivo declarado: re-medir §0.4 sobre el artefacto que se midió y probar que
/// su huella no deriva). Es el ÚNICO consumidor de la v3 que queda en la suite; ningún camino de
/// producción ni fixture de procesamiento la usa. La plantilla vigente (v4) y su huella-gate viven en
/// <see cref="HuellaPlantillaVigenteV4Plan59Tests"/>, que además prueba que la v4 es la v3 más
/// exactamente las dos fórmulas corregidas.
/// </para>
///
/// <para>
/// <b>Sin acoplamiento a los mapas productivos.</b> Este T0 NO referencia
/// <c>WorkbookLeafCellMapQ2</c> ni <c>WorkbookLeafCellMapPorEmpresa</c>: las celdas literales que
/// esos mapas declaran se congelan aquí como dato citado (archivo:línea) y se contrastan contra la
/// resolución por etiqueta. Así H2 puede retirarlos del camino productivo sin tener que editar esta
/// evidencia.
/// </para>
/// </summary>
public sealed class GeometriaDerivadaV3T0Plan59Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // ── G-SHA (§0.3, §0.4): las tres plantillas, byte a byte ────────────────────────────────
    private const string ShaPlantillaV3 = "8199F8647A7ACA84812AD3A61F6607956B2270E0B5C1C999C7690B62032B1295";
    private const string ShaPlantillaV2 = "0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB";
    private const string ShaPlantillaV1 = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";
    private const long LongitudPlantillaV3 = 705_934;
    private const long LongitudPlantillaV2 = 693_400;
    private const long LongitudPlantillaV1 = 692_598;

    // ── Hojas (nombres exactos del workbook; §0.4-a: mismas 40 hojas y mismo orden) ─────────
    private const string HojaR1 = "Reporte Componentes R1";
    private const string HojaR2 = "Rem. Anticipos R2";
    private const string HojaR4 = "Reversion Pagos R4";
    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";

    /// <summary>Hoja DetRetri de la base 2026092 que trae la plantilla (el runtime la renombra por período, Plan 30).</summary>
    private const string HojaDetRetri = "DetRetri2026092";

    /// <summary>Rótulos de identidad de bloque (columna B de R1/R2/R4), en orden 1..5.</summary>
    private static readonly string[] NombresAse =
        ["PROMOAMBIENTAL", "LIME", "CIUDAD LIMPIA", "BOGOTA LIMPIA", "AREA LIMPIA"];

    /// <summary>Nombre largo del ASE en la columna C de CONSOLIDADO / DetRetri (dato, no rama).</summary>
    private static readonly string[] NombresLargoAse =
    [
        "PROMOAMBIENTAL DISTRITO S.A.S. ESP.",
        "LIMPIEZA METROPOLITANA S.A. ESP.",
        "CIUDAD LIMPIA BOGOTA S.A. ESP.",
        "BOGOTA LIMPIA S.A.S. ESP.",
        "AREA LIMPIA DISTRITO CAPITAL S.A.S. ESP."
    ];

    /// <summary>
    /// Nombre largo EXACTO de la columna C de la hoja <c>DetRetri</c>: idéntico al del CONSOLIDADO
    /// salvo <b>ASE3</b>, que en DetRetri va SIN el punto final (`…S.A. ESP` vs `…S.A. ESP.`).
    /// Dato medido, no errata del fixture: la identidad de fila no puede resolverse por igualdad
    /// exacta del rótulo (ver Hallazgo H-59-4 del T0).
    /// </summary>
    private static readonly string[] NombresLargoAseDetRetri =
    [
        "PROMOAMBIENTAL DISTRITO S.A.S. ESP.",
        "LIMPIEZA METROPOLITANA S.A. ESP.",
        "CIUDAD LIMPIA BOGOTA S.A. ESP",
        "BOGOTA LIMPIA S.A.S. ESP.",
        "AREA LIMPIA DISTRITO CAPITAL S.A.S. ESP."
    ];

    private static readonly string[] CarpetasAse =
        ["1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia"];

    // ── §0.4-b: cabeceras de bloque por ASE (fila donde B = rótulo del ASE) ──────────────────
    private static readonly (int Ase, int V2, int V3)[] CabecerasR1 =
        [(1, 4, 4), (2, 108, 122), (3, 287, 356), (4, 448, 517), (5, 601, 725)];

    private static readonly (int Ase, int V2, int V3)[] CabecerasR2 =
        [(1, 1, 1), (2, 65, 75), (3, 173, 196), (4, 291, 317), (5, 402, 441)];

    private static readonly (int Ase, int V2, int V3)[] CabecerasR4 =
        [(1, 1, 1), (2, 91, 97), (3, 192, 198), (4, 229, 235), (5, 344, 350)];

    // ── T0.2: gran total del bloque (A='Total' ∧ B vacío) en la v3 ──────────────────────────
    private static readonly int[] GranTotalR1 = [86, 279, 424, 674, 807];
    private static readonly int[] GranTotalR2 = [43, 147, 231, 389, 483];
    private static readonly int[] GranTotalR4 = [21, 117, 212, 255, 370];

    /// <summary>Primera fila de datos del bloque = cabecera + 2 (fila de encabezado de columnas en medio).</summary>
    private static readonly int[] PrimeraFilaDatosR1 = [6, 124, 358, 519, 727];
    private static readonly int[] PrimeraFilaDatosR2 = [3, 77, 198, 319, 443];
    private static readonly int[] PrimeraFilaDatosR4 = [3, 99, 200, 237, 352];

    // ── T0.2: cierre de SECCIÓN en la v3 (R1/R2: A=sección ∧ B=Total; R4: B=sección ∧ C=Total) ──
    private static readonly (int Fila, string Etiqueta)[][] SeccionesR1V3 =
    [
        [(17, "AFaseo"), (56, "Componente"), (85, "Subs/Cont")],
        [(150, "AFaseo"), (225, "Componente"), (278, "Subs/Cont")],
        [(369, "AFaseo"), (400, "Componente"), (423, "Subs/Cont")],
        [(545, "AFaseo"), (620, "Componente"), (673, "Subs/Cont")],
        [(738, "AFaseo"), (777, "Componente"), (806, "Subs/Cont")]
    ];

    private static readonly (int Fila, string Etiqueta)[][] SeccionesR2V3 =
    [
        [(25, "Componente"), (42, "Subs/Cont")],
        [(117, "Componente"), (146, "Subs/Cont")],
        [(216, "Componente"), (230, "Subs/Cont")],
        [(359, "Componente"), (388, "Subs/Cont")],
        [(465, "Componente"), (482, "Subs/Cont")]
    ];

    private static readonly (int Fila, string Etiqueta)[][] SeccionesR4V3 =
    [
        [(5, "Componente"), (7, "Subs/Cont"), (11, "Componente"), (13, "Subs/Cont"), (17, "Componente"), (19, "Subs/Cont")],
        [(101, "Componente"), (103, "Subs/Cont"), (107, "Componente"), (109, "Subs/Cont"), (113, "Componente"), (115, "Subs/Cont")],
        [(202, "Componente"), (204, "Subs/Cont"), (208, "Componente"), (210, "Subs/Cont")],
        [(239, "Componente"), (241, "Subs/Cont"), (245, "Componente"), (247, "Subs/Cont"), (251, "Componente"), (253, "Subs/Cont")],
        [(354, "Componente"), (356, "Subs/Cont"), (360, "Componente"), (362, "Subs/Cont"), (366, "Componente"), (368, "Subs/Cont")]
    ];

    // ── T0.2: cierre de SUBSECCIÓN en R1 (B=subsección ∧ C=Total). R2/R4 no tienen dimensión. ──
    private static readonly (int Fila, string Etiqueta)[][] SubseccionesR1V3 =
    [
        [(16, "Mes"), (36, "Aplicacion nuevos x reversion"), (55, "Mes"), (70, "Aplicacion nuevos x reversion"), (84, "Mes")],
        [(149, "Mes"), (187, "Aplicacion nuevos x reversion"), (224, "Mes"), (251, "Aplicacion nuevos x reversion"), (277, "Mes")],
        [(368, "Mes"), (384, "Aplicacion nuevos x reversion"), (399, "Mes"), (411, "Aplicacion nuevos x reversion"), (422, "Mes")],
        [(544, "Mes"), (582, "Aplicacion nuevos x reversion"), (619, "Mes"), (646, "Aplicacion nuevos x reversion"), (672, "Mes")],
        [(737, "Mes"), (757, "Aplicacion nuevos x reversion"), (776, "Mes"), (791, "Aplicacion nuevos x reversion"), (805, "Mes")]
    ];

    // ── T0.2: cierre de GRUPO EMPRESA (R1: C=empresa ∧ D=Total; R2: B=empresa ∧ C=Total; R4: A=empresa ∧ B=Total) ──
    private static readonly (int Fila, string Etiqueta)[][] EmpresasR1V3 =
    [
        [(8, "ENEL"), (15, "OCCIDENTE"), (21, "ENEL"), (25, "ENERBIT"), (35, "OCCIDENTE"), (40, "ENEL"), (44, "ENERBIT"),
            (54, "OCCIDENTE"), (59, "ENEL"), (62, "ENERBIT"), (69, "OCCIDENTE"), (73, "ENEL"), (76, "ENERBIT"), (83, "OCCIDENTE")],
        [(126, "ENEL"), (141, "NUEVO ESQUEMA"), (148, "OCCIDENTE"), (154, "ENEL"), (176, "NUEVO ESQUEMA"), (186, "OCCIDENTE"),
            (191, "ENEL"), (213, "NUEVO ESQUEMA"), (223, "OCCIDENTE"), (228, "ENEL"), (243, "NUEVO ESQUEMA"), (250, "OCCIDENTE"),
            (254, "ENEL"), (269, "NUEVO ESQUEMA"), (276, "OCCIDENTE")],
        [(360, "ENEL"), (367, "OCCIDENTE"), (373, "ENEL"), (383, "OCCIDENTE"), (388, "ENEL"), (398, "OCCIDENTE"),
            (403, "ENEL"), (410, "OCCIDENTE"), (414, "ENEL"), (421, "OCCIDENTE")],
        [(521, "ENEL"), (536, "NUEVO ESQUEMA"), (543, "OCCIDENTE"), (549, "ENEL"), (571, "NUEVO ESQUEMA"), (581, "OCCIDENTE"),
            (586, "ENEL"), (608, "NUEVO ESQUEMA"), (618, "OCCIDENTE"), (623, "ENEL"), (638, "NUEVO ESQUEMA"), (645, "OCCIDENTE"),
            (649, "ENEL"), (664, "NUEVO ESQUEMA"), (671, "OCCIDENTE")],
        [(729, "ENEL"), (736, "OCCIDENTE"), (742, "ENEL"), (746, "ENERBIT"), (756, "OCCIDENTE"), (761, "ENEL"), (765, "ENERBIT"),
            (775, "OCCIDENTE"), (780, "ENEL"), (783, "ENERBIT"), (790, "OCCIDENTE"), (794, "ENEL"), (797, "ENERBIT"), (804, "OCCIDENTE")]
    ];

    private static readonly (int Fila, string Etiqueta)[][] EmpresasR2V3 =
    [
        [(6, "EAB"), (10, "ENEL"), (14, "ENERBIT"), (24, "OCCIDENTE"), (28, "EAB"), (31, "ENEL"), (34, "ENERBIT"), (41, "OCCIDENTE")],
        [(80, "EAB"), (84, "ENEL"), (106, "NUEVO ESQUEMA"), (116, "OCCIDENTE"), (120, "EAB"), (123, "ENEL"), (138, "NUEVO ESQUEMA"), (145, "OCCIDENTE")],
        [(201, "EAB"), (205, "ENEL"), (215, "OCCIDENTE"), (219, "EAB"), (222, "ENEL"), (229, "OCCIDENTE")],
        [(322, "EAB"), (326, "ENEL"), (348, "NUEVO ESQUEMA"), (358, "OCCIDENTE"), (362, "EAB"), (365, "ENEL"), (380, "NUEVO ESQUEMA"), (387, "OCCIDENTE")],
        [(446, "EAB"), (450, "ENEL"), (454, "ENERBIT"), (464, "OCCIDENTE"), (468, "EAB"), (471, "ENEL"), (474, "ENERBIT"), (481, "OCCIDENTE")]
    ];

    private static readonly (int Fila, string Etiqueta)[][] EmpresasR4V3 =
    [
        [(8, "ENEL"), (14, "ENERBIT"), (20, "OCCIDENTE")],
        [(104, "ENEL"), (110, "NUEVO ESQUEMA"), (116, "OCCIDENTE")],
        [(205, "ENEL"), (211, "OCCIDENTE")],
        [(242, "ENEL"), (248, "NUEVO ESQUEMA"), (254, "OCCIDENTE")],
        [(357, "ENEL"), (363, "ENERBIT"), (369, "OCCIDENTE")]
    ];

    // ── T0.2: cuántos cierres de GRUPO CANAL declara cada bloque de la v3 (R1: D=canal ∧ E=Total;
    //          R2: C=canal ∧ D=Total; R4: NO tiene dimensión de canal ⇒ 0). Medido en el rango
    //          [cabecera de bloque … gran total] (el cuadro-resumen queda fuera: §2.5 del T0). ──────
    private static readonly int[] CanalesR1V3 = [24, 55, 20, 55, 24];
    private static readonly int[] CanalesR2V3 = [12, 24, 10, 24, 12];

    // ── T0.1: zonas nuevas de la v3 con CERO celdas <f> (§0.4-c) ────────────────────────────
    private static readonly (string Hoja, string Zona, int Inicio, int Fin)[] ZonasNuevasV3 =
    [
        (HojaR1, "R1 ASE2 122-150", 122, 150),
        (HojaR1, "R1 ASE1 55-72", 55, 72),
        (HojaR2, "R2 ASE2 75-106", 75, 106),
        (HojaR4, "R4 ASE1 14-19", 14, 19),
        (HojaR4, "R4 ASE5 350-356", 350, 356)
    ];

    // ── T0.3: las 6 tablas del CONSOLIDADO y la celda destino por ASE (dato citado de
    //          WorkbookLeafCellMapQ2.ConsolidadoProtected, :224-:257) ────────────────────────
    private const string TituloR1 = "R1 - RECAUDO OPORTUNO";
    private const string TituloR2 = "R2 - ANTICIPOS A USUARIOS";
    private const string TituloR3 = "R3 - RECAUDO EXTEMPORANEO";
    private const string TituloR4 = "R4 - REVERSIÓN DE PAGOS";
    private const string TituloR5 = "R5 - REMUNERACIÓN SALDOS A FAVOR POR NOTA";
    private const string TituloTotal = "REPORTE CONSOLIDADO TOTAL DE RECAUDO";

    private static readonly (string Titulo, string PrimeraCeldaLiteral)[] TablasConsolidado =
    [
        (TituloR1, "D9"),
        (TituloR2, "D28"),
        (TituloR3, "D47"),
        (TituloR4, "D66"),
        (TituloR5, "D85"),
        (TituloTotal, "D104")
    ];

    // ── T0.4: grupos `canal A` / empresa `EAB` en R2 (fila del cierre canal → fila del cierre EAB) ──
    private static readonly (int Canal, int Eab)[][] CanalAEabR2V3 =
    [
        [(5, 6), (27, 28)],
        [(79, 80), (119, 120)],
        [(200, 201), (218, 219)],
        [(321, 322), (361, 362)],
        [(445, 446), (467, 468)]
    ];

    private static readonly (int Canal, int Eab)[] CanalAEabR2V2 =
        [(5, 6), (69, 70), (177, 178), (295, 296), (406, 407)];

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.1 / S1.1 — Re-verificación de §0.4 en disco
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void T0_1_Sha256YLongitudDeLasTresPlantillas_CoincidenConElPlan()
    {
        AssertEqualToPlan(PlantillaV3, HuellaSemanticaPlantilla.V3);
        AssertEqualToPlan(PlantillaV2, HuellaSemanticaPlantilla.V2);
        Assert.Equal(HuellaSemanticaPlantilla.V1Sha256Archivo, Sha256(PlantillaV1));
    }

    [Fact]
    public void T0_1_CuarentaHojasConElMismoNombreYOrdenEnV2YV3()
    {
        var v2 = LibroV2.Hojas;
        var v3 = LibroV3.Hojas;

        Assert.Equal(40, v2.Count);
        Assert.Equal(40, v3.Count);
        Assert.Equal(v2.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase), v3.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase));

        // §0.4-a: orden idéntico (no solo el conjunto).
        Assert.Equal(v2.Keys.ToArray(), v3.Keys.ToArray());
    }

    [Fact]
    public void T0_1_DimensionesDeHojaYFilas_CoincidenConLaTabla_0_4_b()
    {
        // §0.4-a: columnas R1 A1:AP, R2 A1:S, R4 A1:R. §0.4-b: total de filas v2 → v3.
        AssertDimension(HojaR1, "A1:AP712", "A1:AP843");
        AssertDimension(HojaR2, "A1:S485", "A1:S527");
        AssertDimension(HojaR4, "A1:R382", "A1:R394");

        // Los otros 37 namespaces de hoja conservan geometría byte-idéntica entre v2 y v3.
        var conGeometriaDistinta = LibroV2.Hojas
            .Where(par => !string.Equals(par.Value.Dimension, LibroV3.Hojas[par.Key].Dimension, StringComparison.OrdinalIgnoreCase))
            .Select(par => par.Key)
            .ToArray();

        Assert.Equal(new[] { HojaR1, HojaR2, HojaR4 }, conGeometriaDistinta);
    }

    [Fact]
    public void T0_1_CabecerasDeBloquePorAse_CoincidenConLaTabla_0_4_b()
    {
        AssertCabeceras(HojaR1, CabecerasR1);
        AssertCabeceras(HojaR2, CabecerasR2);
        AssertCabeceras(HojaR4, CabecerasR4);
    }

    [Fact]
    public void T0_1_ZonasNuevasDeLaV3_SinNingunaCeldaFormula()
    {
        // §0.4-c: las zonas que la v3 estrena son slots de valor puros (regla "solo datos").
        foreach (var (hoja, zona, inicio, fin) in ZonasNuevasV3)
        {
            var h = LibroV3.Hojas[hoja];
            var conFormula = h.Celdas
                .Where(par => par.Value.TieneFormula && FilaDe(par.Key) >= inicio && FilaDe(par.Key) <= fin)
                .Select(par => par.Key)
                .OrderBy(k => k, StringComparer.Ordinal)
                .ToArray();

            Assert.True(
                conFormula.Length == 0,
                $"La zona nueva '{zona}' de la v3 debe tener 0 celdas <f>; encontró [{string.Join(", ", conFormula)}].");
        }
    }

    [Fact]
    public void T0_1_V3EsSuperconjuntoEstrictoDeV2_PorBloqueSeccionYSubseccion()
    {
        foreach (var hoja in new[] { HojaR1, HojaR2, HojaR4 })
        {
            var v2 = MulticonjuntoPorBloqueSeccionSubseccion(LibroV2.Hojas[hoja], hoja, esV3: false);
            var v3 = MulticonjuntoPorBloqueSeccionSubseccion(LibroV3.Hojas[hoja], hoja, esV3: true);

            var faltantes = v2
                .Where(par => !v3.TryGetValue(par.Key, out var enV3) || enV3 < par.Value)
                .Select(par =>
                {
                    v3.TryGetValue(par.Key, out var enV3);
                    return $"'{par.Key}' v2={par.Value} v3={enV3}";
                })
                .OrderBy(x => x, StringComparer.Ordinal)
                .ToArray();

            Assert.True(faltantes.Length == 0, $"{hoja}: la v3 debe contener TODAS las etiquetas de la v2. Faltan: {string.Join(" | ", faltantes)}");

            // Superconjunto ESTRICTO: la v3 declara claves que la v2 no tenía.
            var clavesNuevas = v3.Keys.Count(k => !v2.ContainsKey(k));
            Assert.True(clavesNuevas > 0, $"{hoja}: la v3 no aporta ninguna etiqueta nueva sobre la v2 (esperado superconjunto estricto).");
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.2 / S1.2 — Vocabulario de cierres por hoja (insumo del delimitador de H1)
    // ════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// R1 declara los CUATRO niveles: sección (A/B=Total), subsección (B/C=Total), grupo empresa
    /// (C/D=Total) y grupo canal (D/E=Total) + gran total (A=Total ∧ B vacío). Ambas filas de grupo
    /// son TRAILER: cierran lo que está ARRIBA, nunca encabezan.
    /// </summary>
    [Fact]
    public void T0_2_R1_VocabularioDeCierresPorBloque()
    {
        var hoja = LibroV3.Hojas[HojaR1];
        var bloques = Bloques(hoja, CabecerasR1.Select(c => c.V3).ToArray(), GranTotalR1);

        for (var i = 0; i < bloques.Count; i++)
        {
            var bloque = bloques[i];

            Assert.Equal(i + 1, bloque.AseId);
            Assert.Equal(NombresAse[i], bloque.Rotulo);
            Assert.Equal(PrimeraFilaDatosR1[i], bloque.PrimeraFilaDatos);

            // (e) primera fila de datos = primera fila-concepto tras la cabecera de bloque.
            Assert.Equal("Vlr Servicio", Texto(hoja, "E", bloque.PrimeraFilaDatos));
            Assert.Equal(string.Empty, Texto(hoja, "A", bloque.PrimeraFilaDatos));
            Assert.Equal(string.Empty, Texto(hoja, "B", bloque.PrimeraFilaDatos));
            Assert.Equal(string.Empty, Texto(hoja, "C", bloque.PrimeraFilaDatos));
            Assert.Equal(string.Empty, Texto(hoja, "D", bloque.PrimeraFilaDatos));

            AssertEtiquetas(hoja, "R1 sección", SeccionesR1V3[i], bloque.Secciones);
            AssertEtiquetas(hoja, "R1 subsección", SubseccionesR1V3[i], bloque.Subsecciones);
            AssertEtiquetas(hoja, "R1 grupo empresa", EmpresasR1V3[i], bloque.GruposEmpresa);
            Assert.Equal(CanalesR1V3[i], bloque.GruposCanal.Count);
        }
    }

    /// <summary>
    /// R2 declara TRES niveles: sección (A/B=Total), grupo empresa (B/C=Total) y grupo canal
    /// (C/D=Total). NO tiene subsección: las dos mitades del bloque (anticipos / subsidio) son
    /// secciones distintas (<c>Componente</c> y <c>Subs/Cont</c>), no subsecciones de una misma.
    /// </summary>
    [Fact]
    public void T0_2_R2_VocabularioDeCierresPorBloque()
    {
        var hoja = LibroV3.Hojas[HojaR2];
        var bloques = Bloques(hoja, CabecerasR2.Select(c => c.V3).ToArray(), GranTotalR2);

        for (var i = 0; i < bloques.Count; i++)
        {
            var bloque = bloques[i];

            Assert.Equal(i + 1, bloque.AseId);
            Assert.Equal(NombresAse[i], bloque.Rotulo);
            Assert.Equal(PrimeraFilaDatosR2[i], bloque.PrimeraFilaDatos);
            Assert.Equal("Vlr Servicio", Texto(hoja, "D", bloque.PrimeraFilaDatos));

            AssertEtiquetas(hoja, "R2 sección", SeccionesR2V3[i], bloque.Secciones);
            AssertEtiquetas(hoja, "R2 grupo empresa", EmpresasR2V3[i], bloque.GruposEmpresa);
            Assert.Equal(CanalesR2V3[i], bloque.GruposCanal.Count);

            // Sin dimensión de subsección (medido por bytes): la firma de subsección de R1
            // (`A` vacía ∧ `B`≠'Total' ∧ `C`='Total') coincide EXACTAMENTE con los cierres de EMPRESA
            // de R2 — no existe un tercer nivel entre sección y grupo.
            var firmaSubseccion = Enumerable.Range(bloque.Cabecera, bloque.GranTotal - bloque.Cabecera + 1)
                .Where(f => Texto(hoja, "A", f).Length == 0
                    && Texto(hoja, "B", f).Length > 0
                    && !EsTotal(Texto(hoja, "B", f))
                    && EsTotal(Texto(hoja, "C", f)))
                .ToArray();

            Assert.Equal(bloque.GruposEmpresa.Select(g => g.Fila).ToArray(), firmaSubseccion);
            Assert.DoesNotContain(bloque.GruposEmpresa, g => g.Etiqueta is "Mes" or "Aplicacion nuevos x reversion");
        }
    }

    /// <summary>
    /// R4 no compone: su rótulo de grupo (columna A) YA es la empresa y no hay dimensión de canal.
    /// La sección vive en B/C=Total y cada empresa trae sus dos secciones (Componente, Subs/Cont).
    /// </summary>
    [Fact]
    public void T0_2_R4_VocabularioDeCierresPorBloque_SinDimensionDeCanal()
    {
        var hoja = LibroV3.Hojas[HojaR4];
        var bloques = Bloques(hoja, CabecerasR4.Select(c => c.V3).ToArray(), GranTotalR4);

        for (var i = 0; i < bloques.Count; i++)
        {
            var bloque = bloques[i];

            Assert.Equal(i + 1, bloque.AseId);
            Assert.Equal(NombresAse[i], bloque.Rotulo);
            Assert.Equal(PrimeraFilaDatosR4[i], bloque.PrimeraFilaDatos);
            Assert.Equal("Vlr Servicio", Texto(hoja, "C", bloque.PrimeraFilaDatos));

            AssertEtiquetas(hoja, "R4 sección", SeccionesR4V3[i], bloque.Secciones);
            AssertEtiquetas(hoja, "R4 grupo empresa", EmpresasR4V3[i], bloque.GruposEmpresa);

            // Sin CANAL (medido por bytes): ninguna fila del bloque usa la firma de cierre de canal
            // (`C`≠'' ∧ `D`='Total') y la columna `C` sólo declara CONCEPTOS, nunca canales.
            var firmaCanal = Enumerable.Range(bloque.Cabecera, bloque.GranTotal - bloque.Cabecera + 1)
                .Where(f => Texto(hoja, "C", f).Length > 0 && EsTotal(Texto(hoja, "D", f)))
                .ToArray();

            Assert.Empty(firmaCanal);

            var conceptosR4 = Enumerable.Range(bloque.Cabecera, bloque.GranTotal - bloque.Cabecera + 1)
                .Select(f => Texto(hoja, "C", f))
                .Where(c => c.Length > 0 && !EsTotal(c))
                .Distinct(StringComparer.Ordinal)
                .OrderBy(c => c, StringComparer.Ordinal)
                .ToArray();

            Assert.DoesNotContain(conceptosR4, c => c is "A" or "E" or "H" or "O" or "T" or "1" or "2" or "3" or "4" or "5" or "6" or "9");
        }
    }

    /// <summary>
    /// §0.4-d, evidencia aritmética: en la fuente real R1 de octubre-1, CADA cierre
    /// <c>C=&lt;empresa&gt;/D=Total</c> es la suma exacta (tolerancia de redondeo) de los cierres de
    /// canal <c>D=&lt;canal&gt;/E=Total</c> que lo PRECEDEN dentro de su cuadro. 42/42 cierres.
    /// </summary>
    [Fact]
    public void T0_2_TrailerConfirmadoConAritmeticaDeLaFuenteR1_Octubre1()
    {
        var verificados = 0;
        var totales = 0;

        for (var ase = 1; ase <= 5; ase++)
        {
            var hoja = Cargar(FuenteR1Octubre1(ase)).Hojas.First().Value;
            var max = FilaMaxima(hoja);

            decimal sumaCanales = 0m;
            var canales = new List<string>();

            for (var fila = 1; fila <= max; fila++)
            {
                var a = Texto(hoja, "A", fila);
                var b = Texto(hoja, "B", fila);
                var c = Texto(hoja, "C", fila);
                var d = Texto(hoja, "D", fila);
                var e = Texto(hoja, "E", fila);
                var valor = Numero(Texto(hoja, "F", fila));

                // Cierre de canal: D=<canal> ∧ E=Total, sin A/B/C (dentro del cuadro).
                if (d.Length > 0 && Normalizar(e) == "TOTAL" && a.Length == 0 && b.Length == 0 && c.Length == 0 && Normalizar(d) != "TOTAL")
                {
                    sumaCanales += valor;
                    canales.Add($"{d}={valor.ToString(CultureInfo.InvariantCulture)}");
                    continue;
                }

                // Cierre de empresa: C=<empresa> ∧ D=Total, sin A/B.
                if (c.Length > 0 && Normalizar(d) == "TOTAL" && a.Length == 0 && b.Length == 0 && Normalizar(c) != "TOTAL")
                {
                    totales++;
                    if (Math.Abs(valor - sumaCanales) <= 0.005m)
                    {
                        verificados++;
                    }
                    else
                    {
                        Assert.Fail(
                            $"ASE {ase} ({Path.GetFileName(FuenteR1Octubre1(ase))}), cierre '{c}' fila {fila}: "
                            + $"total={valor.ToString(CultureInfo.InvariantCulture)} ≠ Σ canales={sumaCanales.ToString(CultureInfo.InvariantCulture)} "
                            + $"[{string.Join(", ", canales)}] — el cierre NO es trailer de los canales que lo preceden.");
                    }

                    sumaCanales = 0m;
                    canales.Clear();
                }
            }
        }

        Assert.Equal(42, totales);
        Assert.Equal(42, verificados);
    }

    /// <summary>
    /// Caso citado célula por célula del trailer en la fuente R1 de octubre-1 (ASE2 LIME):
    /// <c>ENEL</c> (fila 16) = canal E (fila 15); <c>OCCIDENTE</c> (fila 22) = H (18) + O (21); y la
    /// <b>subsección</b> <c>Aplicacion nuevos x reversion/Total</c> (fila 23) = ENEL + OCCIDENTE.
    /// Tres niveles de trailer, al centavo.
    /// </summary>
    [Fact]
    public void T0_2_TrailerCasoCitado_Ase2Octubre1_TresNivelesAlCentavo()
    {
        var hoja = Cargar(FuenteR1Octubre1(2)).Hojas.First().Value;

        Assert.Equal("ENEL", Texto(hoja, "C", 16));
        Assert.Equal("Total", Texto(hoja, "D", 16));
        Assert.Equal(11_087_940.08m, Numero(Texto(hoja, "F", 16)));
        Assert.Equal(11_087_940.08m, Numero(Texto(hoja, "F", 15)));

        Assert.Equal("OCCIDENTE", Texto(hoja, "C", 22));
        Assert.Equal("Total", Texto(hoja, "D", 22));
        Assert.Equal(666_142.54m, Numero(Texto(hoja, "F", 22)));
        Assert.Equal(21_770.00m, Numero(Texto(hoja, "F", 18)));
        Assert.Equal(644_372.54m, Numero(Texto(hoja, "F", 21)));
        Assert.Equal(666_142.54m, Numero(Texto(hoja, "F", 18)) + Numero(Texto(hoja, "F", 21)));

        Assert.Equal("Aplicacion nuevos x reversion", Texto(hoja, "B", 23));
        Assert.Equal("Total", Texto(hoja, "C", 23));
        Assert.Equal(11_754_082.62m, Math.Round(Numero(Texto(hoja, "F", 23)), 2));
        Assert.Equal(
            Math.Round(Numero(Texto(hoja, "F", 16)) + Numero(Texto(hoja, "F", 22)), 2),
            Math.Round(Numero(Texto(hoja, "F", 23)), 2));
    }

    /// <summary>
    /// <b>Trampa medida que el delimitador de H1 NO puede ignorar.</b> Debajo del gran total de cada
    /// bloque hay un cuadro-resumen que reutiliza parcialmente el mismo vocabulario:
    /// R1 usa <c>D∈{OPORTUNO,EXTEMPORANEO}/E=Total</c> (no colisiona), pero R2 usa
    /// <c>C=OPORTUNO/D=Total</c> (firma de canal) y R4 usa <c>B=OPORTUNO/C=Total</c> (firma de
    /// sección). Por eso el bloque de datos se delimita [cabecera … gran total] y el resumen queda
    /// FUERA: cualquier regla que escanee hasta el fin del bloque "de filas" los confundiría.
    /// </summary>
    [Fact]
    public void T0_2_ElCuadroResumenReutilizaElVocabulario_PorEsoElBloqueTerminaEnElGranTotal()
    {
        // R2 v3: C='OPORTUNO' ∧ D='Total' debajo del gran total (firma idéntica a un cierre de canal).
        var r2 = LibroV3.Hojas[HojaR2];
        var cabecerasR2V3 = CabecerasR2.Select(c => c.V3).ToArray();
        var resumenR2 = Filas(r2).Where(f => Normalizar(Texto(r2, "C", f)) == "OPORTUNO" && Normalizar(Texto(r2, "D", f)) == "TOTAL").ToArray();
        Assert.NotEmpty(resumenR2);
        foreach (var fila in resumenR2)
        {
            var ase = BordeBloque(cabecerasR2V3, FilaMaxima(r2), fila);
            Assert.True(ase > 0, $"R2 fila {fila}: el cuadro-resumen debe vivir dentro de un bloque.");
            Assert.True(
                fila > GranTotalR2[ase - 1],
                $"R2 fila {fila}: el cuadro-resumen debe quedar DEBAJO del gran total ({GranTotalR2[ase - 1]}) de su bloque.");
        }

        // R4 v3: B='OPORTUNO' ∧ C='Total' (firma idéntica a un cierre de sección).
        var r4 = LibroV3.Hojas[HojaR4];
        var cabecerasR4V3 = CabecerasR4.Select(c => c.V3).ToArray();
        var resumenR4 = Filas(r4).Where(f => Normalizar(Texto(r4, "B", f)) == "OPORTUNO" && Normalizar(Texto(r4, "C", f)) == "TOTAL").ToArray();
        Assert.NotEmpty(resumenR4);
        foreach (var fila in resumenR4)
        {
            var ase = BordeBloque(cabecerasR4V3, FilaMaxima(r4), fila);
            Assert.True(ase > 0, $"R4 fila {fila}: el cuadro-resumen debe vivir dentro de un bloque.");
            Assert.True(
                fila > GranTotalR4[ase - 1],
                $"R4 fila {fila}: el cuadro-resumen debe quedar DEBAJO del gran total ({GranTotalR4[ase - 1]}) de su bloque.");
        }

        // R1 v3: el resumen usa D∈{OPORTUNO,EXTEMPORANEO}/E=Total (columna del canal), nunca D='Total'.
        var r1 = LibroV3.Hojas[HojaR1];
        var cabecerasR1V3 = CabecerasR1.Select(c => c.V3).ToArray();
        var resumenR1 = Filas(r1)
            .Where(f => Normalizar(Texto(r1, "E", f)) == "TOTAL"
                && (Normalizar(Texto(r1, "D", f)) == "OPORTUNO" || Normalizar(Texto(r1, "D", f)) == "EXTEMPORANEO"))
            .ToArray();
        Assert.NotEmpty(resumenR1);
        foreach (var fila in resumenR1)
        {
            var ase = BordeBloque(cabecerasR1V3, FilaMaxima(r1), fila);
            Assert.True(ase > 0, $"R1 fila {fila}: el cuadro-resumen debe vivir dentro de un bloque.");
            Assert.True(
                fila > GranTotalR1[ase - 1],
                $"R1 fila {fila}: el cuadro-resumen debe quedar DEBAJO del gran total ({GranTotalR1[ase - 1]}) de su bloque.");
        }
    }

    /// <summary>
    /// Cardinalidad de subsecciones: la PLANTILLA declara 5 subsecciones en los 5 bloques de R1
    /// (constante), pero la FUENTE no las trae todas. Julio-1 (1.ª quincena) no trae
    /// <c>Aplicacion nuevos x reversion</c> en ningún cuadro (3 subsecciones), mientras octubre-1 sí
    /// trae las 5. El delimitador/lector no puede congelar la cardinalidad de la fuente.
    /// </summary>
    [Fact]
    public void T0_2_SubseccionesAusentesEnLaFuenteQ1_NoSaturanCardinalidad()
    {
        // Plantilla: 5 subsecciones por bloque, siempre las mismas 5 (2 'Aplicacion' + 3 'Mes').
        foreach (var esperadas in SubseccionesR1V3)
        {
            Assert.Equal(5, esperadas.Length);
            Assert.Equal(2, esperadas.Count(s => s.Etiqueta == "Aplicacion nuevos x reversion"));
            Assert.Equal(3, esperadas.Count(s => s.Etiqueta == "Mes"));
        }

        var julio1 = Cargar(FuenteR1Julio1(2)).Hojas.First().Value;
        var octubre1 = Cargar(FuenteR1Octubre1(2)).Hojas.First().Value;

        var subJulio1 = CierresSubseccion(julio1);
        var subOctubre1 = CierresSubseccion(octubre1);

        Assert.Equal(3, subJulio1.Count);
        Assert.Equal(5, subOctubre1.Count);
        Assert.DoesNotContain(subJulio1, s => s.Etiqueta == "Aplicacion nuevos x reversion");
        Assert.Equal(2, subOctubre1.Count(s => s.Etiqueta == "Aplicacion nuevos x reversion"));

        // La fuente SÍ conserva las 3 secciones en ambas quincenas (la ausencia es de subsección).
        Assert.Equal(new[] { "AFaseo", "Componente", "Subs/Cont" }, CierresSeccion(julio1).Select(s => s.Etiqueta).ToArray());
        Assert.Equal(new[] { "AFaseo", "Componente", "Subs/Cont" }, CierresSeccion(octubre1).Select(s => s.Etiqueta).ToArray());
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.3 / S1.3 — ¿Los destinos CONSOLIDADO/DetRetri son derivables por ETIQUETA?
    // ════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Veredicto <b>derivable-por-etiqueta</b> para las 6 tablas del CONSOLIDADO. La regla que
    /// resuelve cada destino sin ningún literal es:
    /// (1) fila-título = la única fila con el rótulo en <c>F</c> que además declara un encabezado
    ///     <c>B='ASE' / C='Concesionarios'</c> DEBAJO;
    /// (2) fila del ASE = la primera fila del bloque cuya <c>B</c> = id del ASE y cuya <c>C</c> = su
    ///     nombre largo;
    /// (3) columna = la columna de valor de esa tabla (D).
    /// El resultado se contrasta contra la celda literal que declara el mapa productivo.
    /// </summary>
    [Fact]
    public void T0_3_Consolidado_DestinosResolublesPorEtiquetaDeTablaYAse()
    {
        var hoja = LibroV3.Hojas[HojaConsolidado];

        foreach (var (titulo, primeraCeldaLiteral) in TablasConsolidado)
        {
            var tabla = ResolverTablaConsolidado(hoja, titulo);

            for (var ase = 1; ase <= 5; ase++)
            {
                var fila = tabla.FilaAse[ase];
                var celda = $"D{fila}";

                Assert.Equal(NombresLargoAse[ase - 1], Texto(hoja, "C", fila));
                Assert.Equal(ase.ToString(CultureInfo.InvariantCulture), Texto(hoja, "B", fila));

                // La celda que resuelve la etiqueta es EXACTAMENTE la que el mapa congeló.
                Assert.Equal(CeldaLiteralEsperada(primeraCeldaLiteral, ase), celda);
            }

            // Fila de total de la tabla: C='TOTAL' (etiqueta que cierra el bloque de ASE).
            var filaTotal = tabla.FilaAse[5] + 1;
            Assert.Equal("TOTAL", Texto(hoja, "C", filaTotal));
            Assert.Equal(string.Empty, Texto(hoja, "B", filaTotal));
        }
    }

    /// <summary>
    /// Ambigüedad real medida: el rótulo <c>R2 - ANTICIPOS A USUARIOS</c> aparece DOS veces en la
    /// columna F (fila 25 = título de tabla; fila 141 = rótulo de columna de la tabla de validación).
    /// La etiqueta sola NO basta: el título de tabla se distingue porque declara el encabezado
    /// <c>B='ASE'/C='Concesionarios'</c> debajo. Es el único título con colisión en la v3.
    /// </summary>
    [Fact]
    public void T0_3_Consolidado_UnTituloColisionaConUnRotuloDeColumna_YSeDesambiguaPorElEncabezado()
    {
        var hoja = LibroV3.Hojas[HojaConsolidado];

        var candidatasR2 = Filas(hoja).Where(f => Normalizar(Texto(hoja, "F", f)) == Normalizar(TituloR2)).ToArray();
        Assert.Equal(new[] { 25, 141 }, candidatasR2);

        // Sólo la fila 25 declara el encabezado de tabla debajo; la 141 es un rótulo de columna.
        Assert.Equal(27, EncabezadoTablaConsolidado(hoja, 25));
        Assert.Throws<InvalidOperationException>(() => EncabezadoTablaConsolidado(hoja, 141));

        foreach (var (titulo, _) in TablasConsolidado)
        {
            var resueltas = Filas(hoja)
                .Where(f => Normalizar(Texto(hoja, "F", f)) == Normalizar(titulo))
                .Where(f => TryEncabezadoTablaConsolidado(hoja, f, out _))
                .ToArray();
            Assert.Single(resueltas);
        }
    }

    /// <summary>
    /// Los destinos del CONSOLIDADO son <b>fórmulas</b> en la v3 (protegidas): la derivación por
    /// etiqueta sirve para IDENTIFICAR la celda, y la protección se verifica por identidad semántica
    /// (no por coordenada literal). <c>D106</c> es seguidora de la compartida <c>si=47</c> cuyo
    /// maestro es <c>D104</c> (<c>ref="D104:F106"</c>): su <c>&lt;f&gt;</c> viene sin texto.
    /// </summary>
    [Fact]
    public void T0_3_Consolidado_DestinosSonFormulaEnLaV3_ConD106SeguidoraDelMaestroD104()
    {
        var hoja = LibroV3.Hojas[HojaConsolidado];

        foreach (var celda in new[]
                 {
                     "D9", "D10", "D11", "D12", "D13",
                     "D28", "D29", "D30", "D31", "D32",
                     "D47", "D48", "D49", "D50", "D51",
                     "D66", "D67", "D68", "D69", "D70",
                     "D85", "D86", "D87", "D88", "D89",
                     "D104", "D105", "D106", "D107", "D108", "D109"
                 })
        {
            Assert.True(hoja.Celdas.TryGetValue(celda, out var c), $"Falta la celda protegida '{HojaConsolidado}!{celda}'.");
            Assert.True(c!.TieneFormula, $"'{HojaConsolidado}!{celda}' debe ser fórmula en la v3 (destino protegido por identidad).");
        }

        var maestro = hoja.Celdas["D104"];
        Assert.Equal("shared", maestro.TipoFormula);
        Assert.Equal("47", maestro.IndiceCompartida);
        Assert.Equal("D104:F106", maestro.RefCompartida);
        Assert.Equal("D9+D28+D47+D66+D85", maestro.Formula);

        var seguidora = hoja.Celdas["D106"];
        Assert.Equal("shared", seguidora.TipoFormula);
        Assert.Equal("47", seguidora.IndiceCompartida);
        Assert.Null(seguidora.RefCompartida);
        Assert.Equal(string.Empty, seguidora.Formula);
    }

    /// <summary>
    /// Veredicto <b>derivable-por-etiqueta</b> para <c>DetRetriDestinoPorAse</c>
    /// (<c>WorkbookLeafCellMapQ2.cs:264</c>: ASE1..5 → D9..D13). La tabla se identifica por su
    /// encabezado <c>B='ASE' / C='CONCESIONARIOS'</c> — que es ÚNICO en la hoja — y cada fila por
    /// <c>B=&lt;id&gt;</c> + <c>C=&lt;nombre largo&gt;</c>. Las celdas son literales (escribibles).
    /// </summary>
    [Fact]
    public void T0_3_DetRetri_DestinosResolublesPorEtiquetaYEsUnicos()
    {
        var hoja = LibroV3.Hojas[HojaDetRetri];

        var encabezados = Filas(hoja)
            .Where(f => Normalizar(Texto(hoja, "B", f)) == "ASE" && Normalizar(Texto(hoja, "C", f)) == "CONCESIONARIOS")
            .ToArray();

        Assert.Equal(new[] { 8 }, encabezados);
        Assert.Equal("RECAUDO TOTAL", Normalizar(Texto(hoja, "D", 8)));

        for (var ase = 1; ase <= 5; ase++)
        {
            var fila = 8 + ase;
            Assert.Equal(ase.ToString(CultureInfo.InvariantCulture), Texto(hoja, "B", fila));
            Assert.Equal(NombresLargoAseDetRetri[ase - 1], Texto(hoja, "C", fila));

            // D9..D13 = el literal congelado por el mapa productivo (dato citado, no dependencia).
            var celda = $"D{fila}";
            Assert.True(hoja.Celdas.TryGetValue(celda, out var contenido), $"Falta 'DetRetri!{celda}'.");
            Assert.False(contenido!.TieneFormula, $"'DetRetri!{celda}' es destino de ESCRITURA: no puede ser fórmula.");
            Assert.Equal("0", contenido.Valor);
        }

        // Hallazgo H-59-4: el rótulo de ASE3 difiere del del CONSOLIDADO por el punto final.
        Assert.Equal("CIUDAD LIMPIA BOGOTA S.A. ESP.", NombresLargoAse[2]);
        Assert.Equal("CIUDAD LIMPIA BOGOTA S.A. ESP", NombresLargoAseDetRetri[2]);
        Assert.Equal(NombresLargoAse[2].TrimEnd('.'), NombresLargoAseDetRetri[2]);
        Assert.NotEqual(NombresLargoAse[2], NombresLargoAseDetRetri[2]);

        // Fila de total (D14 = DetRetriTotal del mapa): C='Total' cierra la tabla.
        Assert.Equal("Total", Texto(hoja, "C", 14));
        Assert.False(hoja.Celdas["D14"].TieneFormula);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.4 / S1.4 — Grupo `canal A` / empresa `EAB` del 2.º cuadro de R2 (hallazgo §0.5)
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void T0_4_R2V3_ElGrupoCanalAEmpresaEabSeDeclaraEnLosCincoAseYDosVeces()
    {
        var hoja = LibroV3.Hojas[HojaR2];
        var cabeceras = CabecerasR2.Select(c => c.V3).ToArray();

        for (var ase = 1; ase <= 5; ase++)
        {
            var canalA = Filas(hoja)
                .Where(f => Normalizar(Texto(hoja, "C", f)) == "A" && Normalizar(Texto(hoja, "D", f)) == "TOTAL")
                .Where(f => f >= cabeceras[ase - 1] && f < (ase < 5 ? cabeceras[ase] : FilaMaxima(hoja) + 1))
                .ToArray();

            Assert.Equal(CanalAEabR2V3[ase - 1].Select(p => p.Canal).ToArray(), canalA);

            foreach (var (canal, eab) in CanalAEabR2V3[ase - 1])
            {
                Assert.Equal("A", Texto(hoja, "C", canal));
                Assert.Equal("Total", Texto(hoja, "D", canal));
                Assert.Equal(string.Empty, Texto(hoja, "B", canal));

                // El cierre EAB es el TRAILER inmediatamente debajo: B='EAB' ∧ C='Total'.
                Assert.Equal(eab, canal + 1);
                Assert.Equal("EAB", Texto(hoja, "B", eab));
                Assert.Equal("Total", Texto(hoja, "C", eab));
            }

            // 1.er grupo bajo la sección `Componente`, 2.º bajo `Subs/Cont` (§0.5: el 2.º cuadro).
            Assert.Equal(SeccionesR2V3[ase - 1][0].Fila, SectionQueCierra(hoja, CanalAEabR2V3[ase - 1][0].Canal));
            Assert.Equal(SeccionesR2V3[ase - 1][1].Fila, SectionQueCierra(hoja, CanalAEabR2V3[ase - 1][1].Canal));
        }
    }

    [Fact]
    public void T0_4_R2V2_DeclaraUnSoloGrupoCanalAEnElPrimerCuadro_EvidenciaElHueco()
    {
        var hoja = LibroV2.Hojas[HojaR2];
        var cabeceras = CabecerasR2.Select(c => c.V2).ToArray();

        for (var ase = 1; ase <= 5; ase++)
        {
            var finBloque = ase < 5 ? cabeceras[ase] - 1 : FilaMaxima(hoja);

            var canalA = Filas(hoja)
                .Where(f => Normalizar(Texto(hoja, "C", f)) == "A" && Normalizar(Texto(hoja, "D", f)) == "TOTAL")
                .Where(f => f >= cabeceras[ase - 1] && f <= finBloque)
                .ToArray();

            Assert.Equal(new[] { CanalAEabR2V2[ase - 1].Canal }, canalA);

            var (canal, eab) = CanalAEabR2V2[ase - 1];
            Assert.Equal("A", Texto(hoja, "C", canal));
            Assert.Equal("EAB", Texto(hoja, "B", eab));
            Assert.Equal("Total", Texto(hoja, "C", eab));

            // En la v2 el bloque declara DOS secciones y el grupo `canal A` sólo existe en la primera:
            // la segunda (el 2.º cuadro) NO lo declara — ése es exactamente el hueco de §0.5.
            var secciones = Filas(hoja)
                .Where(f => f >= cabeceras[ase - 1] && f <= finBloque)
                .Where(f => Texto(hoja, "A", f).Length > 0 && EsTotal(Texto(hoja, "B", f)) && Texto(hoja, "C", f).Length == 0)
                .ToArray();

            Assert.Equal(2, secciones.Length);
            Assert.True(canal < secciones[0], $"R2 v2 ASE {ase}: el grupo `canal A` debe caer en el 1.er cuadro (cierre {secciones[0]}).");
            Assert.DoesNotContain(canalA, f => f > secciones[0]);
        }
    }

    /// <summary>
    /// <c>R2!E285</c>: en la v2 era <c>&lt;f&gt;E192+E203-K192&lt;/f&gt;</c> (el bloqueo
    /// <c>R-MAPA-V2-R2-ASE4</c>); en la v3 la celda existe <b>vacía</b> (sin <c>&lt;f&gt;</c> y sin
    /// <c>&lt;v&gt;</c>). <c>K285</c> sigue el mismo camino (v2: seguidora compartida <c>si=9</c>).
    /// </summary>
    [Fact]
    public void T0_4_E285_NoEsFormulaEnLaV3_YSiLoEraEnLaV2()
    {
        var v2 = LibroV2.Hojas[HojaR2];
        var v3 = LibroV3.Hojas[HojaR2];

        Assert.True(v2.Celdas.TryGetValue("E285", out var e285V2));
        Assert.True(e285V2!.TieneFormula);
        Assert.Equal("E192+E203-K192", e285V2.Formula);

        Assert.True(v3.Celdas.TryGetValue("E285", out var e285V3));
        Assert.False(e285V3!.TieneFormula, "R2!E285 ya no es fórmula en la v3 (§0.5).");
        Assert.Null(e285V3.Valor);

        Assert.True(v2.Celdas.TryGetValue("K285", out var k285V2));
        Assert.True(k285V2!.TieneFormula);
        Assert.Equal("shared", k285V2.TipoFormula);
        Assert.Equal("9", k285V2.IndiceCompartida);

        Assert.True(v3.Celdas.TryGetValue("K285", out var k285V3));
        Assert.False(k285V3!.TieneFormula);
        Assert.Null(k285V3.Valor);
    }

    /// <summary>
    /// <b>Precisión medida sobre §0.5 (no contradice el plan, lo acota).</b> En la v3 el literal
    /// <c>E285</c> dejó de ser fórmula, pero tampoco cae dentro del bloque de ASE4 (317..440): está
    /// en el bloque de ASE3 (196..316), por debajo de su gran total (231) y antes del cuadro-resumen
    /// (299). Es decir: el literal no bloquea la escritura, pero TAMPOCO es el destino del ASE4 —
    /// escribir ahí sería una colocación silenciosa en el bloque equivocado. Conclusión para H2: el
    /// destino DEBE derivarse por etiqueta, nunca "usar E285 porque ya no es fórmula".
    /// </summary>
    [Fact]
    public void T0_4_E285EnLaV3CaeFueraDelBloqueDeAse4_PorEsoElDestinoDebeDerivarsePorEtiqueta()
    {
        var hoja = LibroV3.Hojas[HojaR2];
        var cabeceras = CabecerasR2.Select(c => c.V3).ToArray();

        var bloqueDeE285 = BloqueQueContiene(cabeceras, FilaMaxima(hoja), 285);
        Assert.Equal(3, bloqueDeE285);

        Assert.True(285 < cabeceras[3], "El literal E285 debe quedar por ENCIMA de la cabecera del ASE4 (317).");
        Assert.True(285 > GranTotalR2[2], "El literal E285 debe quedar por DEBAJO del gran total del ASE3 (231).");

        // El cierre `canal A / EAB` del 2.º cuadro de ASE4 (resuelto por etiqueta) vive en la fila
        // 362, no en 285: `E362` es la celda de valor de ese cierre y es un literal escribible.
        Assert.Equal((361, 362), CanalAEabR2V3[3][1]);
        Assert.True(hoja.Celdas.TryGetValue("E362", out var e362), "Falta 'R2!E362' (cierre EAB del 2.º cuadro de ASE4).");
        Assert.False(e362!.TieneFormula, "'R2!E362' debe ser un slot de valor, no una fórmula.");
    }

    /// <summary>
    /// Fuente octubre-1: sólo ASE2 (LIME) trae grupos <c>canal A / empresa EAB</c> — dos, uno por
    /// cuadro — con <c>+149.137,87</c> (1.º) y <c>−18.613,87</c> (2.º). ASE1/3/4/5 traen CERO grupos
    /// <c>canal A</c> en ese período (§0.5: la ausencia es normal y cae en el 0 del template).
    /// </summary>
    [Fact]
    public void T0_4_FuenteOctubre1_SoloAse2TraeDosGruposCanalAConLosValoresDe_0_5()
    {
        var grupos = new Dictionary<int, List<(int Fila, decimal Valor)>>();

        for (var ase = 1; ase <= 5; ase++)
        {
            var hoja = Cargar(FuenteR2Octubre1(ase)).Hojas.First().Value;
            var encontrados = Filas(hoja)
                .Where(f => Normalizar(Texto(hoja, "C", f)) == "A" && Normalizar(Texto(hoja, "D", f)) == "TOTAL")
                .Select(f =>
                {
                    // El cierre de empresa EAB es el trailer inmediatamente debajo.
                    var eab = f + 1;
                    Assert.Equal("EAB", Texto(hoja, "B", eab));
                    Assert.Equal("Total", Texto(hoja, "C", eab));
                    return (Fila: f, Valor: Numero(Texto(hoja, "E", eab)));
                })
                .ToList();

            grupos[ase] = encontrados;
        }

        Assert.Equal(new[] { 6, 35 }, grupos[2].Select(g => g.Fila).ToArray());
        Assert.Equal(149_137.87m, grupos[2][0].Valor);
        Assert.Equal(-18_613.87m, grupos[2][1].Valor);

        foreach (var ase in new[] { 1, 3, 4, 5 })
        {
            Assert.Empty(grupos[ase]);
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // T0.5 / S1.5 — G-SHA: la plantilla NO se toca (antes y después de medir)
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void T0_5_PlantillasByteIdenticasAntesYDespuesDeTodasLasMediciones()
    {
        var rutas = new[]
        {
            (Ruta: PlantillaV1, Huella: HuellaSemanticaPlantilla.V1),
            (Ruta: PlantillaV2, Huella: HuellaSemanticaPlantilla.V2),
            (Ruta: PlantillaV3, Huella: HuellaSemanticaPlantilla.V3),
        };
        var antes = rutas.Select(p => (p.Ruta, Sha: Sha256(p.Ruta), Longitud: new FileInfo(p.Ruta).Length, Huella: HuellaSemanticaPlantilla.Calcular(p.Ruta))).ToArray();

        // Ejercita los cuatro frentes completos (lectura pura) entre las dos mediciones de SHA.
        foreach (var hoja in new[] { HojaR1, HojaR2, HojaR4 })
        {
            _ = Bloques(LibroV3.Hojas[hoja], CabecerasDe(hoja, esV3: true), GranTotalDe(hoja));
            _ = MulticonjuntoPorBloqueSeccionSubseccion(LibroV2.Hojas[hoja], hoja, esV3: false);
            _ = MulticonjuntoPorBloqueSeccionSubseccion(LibroV3.Hojas[hoja], hoja, esV3: true);
        }

        _ = ResolverTablaConsolidado(LibroV3.Hojas[HojaConsolidado], TituloTotal);
        _ = Cargar(FuenteR2Octubre1(2)).Hojas.First().Value;

        var despues = rutas.Select(p => (p.Ruta, Sha: Sha256(p.Ruta), Longitud: new FileInfo(p.Ruta).Length, Huella: HuellaSemanticaPlantilla.Calcular(p.Ruta))).ToArray();

        Assert.Equal(antes, despues);

        for (var i = 0; i < rutas.Length; i++)
        {
            Assert.Equal(rutas[i].Huella, despues[i].Huella);
            Assert.True(despues[i].Longitud > 0, $"La plantilla '{Path.GetFileName(rutas[i].Ruta)}' quedó vacía.");
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // Estructura medida por bloque
    // ════════════════════════════════════════════════════════════════════════════════════════

    private sealed record Bloque(
        int AseId,
        string Rotulo,
        int Cabecera,
        int Fin,
        int GranTotal,
        int PrimeraFilaDatos,
        IReadOnlyList<(int Fila, string Etiqueta)> Secciones,
        IReadOnlyList<(int Fila, string Etiqueta)> Subsecciones,
        IReadOnlyList<(int Fila, string Etiqueta)> GruposEmpresa,
        IReadOnlyList<(int Fila, string Etiqueta)> GruposCanal);

    private static List<Bloque> Bloques(Hoja hoja, int[] cabeceras, int[] granTotales)
    {
        var max = FilaMaxima(hoja);
        var bloques = new List<Bloque>(cabeceras.Length);

        for (var i = 0; i < cabeceras.Length; i++)
        {
            var cabecera = cabeceras[i];
            var fin = i + 1 < cabeceras.Length ? cabeceras[i + 1] - 1 : max;
            var granTotal = granTotales[i];

            var secciones = new List<(int, string)>();
            var subsecciones = new List<(int, string)>();
            var empresas = new List<(int, string)>();
            var canales = new List<(int, string)>();

            for (var fila = cabecera; fila <= granTotal; fila++)
            {
                var a = Texto(hoja, "A", fila);
                var b = Texto(hoja, "B", fila);
                var c = Texto(hoja, "C", fila);
                var d = Texto(hoja, "D", fila);
                var e = Texto(hoja, "E", fila);

                var aVacia = a.Length == 0;
                var bVacia = b.Length == 0;
                var cVacia = c.Length == 0;

                // ── R1 ──
                if (string.Equals(hoja.Nombre, HojaR1, StringComparison.OrdinalIgnoreCase))
                {
                    if (!aVacia && EsTotal(b) && cVacia) secciones.Add((fila, a));
                    else if (aVacia && !bVacia && !EsTotal(b) && EsTotal(c)) subsecciones.Add((fila, b));
                    else if (aVacia && bVacia && !cVacia && EsTotal(d)) empresas.Add((fila, c));
                    else if (aVacia && bVacia && cVacia && d.Length > 0 && !EsTotal(d) && EsTotal(e)) canales.Add((fila, d));
                }
                // ── R2 ──
                else if (string.Equals(hoja.Nombre, HojaR2, StringComparison.OrdinalIgnoreCase))
                {
                    if (!aVacia && EsTotal(b) && cVacia) secciones.Add((fila, a));
                    else if (aVacia && !bVacia && !EsTotal(b) && EsTotal(c)) empresas.Add((fila, b));
                    else if (aVacia && bVacia && !cVacia && !EsTotal(c) && EsTotal(d)) canales.Add((fila, c));
                }
                // ── R4 ──
                else
                {
                    if (!aVacia && !EsTotal(a) && EsTotal(b)) empresas.Add((fila, a));
                    else if (aVacia && !bVacia && !EsTotal(b) && EsTotal(c)) secciones.Add((fila, b));
                }
            }

            bloques.Add(new Bloque(
                i + 1,
                Texto(hoja, "B", cabecera),
                cabecera,
                fin,
                granTotal,
                cabecera + 2,
                secciones,
                subsecciones,
                empresas,
                canales));
        }

        return bloques;
    }

    private static Dictionary<string, int> MulticonjuntoPorBloqueSeccionSubseccion(Hoja hoja, string nombreHoja, bool esV3)
    {
        var cabeceras = CabecerasDe(nombreHoja, esV3);
        var granTotales = GranTotalDe(nombreHoja);
        var max = FilaMaxima(hoja);
        var acumulado = new Dictionary<string, int>(StringComparer.Ordinal);

        for (var i = 0; i < cabeceras.Length; i++)
        {
            var cabecera = cabeceras[i];
            var fin = i + 1 < cabeceras.Length ? cabeceras[i + 1] - 1 : max;

            var seccionPorFila = new string[fin + 2];
            var subseccionPorFila = new string[fin + 2];

            for (var fila = fin; fila >= cabecera; fila--)
            {
                string? cierreSeccion = null;
                string? cierreSubseccion = null;

                if (nombreHoja == HojaR4)
                {
                    if (Texto(hoja, "B", fila).Length > 0 && EsTotal(Texto(hoja, "C", fila))) cierreSeccion = Texto(hoja, "B", fila);
                }
                else
                {
                    if (Texto(hoja, "A", fila).Length > 0 && EsTotal(Texto(hoja, "B", fila))) cierreSeccion = Texto(hoja, "A", fila);

                    // Subsección: la única hoja con dimensión de subsección es R1 (B=Mes|Aplicacion / C=Total).
                    if (nombreHoja == HojaR1
                        && Texto(hoja, "B", fila).Length > 0
                        && !EsTotal(Texto(hoja, "B", fila))
                        && EsTotal(Texto(hoja, "C", fila)))
                    {
                        cierreSubseccion = Texto(hoja, "B", fila);
                    }
                }

                seccionPorFila[fila] = cierreSeccion ?? (fila + 1 < seccionPorFila.Length ? seccionPorFila[fila + 1] : string.Empty);
                subseccionPorFila[fila] = cierreSubseccion ?? (fila + 1 < subseccionPorFila.Length ? subseccionPorFila[fila + 1] : string.Empty);
            }

            for (var fila = cabecera; fila <= fin; fila++)
            {
                var etiqueta = string.Join(
                    '|',
                    Texto(hoja, "A", fila),
                    Texto(hoja, "B", fila),
                    Texto(hoja, "C", fila),
                    Texto(hoja, "D", fila),
                    Texto(hoja, "E", fila));

                if (etiqueta.Replace("|", string.Empty, StringComparison.Ordinal).Length == 0)
                {
                    continue;
                }

                var clave = $"ASE{i + 1}/{seccionPorFila[fila]}/{subseccionPorFila[fila]}::{etiqueta}";
                acumulado[clave] = acumulado.TryGetValue(clave, out var cuenta) ? cuenta + 1 : 1;
            }
        }

        return acumulado;
    }

    private static int[] CabecerasDe(string nombreHoja, bool esV3) => nombreHoja switch
    {
        HojaR1 => CabecerasR1.Select(c => esV3 ? c.V3 : c.V2).ToArray(),
        HojaR2 => CabecerasR2.Select(c => esV3 ? c.V3 : c.V2).ToArray(),
        HojaR4 => CabecerasR4.Select(c => esV3 ? c.V3 : c.V2).ToArray(),
        _ => throw new ArgumentOutOfRangeException(nameof(nombreHoja), nombreHoja, "Hoja sin bloques por ASE.")
    };

    private static int[] GranTotalDe(string nombreHoja) => nombreHoja switch
    {
        HojaR1 => GranTotalR1,
        HojaR2 => GranTotalR2,
        HojaR4 => GranTotalR4,
        _ => throw new ArgumentOutOfRangeException(nameof(nombreHoja), nombreHoja, "Hoja sin gran total por ASE.")
    };

    // ════════════════════════════════════════════════════════════════════════════════════════
    // Resolución por etiqueta del CONSOLIDADO (§0.2: conceptos, no coordenadas)
    // ════════════════════════════════════════════════════════════════════════════════════════

    private sealed record TablaConsolidado(int FilaTitulo, int FilaEncabezado, IReadOnlyDictionary<int, int> FilaAse);

    private static TablaConsolidado ResolverTablaConsolidado(Hoja hoja, string titulo)
    {
        var filaTitulo = Filas(hoja)
            .Where(f => Normalizar(Texto(hoja, "F", f)) == Normalizar(titulo))
            .Where(f => TryEncabezadoTablaConsolidado(hoja, f, out _))
            .DefaultIfEmpty(-1)
            .Single();

        Assert.True(filaTitulo > 0, $"CONSOLIDADO: no se resolvió la tabla con título '{titulo}'.");

        var encabezado = EncabezadoTablaConsolidado(hoja, filaTitulo);
        var filas = new Dictionary<int, int>();

        for (var ase = 1; ase <= 5; ase++)
        {
            var esperado = ase.ToString(CultureInfo.InvariantCulture);
            var fila = Filas(hoja)
                .Where(f => f > encabezado)
                .FirstOrDefault(f => Texto(hoja, "B", f) == esperado && Normalizar(Texto(hoja, "C", f)).StartsWith(Normalizar(NombresLargoAse[ase - 1]), StringComparison.Ordinal));

            Assert.True(fila > 0, $"CONSOLIDADO '{titulo}': falta la fila del ASE {ase} (B={esperado}).");
            filas[ase] = fila;
        }

        return new TablaConsolidado(filaTitulo, encabezado, filas);
    }

    /// <summary>Encabezado de tabla = primera fila debajo del título con <c>B='ASE'</c> y <c>C='Concesionarios'</c>.</summary>
    private static int EncabezadoTablaConsolidado(Hoja hoja, int filaTitulo)
    {
        if (!TryEncabezadoTablaConsolidado(hoja, filaTitulo, out var fila))
        {
            throw new InvalidOperationException($"CONSOLIDADO: la fila {filaTitulo} no declara encabezado de tabla (B=ASE / C=Concesionarios) debajo.");
        }

        return fila;
    }

    private static bool TryEncabezadoTablaConsolidado(Hoja hoja, int filaTitulo, out int filaEncabezado)
    {
        filaEncabezado = Filas(hoja)
            .Where(f => f > filaTitulo)
            .FirstOrDefault(f => Normalizar(Texto(hoja, "B", f)) == "ASE" && Normalizar(Texto(hoja, "C", f)) == "CONCESIONARIOS");

        return filaEncabezado > 0;
    }

    private static string CeldaLiteralEsperada(string primeraCeldaLiteral, int ase)
    {
        var columna = new string(primeraCeldaLiteral.Where(char.IsLetter).ToArray());
        var primeraFila = FilaDe(primeraCeldaLiteral);
        return $"{columna}{primeraFila + ase - 1}";
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // Helpers de etiquetas y bloques
    // ════════════════════════════════════════════════════════════════════════════════════════

    private static void AssertEtiquetas(Hoja hoja, string contexto, (int Fila, string Etiqueta)[] esperadas, IReadOnlyList<(int Fila, string Etiqueta)> observadas)
    {
        Assert.Equal(esperadas.Length, observadas.Count);

        for (var i = 0; i < esperadas.Length; i++)
        {
            Assert.True(
                esperadas[i].Fila == observadas[i].Fila && esperadas[i].Etiqueta == observadas[i].Etiqueta,
                $"{contexto}: se esperaba ({esperadas[i].Fila}, '{esperadas[i].Etiqueta}') y se midió "
                + $"({observadas[i].Fila}, '{observadas[i].Etiqueta}') en '{hoja.Nombre}'.");
        }
    }

    private static bool EsTotal(string texto) => Normalizar(texto) == "TOTAL";

    private static bool EsGranTotal(Hoja hoja, int fila) =>
        Normalizar(Texto(hoja, "A", fila)) == "TOTAL" && Texto(hoja, "B", fila).Length == 0;

    private static int BordeBloque(int[] cabeceras, int max, int fila)
    {
        for (var i = 0; i < cabeceras.Length; i++)
        {
            var fin = i + 1 < cabeceras.Length ? cabeceras[i + 1] - 1 : max;
            if (fila >= cabeceras[i] && fila <= fin)
            {
                return i + 1;
            }
        }

        return 0;
    }

    private static int BloqueQueContiene(int[] cabeceras, int max, int fila) => BordeBloque(cabeceras, max, fila);

    /// <summary>Fila del cierre de sección que CIERRA el cuadro donde vive <paramref name="fila"/>.</summary>
    private static int SectionQueCierra(Hoja hoja, int fila)
    {
        for (var f = fila + 1; f <= FilaMaxima(hoja); f++)
        {
            if (Texto(hoja, "A", f).Length > 0 && EsTotal(Texto(hoja, "B", f)) && Texto(hoja, "C", f).Length == 0)
            {
                return f;
            }

            if (EsGranTotal(hoja, f))
            {
                break;
            }
        }

        return -1;
    }

    private static List<(int Fila, string Etiqueta)> CierresSeccion(Hoja hoja) =>
        Filas(hoja)
            .Where(f => Texto(hoja, "A", f).Length > 0 && EsTotal(Texto(hoja, "B", f)) && Texto(hoja, "C", f).Length == 0)
            .Select(f => (f, Texto(hoja, "A", f)))
            .ToList();

    private static List<(int Fila, string Etiqueta)> CierresSubseccion(Hoja hoja) =>
        Filas(hoja)
            .Where(f => Texto(hoja, "B", f).Length > 0 && !EsTotal(Texto(hoja, "B", f)) && EsTotal(Texto(hoja, "C", f)) && Texto(hoja, "A", f).Length == 0)
            .Select(f => (f, Texto(hoja, "B", f)))
            .ToList();

    // ════════════════════════════════════════════════════════════════════════════════════════
    // Lectura por bytes (BCL puro)
    // ════════════════════════════════════════════════════════════════════════════════════════

    private sealed record Celda(string? Formula, string? Valor, bool TieneFormula, string? TipoFormula, string? IndiceCompartida, string? RefCompartida);

    private sealed class Hoja
    {
        public Hoja(string nombre, string? dimension, Dictionary<string, Celda> celdas)
        {
            Nombre = nombre;
            Dimension = dimension;
            Celdas = celdas;
        }

        public string Nombre { get; }

        public string? Dimension { get; }

        public Dictionary<string, Celda> Celdas { get; }
    }

    private sealed class Libro
    {
        public Dictionary<string, Hoja> Hojas { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private static readonly Lazy<Libro> LibroV1Lazy = new(() => Cargar(PlantillaV1), isThreadSafe: true);
    private static readonly Lazy<Libro> LibroV2Lazy = new(() => Cargar(PlantillaV2), isThreadSafe: true);
    private static readonly Lazy<Libro> LibroV3Lazy = new(() => Cargar(PlantillaV3), isThreadSafe: true);

    private static Libro LibroV1 => LibroV1Lazy.Value;
    private static Libro LibroV2 => LibroV2Lazy.Value;
    private static Libro LibroV3 => LibroV3Lazy.Value;

    private static Libro Cargar(string ruta)
    {
        Assert.True(File.Exists(ruta), $"Falta el archivo: {ruta}");

        using var zip = ZipFile.Open(ruta, ZipArchiveMode.Read);
        var workbook = LeerXml(zip, "xl/workbook.xml");
        var relaciones = LeerXml(zip, "xl/_rels/workbook.xml.rels");
        var compartidas = zip.GetEntry("xl/sharedStrings.xml") is { } entradaCompartidas
            ? LeerXml(entradaCompartidas).Descendants(Ns + "si")
                .Select(si => string.Concat(si.Descendants(Ns + "t").Select(t => t.Value)))
                .ToList()
            : [];

        var destinos = relaciones.Root!.Elements()
            .Where(e => e.Name.LocalName == "Relationship")
            .ToDictionary(
                e => e.Attribute("Id")?.Value ?? string.Empty,
                e => e.Attribute("Target")?.Value ?? string.Empty,
                StringComparer.Ordinal);

        var libro = new Libro();

        foreach (var hoja in workbook.Descendants(Ns + "sheet"))
        {
            var nombre = hoja.Attribute("name")?.Value;
            var id = hoja.Attribute(RelNs + "id")?.Value;

            if (string.IsNullOrEmpty(nombre) || string.IsNullOrEmpty(id) || !destinos.TryGetValue(id, out var destino))
            {
                continue;
            }

            var entrada = zip.GetEntry(RutaDeParte(destino));
            if (entrada is null)
            {
                continue;
            }

            var documento = LeerXml(entrada);
            var dimension = documento.Descendants(Ns + "dimension").FirstOrDefault()?.Attribute("ref")?.Value;
            var celdas = new Dictionary<string, Celda>(StringComparer.OrdinalIgnoreCase);

            foreach (var celda in documento.Descendants(Ns + "c"))
            {
                var referencia = celda.Attribute("r")?.Value;
                if (string.IsNullOrEmpty(referencia))
                {
                    continue;
                }

                var formula = celda.Element(Ns + "f");
                var valor = celda.Element(Ns + "v")?.Value;

                if (string.Equals(celda.Attribute("t")?.Value, "s", StringComparison.Ordinal)
                    && int.TryParse(valor, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indice)
                    && indice >= 0 && indice < compartidas.Count)
                {
                    valor = compartidas[indice];
                }
                else if (string.Equals(celda.Attribute("t")?.Value, "inlineStr", StringComparison.Ordinal))
                {
                    valor = string.Concat(celda.Descendants(Ns + "t").Select(t => t.Value));
                }

                celdas[referencia] = new Celda(
                    formula is null ? null : formula.Value,
                    valor,
                    formula is not null,
                    formula?.Attribute("t")?.Value,
                    formula?.Attribute("si")?.Value,
                    formula?.Attribute("ref")?.Value);
            }

            libro.Hojas[nombre] = new Hoja(nombre, dimension, celdas);
        }

        return libro;
    }

    private static XDocument LeerXml(ZipArchive zip, string entrada) =>
        LeerXml(zip.GetEntry(entrada) ?? throw new InvalidOperationException($"El paquete no trae '{entrada}'."));

    private static XDocument LeerXml(ZipArchiveEntry entrada)
    {
        using var stream = entrada.Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// Ruta de la parte dentro del paquete: los generadores escriben <c>Target</c> relativo
    /// (<c>worksheets/sheet1.xml</c>) o absoluto (<c>/xl/worksheets/sheet.xml</c>).
    /// </summary>
    private static string RutaDeParte(string destino)
    {
        var normalizado = destino.Replace('\\', '/').TrimStart('/');
        return normalizado.StartsWith("xl/", StringComparison.OrdinalIgnoreCase) ? normalizado : $"xl/{normalizado}";
    }

    // ── helpers de celda ──────────────────────────────────────────────────────────────────────

    private static string Texto(Hoja hoja, string columna, int fila) =>
        hoja.Celdas.TryGetValue(columna + fila.ToString(CultureInfo.InvariantCulture), out var celda)
            ? (celda.Valor ?? string.Empty)
            : string.Empty;

    private static decimal Numero(string texto) =>
        decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor) ? valor : 0m;

    private static IEnumerable<int> Filas(Hoja hoja) =>
        hoja.Celdas.Keys.Select(FilaDe).Distinct().OrderBy(f => f);

    private static int FilaMaxima(Hoja hoja)
    {
        // La dimensión declara el último row del sheet ("A1:AP843"); se re-mide además por celdas.
        var porDimension = int.Parse(new string((hoja.Dimension ?? "A1").Split(':').Last().Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);

        var porCeldas = hoja.Celdas.Keys.Select(FilaDe).DefaultIfEmpty(0).Max();
        Assert.True(porCeldas <= porDimension, $"'{hoja.Nombre}': hay celdas fuera de la dimensión declarada ({porCeldas} > {porDimension}).");
        return porDimension;
    }

    private static int FilaDe(string referencia) =>
        int.Parse(new string(referencia.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);

    private static string Normalizar(string texto)
    {
        var sinDiacriticos = texto.Trim().ToUpperInvariant().Normalize(NormalizationForm.FormD);
        var limpio = new string(sinDiacriticos.Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark).ToArray());
        return string.Join(' ', limpio.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
    }

    // ── rutas explícitas (prohibido Insumos.Raiz() en este T0) ────────────────────────────────

    private static string RaizRepo()
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

    private static string Ruta(params string[] partes) => Path.Combine([RaizRepo(), .. partes]);

    private static string PlantillaV3 => Ruta("Docs", "Plantilla estandarizada", "v3", "Formato Remuneracion 2026xx-x Total v3.xlsx");

    private static string PlantillaV2 => Ruta("Docs", "Plantilla estandarizada", "v2", "Formato Remuneracion 2026xx-x Total vf.xlsx");

    private static string PlantillaV1 => Ruta("Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string FuenteR1Octubre1(int aseId) =>
        Ruta("Docs", "Prueba Octubre-1", "Insumos", CarpetasAse[aseId - 1], NombreFuenteR1Octubre1(aseId));

    private static string FuenteR2Octubre1(int aseId) =>
        Ruta("Docs", "Prueba Octubre-1", "Insumos", CarpetasAse[aseId - 1], NombreFuenteR2Octubre1(aseId));

    private static string FuenteR1Julio1(int aseId) =>
        Ruta("Docs", "Prueba Julio-1", "Insumos", CarpetasAse[aseId - 1], NombreFuenteR1Julio1(aseId));

    // Nombres EXACTOS en disco (octubre-1 y julio-1): sin glob, sin prefijo.
    private static string NombreFuenteR1Octubre1(int aseId) => aseId switch
    {
        1 => "Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174133882.xlsx",
        2 => "Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174350151.xlsx",
        3 => "Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174753997.xlsx",
        4 => "Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174948502.xlsx",
        5 => "Recaudoporcomponente_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016175148433.xlsx",
        _ => throw new ArgumentOutOfRangeException(nameof(aseId), aseId, "ASE fuera de 1..5.")
    };

    private static string NombreFuenteR2Octubre1(int aseId) => aseId switch
    {
        1 => "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174159251.xlsx",
        2 => "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___2026101617441360.xlsx",
        3 => "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___20261016174813931.xlsx",
        4 => "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___202610161750610.xlsx",
        5 => "RerpoteDetalleSaldosaFavor_to_date01102026ddMMyyyy_to_date15102026ddMMyyyy___202610161752844.xlsx",
        _ => throw new ArgumentOutOfRangeException(nameof(aseId), aseId, "ASE fuera de 1..5.")
    };

    private static string NombreFuenteR1Julio1(int aseId) => aseId switch
    {
        2 => "Recaudoporcomponente_to_date01072026ddMMyyyy_to_date15072026ddMMyyyy___202671616523961.xlsx",
        _ => throw new ArgumentOutOfRangeException(nameof(aseId), aseId, "El T0 solo mide julio-1 para ASE2 (caso de subsecciones ausentes).")
    };

    // ── G-SHA ────────────────────────────────────────────────────────────────────────────────

    private static void AssertEqualToPlan(string ruta, string huellaEsperada)
    {
        Assert.True(File.Exists(ruta), $"Falta la plantilla: {ruta}");

        // Plan 59 §0.6 (G-HUELLA): el gate es la huella de CONTENIDO (hojas, dimension, <f>, <v> y
        // fórmulas compartidas normalizadas y ordenadas). El SHA y la longitud del ARCHIVO cambian cada
        // vez que un humano guarda la plantilla en Excel y por eso NO son gate: quedan como dato.
        Assert.Equal(huellaEsperada, HuellaSemanticaPlantilla.Calcular(ruta));
    }

    private static void AssertDimension(string hoja, string dimensionV2, string dimensionV3)
    {
        Assert.Equal(dimensionV2, LibroV2.Hojas[hoja].Dimension);
        Assert.Equal(dimensionV3, LibroV3.Hojas[hoja].Dimension);
    }

    private static void AssertCabeceras(string hoja, (int Ase, int V2, int V3)[] esperadas)
    {
        foreach (var (nombre, libro, columnaFila) in new[] { ("v2", LibroV2, 1), ("v3", LibroV3, 2) })
        {
            var h = libro.Hojas[hoja];
            var cabeceras = Filas(h)
                .Where(f => NombresAse.Contains(Texto(h, "B", f).Trim(), StringComparer.OrdinalIgnoreCase))
                .ToArray();

            var esperado = esperadas.Select(c => columnaFila == 1 ? c.V2 : c.V3).ToArray();
            Assert.Equal(esperado, cabeceras);

            for (var i = 0; i < esperado.Length; i++)
            {
                Assert.Equal(NombresAse[i], Texto(h, "B", esperado[i]));
            }
        }
    }

    private static string Sha256(string ruta)
    {
        using var stream = File.OpenRead(ruta);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}
