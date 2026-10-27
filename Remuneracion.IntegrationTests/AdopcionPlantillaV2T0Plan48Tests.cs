using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using Remuneracion.Core.Models;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 48 — T0 read-only de adopción de la plantilla v2 y paridad con el manual julio-2.
///
/// <para>
/// Congela por bytes (BCL <see cref="ZipArchive"/> + <see cref="XDocument"/>; sin Excel/COM ni
/// OpenXML de producción) la matriz v1→v2 (P1), la clasificación P2 (geometría / cachés /
/// contenido real), el inventario P3 del delta del manual (492 fórmulas de R1), el veredicto P4
/// (qué cadena congelada contra v1 falla en v2), las preguntas P5 y la tabla P6 de las 40 hojas.
/// Evidencia: <c>Docs/T0-Plan48-Adopcion-Plantilla-v2.md</c>.
/// </para>
///
/// <para>
/// H0 NO CORRIGE NADA: no se crea/modifica/borra ninguna <c>&lt;f&gt;</c>, no se inserta/borra
/// ninguna fila/columna, no se escribe sobre la v1 ni la v2 (ambas se verifican por SHA256 antes y
/// después). El manual v2 (nuevo oráculo) solo se lee.
/// </para>
/// </summary>
public sealed class AdopcionPlantillaV2T0Plan48Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string ShaV1 = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";
    private const string ShaV2 = "FF5CF7A819110A1E1E2CA7BDAD553DD7A25141ADD472B1904BD496F082642632";

    private const string HojaR1 = "Reporte Componentes R1";
    private const string HojaR2 = "Rem. Anticipos R2";
    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";
    private const string HojaValidaControl = "Valida - Control Recaudo";

    private static readonly string[] Hojas40 =
    [
        "Recaudo EAAB Reciprocidad", "Recaudo ENEL", "Recaudo ENERBIT", "Recaudo EAAB + Ciud Limp",
        "Recaudo Directa Occidente", "REPORTE RECAUDO x BANCO", "ANT EXT-REV", "ANTICIPOS USUARIOS",
        "BCE SC POR FACT.", "Reporte Componentes R1", "Rem. Anticipos R2", "Reversion Pagos R4",
        "SALDOS POR NOTA", "RETRIBUCION NEGATIVA", "AJUSTES - SF-T", "INTERVENTORIA",
        "Informe AFaseo Recaudo 202609-2", "CONSOLIDADO_TOTAL RECAUDO", "VALIDACION_TOTAL",
        "GERENTES_TOTAL", "REMUNERACION Reciprocidad EAAB", "VALIDACION_RECIP",
        "GERENTES_Reciprocidad_ EAAB", "REMUNERACION_ENEL", "VALIDACION_ENEL", "GERENTES_ENEL",
        "REMUNERACION_ENERBIT", "VALIDACION_ENERBIT", "GERENTES_ENERBIT",
        "REMUNERACION_OCCIDENTE_ Directa", "VALIDACION_OCCIDENTE", "GERENTES_OCCIDENTE _ Directa",
        "REMUNERACION_EAAB-CL", "VALIDACION_EAAB-CL", "GERENTES_EAAB-CL", "DetRetri2026092",
        "DetValiRetri2026092", "Valida -Remunera", "Valida - Anticipos", "Valida - Control Recaudo",
    ];

    /// <summary>P6: clase de delta probada v1→v2 por hoja (nombres = <see cref="Hojas40"/>).</summary>
    private static readonly DeltaHoja[] MatrizP6 =
    [
        new("Recaudo EAAB Reciprocidad", 0, 0), new("Recaudo ENEL", 0, 0),
        new("Recaudo ENERBIT", 0, 0), new("Recaudo EAAB + Ciud Limp", 0, 0),
        new("Recaudo Directa Occidente", 0, 0), new("REPORTE RECAUDO x BANCO", 0, 0),
        new("ANT EXT-REV", 0, 0), new("ANTICIPOS USUARIOS", 0, 0), new("BCE SC POR FACT.", 0, 0),
        new("Reporte Componentes R1", 0, 16), new("Rem. Anticipos R2", 352, 0),
        new("Reversion Pagos R4", 0, 0), new("SALDOS POR NOTA", 0, 0),
        new("RETRIBUCION NEGATIVA", 0, 0), new("AJUSTES - SF-T", 0, 0), new("INTERVENTORIA", 0, 0),
        new("Informe AFaseo Recaudo 202609-2", 0, 0), new(HojaConsolidado, 61, 0),
        new("VALIDACION_TOTAL", 0, 2), new("GERENTES_TOTAL", 0, 0),
        new("REMUNERACION Reciprocidad EAAB", 60, 0), new("VALIDACION_RECIP", 0, 2),
        new("GERENTES_Reciprocidad_ EAAB", 0, 0), new("REMUNERACION_ENEL", 60, 0),
        new("VALIDACION_ENEL", 0, 2), new("GERENTES_ENEL", 0, 0),
        new("REMUNERACION_ENERBIT", 60, 0), new("VALIDACION_ENERBIT", 0, 2),
        new("GERENTES_ENERBIT", 0, 0), new("REMUNERACION_OCCIDENTE_ Directa", 60, 0),
        new("VALIDACION_OCCIDENTE", 0, 2), new("GERENTES_OCCIDENTE _ Directa", 0, 0),
        new("REMUNERACION_EAAB-CL", 56, 0), new("VALIDACION_EAAB-CL", 0, 0),
        new("GERENTES_EAAB-CL", 0, 0), new("DetRetri2026092", 40, 0),
        new("DetValiRetri2026092", 51, 0), new("Valida -Remunera", 0, 0),
        new("Valida - Anticipos", 0, 0), new(HojaValidaControl, 3, 7),
    ];

    private sealed record DeltaHoja(string Hoja, int FormulaDiff, int NumDiff);

    // ── G-SHA: ambas plantillas intocables ──────────────────────────────────────────────────────

    [Fact]
    public void T0P48_Sha256_V1_V2_Intactas()
    {
        Assert.Equal(ShaV1, Sha256(V1()));
        // Plan 59 §0.6 (G-HUELLA): la v2 fue re-guardada en Excel; el gate es la huella de CONTENIDO
        // y el SHA del archivo queda informativo.
        Assert.Equal(HuellaSemanticaPlantilla.V2, HuellaSemanticaPlantilla.Calcular(V2()));
    }

    [Fact]
    public void T0P48_ManualV2_ExisteYEsElOraculo()
    {
        Assert.True(File.Exists(Manual()), $"Falta el manual v2 (oráculo): {Manual()}");
    }

    // ── P1: 40 hojas, mismo orden y nombres ─────────────────────────────────────────────────────

    [Fact]
    public void T0P48_P1_40Hojas_MismoOrdenYNombre()
    {
        var v1 = NombresDeHoja(V1());
        var v2 = NombresDeHoja(V2());

        Assert.Equal(40, v1.Length);
        Assert.Equal(40, v2.Length);
        Assert.Equal(v1, v2);
        Assert.Equal(Hojas40, v1);
    }

    // ── P1: R1 — fila 649 (16 residuales limpiados en v2; el mirror la reescribe) ─────────────

    [Fact]
    public void T0P48_P1_R1_Fila649_V1Con16Valores_V2EnCero()
    {
        var v1 = Leer(V1(), HojaR1);
        var v2 = Leer(V2(), HojaR1);

        string[] refs = ["F649", "G649", "H649", "I649", "J649", "K649", "L649", "M649",
                         "N649", "O649", "P649", "Q649", "R649", "S649", "T649", "Z649"];

        foreach (var r in refs)
        {
            Assert.True(v1.NumeroOrNull(r) is not null, $"V1 debe traer valor en {r}.");
            Assert.NotEqual(0m, v1.Numero(r));
            Assert.Equal(0m, v2.Numero(r));
            Assert.False(v2.TieneFormula(r));
        }

        AssertCerca(13359789539.23m, v1.Numero("F649"));
        AssertCerca(2072913975.94m, v1.Numero("P649"));
        AssertCerca(173121721.61m, v1.Numero("Q649"));
    }

    // ── P1: R2 — bloques por ASE, rangos nuevos y crecimiento +11 en ASE2/ASE4 ────────────────

    [Fact]
    public void T0P48_P1_R2_BloquesPorAse_V1VsV2()
    {
        var v1 = Leer(V1(), HojaR2);
        var v2 = Leer(V2(), HojaR2);

        string[] ase = ["PROMOAMBIENTAL", "LIME", "CIUDAD LIMPIA", "BOGOTA LIMPIA", "AREA LIMPIA"];

        (int Inicio, int Fin)[] esperadoV1 = [(3, 33), (67, 113), (164, 194), (282, 328), (382, 419)];
        (int Inicio, int Fin)[] esperadoV2 = [(3, 33), (67, 124), (175, 205), (293, 350), (404, 441)];

        for (var i = 0; i < ase.Length; i++)
        {
            Assert.Equal(esperadoV1[i], BloqueR2(v1, ase[i]));
            Assert.Equal(esperadoV2[i], BloqueR2(v2, ase[i]));
        }

        // Crecimiento: ASE2 +11 y ASE4 +11; ASE1/3/5 idénticos en tamaño.
        Assert.Equal(0, Tamano(esperadoV2[0]) - Tamano(esperadoV1[0]));
        Assert.Equal(11, Tamano(esperadoV2[1]) - Tamano(esperadoV1[1]));
        Assert.Equal(0, Tamano(esperadoV2[2]) - Tamano(esperadoV1[2]));
        Assert.Equal(11, Tamano(esperadoV2[3]) - Tamano(esperadoV1[3]));
        Assert.Equal(0, Tamano(esperadoV2[4]) - Tamano(esperadoV1[4]));

        // Desplazamiento acumulado de los cuadros OPORTUNO: +11 (ASE2/3) y +22 (ASE4/5).
        Assert.Equal(47, FilaOportunoTotal(v1, 0));
        Assert.Equal(155, FilaOportunoTotal(v2, 1));
        Assert.Equal(273, FilaOportunoTotal(v2, 2));
        Assert.Equal(384, FilaOportunoTotal(v2, 3));
        Assert.Equal(468, FilaOportunoTotal(v2, 4));
        Assert.Equal(47, FilaOportunoTotal(v1, 0));
        Assert.Equal(144, FilaOportunoTotal(v1, 1));
        Assert.Equal(262, FilaOportunoTotal(v1, 2));
        Assert.Equal(362, FilaOportunoTotal(v1, 3));
        Assert.Equal(446, FilaOportunoTotal(v1, 4));
    }

    [Fact]
    public void T0P48_P1_R2_Ase2Ase4_GananCanal3YCanal9_QueV1NoTiene()
    {
        var v1 = Leer(V1(), HojaR2);
        var v2 = Leer(V2(), HojaR2);

        foreach (var (inicio, fin, canal3, canal9) in new[]
                 {
                     (67, 113, 83, 89), (282, 328, 309, 315),
                 })
        {
            Assert.DoesNotContain(Enumerable.Range(inicio, fin - inicio + 1),
                n => v1.Texto("C", n) == "9");
            Assert.DoesNotContain(Enumerable.Range(inicio, fin - inicio + 1),
                n => v1.Texto("C", n) == "3");

            Assert.Equal("3", v2.Texto("C", canal3));
            Assert.Equal("Total", v2.Texto("D", canal3));
            Assert.Equal("9", v2.Texto("C", canal9));
            Assert.Equal("Total", v2.Texto("D", canal9));
            Assert.Equal("NUEVO ESQUEMA", v2.Texto("B", canal9 + 1));
        }
    }

    // ── P1 / R-OPORTUNO: cuadro OPORTUNO — fórmula vs literal por fila y por ASE ──────────────

    [Fact]
    public void T0P48_P1_R2_CuadroOportuno_FormulaVsLiteral_PorAse()
    {
        var v2 = Leer(V2(), HojaR2);

        // (fila-empresa, empresa, ¿E es fórmula?, fila-subsidio)
        (int Fila, string Empresa, bool Formula, int Subsidio)[] ase1 =
        [
            (47, "TOTAL", true, 48), (50, "RECIPROCIDAD", false, 51), (53, "ENEL", true, 54),
            (56, "ENERBIT", false, 57), (59, "OCCIDENTE", true, 60), (62, "CIUDAD LIMPIA-ACUEDUCTO", false, 63),
        ];
        (int Fila, string Empresa, bool Formula, int Subsidio)[] ase2 =
        [
            (155, "TOTAL", true, 156), (158, "RECIPROCIDAD", true, 159), (161, "ENEL", true, 162),
            (164, "ENERBIT", false, 165), (167, "OCCIDENTE", true, 168), (170, "CIUDAD LIMPIA-ACUEDUCTO", false, 171),
        ];
        (int Fila, string Empresa, bool Formula, int Subsidio)[] ase3 =
        [
            (273, "TOTAL", true, 274), (276, "RECIPROCIDAD", false, 277), (279, "ENEL", true, 280),
            (282, "ENERBIT", false, 283), (285, "OCCIDENTE", true, 286), (288, "CIUDAD LIMPIA-ACUEDUCTO", false, 289),
        ];
        (int Fila, string Empresa, bool Formula, int Subsidio)[] ase4 =
        [
            (384, "TOTAL", true, 385), (387, "RECIPROCIDAD", true, 388), (390, "ENEL", true, 391),
            (393, "ENERBIT", false, 394), (396, "OCCIDENTE", true, 397), (399, "CIUDAD LIMPIA-ACUEDUCTO", false, 400),
        ];
        (int Fila, string Empresa, bool Formula, int Subsidio)[] ase5 =
        [
            (468, "TOTAL", true, 469), (471, "RECIPROCIDAD", false, 472), (474, "ENEL", true, 475),
            (477, "ENERBIT", true, 478), (480, "OCCIDENTE", true, 481), (483, "CIUDAD LIMPIA-ACUEDUCTO", false, 484),
        ];

        foreach (var tabla in new[] { ase1, ase2, ase3, ase4, ase5 })
        {
            foreach (var (fila, empresa, formula, subsidio) in tabla)
            {
                Assert.Equal("OPORTUNO", v2.Texto("C", fila));
                Assert.Equal("Total", v2.Texto("D", fila));
                if (empresa != "TOTAL")
                {
                    Assert.Equal(empresa, v2.Texto("B", fila));
                }

                Assert.Equal(formula, v2.TieneFormula("E" + fila.ToString(CultureInfo.InvariantCulture)));
                Assert.Equal("Subsidio(-)/Contribucion(+)", v2.Texto("D", subsidio));
                Assert.Equal(formula, v2.TieneFormula("E" + subsidio.ToString(CultureInfo.InvariantCulture)));
            }
        }

        // Fórmulas testigo del TOTAL (una por ASE, la geometría del cuadro en v2).
        Assert.Equal("E21+E32-K21", v2.Formula("E47"));
        Assert.Equal("E101+E123-K101", v2.Formula("E155"));
        Assert.Equal("E193+E204-K193", v2.Formula("E273"));
        Assert.Equal("E327+E349-K327", v2.Formula("E384"));
        Assert.Equal("E440+E426-K426", v2.Formula("E468"));
    }

    // ── P1: Valida - Control Recaudo (E7/E10/F10 pasan de valor a fórmula) ────────────────────

    [Fact]
    public void T0P48_P1_ValidaControlRecaudo_V2_TraeFormulasE7E10F10()
    {
        var v1 = Leer(V1(), HojaValidaControl);
        var v2 = Leer(V2(), HojaValidaControl);

        Assert.False(v1.TieneFormula("E7"));
        Assert.False(v1.TieneFormula("E10"));
        Assert.False(v1.TieneFormula("F10"));
        Assert.Equal(68245142592m, v1.Numero("E7"));

        Assert.Equal("SUM(D7:D9)", v2.Formula("E7"));
        Assert.Equal("SUM(D10)", v2.Formula("E10"));
        Assert.Equal("E7+E10", v2.Formula("F10"));
        Assert.Equal(0m, v2.Numero("E7"));
        Assert.Equal(0m, v2.Numero("E10"));
        Assert.Equal(0m, v2.Numero("F10"));
    }

    // ── P1: VALIDACION_* — C15 prefijado en 1 y C14 sin caché ─────────────────────────────────

    [Theory]
    [InlineData("VALIDACION_ENEL")]
    [InlineData("VALIDACION_ENERBIT")]
    [InlineData("VALIDACION_OCCIDENTE")]
    [InlineData("VALIDACION_RECIP")]
    [InlineData("VALIDACION_TOTAL")]
    public void T0P48_P1_ValidacionX_C15_Uno_C14_SinCache(string hoja)
    {
        var v1 = Leer(V1(), hoja);
        var v2 = Leer(V2(), hoja);

        Assert.Equal(0m, v1.Numero("C15"));
        Assert.NotEqual(0m, v1.Numero("C14"));

        Assert.Equal(1m, v2.Numero("C15"));
        Assert.Equal(0m, v2.Numero("C14"));
    }

    // ── P1: consumidores ya apuntan a las filas nuevas de R2 en v2 ────────────────────────────

    [Fact]
    public void T0P48_P1_Consumidores_ApuntanFilasNuevasR2()
    {
        var v1 = Leer(V1(), HojaConsolidado);
        var v2 = Leer(V2(), HojaConsolidado);

        string[] v1Refs = ["E47", "E144", "E262", "E362", "E446"];
        string[] v2Refs = ["E47", "E155", "E273", "E384", "E468"];

        for (var i = 0; i < 5; i++)
        {
            Assert.Equal($"'Rem. Anticipos R2'!{v1Refs[i]}", v1.Formula("D" + (28 + i).ToString(CultureInfo.InvariantCulture)));
            Assert.Equal($"'Rem. Anticipos R2'!{v2Refs[i]}", v2.Formula("D" + (28 + i).ToString(CultureInfo.InvariantCulture)));
        }

        // El arrastre del período también quedó realineado (muestras testigo del T0).
        Assert.Equal("'Rem. Anticipos R2'!I144", v1.Formula("H29"));
        Assert.Equal("'Rem. Anticipos R2'!I155", v2.Formula("H29"));
        Assert.Equal("'Rem. Anticipos R2'!H362", v1.Formula("G31"));
        Assert.Equal("'Rem. Anticipos R2'!H384", v2.Formula("G31"));
        Assert.Equal("'Rem. Anticipos R2'!E447", v1.Formula("AE32"));
        Assert.Equal("'Rem. Anticipos R2'!E469", v2.Formula("AE32"));
    }

    // ── P4: la cadena de fórmulas congelada contra v1 ya no existe en v2 (fail-fast) ──────────

    [Fact]
    public void T0P48_P4_CadenaFormulaV1_E144_YaNoEsFormulaEnV2()
    {
        var v1 = Leer(V1(), HojaR2);
        var v2 = Leer(V2(), HojaR2);

        Assert.True(v1.TieneFormula("E144"));
        Assert.Equal("E94+E112-K94", v1.Formula("E144"));

        // En v2 la misma celda lógica se desplazó al cuadro OPORTUNO de ASE2 (R155) y E144
        // quedó vacía: el guard anti-fórmula del writer la reporta como ERR-PLANTILLA.
        Assert.False(v2.TieneFormula("E144"));
        Assert.True(v2.TieneFormula("E155"));
        Assert.Equal("E101+E123-K101", v2.Formula("E155"));
    }

    // ── P2: la diferencia de contenido R2 es EXACTAMENTE el falso hueco documentado ───────────

    [Fact]
    public void T0P48_P2_R2_DeltaContenido_EsElFalsoHuecoDeCerosDocumentados()
    {
        // El T0 midió (suma de columnas por bloque, geometrías alineadas) que las 9 filas del falso
        // hueco suman 53.283.379,41: ASE2 ΔE = 1.664.245,00 (406,42 + 831.919,29 + 831.919,29) y
        // ASE4 ΔE = 51.619.134,41 (-8.785,59 + 20.152.773,11 - 2.377,86 + 20.150.395,25
        // + 5.663.564,75 + 5.663.564,75). El Plan 47 las declaraba "sin slot"; el H1 del Plan 48
        // probó que la v2 SÍ las tiene y las RETIRÓ de `CerosDocumentados` (solo quedan las 3 claves
        // A/EAB de agosto-2, que siguen ausentes en la mitad Subs/Cont de ASE4).
        Assert.Equal(3, CerosDocumentadosDetalleR2R4.TotalFilasDocumentadas);
        Assert.Equal(3, CerosDocumentadosDetalleR2R4.Entradas.Count(e => e.Hoja == HojaDetalleR2R4.R2));
        Assert.All(CerosDocumentadosDetalleR2R4.Entradas, e => Assert.True(e.Canal is null || e.Canal.Contains("A/EAB")));

        foreach (var (ase, firma, canal) in new[]
                 {
                     (2, "|||Vlr Servicio", "9/NUEVO ESQUEMA"),
                     (2, "|||Vlr Intereses", "2/NUEVO ESQUEMA"),
                     (2, "||9|Total", "9/NUEVO ESQUEMA"),
                     (4, "|||Vlr Servicio", "9/NUEVO ESQUEMA"),
                     (4, "|||Vlr Intereses", "9/NUEVO ESQUEMA"),
                     (4, "||9|Total", "9/NUEVO ESQUEMA"),
                     (4, "|||Subsidio(-)/Contribucion(+)", "9/NUEVO ESQUEMA"),
                 })
        {
            Assert.Equal(0, CerosDocumentadosDetalleR2R4.CantidadDocumentada(HojaDetalleR2R4.R2, ase, firma, canal));
        }

        // El manual (y la v2) llenan esos slots: el falso hueco era real y quedó cerrado en H1.
        var manual = Leer(Manual(), HojaR2);
        AssertCerca(831919.29m, manual.Numero("E87"));    // ASE2 Vlr Servicio canal 9
        AssertCerca(831919.29m, manual.Numero("E89"));    // ASE2 9|Total
        AssertCerca(20152773.11m, manual.Numero("E313")); // ASE4 Vlr Servicio canal 9
        AssertCerca(-2377.86m, manual.Numero("E314"));    // ASE4 Vlr Intereses canal 9
        AssertCerca(20150395.25m, manual.Numero("E315")); // ASE4 9|Total (Anticipos)

        var deltaE = 406.42m + 831919.29m + 831919.29m
                     + -8785.59m + 20152773.11m + -2377.86m
                     + 20150395.25m + 5663564.75m + 5663564.75m;
        AssertCerca(53283379.41m, deltaE);
    }

    // ── P3: inventario del delta del manual en R1 (492 fórmulas) ───────────────────────────────

    [Fact]
    public void T0P48_P3_R1_Manual_492_Formulas_DosPatrones()
    {
        var v2 = Leer(V2(), HojaR1);
        var manual = Leer(Manual(), HojaR1);

        var cambios = new List<(string Ref, string Antes, string Despues)>();
        foreach (var (referencia, formula) in v2.Formulas)
        {
            if (!manual.TieneFormula(referencia))
            {
                continue;
            }

            var otra = manual.Formula(referencia);
            if (!string.Equals(formula, otra, StringComparison.Ordinal))
            {
                cambios.Add((referencia, formula, otra));
            }
        }

        // Formulas NUEVAS en el manual (v2 no tenía fórmula).
        var nuevas = manual.Formulas.Where(f => !v2.TieneFormula(f.Celda)).ToArray();

        Assert.Equal(492, cambios.Count + nuevas.Length);
        Assert.Equal(77, nuevas.Length);

        var texto = cambios.Count;
        Assert.Equal(415, texto);

        var (gH, fG, otros) = ClasificarCambios(cambios);
        Assert.Equal(337, gH);
        Assert.Equal(41, fG);
        Assert.Equal(37, otros);

        // Las 77 nuevas se concentran en las filas 682/692/702 (+ 683/693/703), zona del residual.
        var nuevasRefs = nuevas.Select(n => n.Celda).ToHashSet(StringComparer.Ordinal);
        Assert.Contains("F682", nuevasRefs);
        Assert.Contains("F692", nuevasRefs);
        Assert.Contains("F702", nuevasRefs);
        Assert.Contains("W682", nuevasRefs);
        Assert.Contains("S682", nuevasRefs);
        Assert.Contains("AN682", nuevasRefs);
        Assert.Contains("F683", nuevasRefs);
        Assert.Contains("F693", nuevasRefs);
        Assert.Contains("F703", nuevasRefs);

        Assert.Equal("F660+F629-L629", manual.Formula("F682"));
        Assert.Equal("F652+F618-L618", manual.Formula("F692"));
        Assert.Equal("F659+F628-L628", manual.Formula("F702"));
        Assert.Equal("G629", manual.Formula("G682"));
        Assert.Equal("G618", manual.Formula("G692"));
        Assert.Equal("G628", manual.Formula("G702"));
        Assert.Equal("F660", manual.Formula("F683"));
        Assert.Equal("F652", manual.Formula("F693"));
        Assert.Equal("F659", manual.Formula("F703"));
        Assert.Equal("H629", manual.Formula("W682"));
        Assert.Equal("H629", manual.Formula("S682"));
        Assert.Equal("H629", manual.Formula("AN682"));

        #region DEUDA-DECLARADA Plan 49 (T0-EXTEMP-Ase5-Plan49 §4.3/§8.1)
        // La celda F682/F702 es ancla consumidora (CONSOLIDADO!D51). El assert congela el
        // defecto de la plantilla v2 (literal 0), NO un comportamiento esperado.
        // Condición de retiro: v3 de la propietaria con la fórmula restaurada + T0 nuevo.
        // En la v2 esas celdas están vacías (sin fórmula): el delta es del manual, no de la base.
        foreach (var r in new[] { "F682", "F692", "F702", "W682", "S682", "AN682", "F683", "F693", "F703" })
        {
            Assert.False(v2.TieneFormula(r), $"V2 no debe traer fórmula en {r}.");
        }
        #endregion
    }

    // ── S1.6: ancla del oráculo reproducida (manual v2) ───────────────────────────────────────

    [Fact]
    public void T0P48_S1_6_AnclaOraculo_ManualV2()
    {
        var r1 = Leer(Manual(), HojaR1);
        var r2 = Leer(Manual(), HojaR2);

        // R1 ASE1 R6-R20 (canal E/H/O/T y el Mes/AFaseo).
        AssertCerca(-7108.82m, r1.Numero("F6"));
        AssertCerca(-213.69m, r1.Numero("H6"));
        AssertCerca(-2331.79m, r1.Numero("O6"));
        AssertCerca(0m, r1.Numero("T6"));
        AssertCerca(-11217.37m, r1.Numero("F16"));
        AssertCerca(-11217.37m, r1.Numero("F17"));

        // R1 ASE5: 603..614 = 0; 619/621 = 52.706,16; 622-624 = 49.921,06 + 402,40;
        // 625-627 = 306.792,77 + 7.697,23; 628/629 = 417.519,62.
        foreach (var fila in new[] { 603, 606, 608, 610, 613, 614 })
        {
            AssertCerca(0m, r1.Numero("F" + fila.ToString(CultureInfo.InvariantCulture)));
        }

        AssertCerca(52706.16m, r1.Numero("F619"));
        AssertCerca(52706.16m, r1.Numero("F621"));
        AssertCerca(49921.06m, r1.Numero("F622"));
        AssertCerca(402.40m, r1.Numero("F623"));
        AssertCerca(50323.46m, r1.Numero("F624"));
        AssertCerca(306792.77m, r1.Numero("F625"));
        AssertCerca(7697.23m, r1.Numero("F626"));
        AssertCerca(314490m, r1.Numero("F627"));
        AssertCerca(417519.62m, r1.Numero("F628"));
        AssertCerca(417519.62m, r1.Numero("F629"));

        // R2 ASE1: R3/R4 = 0; R7/R8 = 104.754.634,94 / 15.816,77.
        AssertCerca(0m, r2.Numero("E3"));
        AssertCerca(0m, r2.Numero("E4"));
        AssertCerca(104754634.94m, r2.Numero("E7"));
        AssertCerca(15816.77m, r2.Numero("E8"));

        // R2 LIME: R67/R68 = 0; R71/R72 = 71.702.408,71 / 95.038,85.
        AssertCerca(0m, r2.Numero("E67"));
        AssertCerca(0m, r2.Numero("E68"));
        AssertCerca(71702408.71m, r2.Numero("E71"));
        AssertCerca(95038.85m, r2.Numero("E72"));
    }

    [Fact]
    public void T0P48_S1_6_CorridaP47H2c_IgualAlManual_EnElAncla()
    {
        var corrida = Leer(CorridaP47H2c(), HojaR2);
        var manual = Leer(Manual(), HojaR2);

        foreach (var r in new[] { "E3", "E4", "E7", "E8", "E67", "E68", "E71", "E72" })
        {
            AssertCerca(manual.Numero(r), corrida.Numero(r));
        }

        var r1Corrida = Leer(CorridaP47H2c(), HojaR1);
        var r1Manual = Leer(Manual(), HojaR1);
        foreach (var r in new[] { "F6", "H6", "O6", "F16", "F17", "F619", "F622", "F623",
                                  "F624", "F625", "F626", "F627", "F628", "F629" })
        {
            AssertCerca(r1Manual.Numero(r), r1Corrida.Numero(r));
        }
    }

    // ── P6: tabla de paridad por hoja (las 40) v1→v2 ──────────────────────────────────────────

    [Fact]
    public void T0P48_P6_TablaParidad_40Hojas_ConteosReales()
    {
        var v1 = LeerLibro(V1());
        var v2 = LeerLibro(V2());

        var fallos = new List<string>();
        foreach (var esperado in MatrizP6)
        {
            Assert.True(v1.ContainsKey(esperado.Hoja), $"V1 no trae la hoja '{esperado.Hoja}'.");
            Assert.True(v2.ContainsKey(esperado.Hoja), $"V2 no trae la hoja '{esperado.Hoja}'.");

            var (fDiff, numDiff, textDiff, mixto, onlyA, onlyB) = Comparar(v1[esperado.Hoja], v2[esperado.Hoja]);

            if (fDiff != esperado.FormulaDiff)
            {
                fallos.Add($"{esperado.Hoja}: fDiff esperado {esperado.FormulaDiff}, observado {fDiff}");
            }

            if (numDiff != esperado.NumDiff)
            {
                fallos.Add($"{esperado.Hoja}: numDiff esperado {esperado.NumDiff}, observado {numDiff}");
            }

            if (esperado.FormulaDiff == 0 && esperado.NumDiff == 0
                && (textDiff != 0 || mixto != 0 || onlyA != 0 || onlyB != 0))
            {
                fallos.Add($"{esperado.Hoja}: sin delta esperado pero textDiff={textDiff} mixto={mixto} onlyA={onlyA} onlyB={onlyB}");
            }
        }

        Assert.True(fallos.Count == 0, "P6 desviaciones:" + Environment.NewLine + string.Join(Environment.NewLine, fallos));
    }

    [Fact]
    public void T0P48_P6_R2_ConteosCmpnum_Documentados()
    {
        var (fDiff, numDiff, textDiff, mixto, onlyA, onlyB) = Comparar(Leer(V1(), HojaR2), Leer(V2(), HojaR2));

        Assert.Equal(352, fDiff);
        Assert.Equal(0, numDiff);
        Assert.Equal(94, textDiff);
        Assert.Equal(45, mixto);
        Assert.Equal(985, onlyA);
        Assert.Equal(1323, onlyB);
    }

    // ── Comparador (espejo de cmpnum + diffv: shared expandida, tolerancia ±0.5) ──────────────

    private static (int FDiff, int NumDiff, int TextDiff, int Mixto, int OnlyA, int OnlyB) Comparar(Hoja a, Hoja b)
    {
        // fDiff del T0 (espejo de diffv): para toda celda presente en AMBOS libros se compara el
        // texto de fórmula efectivo; ausencia de fórmula = cadena vacía (fórmula↔literal/ausente
        // cuenta como diferencia). Así se detectan los 3 de `Valida - Control Recaudo` y los 61 del
        // CONSOLIDADO, no solo los cambios fórmula-vs-fórmula.
        var fDiff = 0;
        foreach (var referencia in a.Celdas.Keys)
        {
            if (!b.Celdas.ContainsKey(referencia))
            {
                continue;
            }

            var fa = a.TieneFormula(referencia) ? a.Formula(referencia) : string.Empty;
            var fb = b.TieneFormula(referencia) ? b.Formula(referencia) : string.Empty;
            if (!string.Equals(fa, fb, StringComparison.Ordinal))
            {
                fDiff++;
            }
        }

        // cmpnum: solo celdas con valor no vacío.
        var va = ValoresNoVacios(a);
        var vb = ValoresNoVacios(b);
        var numDiff = 0;
        var textDiff = 0;
        var mixto = 0;
        var onlyA = 0;
        var onlyB = 0;

        foreach (var (referencia, (valorA, esTextoA)) in va)
        {
            if (!vb.TryGetValue(referencia, out var v))
            {
                onlyA++;
                continue;
            }

            var (valorB, esTextoB) = v;
            var na = esTextoA ? (decimal?)null : ToNum(valorA);
            var nb = esTextoB ? (decimal?)null : ToNum(valorB);

            if (na is not null && nb is not null)
            {
                if (Math.Abs((double)(na.Value - nb.Value)) > 0.5)
                {
                    numDiff++;
                }
            }
            else if (esTextoA && esTextoB)
            {
                if (!string.Equals(valorA, valorB, StringComparison.Ordinal))
                {
                    textDiff++;
                }
            }
            else if (!string.Equals(valorA, valorB, StringComparison.Ordinal))
            {
                mixto++;
            }
        }

        foreach (var referencia in vb.Keys)
        {
            if (!va.ContainsKey(referencia))
            {
                onlyB++;
            }
        }

        return (fDiff, numDiff, textDiff, mixto, onlyA, onlyB);
    }

    private static Dictionary<string, (string Valor, bool EsTexto)> ValoresNoVacios(Hoja hoja)
    {
        var mapa = new Dictionary<string, (string, bool)>(StringComparer.OrdinalIgnoreCase);
        foreach (var (referencia, celda) in hoja.Celdas)
        {
            if (celda.Valor.Length > 0)
            {
                mapa[referencia] = (celda.Valor, celda.EsTexto);
            }
        }

        return mapa;
    }

    private static decimal? ToNum(string texto) =>
        decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : null;

    /// <summary>Tolerancia de comparación del T0 (±0.5): los cachés de Excel serializan los mismos
    /// números con distinta representación (`71.702.408,709999993` vs `71.702.408,71`).</summary>
    private static void AssertCerca(decimal esperado, decimal observado)
    {
        var delta = Math.Abs((double)(esperado - observado));
        Assert.True(delta <= 0.5, $"Se esperaba ≈{esperado} (±0.5) y se observó {observado} (Δ={delta}).");
    }

    private static (int GH, int FG, int Otros) ClasificarCambios(
        IReadOnlyList<(string Ref, string Antes, string Despues)> cambios)
    {
        var gH = 0;
        var fG = 0;
        var otros = 0;
        foreach (var (_, antes, despues) in cambios)
        {
            if (string.Equals(antes.Replace("G", "H", StringComparison.Ordinal), despues, StringComparison.Ordinal))
            {
                gH++;
            }
            else if (string.Equals(antes.Replace("F", "G", StringComparison.Ordinal), despues, StringComparison.Ordinal))
            {
                fG++;
            }
            else
            {
                otros++;
            }
        }

        return (gH, fG, otros);
    }

    // ── Lectura BCL ZIP+XML (read-only) ────────────────────────────────────────────────────────

    private sealed record Celda(string Valor, bool EsTexto, string? Formula);

    private sealed class Hoja
    {
        public Dictionary<string, Celda> Celdas { get; } = new(StringComparer.OrdinalIgnoreCase);

        public int MaxRow { get; private set; }

        public IEnumerable<(string Celda, string Formula)> Formulas =>
            Celdas.Where(c => c.Value.Formula is { Length: > 0 }).Select(c => (c.Key, c.Value.Formula!));

        // Igual que el volcado del T0: una etiqueta numérica (`C=3`, `C=9`) se devuelve como "3"/"9".
        public string Texto(string columna, int fila) => Celdas.TryGetValue($"{columna}{fila}", out var c) ? c.Valor : string.Empty;

        public decimal Numero(string referencia) => NumeroOrNull(referencia) ?? 0m;

        public decimal? NumeroOrNull(string referencia) =>
            Celdas.TryGetValue(referencia, out var c) && !c.EsTexto ? ToNum(c.Valor) : null;

        public bool TieneFormula(string referencia) =>
            Celdas.TryGetValue(referencia, out var c) && c.Formula is { Length: > 0 };

        public string Formula(string referencia) =>
            Celdas.TryGetValue(referencia, out var c) ? c.Formula ?? string.Empty : string.Empty;

        public void Agregar(string referencia, Celda celda)
        {
            Celdas[referencia] = celda;
            var fila = int.Parse(new string(referencia.Where(char.IsAsciiDigit).ToArray()), CultureInfo.InvariantCulture);
            if (fila > MaxRow)
            {
                MaxRow = fila;
            }
        }
    }

    private static Hoja Leer(string ruta, string nombreHoja) => LeerLibro(ruta)[nombreHoja];

    private static Dictionary<string, Hoja> LeerLibro(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var wb = Cargar(zip, "xl/workbook.xml");
        var rels = Cargar(zip, "xl/_rels/workbook.xml.rels");
        var relMap = rels.Root!.Elements()
            .ToDictionary(r => r.Attribute("Id")!.Value, r => r.Attribute("Target")!.Value);

        var compartidas = new List<string>();
        if (zip.GetEntry("xl/sharedStrings.xml") is not null)
        {
            compartidas = Cargar(zip, "xl/sharedStrings.xml").Root!.Elements(Ns + "si")
                .Select(si => string.Concat(si.Descendants(Ns + "t").Select(t => t.Value)))
                .ToList();
        }

        var libro = new Dictionary<string, Hoja>(StringComparer.Ordinal);
        foreach (var sh in wb.Root!.Element(Ns + "sheets")!.Elements(Ns + "sheet"))
        {
            var nombre = sh.Attribute("name")!.Value;
            var id = sh.Attribute(RelNs + "id")!.Value;
            var destino = relMap[id];
            if (!destino.StartsWith("xl/", StringComparison.Ordinal))
            {
                destino = "xl/" + destino.TrimStart('/');
            }

            var doc = Cargar(zip, destino);
            var hoja = new Hoja();
            var maestros = new Dictionary<string, string>(StringComparer.Ordinal);
            var pendientes = new List<(string Ref, string Si)>();

            foreach (var row in doc.Root!.Element(Ns + "sheetData")!.Elements(Ns + "row"))
            {
                foreach (var c in row.Elements(Ns + "c"))
                {
                    var referencia = c.Attribute("r")?.Value;
                    if (referencia is null)
                    {
                        continue;
                    }

                    var tipo = c.Attribute("t")?.Value;
                    var v = c.Element(Ns + "v")?.Value;
                    var f = c.Element(Ns + "f");
                    string valor;
                    var esTexto = false;

                    if (tipo == "s" && v is not null && int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx)
                        && idx >= 0 && idx < compartidas.Count)
                    {
                        valor = compartidas[idx];
                        esTexto = true;
                    }
                    else if (tipo == "inlineStr")
                    {
                        valor = string.Concat(c.Element(Ns + "is")?.Descendants(Ns + "t").Select(t => t.Value) ?? []);
                        esTexto = true;
                    }
                    else if (tipo == "str" && v is not null)
                    {
                        valor = v;
                        esTexto = true;
                    }
                    else
                    {
                        valor = v ?? string.Empty;
                    }

                    string? formula = null;
                    if (f is not null)
                    {
                        formula = f.Value;
                        if (f.Attribute("t")?.Value == "shared")
                        {
                            var si = f.Attribute("si")?.Value;
                            if (si is not null)
                            {
                                if (formula.Length > 0)
                                {
                                    maestros[si] = formula;
                                }
                                else
                                {
                                    pendientes.Add((referencia, si));
                                }
                            }
                        }
                    }

                    hoja.Agregar(referencia, new Celda(valor, esTexto, formula));
                }
            }

            foreach (var (referencia, si) in pendientes)
            {
                if (maestros.TryGetValue(si, out var maestro))
                {
                    hoja.Agregar(referencia, hoja.Celdas[referencia] with { Formula = maestro });
                }
            }

            libro[nombre] = hoja;
        }

        return libro;
    }

    private static XDocument Cargar(ZipArchive zip, string entrada)
    {
        var e = zip.GetEntry(entrada) ?? throw new InvalidOperationException($"El xlsx no trae '{entrada}'.");
        using var s = e.Open();
        return XDocument.Load(s);
    }

    private static string[] NombresDeHoja(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        return Cargar(zip, "xl/workbook.xml").Root!.Element(Ns + "sheets")!.Elements(Ns + "sheet")
            .Select(s => s.Attribute("name")!.Value)
            .ToArray();
    }

    // ── Derivación de geometría R2 ─────────────────────────────────────────────────────────────

    private static (int Inicio, int Fin) BloqueR2(Hoja hoja, string ase)
    {
        var cabecera = Enumerable.Range(1, hoja.MaxRow)
            .First(n => string.Equals(hoja.Texto("B", n), ase, StringComparison.Ordinal));
        var inicio = cabecera + 2;
        var fin = Enumerable.Range(inicio, hoja.MaxRow - inicio + 1)
            .First(n => string.Equals(hoja.Texto("A", n), "Total", StringComparison.Ordinal));
        return (inicio, fin);
    }

    private static int Tamano((int Inicio, int Fin) bloque) => bloque.Fin - bloque.Inicio + 1;

    private static int FilaOportunoTotal(Hoja hoja, int ase)
    {
        var totales = Enumerable.Range(1, hoja.MaxRow)
            .Where(n => hoja.Texto("B", n) == "TOTAL" && hoja.Texto("C", n) == "OPORTUNO")
            .ToArray();
        return totales[ase];
    }

    private static string Sha256(string ruta)
    {
        using var stream = File.OpenRead(ruta);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

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

    private static string V1() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string V2() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "v2", "Formato Remuneracion 2026xx-x Total vf.xlsx");

    private static string Manual() =>
        Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Resultado",
            "Resultado Administrativo con plantilla estandar v2", "Formato Remuneracion 202607-2 Total.xlsx");

    private static string CorridaP47H2c() =>
        Path.Combine(Path.GetTempPath(), "opencode", "p47-h2c", "Remuneración 202607-2 Total.xlsx");
}
