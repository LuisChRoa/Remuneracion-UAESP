using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 55 (H2.1, S3.1): cobertura por ASTILLA del gate <c>ValidarSigmaEmpresas</c> (que no tenía
/// ninguna — F5) + un test por delta D1-D5 (causa (i), S0.5/D2) + el candado anti-tautología (CA8).
///
/// <para>
/// Doctrina vinculante: el agregado por empresa se deriva de la FUENTE, agrupado por etiqueta, sobre el
/// mismo conjunto de roles/secciones que alimenta el visible de bloque. <b>Prohibido</b> definirlo como
/// la suma de lo escrito en los términos de la fórmula: eso haría el gate trivialmente verde. Por eso
/// los tests de delta comparan los TRES lados (fuente, dominio por empresa, visible de bloque) y el
/// test de anti-tautología exige que el agregado NO sea la suma de las celdas escritas.
/// </para>
/// </summary>
public sealed class ConciliacionEmpresaSigmaPlan55Tests
{
    private const decimal Tolerancia = 0.5m;

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // CA1 — la corrida 2026101 cierra el gate en los 5 ASE (Σ empresas == visible de bloque)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void H2_1_Octubre1_PorAse_LaSigmaCierraContraElVisibleDeBloque(int aseId)
    {
        var (leaf, conciliacion) = Leer(aseId, "2026101", "Prueba Octubre-1");

        Assert.NotEmpty(conciliacion);

        // Astilla R1 / R2 / R4 del gate, con la MISMA tolerancia (±0,5) y el mismo agregado.
        Assert.InRange(conciliacion.Sum(c => c.VisibleR1), leaf.R1.TotalOportunoEsperadoPorAse - Tolerancia, leaf.R1.TotalOportunoEsperadoPorAse + Tolerancia);
        Assert.InRange(conciliacion.Sum(c => c.VisibleR2), leaf.R2.TotalOportunoEsperado - Tolerancia, leaf.R2.TotalOportunoEsperado + Tolerancia);
        Assert.InRange(conciliacion.Sum(c => c.VisibleR4), leaf.R4.TotalReversionEsperada - Tolerancia, leaf.R4.TotalReversionEsperada + Tolerancia);

        // Y el gate real (el que corre el writer antes de escribir) no lanza.
        WorkbookLeafCoherence.ValidarSigmaEmpresas(conciliacion, leaf);
    }

    /// <summary>DetRetri esperado (R10 oficial) por ASE para 2026101 — CA2 del Plan 55 (Plan 54 S8).</summary>
    private static readonly (int Ase, decimal DetRetri)[] R10Oficial =
    [
        (1, 14_842_425_136m),
        (2, 11_302_637_605m),
        (3, 8_823_787_300m),
        (4, 10_483_172_762m),
        (5, 7_550_668_665m),
    ];

    [Theory]
    [MemberData(nameof(DetRetriEsperado))]
    public void H2_1_Octubre1_ElDetRetriNoSeMueveNiUnPeso(int aseId, decimal esperado)
    {
        var (leaf, _) = Leer(aseId, "2026101", "Prueba Octubre-1");

        // DetRetri = ROUND(D104:D108,0) = R1 + R2 + EXTEMP + R4 (Q1: sin AJUSTES-SF-T).
        var totalD104 = leaf.R1.TotalOportunoEsperadoPorAse
            + leaf.R2.TotalOportunoEsperado
            + leaf.R1.ExtemporaneoEsperadoPorAse
            + leaf.R4.TotalReversionEsperada;

        Assert.Equal(esperado, DetRetriRounder.Round(totalD104));
    }

    public static TheoryData<int, decimal> DetRetriEsperado()
    {
        var data = new TheoryData<int, decimal>();
        foreach (var (ase, detRetri) in R10Oficial)
        {
            data.Add(ase, detRetri);
        }

        return data;
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // CA3 — una prueba por delta D1..D5 (causa (i), S0.5/D2): el delta queda atribuido y cerrado
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// D1 (ASE1·R1): la PRIMERA sección Oportuno (<c>F8</c> = −2.015,07) entra en el bloque
    /// (<c>F75 = F47+F70+F16-L16-L47</c>) y antes NO entraba en el agregado por empresa (cadena de 3
    /// términos). Ahora el agregado cubre TODAS las secciones Oportuno ⇒ Δ = 0.
    /// </summary>
    [Fact]
    public void H2_1_D1_Ase1R1_PrimeraSeccionOportuno_EntraEnElAgregadoYElDeltaEsCero()
    {
        var (leaf, conciliacion) = Leer(1, "2026101", "Prueba Octubre-1");
        var enel = conciliacion.Single(c => c.Empresa.Id == 2);

        Assert.Equal(-2015.07m, enel.CeldasR1["F8"]);
        Assert.Equal(14_471_399_690.90m, enel.VisibleR1);
        Assert.Equal(leaf.R1.TotalOportunoEsperadoPorAse, conciliacion.Sum(c => c.VisibleR1));
    }

    /// <summary>
    /// D2 (ASE4·R1): Reciprocidad (rótulo de fuente <c>NUEVO ESQUEMA</c>) solo trae sección
    /// <c>Aplicacion</c>; el TOT_OPT del bloque suma únicamente secciones <c>Oportuno</c> ⇒ su
    /// contribución al visible es 0 (antes: 309.350,00 absorbidos por <c>UltimaOportuna</c>).
    /// </summary>
    [Fact]
    public void H2_1_D2_Ase4R1_ReciprocidadSoloAplicacion_NoAportaAlVisible()
    {
        var (leaf, conciliacion) = Leer(4, "2026101", "Prueba Octubre-1");
        var recip = conciliacion.Single(c => c.Empresa.Id == 1);

        Assert.Equal(0m, recip.VisibleR1);
        Assert.Equal(9_882_391_541.17m, conciliacion.Single(c => c.Empresa.Id == 2).VisibleR1);
        Assert.Equal(591_576_272.25m, conciliacion.Single(c => c.Empresa.Id == 4).VisibleR1);
        Assert.Equal(leaf.R1.TotalOportunoEsperadoPorAse, conciliacion.Sum(c => c.VisibleR1));
    }

    /// <summary>
    /// D3 (ASE2·R2): la etiqueta <c>EAB</c> la DECLARA el cuadro de la plantilla (<c>R2!B70</c>,
    /// <c>E70</c>/<c>K70</c> literales). Deja de ser "etiqueta sin dueño": entra al agregado con su
    /// importe exacto (149.137,87 − 18.613,87 = 130.524,00) y el Δ cierra.
    /// </summary>
    [Fact]
    public void H2_1_D3_Ase2R2_EabDeclaradaPorElCuadro_EntraAlAgregado()
    {
        var (leaf, conciliacion) = Leer(2, "2026101", "Prueba Octubre-1");
        var eab = conciliacion.Single(c => c.Empresa.Nombre == "EAB");

        Assert.Equal(130_524.00m, eab.VisibleR2);
        Assert.Equal(149_137.87m, eab.CeldasR2["E70"]);
        Assert.Equal(leaf.R2.TotalOportunoEsperado, conciliacion.Sum(c => c.VisibleR2));

        // Exceso registrado (no descarte silencioso): EAB no tiene fila en el cuadro Subs/Cont del
        // ASE2 ⇒ su importe de subsidio queda sin celda declarada y se audita.
        Assert.Contains(eab.DestinosSinCelda, d => d.StartsWith("R2:EAB:Subs/Cont", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// D4 (ASE1·R4): <c>ENEL</c> tiene slot en el cuadro (<c>R4!A8/B8</c>, <c>D8</c>/<c>P8</c>
    /// literales) y antes no estaba en el mapa ⇒ su importe vuelve al agregado (−3.698.870,00).
    /// </summary>
    [Fact]
    public void H2_1_D4_Ase1R4_EnelResuelveSuSlotYElDeltaEsCero()
    {
        var (leaf, conciliacion) = Leer(1, "2026101", "Prueba Octubre-1");
        var enel = conciliacion.Single(c => c.Empresa.Id == 2);

        Assert.Equal(-3_698_870.00m, enel.CeldasR4["D8"]);
        Assert.Equal(-3_698_870.00m, enel.VisibleR4);
        Assert.Equal(leaf.R4.TotalReversionEsperada, conciliacion.Sum(c => c.VisibleR4));
    }

    /// <summary>
    /// D5 (ASE2·R4): <c>OCCIDENTE</c> resuelve su slot declarado (<c>R4!A110/B110</c>) y su importe
    /// (−945.063,00) vuelve al agregado ⇒ Δ = 0.
    /// </summary>
    [Fact]
    public void H2_1_D5_Ase2R4_OccidenteResuelveSuSlotYElDeltaEsCero()
    {
        var (leaf, conciliacion) = Leer(2, "2026101", "Prueba Octubre-1");
        var occidente = conciliacion.Single(c => c.Empresa.Id == 4);

        Assert.Equal(-945_063.00m, occidente.CeldasR4["D110"]);
        Assert.Equal(leaf.R4.TotalReversionEsperada, conciliacion.Sum(c => c.VisibleR4));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // CA8 — el agregado NO es tautológico: no es la suma de las celdas escritas
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// CA8: el visible por empresa se deriva de la FUENTE sobre el MISMO conjunto de roles del bloque
    /// (secciones <c>Oportuno</c>; Especiales de todas menos la última). Las celdas ESCRITAS incluyen
    /// además las secciones <c>Aplicacion</c> (por eso <c>Σ celdas ≠ VisibleR1</c>): si el agregado se
    /// definiera como "la suma de lo que escribimos", el gate sería trivialmente verde y este aserto
    /// fallaría.
    /// </summary>
    [Fact]
    public void H2_1_CA8_ElAgregadoSeDerivaDeLaFuenteYNoDeLoEscrito()
    {
        var (leaf, conciliacion) = Leer(1, "2026101", "Prueba Octubre-1");
        var enel = conciliacion.Single(c => c.Empresa.Id == 2);

        var sumaDeLoEscrito = enel.CeldasR1.Values.Sum();
        Assert.NotEqual(enel.VisibleR1, sumaDeLoEscrito);

        // Derivación al peso: Σ Total de las secciones Oportuno − Especiales de todas menos la última.
        // ENEL: O#1 (F8 = −2.015,07) + O#2 (F36 = 17.734.947.270,61) + O#3 (F62 = −3.253.607.640,54)
        //       − L(O#2) (9.937.924,10) = 14.471.399.690,90.
        Assert.Equal(-2015.07m + 17_734_947_270.61m - 3_253_607_640.54m - 9_937_924.10m, enel.VisibleR1);

        // El bloque (camino independiente: roles Mes del template) cierra al peso con la Σ por empresa.
        Assert.Equal(leaf.R1.TotalOportunoEsperadoPorAse, conciliacion.Sum(c => c.VisibleR1));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // CA7 / S2.4 — ningún destino resuelto es fórmula (H1.4) y los 6 fórmula NO se tocan
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H2_1_H1_4_NingunDestinoResueltoEsFormula_NiTocaLosSeisFormulaDeLaV2()
    {
        var seisFormula = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "R1:F86", "R1:F95", "R1:L95", "R1:F416", "R2:E285", "R2:K285"
        };

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var (_, conciliacion) = Leer(aseId, "2026101", "Prueba Octubre-1");
            foreach (var conc in conciliacion)
            {
                Assert.Empty(conc.DestinosFormula);
                foreach (var (celda, _) in conc.CeldasR1)
                {
                    Assert.DoesNotContain($"R1:{celda}", seisFormula);
                }

                foreach (var (celda, _) in conc.CeldasR2)
                {
                    Assert.DoesNotContain($"R2:{celda}", seisFormula);
                }
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Astillas del gate con chips (sin workbooks): early-returns, tolerancia en el borde y mismatch
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H2_1_Gate_ListaVacia_NoEvalua()
    {
        var leaf = LeafQ1(visibleR1: 100m, visibleR2: 100m, visibleR4: 100m);
        WorkbookLeafCoherence.ValidarSigmaEmpresas([], leaf);
    }

    [Fact]
    public void H2_1_Gate_Q2_NoEvalua()
    {
        var leaf = Leaf(quincena: 2, visibleR1: 100m, visibleR2: 100m, visibleR4: 100m);
        var conciliacion = new[] { Conc(1m, 1m, 1m) };

        WorkbookLeafCoherence.ValidarSigmaEmpresas(conciliacion, leaf);
    }

    [Fact]
    public void H2_1_Gate_ToleranciaEnElBorde_MedioPuntoPasa_MedioPuntoYMedioNo()
    {
        var leaf = LeafQ1(visibleR1: 100m, visibleR2: 100m, visibleR4: 100m);

        WorkbookLeafCoherence.ValidarSigmaEmpresas([Conc(100.5m, 99.5m, 100m)], leaf);

        var ex = Assert.Throws<Remuneracion.Core.Exceptions.CalculoInvalidoException>(
            () => WorkbookLeafCoherence.ValidarSigmaEmpresas([Conc(100.5001m, 100m, 100m)], leaf));
        Assert.Contains("R1", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void H2_1_Gate_MismatchR1_NombraElAse()
    {
        var leaf = LeafQ1(visibleR1: 100m, visibleR2: 100m, visibleR4: 100m);
        var ex = Assert.Throws<Remuneracion.Core.Exceptions.CalculoInvalidoException>(
            () => WorkbookLeafCoherence.ValidarSigmaEmpresas([Conc(10m, 100m, 100m)], leaf));
        Assert.Contains($"ASE {leaf.Ase.Id}", ex.Message, StringComparison.Ordinal);
        Assert.Contains("R1", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void H2_1_Gate_MismatchR2_NombraElAse()
    {
        var leaf = LeafQ1(visibleR1: 100m, visibleR2: 100m, visibleR4: 100m);
        var ex = Assert.Throws<Remuneracion.Core.Exceptions.CalculoInvalidoException>(
            () => WorkbookLeafCoherence.ValidarSigmaEmpresas([Conc(100m, 10m, 100m)], leaf));
        Assert.Contains($"ASE {leaf.Ase.Id}", ex.Message, StringComparison.Ordinal);
        Assert.Contains("R2", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void H2_1_Gate_MismatchR4_NombraElAse()
    {
        var leaf = LeafQ1(visibleR1: 100m, visibleR2: 100m, visibleR4: 100m);
        var ex = Assert.Throws<Remuneracion.Core.Exceptions.CalculoInvalidoException>(
            () => WorkbookLeafCoherence.ValidarSigmaEmpresas([Conc(100m, 100m, 10m)], leaf));
        Assert.Contains($"ASE {leaf.Ase.Id}", ex.Message, StringComparison.Ordinal);
        Assert.Contains("R4", ex.Message, StringComparison.Ordinal);
    }

    // ── helpers ────────────────────────────────────────────────────────────────────────────────

    private static ConciliacionEmpresaInputs Conc(decimal r1, decimal r2, decimal r4) => new()
    {
        Empresa = EmpresaFacturacion.Obtener(2),
        Ase = AseFactory.DesdeId(1),
        VisibleR1 = r1,
        VisibleR2 = r2,
        VisibleR4 = r4
    };

    private static WorkbookLeafInputs LeafQ1(decimal visibleR1, decimal visibleR2, decimal visibleR4) =>
        Leaf(1, visibleR1, visibleR2, visibleR4);

    private static WorkbookLeafInputs Leaf(int quincena, decimal visibleR1, decimal visibleR2, decimal visibleR4) => new()
    {
        Ase = AseFactory.DesdeId(1),
        Periodo = new Periodo { CodigoAAAAMM = "202610", NumeroQuincena = quincena },
        R1 = new WorkbookLeafInputsR1 { TotalOportunoEsperadoPorAse = visibleR1 },
        R2 = new WorkbookLeafInputsR2 { E15 = visibleR2, E26 = 0m, K15 = 0m },
        R4 = new WorkbookLeafInputsR4 { D9 = visibleR4, P9 = 0m }
    };

    private static (WorkbookLeafInputs Leaf, IReadOnlyList<ConciliacionEmpresaInputs> Conciliacion) Leer(
        int aseId, string codigo, string carpetaPeriodo)
    {
        var (r1, r2, r4) = Rutas(carpetaPeriodo, aseId);
        var periodo = Periodo.Parse(codigo);

        // Plan 55 (H1.1/D4): la PLANTILLA es la autoridad de los destinos (resolución por etiquetas).
        var reader = new ExcelDataReaderWorkbookLeafInputReader { RutaPlantilla = Insumos.PlantillaOficial() };
        var ase = AseFactory.DesdeId(aseId);
        var leaf = reader.LeerLeafInputs(ase, periodo, r1, r2, r4);
        var conciliacion = reader.LeerConciliacionEmpresas(ase, periodo, r1, r2, r4);

        return (leaf, conciliacion);
    }

    private static (string R1, string R2, string R4) Rutas(string carpetaPeriodo, int aseId)
    {
        var carpeta = Path.Combine(Insumos.Raiz(), "Docs", carpetaPeriodo, "Insumos", CarpetasAse[aseId - 1]);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta explícita de insumos: {carpeta}");

        return (
            Unico(carpeta, "Recaudoporcomponente"),
            Unico(carpeta, "RerpoteDetalleSaldosaFavor"),
            Unico(carpeta, "Reversi"));
    }

    private static string Unico(string carpeta, string prefijo)
    {
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta '{prefijo}*' en {carpeta}.");
        return archivo!;
    }
}
