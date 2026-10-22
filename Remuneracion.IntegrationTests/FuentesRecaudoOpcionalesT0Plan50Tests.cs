using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 50 (H0, S1.1-S1.5): T0 READ-ONLY, con CERO código de producción.
///
/// Congela por BYTES (BCL <see cref="ZipArchive"/> + <see cref="XDocument"/>; sin Excel/COM, sin
/// OpenXml SDK, sin ExcelDataReader) los cinco veredictos del plan:
/// <list type="bullet">
/// <item>S1.1/S1.2 — inventario de ausencia por (período × ASE × reporte × label) en octubre-1,
/// julio-2 y agosto-2: qué labels de los mapas por empresa no traen NINGUNA fila (ausencia total →
/// 0 + huella por A1) y cuáles traen filas (presencia);</item>
/// <item>S1.2 — el ÚNICO caso parcial con label presente: octubre-1 ASE3 R1 OCCIDENTE tiene una sola
/// fila <c>Total</c> y el mapa exige <c>TotalMedio</c> (índice 1) → fail-fast vigente, residuo
/// declarado con su propio T0 (no se mueve a 0 sin veredicto de negocio);</item>
/// <item>S1.3 — A3: la aritmética "Σ fuentes presentes vs total del reporte" NO es viable en R1
/// (el total del cuadro <c>Componente</c> no es la suma de las filas-etiqueta: Δ ≠ 0 medido) ni en R2
/// (el total incluye la etiqueta <c>EAB</c>, que no tiene destino en el mapa → Σ-reconocidas ≠ total).
/// En R4 sí cierra exacto y por construcción (Σ filas-etiqueta == fila <c>Total</c>), por lo que
/// aporta la misma información que el control de etiquetas. Veredicto: control mínimo alternativo
/// (§2.5) = <b>control de etiquetas no reconocidas</b>, con T0 propio (este fixture);</item>
/// <item>S1.4 — residuo Q2: la ÚNICA divergencia Q2 es la ausencia del label en R4 por empresa
/// (cubierta por A1), y ningún reporte Q2 tiene parciales;</item>
/// <item>S1.5 — el inventario de etiquetas NO reconocidas por (período, ASE, reporte) que alimenta el
/// control A3: hoy solo <c>EAB</c> en R2 (ASE2 octubre-1; ASE3 julio-2/agosto-2; ASE4 agosto-2),
/// siempre SIN ausencias declaradas en el mismo reporte (por eso no dispara fail-fast).</item>
/// </list>
///
/// El fixture NO usa los mapas de producción ni el reader: es una derivación INDEPENDIENTE (espejo
/// local de las reglas de localización del plan §2.1) y su matriz congelada es la evidencia del T0.
/// La puerta H0 → H1 exige este fixture verde.
/// </summary>
public sealed class FuentesRecaudoOpcionalesT0Plan50Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    /// <summary>Label de la fuente por (empresa, ASE) — espejo de <c>ObtenerLabelFuente</c>.</summary>
    private static string Label(int empresa, int ase) => empresa == 1 && ase is 2 or 4
        ? "NUEVO ESQUEMA"
        : empresa switch
        {
            1 => "RECIPROCIDAD",
            2 => "ENEL",
            3 => "ENERBIT",
            4 => "OCCIDENTE",
            _ => "CIUDAD LIMPIA-ACUEDUCTO",
        };

    /// <summary>Empresas con operando declarado en cada mapa por empresa (espejo de los diccionarios T0).</summary>
    private static readonly Dictionary<(string Reporte, int Ase), int[]> EmpresasPorMapa = new()
    {
        [("R1", 1)] = [2, 4],
        [("R1", 2)] = [1, 2, 4],
        [("R1", 3)] = [2, 4],
        [("R1", 4)] = [1, 2, 4],
        [("R1", 5)] = [2, 3, 4],
        [("R2", 1)] = [2, 4],
        [("R2", 2)] = [1, 2, 4],
        [("R2", 3)] = [2, 4],
        [("R2", 4)] = [1, 2, 4],
        [("R2", 5)] = [2, 4],
        [("R4", 1)] = [4],
        [("R4", 2)] = [1, 2],
        [("R4", 3)] = [2, 4],
        [("R4", 4)] = [1, 2, 4],
        [("R4", 5)] = [2, 4],
    };

    // ── S1.1/S1.2 — matriz de presencia/ausencia total por (período, ASE, reporte) ───────────────

    /// <summary>
    /// Líneas congeladas: <c>período|ASE|reporte|presentes|ausentes</c>. "Ausente" = CERO filas del
    /// label en el reporte (condición de A1: 0 explícito + huella). "Presente" = al menos una fila
    /// (si además falta un bloque que un operando exige, es el caso parcial de S1.2).
    /// </summary>
    public static TheoryData<string, int, string, string> MatrizAusenciasData()
    {
        const string P1 = "Prueba Octubre-1";
        const string P2 = "Prueba Julio-2";
        const string P3 = "Prueba Agosto-2";
        string[] lineas =
        [
            $"{P1}|1|R1|ENEL,OCCIDENTE|",
            $"{P1}|1|R2|ENEL,OCCIDENTE|",
            $"{P1}|1|R4|OCCIDENTE|",
            $"{P1}|2|R1|ENEL,OCCIDENTE|NUEVO ESQUEMA",
            $"{P1}|2|R2|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P1}|2|R4|ENEL|NUEVO ESQUEMA",
            $"{P1}|3|R1|ENEL,OCCIDENTE|",
            $"{P1}|3|R2|ENEL,OCCIDENTE|",
            $"{P1}|3|R4|ENEL|OCCIDENTE",
            $"{P1}|4|R1|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P1}|4|R2|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P1}|4|R4|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P1}|5|R1|ENEL,ENERBIT,OCCIDENTE|",
            $"{P1}|5|R2|ENEL,OCCIDENTE|",
            $"{P1}|5|R4|ENEL|OCCIDENTE",
            $"{P2}|1|R1|ENEL,OCCIDENTE|",
            $"{P2}|1|R2|ENEL,OCCIDENTE|",
            $"{P2}|1|R4|OCCIDENTE|",
            $"{P2}|2|R1|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P2}|2|R2|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P2}|2|R4|ENEL|NUEVO ESQUEMA",
            $"{P2}|3|R1|ENEL,OCCIDENTE|",
            $"{P2}|3|R2|ENEL,OCCIDENTE|",
            $"{P2}|3|R4|ENEL,OCCIDENTE|",
            $"{P2}|4|R1|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P2}|4|R2|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P2}|4|R4|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P2}|5|R1|ENEL,ENERBIT,OCCIDENTE|",
            $"{P2}|5|R2|ENEL,OCCIDENTE|",
            $"{P2}|5|R4|ENEL,OCCIDENTE|",
            $"{P3}|1|R1|ENEL,OCCIDENTE|",
            $"{P3}|1|R2|ENEL,OCCIDENTE|",
            $"{P3}|1|R4||OCCIDENTE",
            $"{P3}|2|R1|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P3}|2|R2|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P3}|2|R4|ENEL|NUEVO ESQUEMA",
            $"{P3}|3|R1|ENEL,OCCIDENTE|",
            $"{P3}|3|R2|ENEL,OCCIDENTE|",
            $"{P3}|3|R4|ENEL|OCCIDENTE",
            $"{P3}|4|R1|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P3}|4|R2|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P3}|4|R4|ENEL,NUEVO ESQUEMA,OCCIDENTE|",
            $"{P3}|5|R1|ENEL,ENERBIT,OCCIDENTE|",
            $"{P3}|5|R2|ENEL,OCCIDENTE|",
            $"{P3}|5|R4|ENEL,OCCIDENTE|",
        ];

        var data = new TheoryData<string, int, string, string>();
        foreach (var linea in lineas)
        {
            var partes = linea.Split('|');
            data.Add(partes[0], int.Parse(partes[1], CultureInfo.InvariantCulture), partes[2], partes[3] + "|" + partes[4]);
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(MatrizAusenciasData))]
    public void T0P50_MatrizAusenciaTotal_PorPeriodoAseReporte(string periodo, int ase, string reporte, string esperado)
    {
        var separador = esperado.IndexOf('|');
        var presentesEsperados = esperado[..separador];
        var ausentesEsperados = esperado[(separador + 1)..];

        var filasPorLabel = ContarFilasPorLabel(periodo, ase, reporte);
        var presentes = string.Join(",", filasPorLabel.Where(kv => kv.Value > 0).Select(kv => kv.Key).OrderBy(x => x, StringComparer.Ordinal));
        var ausentes = string.Join(",", filasPorLabel.Where(kv => kv.Value == 0).Select(kv => kv.Key).OrderBy(x => x, StringComparer.Ordinal));

        Assert.Equal(presentesEsperados, presentes);
        Assert.Equal(ausentesEsperados, ausentes);
    }

    // ── S1.2 — el único parcial con label presente (fail-fast vigente) ───────────────────────────

    [Fact]
    public void T0P50_Parcial_OctubreAse3R1Occidente_UnaSolaFilaTotal_TotalMedioSinFila()
    {
        var filas = LeerFilas(RutaFuente("Prueba Octubre-1", 3, "Recaudoporcomponente"));
        var indiceComponente = filas.FindIndex(f => Txt(f, 0).Equals("Componente", StringComparison.OrdinalIgnoreCase));
        Assert.True(indiceComponente > 0, "La fuente R1 de octubre-1 ASE3 debe traer la fila invariante 'Componente'.");

        // Evidencia: OCCIDENTE trae UNA fila Total en el bloque Componente y UNA en Subs/Cont.
        var totalesOccidente = filas
            .Select((fila, indice) => (fila, indice))
            .Where(t => t.indice < indiceComponente && Txt(t.fila, 2).Equals("OCCIDENTE", StringComparison.OrdinalIgnoreCase)
                        && Txt(t.fila, 3).Equals("Total", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var subsOccidente = filas
            .Select((fila, indice) => (fila, indice))
            .Where(t => t.indice > indiceComponente && Txt(t.fila, 2).Equals("OCCIDENTE", StringComparison.OrdinalIgnoreCase)
                        && Txt(t.fila, 3).Equals("Total", StringComparison.OrdinalIgnoreCase))
            .ToList();

        Assert.Single(totalesOccidente);
        Assert.Single(subsOccidente);
        Assert.Equal(286711907.53m, Num(totalesOccidente[0].fila, 5));
        Assert.Equal(145489038.47m, Num(subsOccidente[0].fila, 5));

        // El mapa (4,3) exige TotalMedio = índice 1 de la lista de filas Total → no existe.
        // Regla ratificada (§7.3): parcial con label presente = FAIL-FAST vigente, no 0 silencioso.
        Assert.True(totalesOccidente.Count - 1 < 1, "TotalMedio (índice 1) no resuelve con una sola fila Total.");
    }

    // ── S1.3 — aritmética A3: R1 no cierra, R2 reconocidas no cierra, R4 cierra por construcción ─

    [Fact]
    public void T0P50_A3_R1_SumaEtiquetasVsComponenteTotal_NoCierra()
    {
        // Δ medidos (octubre-1): el cuadro Componente/Total NO es la suma de las filas-etiqueta
        // (incluye la fila frontera 'Aplicacion nuevos x reversion' y excluye la sección primera).
        Assert.Equal(2015.07m, decimal.Round(DeltaR1("Prueba Octubre-1", 1), 2));
        Assert.Equal(243714.65m, decimal.Round(DeltaR1("Prueba Octubre-1", 2), 2));
        Assert.Equal(558.05m, decimal.Round(DeltaR1("Prueba Octubre-1", 3), 2));
        Assert.Equal(30180.94m, decimal.Round(DeltaR1("Prueba Octubre-1", 4), 2));

        // Δ medidos (julio-2 y agosto-2): tampoco cierra → la aritmética de R1 no es viable.
        Assert.Equal(11217.37m, decimal.Round(DeltaR1("Prueba Julio-2", 1), 2));
        Assert.Equal(405906.75m, decimal.Round(DeltaR1("Prueba Julio-2", 2), 2));
        Assert.Equal(16458.46m, decimal.Round(DeltaR1("Prueba Julio-2", 3), 2));
        Assert.Equal(1443509.98m, decimal.Round(DeltaR1("Prueba Agosto-2", 2), 2));
        Assert.Equal(4456.11m, decimal.Round(DeltaR1("Prueba Agosto-2", 3), 2));
    }

    [Fact]
    public void T0P50_A3_R2_SumaReconocidasVsComponenteTotal_NoCierraPorEtiquetaEab()
    {
        // R2: Σ de TODAS las filas-etiqueta == Componente/Total (cierra por construcción: no aporta
        // poder de detección). Σ de las RECONOCIDAS excluye 'EAB' (fuente real sin destino en el mapa,
        // deuda del Plan 49) → Δ ≠ 0 en corridas sanas → la aritmética de R2 no es viable como
        // anti-falso-0; el control mínimo es el de etiquetas.
        Assert.Equal(149137.87m, decimal.Round(DeltaR2Reconocidas("Prueba Octubre-1", 2), 2));
        Assert.Equal(71689.77m, decimal.Round(DeltaR2Reconocidas("Prueba Agosto-2", 3), 2));
        Assert.Equal(71673.26m, decimal.Round(DeltaR2Reconocidas("Prueba Agosto-2", 4), 2));
        Assert.Equal(0m, decimal.Round(DeltaR2Reconocidas("Prueba Octubre-1", 1), 2));
    }

    [Theory]
    [InlineData("Prueba Octubre-1")]
    [InlineData("Prueba Julio-2")]
    [InlineData("Prueba Agosto-2")]
    public void T0P50_A3_R4_SumaEtiquetasIgualFilaTotal_CierraPorConstruccion(string periodo)
    {
        for (var ase = 1; ase <= 5; ase++)
        {
            var filas = LeerFilas(RutaFuente(periodo, ase, "Reversi"));
            var totalReporte = filas.First(f => Txt(f, 0).Equals("Total", StringComparison.OrdinalIgnoreCase));
            var suma = filas
                .Where(f => Txt(f, 1).Equals("Total", StringComparison.OrdinalIgnoreCase) && Txt(f, 0).Length > 0)
                .Sum(f => Num(f, 3));

            Assert.Equal(Num(totalReporte, 3), suma);
        }
    }

    // ── S1.5 — inventario de etiquetas NO reconocidas (insumo del control A3) ────────────────────

    [Theory]
    [InlineData("Prueba Octubre-1", 1, "R2", "")]
    [InlineData("Prueba Octubre-1", 2, "R1", "")]
    [InlineData("Prueba Octubre-1", 2, "R2", "EAB")]
    [InlineData("Prueba Octubre-1", 2, "R4", "")]
    [InlineData("Prueba Octubre-1", 3, "R4", "")]
    [InlineData("Prueba Octubre-1", 5, "R2", "")]
    [InlineData("Prueba Julio-2", 2, "R4", "")]
    [InlineData("Prueba Julio-2", 3, "R2", "EAB")]
    [InlineData("Prueba Julio-2", 5, "R2", "")]
    [InlineData("Prueba Agosto-2", 1, "R4", "")]
    [InlineData("Prueba Agosto-2", 2, "R4", "")]
    [InlineData("Prueba Agosto-2", 3, "R2", "EAB")]
    [InlineData("Prueba Agosto-2", 4, "R2", "EAB")]
    [InlineData("Prueba Agosto-2", 5, "R2", "")]
    public void T0P50_EtiquetasNoReconocidas_PorPeriodoAseReporte(string periodo, int ase, string reporte, string esperadas)
    {
        var observadas = EtiquetasNoReconocidas(periodo, ase, reporte);
        Assert.Equal(esperadas, string.Join(",", observadas));
    }

    [Fact]
    public void T0P50_EtiquetasNoReconocidas_NoCoincidenConAusenciasEnElMismoReporte()
    {
        // Regla §2.5 (control mínimo): sólo hay fail-fast si una etiqueta no reconocida coexiste con
        // una ausencia declarada en el MISMO reporte del MISMO ASE. Hoy no ocurre en ninguna muestra.
        string[] muestras = ["Prueba Octubre-1", "Prueba Julio-2", "Prueba Agosto-2"];
        foreach (var periodo in muestras)
        {
            for (var ase = 1; ase <= 5; ase++)
            {
                foreach (var reporte in new[] { "R1", "R2", "R4" })
                {
                    var noReconocidas = EtiquetasNoReconocidas(periodo, ase, reporte);
                    var ausentes = ContarFilasPorLabel(periodo, ase, reporte).Where(kv => kv.Value == 0).Select(kv => kv.Key).ToList();
                    Assert.False(noReconocidas.Count > 0 && ausentes.Count > 0,
                        $"{periodo} ASE {ase} {reporte}: no reconocidas [{string.Join(",", noReconocidas)}] + ausentes [{string.Join(",", ausentes)}]");
                }
            }
        }
    }

    // ── S1.4 — residuo Q2: ausencias R4 cubiertas por A1 + el gate Σ no es válido en Q2 ─────────

    [Fact]
    public void T0P50_Q2_AusenciasR4PorEmpresa_SonLaUnicaDivergenciaQueA1Absorbe()
    {
        foreach (var periodo in new[] { "Prueba Julio-2", "Prueba Agosto-2" })
        {
            for (var ase = 1; ase <= 5; ase++)
            {
                // R1 y R2 no tienen ausencias en Q2: todos sus labels traen filas.
                foreach (var reporte in new[] { "R1", "R2" })
                {
                    var ausentes = ContarFilasPorLabel(periodo, ase, reporte).Where(kv => kv.Value == 0).Select(kv => kv.Key).ToList();
                    Assert.Empty(ausentes);
                }
            }
        }

        Assert.Equal(["R4:NUEVO ESQUEMA"], ContarFilasPorLabel("Prueba Julio-2", 2, "R4")
            .Where(kv => kv.Value == 0).Select(kv => "R4:" + kv.Key).ToArray());
        Assert.Equal(["R4:OCCIDENTE"], ContarFilasPorLabel("Prueba Agosto-2", 1, "R4")
            .Where(kv => kv.Value == 0).Select(kv => "R4:" + kv.Key).ToArray());
        Assert.DoesNotContain(ContarFilasPorLabel("Prueba Agosto-2", 4, "R4"), kv => kv.Value == 0);
    }

    [Fact]
    public void T0P50_Q2_GateSigmaEmpresasVsVisibleDeBloque_NoCierra_ResiduoRSigmaQ2()
    {
        // Residuo declarado R-SIGMA-Q2: el gate Σ (HU-08 §2.5) se definió sobre los mapas Q1 y no es
        // una descomposición válida del visible de bloque en Q2. A2 puebla la conciliación en Q2 sin
        // reactivar el gate (ver `Docs/T0-Plan50-FuentesRecaudoOpcionales.md` §4.2).
        Assert.Equal(15533961.70m, decimal.Round(DeltaSigmaR4("Prueba Julio-2", 1), 2));   // ENEL sin celda destino (R4 ASE1)
        Assert.Equal(290750m, decimal.Round(DeltaSigmaR4("Prueba Julio-2", 2), 2));       // OCCIDENTE sin celda destino (R4 ASE2)
        Assert.Equal(73090m, decimal.Round(DeltaSigmaR4("Prueba Agosto-2", 1), 2));       // OCCIDENTE ausente + ENEL sin destino
        Assert.Equal(0m, decimal.Round(DeltaSigmaR4("Prueba Julio-2", 4), 2));            // R4 ASE4 cubre las 3 fuentes

        Assert.Equal(73523.10m, decimal.Round(DeltaSigmaR2("Prueba Julio-2", 3), 2));      // etiqueta EAB sin destino
        Assert.Equal(6242402.48m, decimal.Round(DeltaSigmaR2("Prueba Julio-2", 5), 2));    // etiqueta ENERBIT sin destino (R2 ASE5)
        Assert.Equal(0m, decimal.Round(DeltaSigmaR2("Prueba Julio-2", 1), 2));

        Assert.Equal(12823613987.88m, decimal.Round(DeltaSigmaR1Q2("Prueba Julio-2", 5), 2)); // TotalMain == TotalMenor (1 sola fila)
        Assert.Equal(11217.37m, decimal.Round(DeltaSigmaR1Q2("Prueba Julio-2", 1), 2));
        Assert.Equal(0m, decimal.Round(DeltaSigmaR1Q2("Prueba Julio-2", 4), 2));
    }

    private static decimal DeltaSigmaR4(string periodo, int ase)
    {
        var filas = LeerFilas(RutaFuente(periodo, ase, "Reversi"));
        var total = filas.First(f => Txt(f, 0).Equals("Total", StringComparison.OrdinalIgnoreCase) && Txt(f, 1).Length == 0);
        var suma = EmpresasPorMapa[("R4", ase)]
            .Select(empresa => filas.FirstOrDefault(f =>
                Txt(f, 0).Equals(Label(empresa, ase), StringComparison.OrdinalIgnoreCase)
                && Txt(f, 1).Equals("Total", StringComparison.OrdinalIgnoreCase)))
            .Where(fila => fila is not null)
            .Sum(fila => Num(fila!, 3));
        return Math.Abs(suma - Num(total, 3));
    }

    private static decimal DeltaSigmaR2(string periodo, int ase)
    {
        var filas = LeerFilas(RutaFuente(periodo, ase, "RerpoteDetalleSaldosaFavor"));
        var idxEspeciales = IndiceEncabezado(filas, "Especiales");
        var componente = filas.FindIndex(f => Txt(f, 0).Equals("Componente", StringComparison.OrdinalIgnoreCase)
                                               && Txt(f, 1).Equals("Total", StringComparison.OrdinalIgnoreCase));
        var subs = filas.FindIndex(f => Txt(f, 0).Equals("Subs/Cont", StringComparison.OrdinalIgnoreCase)
                                        && Txt(f, 1).Equals("Total", StringComparison.OrdinalIgnoreCase));
        Assert.True(componente > 0 && subs > 0, $"{periodo} ASE {ase}: faltan las invariantes del R2.");
        var bloque = Num(filas[componente], 4) + Num(filas[subs], 4)
                     - (idxEspeciales >= 0 ? Num(filas[componente], idxEspeciales) : 0m);

        var suma = 0m;
        foreach (var empresa in EmpresasPorMapa[("R2", ase)])
        {
            var label = Label(empresa, ase);
            var filasEmpresa = filas.Select((f, i) => (f, i))
                .Where(t => Txt(t.f, 1).Equals(label, StringComparison.OrdinalIgnoreCase)
                            && Txt(t.f, 2).Equals("Total", StringComparison.OrdinalIgnoreCase))
                .Select(t => (t.i, Valor: Num(t.f, 4)))
                .OrderByDescending(t => Math.Abs(t.Valor))
                .ToList();
            if (filasEmpresa.Count == 0)
            {
                continue;
            }

            var total = filasEmpresa[0];
            var subsCont = filasEmpresa[^1];
            var especiales = idxEspeciales >= 0 ? Num(filas[total.i], idxEspeciales) : 0m;
            suma += total.Valor + subsCont.Valor - especiales;
        }

        return Math.Abs(suma - bloque);
    }

    /// <summary>Operandos R1 declarados por el mapa (espejo de <c>EditablesR1PorEmpresa</c>).</summary>
    private static readonly Dictionary<(int Empresa, int Ase), string[]> OperandosR1 = new()
    {
        [(2, 1)] = ["SubsMain", "TotalMain", "EspecialesMain"],
        [(4, 1)] = ["SubsMain", "TotalMain", "EspecialesMain", "SubsMenor", "TotalMenor", "EspecialesMenor"],
        [(1, 2)] = ["TotalMain", "SubsMain", "Block5Total"],
        [(2, 2)] = ["SubsMain", "TotalMain", "TotalMenor", "EspecialesMain"],
        [(4, 2)] = ["SubsMain", "TotalMain", "TotalMenor", "EspecialesMain"],
        [(2, 3)] = ["SubsMain", "TotalMain", "TotalMenor", "EspecialesMain"],
        [(4, 3)] = ["SubsMain", "TotalMain", "TotalMenor", "EspecialesMain", "SubsMenor", "TotalMedio"],
        [(1, 4)] = ["SubsMain", "TotalMain", "EspecialesMain"],
        [(2, 4)] = ["SubsMain", "TotalMain", "TotalMenor", "EspecialesMain", "SubsMenor", "TotalMedio"],
        [(4, 4)] = ["SubsMain", "TotalMain", "TotalMenor", "EspecialesMain", "SubsMenor", "TotalMedio"],
        [(2, 5)] = ["SubsMain", "TotalMain", "TotalMenor", "EspecialesMain"],
        [(3, 5)] = ["SubsMain", "TotalMain", "EspecialesMain"],
        [(4, 5)] = ["SubsMain", "TotalMain", "TotalMenor", "EspecialesMain"],
    };

    private static decimal DeltaSigmaR1Q2(string periodo, int ase)
    {
        // Σ visibles por empresa (cadena CalcularVisibleR1) vs totOpt Q2 (ΣF(Mes) − ΣEsp(Mes − última)).
        var filas = LeerFilas(RutaFuente(periodo, ase, "Recaudoporcomponente"));
        var idxEspeciales = IndiceEncabezado(filas, "SERVICIO ESPECIALES");
        var indiceComponente = filas.FindIndex(f => Txt(f, 0).Equals("Componente", StringComparison.OrdinalIgnoreCase));
        Assert.True(indiceComponente > 0, $"{periodo} ASE {ase}: falta la fila invariante 'Componente'.");

        var filasTotal = Enumerable.Range(0, indiceComponente)
            .Where(i => Txt(filas[i], 2).Length > 0 && Txt(filas[i], 3).Equals("Total", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var filasSubs = Enumerable.Range(indiceComponente + 1, filas.Count - indiceComponente - 1)
            .Where(i => Txt(filas[i], 2).Length > 0 && Txt(filas[i], 3).Equals("Total", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var suma = 0m;
        foreach (var ((empresa, aseMapa), operandos) in OperandosR1.Where(k => k.Key.Ase == ase))
        {
            Assert.Equal(ase, aseMapa);
            var label = Label(empresa, ase);
            var totales = filasTotal.Where(i => Txt(filas[i], 2).Equals(label, StringComparison.OrdinalIgnoreCase))
                .Select(i => (i, Valor: Num(filas[i], 5))).OrderByDescending(t => Math.Abs(t.Valor)).ToList();
            var subs = filasSubs.Where(i => Txt(filas[i], 2).Equals(label, StringComparison.OrdinalIgnoreCase))
                .Select(i => (i, Valor: Num(filas[i], 5))).OrderByDescending(t => Math.Abs(t.Valor)).ToList();

            // Sólo los operandos DECLARADOS por el mapa entran en la cadena (como el reader).
            var totalMain = operandos.Contains("TotalMain") && totales.Count > 0 ? totales[0].Valor : 0m;
            var especialesMain = operandos.Contains("EspecialesMain") && totales.Count > 0 && idxEspeciales >= 0
                ? Num(filas[totales[0].i], idxEspeciales)
                : 0m;
            var subsMain = operandos.Contains("SubsMain") && subs.Count > 0 ? subs[0].Valor : 0m;
            var totalMenor = operandos.Contains("TotalMenor") && totales.Count > 0 ? totales[^1].Valor : 0m;
            var block5 = operandos.Contains("Block5Total")
                ? filas.Where(f => Txt(f, 3).Equals("5", StringComparison.OrdinalIgnoreCase)
                                   && Txt(f, 4).Equals("Total", StringComparison.OrdinalIgnoreCase))
                    .Select(f => Num(f, 5)).FirstOrDefault()
                : 0m;

            suma += ase switch
            {
                1 => totalMain + subsMain - especialesMain,
                2 when empresa == 1 => totalMain + subsMain + block5,
                2 => totalMain + subsMain + totalMenor - especialesMain,
                3 => totalMain + subsMain + totalMenor - especialesMain,
                4 when empresa == 1 => totalMain + subsMain - especialesMain,
                _ => totalMain + subsMain + totalMenor - especialesMain,
            };
        }

        var filasMes = Enumerable.Range(0, filas.Count)
            .Where(i => (Txt(filas[i], 0).Equals("Mes", StringComparison.OrdinalIgnoreCase) && Txt(filas[i], 1).Equals("Total", StringComparison.OrdinalIgnoreCase))
                        || (Txt(filas[i], 1).Equals("Mes", StringComparison.OrdinalIgnoreCase) && Txt(filas[i], 2).Equals("Total", StringComparison.OrdinalIgnoreCase)))
            .ToList();
        var totOpt = filasMes.Sum(i => Num(filas[i], 5))
                     - (filasMes.Count > 1 ? filasMes.Take(filasMes.Count - 1).Sum(i => idxEspeciales >= 0 ? Num(filas[i], idxEspeciales) : 0m) : 0m);

        return Math.Abs(suma - totOpt);
    }

    private static int IndiceEncabezado(List<Dictionary<int, object?>> filas, string texto)
    {
        foreach (var fila in filas.Take(10))
        {
            foreach (var (columna, valor) in fila)
            {
                if (valor is string s && s.Trim().Equals(texto, StringComparison.OrdinalIgnoreCase))
                {
                    return columna;
                }
            }
        }

        return -1;
    }
    // ── Helpers de bytes ────────────────────────────────────────────────────────────────────────

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md"))) return dir.FullName;
            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }

    private static string RutaFuente(string periodo, int ase, string prefijo)
    {
        var dir = Path.Combine(Raiz(), "Docs", periodo, "Insumos", CarpetasAse[ase - 1]);
        Assert.True(Directory.Exists(dir), $"Falta la carpeta explícita de insumos: {dir}");
        var archivo = Directory.EnumerateFiles(dir, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta el reporte '{prefijo}*' del ASE {ase} en {dir}");
        return archivo!;
    }

    /// <summary>Cuenta filas-etiqueta por label del mapa según la regla de localización de cada reporte.</summary>
    private static Dictionary<string, int> ContarFilasPorLabel(string periodo, int ase, string reporte)
    {
        var archivo = RutaFuente(periodo, ase, Prefijo(reporte));
        var filas = LeerFilas(archivo);

        var conteo = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var empresa in EmpresasPorMapa[(reporte, ase)])
        {
            var label = Label(empresa, ase);
            var total = reporte switch
            {
                "R1" => filas.Count(f => Txt(f, 2).Equals(label, StringComparison.OrdinalIgnoreCase)
                                         && Txt(f, 3).Equals("Total", StringComparison.OrdinalIgnoreCase)),
                "R2" => filas.Count(f => Txt(f, 1).Equals(label, StringComparison.OrdinalIgnoreCase)
                                         && Txt(f, 2).Equals("Total", StringComparison.OrdinalIgnoreCase)),
                _ => filas.Count(f => Txt(f, 0).Equals(label, StringComparison.OrdinalIgnoreCase)
                                       && Txt(f, 1).Equals("Total", StringComparison.OrdinalIgnoreCase)),
            };
            conteo[label] = total;
        }

        return conteo;
    }

    /// <summary>Etiquetas observadas que no pertenecen al vocabulario de fuentes del negocio.</summary>
    private static List<string> EtiquetasNoReconocidas(string periodo, int ase, string reporte)
    {
        var filas = LeerFilas(RutaFuente(periodo, ase, Prefijo(reporte)));
        // Vocabulario de FUENTES (catálogo de empresas + rótulo vigente por ASE), espejo de producción.
        var conocidas = Enumerable.Range(1, 5).Select(e => Label(e, ase)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        // Sinonimia citada (Plan 34): RECIPROCIDAD ↔ NUEVO ESQUEMA son la misma EFC.
        if (conocidas.Contains("RECIPROCIDAD")) conocidas.Add("NUEVO ESQUEMA");
        if (conocidas.Contains("NUEVO ESQUEMA")) conocidas.Add("RECIPROCIDAD");

        var observadas = reporte switch
        {
            "R1" => filas.Where(f => Txt(f, 3).Equals("Total", StringComparison.OrdinalIgnoreCase)).Select(f => Normalizar(Txt(f, 2))),
            "R2" => filas.Where(f => Txt(f, 2).Equals("Total", StringComparison.OrdinalIgnoreCase)).Select(f => Normalizar(Txt(f, 1))),
            _ => filas.Where(f => Txt(f, 1).Equals("Total", StringComparison.OrdinalIgnoreCase)).Select(f => Normalizar(Txt(f, 0))),
        };

        return observadas
            .Where(e => e.Length > 0 && !conocidas.Contains(e))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(e => e, StringComparer.Ordinal)
            .ToList();
    }

    private static decimal DeltaR1(string periodo, int ase)
    {
        var filas = LeerFilas(RutaFuente(periodo, ase, "Recaudoporcomponente"));
        var indiceComponente = filas.FindIndex(f => Txt(f, 0).Equals("Componente", StringComparison.OrdinalIgnoreCase));
        Assert.True(indiceComponente > 0, $"{periodo} ASE {ase}: falta la fila invariante 'Componente'.");
        var suma = filas.Take(indiceComponente)
            .Where(f => Txt(f, 3).Equals("Total", StringComparison.OrdinalIgnoreCase) && Txt(f, 2).Length > 0)
            .Sum(f => Num(f, 5));
        return Math.Abs(suma - Num(filas[indiceComponente], 5));
    }

    private static decimal DeltaR2Reconocidas(string periodo, int ase)
    {
        var filas = LeerFilas(RutaFuente(periodo, ase, "RerpoteDetalleSaldosaFavor"));
        var indiceComponente = filas.FindIndex(f => Txt(f, 0).Equals("Componente", StringComparison.OrdinalIgnoreCase)
                                                     && Txt(f, 1).Equals("Total", StringComparison.OrdinalIgnoreCase));
        Assert.True(indiceComponente > 0, $"{periodo} ASE {ase}: falta Componente/Total del R2.");
        var reconocidas = EmpresasPorMapa[("R2", ase)].Select(e => Label(e, ase)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var suma = filas.Take(indiceComponente)
            .Where(f => Txt(f, 2).Equals("Total", StringComparison.OrdinalIgnoreCase) && reconocidas.Contains(Normalizar(Txt(f, 1))))
            .Sum(f => Num(f, 4));
        return Math.Abs(suma - Num(filas[indiceComponente], 4));
    }

    private static string Prefijo(string reporte) => reporte switch
    {
        "R1" => "Recaudoporcomponente",
        "R2" => "RerpoteDetalleSaldosaFavor",
        _ => "Reversi",
    };

    private static string Normalizar(string texto) => texto.Trim().ToUpperInvariant();

    private static List<Dictionary<int, object?>> LeerFilas(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var compartidas = new List<string>();
        var entrada = zip.GetEntry("xl/sharedStrings.xml");
        if (entrada is not null)
        {
            using var stream = entrada.Open();
            var doc = XDocument.Load(stream);
            foreach (var si in doc.Root!.Elements(Ns + "si"))
            {
                compartidas.Add(string.Concat(si.Descendants(Ns + "t").Select(t => t.Value)));
            }
        }

        var hoja = zip.Entries.First(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                                         && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        using var hojaStream = hoja.Open();
        var hojaDoc = XDocument.Load(hojaStream);
        var filas = new List<Dictionary<int, object?>>();
        foreach (var row in hojaDoc.Root!.Descendants(Ns + "row"))
        {
            var celdas = new Dictionary<int, object?>();
            foreach (var c in row.Elements(Ns + "c"))
            {
                var referencia = (string?)c.Attribute("r") ?? string.Empty;
                var indice = IndiceColumna(referencia);
                var tipo = (string?)c.Attribute("t");
                var v = c.Element(Ns + "v")?.Value;
                object? valor = tipo switch
                {
                    "s" when v is not null => compartidas[int.Parse(v, CultureInfo.InvariantCulture)],
                    "inlineStr" => string.Concat(c.Descendants(Ns + "t").Select(t => t.Value)),
                    _ => v is null ? null : decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var numero) ? numero : v,
                };
                if (valor is not null) celdas[indice] = valor;
            }

            filas.Add(celdas);
        }

        return filas;
    }

    private static int IndiceColumna(string referencia)
    {
        var n = 0;
        foreach (var ch in referencia)
        {
            if (ch is < 'A' or > 'Z') break;
            n = n * 26 + (ch - 'A' + 1);
        }

        return n - 1;
    }

    private static string Txt(Dictionary<int, object?> fila, int columna) => fila.GetValueOrDefault(columna) switch
    {
        null => string.Empty,
        string s => s.Trim(),
        decimal d => d.ToString(CultureInfo.InvariantCulture),
        var o => o.ToString()?.Trim() ?? string.Empty,
    };

    private static decimal Num(Dictionary<int, object?> fila, int columna) => fila.GetValueOrDefault(columna) switch
    {
        decimal d => d,
        double d => (decimal)d,
        string s when decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var r) => r,
        _ => 0m,
    };
}
