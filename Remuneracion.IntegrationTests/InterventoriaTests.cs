using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-16 (Plan 16 §4 Fase 0 — Unidad 0 + Req 1/3): EVIDENCIA T0 de INTERVENTORIA (D2b) y de las
/// L-Especiales menores (D3a/D3b) contra ambos canónicos y las fuentes R1 reales. Nada HU-16
/// entra al código sin estas tablas:
/// - T0-INTERVENTORIA: bloque por ASE (filas 26..30) en VALORES con L=Id ASE; totales 31 y gran
///   total 32 en FÓRMULA; idéntico en Q1/Q2/Q2-raíz; SIN fuente en Docs/Insumos → D2b declarado.
/// - T0-L-ESPECIALES: la columna L del template espeja la columna SERVICIO ESPECIALES de la
///   fuente R1; toda celda L numérica fuera de V4/HU-08/Q2-map/mapa D3a es CERO en ambos
///   canónicos (A8 stale-guard).
/// </summary>
public sealed class InterventoriaTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    [Fact]
    public void T0_Interventoria_BloquePorAse_ValoresYFormulas_AmbosCanonicosYControl()
    {
        foreach (var (ruta, etiqueta) in new[]
                 { (Insumos.Plantilla, "Q1"), (Insumos.PlantillaQ2, "Q2CANON"), (Insumos.GoldenQ2, "Q2ROOT") })
        {
            for (var i = 0; i < 5; i++)
            {
                var aseId = i + 1;
                var fila = InterventoriaDeclarada.FilaPrimerAse + i;
                // K/M/N son VALORES (no fórmula) y cierran contra la declaración T0.
                Assert.InRange(TestHelpers.LeerCeldaNumerica(ruta, InterventoriaDeclarada.Hoja, $"K{fila}") - InterventoriaDeclarada.ValorOficialMesPorAse[aseId], -Tolerancia, Tolerancia);
                Assert.InRange(TestHelpers.LeerCeldaNumerica(ruta, InterventoriaDeclarada.Hoja, $"M{fila}") - InterventoriaDeclarada.SegundaQuincenaPorAse[aseId], -Tolerancia, Tolerancia);
                Assert.InRange(TestHelpers.LeerCeldaNumerica(ruta, InterventoriaDeclarada.Hoja, $"N{fila}") - InterventoriaDeclarada.PrimeraQuincenaPorAse[aseId], -Tolerancia, Tolerancia);
                Assert.InRange(TestHelpers.LeerCeldaNumerica(ruta, InterventoriaDeclarada.Hoja, $"L{fila}") - aseId, -Tolerancia, Tolerancia);
                Assert.False(TestHelpers.CeldaEsFormula(ruta, InterventoriaDeclarada.Hoja, $"K{fila}"), $"{etiqueta}: K{fila} debió ser VALOR estático (D2b).");
                // Partición por mitades: M+N = K ±1 (redondeo).
                var k = InterventoriaDeclarada.ValorOficialMesPorAse[aseId];
                var sumaMitades = InterventoriaDeclarada.SegundaQuincenaPorAse[aseId] + InterventoriaDeclarada.PrimeraQuincenaPorAse[aseId];
                Assert.InRange(k - sumaMitades, -1m, 1m);
            }

            // Totales en FÓRMULA (K31/M31/N31 = SUM; K32 = SUM(M31:N31)) con caché == Σ.
            Assert.True(TestHelpers.CeldaEsFormula(ruta, InterventoriaDeclarada.Hoja, "K31"), $"{etiqueta}: K31 debió ser fórmula SUM.");
            Assert.True(TestHelpers.CeldaEsFormula(ruta, InterventoriaDeclarada.Hoja, "M31"), $"{etiqueta}: M31 debió ser fórmula SUM.");
            Assert.True(TestHelpers.CeldaEsFormula(ruta, InterventoriaDeclarada.Hoja, "N31"), $"{etiqueta}: N31 debió ser fórmula SUM.");
            Assert.True(TestHelpers.CeldaEsFormula(ruta, InterventoriaDeclarada.Hoja, "K32"), $"{etiqueta}: K32 debió ser fórmula SUM(M31:N31).");
            Assert.InRange(TestHelpers.LeerCeldaNumerica(ruta, InterventoriaDeclarada.Hoja, "K31") - InterventoriaDeclarada.ValorOficialMesPorAse.Values.Sum(), -Tolerancia, Tolerancia);
        }
    }

    [Fact]
    public void T0_Interventoria_SinFuenteEnInsumos_V8()
    {
        // T0-0.3: búsqueda exhaustiva normalizada de archivos de interventoría en Docs/Insumos
        // (ambos períodos). Ninguna fuente → D2b (insumo externo declarado), sin finder.
        var raiz = Insumos.Raiz();
        var carpeta = Path.Combine(raiz, "Docs", "Insumos");
        Assert.True(Directory.Exists(carpeta), "Falta Docs/Insumos.");
        var encontrados = Directory.EnumerateFiles(carpeta, "*.*", SearchOption.AllDirectories)
            .Where(f => Normalizar(Path.GetFileNameWithoutExtension(f)).Contains("nterventoria", StringComparison.Ordinal))
            .Select(f => Path.GetFileName(f))
            .ToArray();
        Assert.Empty(encontrados);
    }

    [Fact]
    public void D3a_LEspecialesMenores_LecturaCierraContraGolden_AmbosPeriodos()
    {
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var q1 = (periodo: Insumos.Periodo(), fuente: (Func<int, string>)Insumos.R1, golden: Insumos.Plantilla, etiqueta: "Q1");
        var q2 = (periodo: Insumos.PeriodoQ2(), fuente: (Func<int, string>)Insumos.R1Q2, golden: Insumos.GoldenQ2, etiqueta: "Q2");

        foreach (var caso in new[] { q1, q2 })
        {
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var ase = Insumos.Ase(aseId);
                var inputs = reader.LeerLEspecialesMenores(ase, caso.periodo, caso.fuente(aseId));
                Assert.True(inputs.TieneCeldas, $"{caso.etiqueta} ASE{aseId}: el mapa D3a exige celdas L-menores.");
                foreach (var (celda, valor) in inputs.Celdas)
                {
                    var golden = TestHelpers.LeerCeldaNumerica(caso.golden, WorkbookLeafCellMapInterventoria.HojaR1, celda);
                    Assert.True(Math.Abs(valor - golden) <= Tolerancia,
                        $"{caso.etiqueta} ASE{aseId} {WorkbookLeafCellMapInterventoria.HojaR1}!{celda}: reader={valor} vs golden={golden} (tolerancia ±{Tolerancia}).");
                }
            }
        }
    }

    [Fact]
    public void D3b_LEspecialesMenores_TodaLNumericaFueraDeMapas_EsCero_EnAmbosCanonicos()
    {
        // A8 stale-guard: TODA celda L numérica del R1 fuera de (V4 ∪ HU-08 ∪ Q2-map ∪ mapa D3a)
        // es CERO en ambos canónicos (Q1 golden y Q2 caché golden). Si un canónico futuro trae
        // no-cero en una de estas celdas, el test FALLA nombrando ASE/celda (nunca 0 silencioso).
        foreach (var (ruta, periodo, etiqueta) in new[]
                 { (Insumos.Plantilla, 1, "Q1"), (Insumos.GoldenQ2, 2, "Q2") })
        {
            var mapeadas = CeldasLMenoresMapeadas(periodo);
            foreach (var (celda, valor) in TestHelpers.EnumerarCeldaLNumericas(ruta, WorkbookLeafCellMapInterventoria.HojaR1))
            {
                if (mapeadas.Contains(celda, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                Assert.True(Math.Abs(valor) <= Tolerancia,
                    $"{etiqueta} {WorkbookLeafCellMapInterventoria.HojaR1}!{celda} fuera del set mapeado vale {valor} (D3b/A8: debe ser 0).");
            }
        }
    }

    /// <summary>
    /// Conjunto de celdas L con estatuto mapeado para el período (V4 HU-07T0 + HU-08 + mapa Q2 +
    /// mapa D3a HU-16). Lo demás de la columna L numérica debe ser 0 (D3b, A8).
    /// </summary>
    private static HashSet<string> CeldasLMenoresMapeadas(int numeroQuincena)
    {
        var celdas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Set V4 (HU-07 T0 §0.1/§0.3): operandos L del R1 por ASE + EXTEMP explícitos.
        foreach (var celda in new[]
                 { "L25", "L10", "L113", "L90", "L217", "L238", "L223", "L357", "L391", "L369", "L478", "L498" })
        {
            celdas.Add(celda);
        }

        // HU-08 (mapa por empresa): L de las filas Total por empresa.
        foreach (var entrada in WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.Values)
        {
            foreach (var (celda, _) in entrada)
            {
                if (celda.StartsWith("L", StringComparison.OrdinalIgnoreCase))
                {
                    celdas.Add(celda);
                }
            }
        }

        if (numeroQuincena == 2)
        {
            // HU-12 (mapa Q2): L de los operandos R1-Q2.
            foreach (var entrada in WorkbookLeafCellMapQ2.R1Q2EditablesPorAse.Values)
            {
                foreach (var (celda, _) in entrada)
                {
                    if (celda.StartsWith("L", StringComparison.OrdinalIgnoreCase))
                    {
                        celdas.Add(celda);
                    }
                }
            }
        }

        foreach (var mapa in numeroQuincena == 2
                     ? WorkbookLeafCellMapInterventoria.LMenoresPorAseQ2.Values
                     : WorkbookLeafCellMapInterventoria.LMenoresPorAse.Values)
        {
            foreach (var (celda, _, _) in mapa)
            {
                celdas.Add(celda);
            }
        }

        return celdas;
    }

    private static string Normalizar(string texto)
    {
        var normalizado = texto.Normalize(System.Text.NormalizationForm.FormD);
        var sb = new System.Text.StringBuilder(normalizado.Length);
        foreach (var ch in normalizado)
        {
            if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch) != System.Globalization.UnicodeCategory.NonSpacingMark)
            {
                sb.Append(char.ToLowerInvariant(ch));
            }
        }

        return sb.ToString();
    }
}
