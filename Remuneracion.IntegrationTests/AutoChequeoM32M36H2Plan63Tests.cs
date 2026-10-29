using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 63 (H2 — T2.3 / T4, CA-a/CA-b/CA-c): el contrato de TEXTO-<c>&lt;f&gt;</c> de la salida de la
/// app para los auto-chequeos <c>DetRetri{per}!M32:M36</c> y la integridad del resto del workbook.
///
/// <para>
/// <b>G-WORKBOOK (R-ORACULO-STALE):</b> el caché <c>&lt;v&gt;</c> de estas celdas NO es oráculo (la
/// plantilla lo trae en 1 y viajaría stale). Por eso:
/// </para>
/// <list type="bullet">
/// <item>Q1: el texto-<c>&lt;f&gt;</c> resta <c>INTERVENTORIA!N26:N30</c> — es decir, el MISMO operando
/// que el manual del administrativo (cuyo SELL medido es VERDADERO ×5), con el cuerpo verificado
/// equivalente salvo filas de layout (T0-03). <b>La equivalencia booleana es la demostración</b>; el
/// recálculo real de la salida se ejecuta en el Gate (G.2).</item>
/// <item>Q2: el texto resta <c>M26:M30</c>, que es valor a valor <c>F15:F19</c> (T0-05) ⇒ sin impacto
/// numérico, mismos cierres que el manual Q2.</item>
/// </list>
///
/// <para>
/// <b>Integridad (CA-c):</b> la reescritura toca SOLO el texto de esas 5 fórmulas (T1.2) y las
/// fórmulas compartidas vecinas (<c>L145</c>/<c>M142</c> del CONSOLIDADO) y la hoja
/// <c>AJUSTES - SF-T</c> quedan byte-idénticas a la plantilla.
/// </para>
/// </summary>
public sealed class AutoChequeoM32M36H2Plan63Tests
{
    private const string PeriodoQ1 = "2026091";
    private const string PeriodoQ2 = "2026072";
    private const string PeriodoBasePlantilla = "2026092";
    private const string HojaConsolidado = WorkbookLeafCellMap.HojaConsolidado;
    private const string HojaAjustes = "AJUSTES - SF-T";
    private const string HojaInterventoria = "INTERVENTORIA";

    private static readonly (string Celda, int FilaCruda)[] Bloque =
    [
        ("M32", 26), ("M33", 27), ("M34", 28), ("M35", 29), ("M36", 30),
    ];

    private static readonly Dictionary<string, int> FilaRedondeada = new(StringComparer.Ordinal)
    {
        ["M32"] = 15, ["M33"] = 16, ["M34"] = 17, ["M35"] = 18, ["M36"] = 19,
    };

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T4 / CA-a — Q1: texto N26:N30 (misma referencia que el manual VERDADERO)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T4 (CA-a): en la base Q1, <c>M32:M36</c> restan <c>INTERVENTORIA!N26:N30</c> — la MISMA
    /// referencia del manual del administrativo (cuyos 5 SELL están en VERDADERO) — y el cuerpo de la
    /// fórmula es equivalente al del manual salvo las filas del layout (normalización <c>!COL#</c>).
    /// Con el operando idéntico y el cuerpo equivalente, el SELL de la app coincide con el del manual:
    /// el recálculo de la salida se ejecuta en el Gate (G.2).
    /// </summary>
    [Fact]
    public void T4_Q1_TextoN26N30_YEquivalenciaDeCuerpoConElManual()
    {
        var periodo = Periodo.Parse(PeriodoQ1);
        var hoja = NombresHojaPeriodo.DetRetri(periodo);
        var producido = HojaXml(Insumos.BaseDelPeriodoV4(PeriodoQ1), hoja);
        var manual = HojaXml(RutaManualQ1(), hoja);

        for (var i = 0; i < Bloque.Length; i++)
        {
            var (celda, fila) = Bloque[i];
            var formula = FormulaTexto(producido, celda)!;
            var formulaManual = FormulaTexto(manual, celda)!;

            // (a) Texto: la variante del run, exactamente una vez, sin la redondeada.
            Assert.Equal(1, ContarReferencia(formula, $"INTERVENTORIA!N{fila}"));
            Assert.Equal(0, ContarReferencia(formula, $"INTERVENTORIA!F{FilaRedondeada[celda]}"));

            // (b) Término AJUSTES y comparando intactos (solo cambió la referencia de interventoría).
            Assert.Equal(1, ContarReferencia(formula, $"'AJUSTES - SF-T'!M{47 + i}"));
            Assert.EndsWith($"=M{9 + i}", formula, StringComparison.Ordinal);

            // (c) Cuerpo equivalente al manual: mismo operando de interventoría y mismos términos,
            //     solo cambian las filas del layout (R-TRAMPA-ANCLA).
            Assert.Equal(NormalizarCuerpo(formulaManual, i), NormalizarCuerpo(formula, i));
            Assert.Equal(
                ContarReferencia(formulaManual, $"INTERVENTORIA!N{fila}"),
                ContarReferencia(formula, $"INTERVENTORIA!N{fila}"));
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T4 / CA-b — Q2: texto M26:M30, sin impacto numérico (M ≡ F)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T4 (CA-b): en la base Q2, <c>M32:M36</c> restan <c>INTERVENTORIA!M26:M30</c>; en Q2
    /// <c>M26:M30 ≡ F15:F19</c> valor a valor, así que el cambio es SOLO de referencia (cero impacto
    /// numérico respecto del manual Q2, que resta <c>F15:F19</c>).
    /// </summary>
    [Fact]
    public void T4_Q2_TextoM26M30_SinImpactoNumerico()
    {
        var periodo = Periodo.Parse(PeriodoQ2);
        var hoja = NombresHojaPeriodo.DetRetri(periodo);
        var producido = HojaXml(Insumos.BaseDelPeriodoV4(PeriodoQ2), hoja);
        var interv = HojaXml(Insumos.BaseDelPeriodoV4(PeriodoQ2), HojaInterventoria);

        for (var i = 0; i < Bloque.Length; i++)
        {
            var (celda, fila) = Bloque[i];
            var formula = FormulaTexto(producido, celda)!;

            Assert.Equal(1, ContarReferencia(formula, $"INTERVENTORIA!M{fila}"));
            Assert.Equal(0, ContarReferencia(formula, $"INTERVENTORIA!F{FilaRedondeada[celda]}"));
            Assert.Equal(1, ContarReferencia(formula, $"'AJUSTES - SF-T'!M{47 + i}"));
            Assert.EndsWith($"=M{9 + i}", formula, StringComparison.Ordinal);

            // M26:M30 ≡ F15:F19 ⇒ el operando nuevo vale lo mismo que el del manual Q2.
            Assert.Equal(
                ValorNumerico(interv, $"F{FilaRedondeada[celda]}"),
                ValorNumerico(interv, $"M{fila}"));
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T4 / CA-c — resto del workbook sin deltas
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T4 (CA-c, R5-GATE): el resto del workbook no se mueve. En particular:
    /// <list type="bullet">
    /// <item>la hoja <c>AJUSTES - SF-T</c> queda BYTE-idéntica a la plantilla (el término que el
    /// auto-chequeo conserva no se toca);</item>
    /// <item>los maestros de fórmula compartida <c>L145</c> (<c>si</c>) y <c>M142</c> (<c>si</c>) del
    /// CONSOLIDADO conservan <c>t</c>/<c>ref</c>/<c>si</c> y su texto.</item>
    /// </list>
    /// </summary>
    [Theory]
    [InlineData(PeriodoQ1)]
    [InlineData(PeriodoQ2)]
    public void T4_RestoDelWorkbookSinDeltas(string codigo)
    {
        var oficialAjustes = NormalizarPeriodo(HojaXml(Insumos.PlantillaOficialV4(), HojaAjustes), codigo);
        var producidoAjustes = NormalizarPeriodo(HojaXml(Insumos.BaseDelPeriodoV4(codigo), HojaAjustes), codigo);
        Assert.Equal(oficialAjustes, producidoAjustes);

        var oficialConsolidado = HojaXml(Insumos.PlantillaOficialV4(), HojaConsolidado);
        var producidoConsolidado = HojaXml(Insumos.BaseDelPeriodoV4(codigo), HojaConsolidado);
        foreach (var maestro in new[] { "L145", "M142" })
        {
            foreach (var atributo in new[] { "t", "ref", "si" })
            {
                Assert.Equal(
                    AtributoFormula(oficialConsolidado, maestro, atributo),
                    AtributoFormula(producidoConsolidado, maestro, atributo));
            }

            Assert.Equal(FormulaTexto(oficialConsolidado, maestro), FormulaTexto(producidoConsolidado, maestro));
        }

        // El comparando del auto-chequeo (M9:M13, etiqueta del ASE) sigue siendo el de la plantilla.
        var hoja = NombresHojaPeriodo.DetRetri(Periodo.Parse(codigo));
        var oficialDet = HojaXml(Insumos.PlantillaOficialV4(), NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoBasePlantilla)));
        var producidoDet = HojaXml(Insumos.BaseDelPeriodoV4(codigo), hoja);
        for (var i = 0; i < 5; i++)
        {
            Assert.Equal(CeldaCruda(oficialDet, $"M{9 + i}"), CeldaCruda(producidoDet, $"M{9 + i}"));
            Assert.False(FormulaTexto(producidoDet, $"M{47 + i}") is not null);
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T1.4 / H2 — Allow-list Phase 0: 73 → 78 celdas (la extensión es de 5 celdas exactas)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T1.4: la allow-list de Phase 0 pasa de 73 a 78 celdas con EXACTAMENTE las 5 de esta familia, y
    /// cada una existe como <c>&lt;f&gt;</c> en la plantilla oficial. Se aserta contra la fuente única
    /// de la geometría (no contra una copia del número). El inventario por nombre de hoja del grep de
    /// cierre vive en <c>FNew1Q1BlockerT0Tests.EsCeldaDeAllowList</c> (Plan 63 le añade M32:M36).
    /// </summary>
    [Fact]
    public void T1_4_AllowListFase0_CreceDe73A78ConLas5CeldasM32M36()
    {
        const int allowListPlan62 = 73;
        const int allowListConPlan63 = 78;

        var celdas = PreparadorBasePipeline.CeldasAutoChequeoDeDetRetriM;

        Assert.Equal(allowListConPlan63, allowListPlan62 + celdas.Count);
        Assert.Equal(
            new[] { "M32", "M33", "M34", "M35", "M36" },
            celdas.Select(c => c.Celda).ToArray());
        Assert.Equal(new[] { 26, 27, 28, 29, 30 }, celdas.Select(c => c.FilaInterventoria).ToArray());

        var oficial = HojaXml(
            Insumos.PlantillaOficialV4(),
            NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoBasePlantilla)));
        foreach (var (celda, _) in celdas)
        {
            Assert.False(string.IsNullOrEmpty(FormulaTexto(oficial, celda)), $"{celda} no trae <f> en la plantilla v4.");
        }
    }

    // ── Helpers ─────────────────────────────────────────────────────────────────────────────────

    private static string RutaManualQ1() => Path.Combine(
        Raiz(), "Docs", "Prueba Septiembre 1", "Resultados", "Administrativo", "Remuneracion 202609-1 Total.xlsx");

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

    private static int ContarReferencia(string formula, string referencia) =>
        Regex.Matches(formula, $"(?<![A-Za-z0-9_]){Regex.Escape(referencia)}(?![0-9])").Count;

    /// <summary>
    /// Neutraliza los tokens de período (el de la base del template y el del run): Phase 0 los
    /// reescribe en TODO el paquete por contrato (Plan 39), no solo en la hoja del auto-chequeo.
    /// </summary>
    private static string NormalizarPeriodo(string xml, string codigo) =>
        xml.Replace(PeriodoBasePlantilla, "@PERIODO@", StringComparison.Ordinal)
            .Replace(codigo, "@PERIODO@", StringComparison.Ordinal);

    /// <summary>
    /// Cuerpo comparable: quita el comparando, el término AJUSTES (el manual Q1 no lo trae y en Q1
    /// vale 0) y normaliza la referencia de interventoría + el número de fila de cada referencia.
    /// </summary>
    private static string NormalizarCuerpo(string formula, int ase)
    {
        var cuerpo = formula;
        var comparando = $"=M{9 + ase}";
        if (cuerpo.EndsWith(comparando, StringComparison.Ordinal))
        {
            cuerpo = cuerpo[..^comparando.Length];
        }

        cuerpo = cuerpo.Replace($"+'AJUSTES - SF-T'!M{47 + ase}", string.Empty, StringComparison.Ordinal);
        cuerpo = Regex.Replace(cuerpo, @"INTERVENTORIA![A-Z]+\d+", "INTERVENTORIA!#", RegexOptions.None);
        cuerpo = Regex.Replace(cuerpo, @"!(?<col>[A-Z]{1,3})\d+", "!${col}#", RegexOptions.None);
        return cuerpo;
    }

    private static decimal ValorNumerico(string xml, string referencia)
    {
        var crudo = CeldaCruda(xml, referencia) ?? throw new InvalidOperationException($"Falta {referencia}.");
        var v = Regex.Match(crudo, "<v>(?<v>[^<]*)</v>");
        Assert.True(v.Success, $"{referencia} no trae <v>.");
        return decimal.Parse(v.Groups["v"].Value, NumberStyles.Float, CultureInfo.InvariantCulture);
    }

    private static string HojaXml(string ruta, string nombreHoja)
    {
        using var zip = ZipFile.OpenRead(ruta);
        return LeerParte(zip, ParteDeHoja(LeerParte(zip, "xl/workbook.xml"), LeerParte(zip, "xl/_rels/workbook.xml.rels"), nombreHoja));
    }

    private static string ParteDeHoja(string workbook, string rels, string nombreHoja)
    {
        var sheet = Regex.Match(
            workbook,
            $"<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?sheet name=\"{Regex.Escape(nombreHoja)}\"[^>]*r:id=\"(?<rid>[^\"]+)\"");
        Assert.True(sheet.Success, $"El workbook no trae la hoja '{nombreHoja}'.");

        var rid = sheet.Groups["rid"].Value;
        foreach (Match relacion in Regex.Matches(rels, "<Relationship[^>]*>"))
        {
            var id = Regex.Match(relacion.Value, "\\bId=\"(?<v>[^\"]+)\"");
            if (!id.Success || !string.Equals(id.Groups["v"].Value, rid, StringComparison.Ordinal))
            {
                continue;
            }

            var target = Regex.Match(relacion.Value, "\\bTarget=\"(?<v>[^\"]+)\"").Groups["v"].Value;
            return target.StartsWith('/') ? target[1..] : "xl/" + target;
        }

        throw new InvalidOperationException($"No se resolvió la relación '{rid}' de '{nombreHoja}'.");
    }

    private static string LeerParte(ZipArchive zip, string parte)
    {
        var entrada = zip.GetEntry(parte) ?? throw new FileNotFoundException($"El paquete no trae '{parte}'.");
        using var stream = entrada.Open();
        using var lector = new StreamReader(stream);
        return lector.ReadToEnd();
    }

    private static string? CeldaCruda(string xml, string referencia)
    {
        var match = Regex.Match(
            xml,
            $"<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c r=\"{Regex.Escape(referencia)}\"[^>]*(?:/>|>(?<inner>.*?)</(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c>)",
            RegexOptions.Singleline);
        return match.Success ? match.Value : null;
    }

    private static string? FormulaTexto(string xml, string referencia)
    {
        var crudo = CeldaCruda(xml, referencia);
        if (crudo is null)
        {
            return null;
        }

        var match = Regex.Match(crudo, "<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?f(?<attrs>[^>]*)>(?<texto>[^<]*)</");
        return match.Success ? match.Groups["texto"].Value : null;
    }

    private static string? AtributoFormula(string xml, string referencia, string atributo)
    {
        var crudo = CeldaCruda(xml, referencia);
        if (crudo is null)
        {
            return null;
        }

        var match = Regex.Match(crudo, "<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?f(?<attrs>[^>]*)>", RegexOptions.Singleline);
        if (!match.Success)
        {
            return null;
        }

        var attr = Regex.Match(match.Groups["attrs"].Value, $"\\b{Regex.Escape(atributo)}=\"(?<v>[^\"]*)\"");
        return attr.Success ? attr.Groups["v"].Value : null;
    }
}
