using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 40 (T3, R-F-1 + E1-4): T0 propio de la familia F-NEW-1 y del bloqueante Q1, y contrato de
/// la reescritura Phase 0 por quincena.
///
/// <para>
/// <b>Método:</b> lectura SOLO con BCL (<see cref="ZipArchive"/> + regex; sin Excel/COM, sin
/// DocumentFormat.OpenXml, sin ExcelDataReader) de la plantilla OFICIAL (base interna 2026092).
/// El fixture congela el TEXTO <c>&lt;f&gt;</c> exacto de las 3 familias por-columna y la
/// referencia del bloque <c>'Recaudo *'</c>; la corrida de Phase 0 (producción) se contrasta
/// contra el veredicto por quincena.
/// </para>
///
/// <para>
/// <b>Regla honesta (T0 — citada, sin supuestos silenciosos):</b>
/// <list type="bullet">
/// <item><c>VALIDACION_TOTAL!C3:C8</c> y <c>VALIDACION_*!C3:C8</c>: el oficial fija la columna del
/// bloque <c>'Recaudo *'</c> (con refs <c>F</c> y <c>D</c> mezcladas); el recaudo se escribe en la
/// columna <b>D</b> en Q1 y <b>F</b> en Q2. Sin normalizar, <c>C9/D24 = 0</c> en Q1 ⇒ <c>C15</c> y
/// <c>D25</c> FALSE (bloqueante E1-4). <b>Reescritura Phase 0 por quincena.</b></item>
/// <item><c>DetRetri{per}!L32:L36</c>, <c>DetRetri{per}!O32:O36</c>, <c>DetValiRetri{per}!I24:I29</c>:
/// hardcodean <c>INTERVENTORIA!M26:M31</c> (Segunda). R-INTERV-Q conmuta <c>L9:L13</c> → N en Q1;
/// el chequeo <c>L32</c> evalúa FALSE por ±1 en Q1 (medido). <b>Reescritura Phase 0 por quincena.</b></item>
/// <item><c>INTERVENTORIA!S26:S31</c> = <c>Q{26..31}=R{26..31}</c>: NO referencia
/// <c>INTERVENTORIA!M/N</c>; <c>Q</c> = <c>CONSOLIDADO!L{104..109}</c> (sigue la selección L) y
/// <c>R</c> = <c>J15:J20</c> (la tabla anual = los valores de la 2.ª quincena). <b>Plan 41 (T3): NO
/// es invariante</b> — en Q1 <c>Q26 = N26 = 189185987 ≠ R26 = J15 = 189185988</c> ⇒ <c>S26</c> FALSE
/// real. Phase 0 reescribe el maestro <c>R26</c> por quincena (Q2 → <c>J15</c> identidad; Q1 →
/// <c>N26</c>) para que el auto-chequeo sea honesto (D-G); el sellado cosmético sigue prohibido.</item>
/// </list>
/// </para>
/// </summary>
public sealed class FNew1Q1BlockerT0Tests
{
    private const string PartWorkbook = "xl/workbook.xml";
    private static readonly Regex CeldaRegex = new(
        "<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c r=\"(?<ref>[A-Z]+\\d+)\"(?<attrs>[^>]*?)(?:/>|>(?<inner>.*?)</(?:[A-Za-z_][A-Za-z0-9_]*:)?c>)",
        RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex FormulaRegex = new(
        "<(?:[A-Za-z_][A-Za-z0-9_]*:)?f(?<attrs>[^>]*?)>(?<texto>[^<]*)</(?:[A-Za-z_][A-Za-z0-9_]*:)?f>",
        RegexOptions.Singleline | RegexOptions.Compiled);

    private static string PlantillaOficial() => Insumos.PlantillaOficial();

    // ── T0: congelamiento del texto-<f> EXACTO del oficial (base 2026092) ─────────────────────────

    [Fact]
    public void T3_T0_CongelaTextosOficial()
    {
        using var zip = ZipFile.OpenRead(PlantillaOficial());

        // (a) VALIDACION_TOTAL!C3:C8 = Σ 'Recaudo *'!{col}{21..26} (5 refs por celda).
        var hojaTotal = LeerHoja(zip, "VALIDACION_TOTAL");
        for (var k = 3; k <= 8; k++)
        {
            var terminos = SepararSuma(FormulaDe(hojaTotal, $"C{k}"));
            Assert.Equal(5, terminos.Count);
            foreach (var termino in terminos)
            {
                Assert.Matches("^'(Recaudo [^']+)'!(F|D)" + (18 + k) + "$", termino);
            }
        }

        // (b) DetRetri2026092!L32:L36 = INTERVENTORIA!M{26..30}+L{23..27}=L{9..13}
        var hojaDet = LeerHoja(zip, "DetRetri2026092");
        for (var i = 0; i < 5; i++)
        {
            Assert.Equal($"INTERVENTORIA!M{26 + i}+L{23 + i}=L{9 + i}", FormulaDe(hojaDet, $"L{32 + i}"));
        }

        // (c) DetRetri2026092!O32:O36 contiene -INTERVENTORIA!M{26..30}
        for (var i = 0; i < 5; i++)
        {
            Assert.Contains($"-INTERVENTORIA!M{26 + i}", FormulaDe(hojaDet, $"O{32 + i}"), StringComparison.Ordinal);
        }

        // (d) DetValiRetri2026092!I24:I29 = -INTERVENTORIA!M{26..31}+I{16..21}=I{9..14}
        var hojaVali = LeerHoja(zip, "DetValiRetri2026092");
        for (var i = 0; i < 6; i++)
        {
            Assert.Equal($"-INTERVENTORIA!M{26 + i}+I{16 + i}=I{9 + i}", FormulaDe(hojaVali, $"I{24 + i}"));
        }

        // (e) INTERVENTORIA!S26:S31 = Q{26..31}=R{26..31}; el maestro R26 = J15 (la tabla anual =
        //     la 2.ª quincena). S26/S27 traen texto explícito; S28:S31 son followers de la fórmula
        //     compartida si=2 (maestro S28, ref=S28:S31): no traen texto propio.
        var hojaInterv = LeerHoja(zip, "INTERVENTORIA");
        Assert.Equal("Q26=R26", FormulaDe(hojaInterv, "S26"));
        Assert.Equal("Q27=R27", FormulaDe(hojaInterv, "S27"));
        Assert.Equal("Q28=R28", FormulaDe(hojaInterv, "S28"));
        for (var i = 3; i < 6; i++)
        {
            Assert.Equal(string.Empty, FormulaDe(hojaInterv, $"S{26 + i}"));
        }

        // (f) Plan 41 (T3): el maestro de la tabla J (R26, shared ref R26:R31) es J15 en el oficial
        //     (= la variante Q2). Q26:Q31 = CONSOLIDADO!L104:L109 (sigue la selección L9:L13 por
        //     quincena) ⇒ en Q1 el chequeo S26 sería FALSE real; Phase 0 lo reescribe por quincena.
        Assert.Equal("J15", FormulaDe(hojaInterv, "R26"));
    }

    // ── Phase 0: variante por quincena (dominio) ──────────────────────────────────────────────────

    [Theory]
    [InlineData("2026071", "D", "N")]
    [InlineData("2026072", "F", "M")]
    public void T3_Phase0_ReescribeLaVarianteDeLaQuincena(string codigo, string columnaRecaudo, string columnaInterventoria)
    {
        var periodo = Periodo.Parse(codigo);
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "p40-t3-t0-" + Guid.NewGuid().ToString("N"))).FullName;
        var salida = Path.Combine(dir, periodo.NombreArchivo);
        try
        {
            PreparadorBasePipeline.Ejecutar(PlantillaOficial(), periodo, salida);

            using var zip = ZipFile.OpenRead(salida);

            // (a) VALIDACION_TOTAL!C3:C8 y VALIDACION_*!C3:C8 → columna del run.
            foreach (var hoja in new[] { "VALIDACION_TOTAL" }.Concat(EmpresaFacturacion.Catalogo.Select(e => e.HojaValidacion)))
            {
                var xml = LeerHoja(zip, hoja);
                for (var k = 3; k <= 8; k++)
                {
                    foreach (var termino in PreparadorBasePipeline.SepararTerminosSuma(FormulaDe(xml, $"C{k}")))
                    {
                        Assert.Matches($"^'Recaudo [^']+'!{columnaRecaudo}{18 + k}$", termino);
                    }
                }
            }

            // (b) Auto-chequeos F-NEW-1 → columna INTERVENTORIA del run.
            var det = LeerHoja(zip, NombresHojaPeriodo.DetRetri(periodo));
            for (var i = 0; i < 5; i++)
            {
                Assert.Equal($"INTERVENTORIA!{columnaInterventoria}{26 + i}+L{23 + i}=L{9 + i}", FormulaDe(det, $"L{32 + i}"));
                Assert.Contains($"-INTERVENTORIA!{columnaInterventoria}{26 + i}", FormulaDe(det, $"O{32 + i}"), StringComparison.Ordinal);
            }

            var vali = LeerHoja(zip, NombresHojaPeriodo.DetValiRetri(periodo));
            for (var i = 0; i < 6; i++)
            {
                Assert.Equal($"-INTERVENTORIA!{columnaInterventoria}{26 + i}+I{16 + i}=I{9 + i}", FormulaDe(vali, $"I{24 + i}"));
            }

            // (c) Plan 41 (T3): INTERVENTORIA!S26:S31 (Q=R) es invariante en su FORMA, pero el maestro
            //     R26 referencia la tabla J de la quincena del run: Q2 → J15 (identidad); Q1 → N26.
            var interv = LeerHoja(zip, "INTERVENTORIA");
            Assert.Equal("Q26=R26", FormulaDe(interv, "S26"));
            Assert.Equal("Q27=R27", FormulaDe(interv, "S27"));
            Assert.Equal("Q28=R28", FormulaDe(interv, "S28"));
            for (var i = 3; i < 6; i++)
            {
                Assert.Equal(string.Empty, FormulaDe(interv, $"S{26 + i}"));
            }

            var esperadoR26 = periodo.NumeroQuincena == 1 ? "N26" : "J15";
            Assert.Equal(esperadoR26, FormulaDe(interv, "R26"));
        }
        finally
        {
            try
            {
                Directory.Delete(dir, recursive: true);
            }
            catch
            {
                // best-effort
            }
        }
    }

    // ── Allow-list: Phase 0 solo toca las celdas declaradas (diff completo por índice de hoja) ──

    [Theory]
    [InlineData("2026071")]
    [InlineData("2026072")]
    public void T3_Phase0_AllowList_SoloCeldasDeclaradas(string codigo)
    {
        var periodo = Periodo.Parse(codigo);
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "p40-t3-al-" + Guid.NewGuid().ToString("N"))).FullName;
        var salida = Path.Combine(dir, periodo.NombreArchivo);
        try
        {
            PreparadorBasePipeline.Ejecutar(PlantillaOficial(), periodo, salida);

            var oficial = MapaFormulas(PlantillaOficial());
            var producido = MapaFormulas(salida);
            Assert.Equal(oficial.Count, producido.Count);

            var noPermitidas = new List<string>();
            for (var i = 0; i < oficial.Count; i++)
            {
                foreach (var (celda, texto) in oficial[i].Celdas)
                {
                    var nuevo = producido[i].Celdas.GetValueOrDefault(celda);
                    if (!string.Equals(texto, nuevo, StringComparison.Ordinal) && !EsCeldaDeAllowList(oficial[i].Hoja, celda))
                    {
                        noPermitidas.Add($"{oficial[i].Hoja}!{celda}");
                    }
                }
            }

            Assert.True(noPermitidas.Count == 0,
                $"Phase 0 tocó <f> fuera de la allow-list: {string.Join(", ", noPermitidas)}");
        }
        finally
        {
            try
            {
                Directory.Delete(dir, recursive: true);
            }
            catch
            {
                // best-effort
            }
        }
    }

    /// <summary>Allow-list de Phase 0 (Plan 39 §4 + Plan 40 T3): por NOMBRE de hoja de la base.</summary>
    private static bool EsCeldaDeAllowList(string hoja, string celda)
    {
        if (string.Equals(hoja, "CONSOLIDADO_TOTAL RECAUDO", StringComparison.Ordinal))
        {
            return celda is "L9" or "L10" or "L11" or "L12" or "L13";
        }

        if (string.Equals(hoja, "BCE SC POR FACT.", StringComparison.Ordinal))
        {
            return celda is "H3" or "H4" or "H5" or "H6" or "H7" or "H10";
        }

        if (string.Equals(hoja, "VALIDACION_TOTAL", StringComparison.Ordinal) || hoja.StartsWith("VALIDACION_", StringComparison.Ordinal))
        {
            return celda is "C3" or "C4" or "C5" or "C6" or "C7" or "C8";
        }

        if (hoja.StartsWith("DetRetri", StringComparison.Ordinal))
        {
            return celda is "L32" or "L33" or "L34" or "L35" or "L36"
                or "O32" or "O33" or "O34" or "O35" or "O36";
        }

        if (hoja.StartsWith("DetValiRetri", StringComparison.Ordinal))
        {
            return celda is "K24" or "K25" or "K26" or "K27" or "K28"
                or "I24" or "I25" or "I26" or "I27" or "I28" or "I29";
        }

        // Plan 41 (T3): el maestro del auto-chequeo S26:S31 (tabla J) se reescribe por quincena.
        if (string.Equals(hoja, "INTERVENTORIA", StringComparison.Ordinal))
        {
            return celda is "R26";
        }

        return false;
    }

    // ── BCL helpers ──────────────────────────────────────────────────────────────────────────────

    private static List<(string Hoja, Dictionary<string, string> Celdas)> MapaFormulas(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var workbook = LeerParte(zip, PartWorkbook);
        var rels = LeerParte(zip, "xl/_rels/workbook.xml.rels");
        var targets = Regex.Matches(rels, "<Relationship[^>]*Id=\"(?<id>[^\"]+)\"[^>]*Target=\"(?<t>[^\"]+)\"")
            .ToDictionary(m => m.Groups["id"].Value, m => m.Groups["t"].Value);

        var lista = new List<(string, Dictionary<string, string>)>();
        foreach (Match hoja in Regex.Matches(workbook, "<(?:[A-Za-z_][A-Za-z0-9_]*:)?sheet name=\"(?<n>[^\"]*)\"[^>]*r:id=\"(?<r>[^\"]+)\""))
        {
            var rid = hoja.Groups["r"].Value;
            var target = targets[rid];
            var parte = target.StartsWith('/') ? target[1..] : "xl/" + target;
            var xml = LeerParte(zip, parte);
            var mapa = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (Match celda in CeldaRegex.Matches(xml))
            {
                if (!celda.Groups["inner"].Success)
                {
                    continue;
                }

                var f = FormulaRegex.Match(celda.Groups["inner"].Value);
                if (f.Success)
                {
                    mapa[celda.Groups["ref"].Value] = f.Groups["texto"].Value;
                }
            }

            lista.Add((hoja.Groups["n"].Value, mapa));
        }

        return lista;
    }

    /// <summary>Parte una suma por '+' respetando los nombres de hoja entrecomillados (BCL puro).</summary>
    private static List<string> SepararSuma(string formula)
    {
        var terminos = new List<string>();
        var enComillas = false;
        var inicio = 0;
        for (var i = 0; i < formula.Length; i++)
        {
            if (formula[i] == '\'')
            {
                enComillas = !enComillas;
            }
            else if (formula[i] == '+' && !enComillas)
            {
                terminos.Add(formula[inicio..i]);
                inicio = i + 1;
            }
        }

        terminos.Add(formula[inicio..]);
        return terminos;
    }

    private static string LeerHoja(ZipArchive zip, string nombreHoja)
    {
        var workbook = LeerParte(zip, PartWorkbook);
        var rels = LeerParte(zip, "xl/_rels/workbook.xml.rels");
        var sheet = Regex.Match(workbook, $"<(?:[A-Za-z_][A-Za-z0-9_]*:)?sheet name=\"{Regex.Escape(nombreHoja)}\"[^>]*r:id=\"(?<r>[^\"]+)\"");
        Assert.True(sheet.Success, $"La plantilla no trae la hoja '{nombreHoja}'.");
        var rid = sheet.Groups["r"].Value;
        var rel = Regex.Match(rels, $"<Relationship[^>]*Id=\"{Regex.Escape(rid)}\"[^>]*Target=\"(?<t>[^\"]+)\"");
        Assert.True(rel.Success, $"No se resolvió '{nombreHoja}'.");
        var target = rel.Groups["t"].Value;
        return LeerParte(zip, target.StartsWith('/') ? target[1..] : "xl/" + target);
    }

    private static string LeerParte(ZipArchive zip, string nombre)
    {
        var entrada = zip.GetEntry(nombre) ?? throw new InvalidOperationException($"Falta {nombre} en el zip.");
        using var stream = entrada.Open();
        using var lector = new StreamReader(stream, Encoding.UTF8);
        return lector.ReadToEnd();
    }

    private static string FormulaDe(string xml, string celda)
    {
        foreach (Match match in CeldaRegex.Matches(xml))
        {
            if (!string.Equals(match.Groups["ref"].Value, celda, StringComparison.Ordinal))
            {
                continue;
            }

            if (!match.Groups["inner"].Success)
            {
                return string.Empty;
            }

            var f = FormulaRegex.Match(match.Groups["inner"].Value);
            return f.Success ? f.Groups["texto"].Value : string.Empty;
        }

        return string.Empty;
    }
}
