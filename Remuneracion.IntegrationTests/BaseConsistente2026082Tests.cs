using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 30 (T3, R-B-1/S4): consistencia de la base canónica de agosto
/// <c>Docs/Prueba2/Plantilla_Remuneracion_2026082.xlsx</c>, producida una sola vez (offline) por la
/// herramienta <c>Herramientas/PreparadorBasePeriodo</c>.
///
/// Verificación estructural ZIP+XML (mismo método del T0 del Plan 30): hojas nuevas, 0 apariciones
/// del token viejo, 11 &lt;f&gt; con el token nuevo, definedNames limpios y metadatos de período
/// idénticos al manual. NO se comparan valores de corrida (difieren del manual por construcción).
/// </summary>
public sealed class BaseConsistente2026082Tests
{
    private const string TokenViejo = "2026072";
    private const string TokenNuevo = "2026082";
    private const string InformeViejo = "Informe AFaseo Recaudo 202607-2";
    private const string InformeNuevo = "Informe AFaseo Recaudo 202608-2";
    private const string HojaConsolidado = "xl/worksheets/sheet18.xml";
    private const string HojaDetRetri = "xl/worksheets/sheet36.xml";

    private static readonly string[] HojasRecaudoQ1 =
    [
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
        "xl/worksheets/sheet3.xml",
        "xl/worksheets/sheet5.xml"
    ];

    private static readonly Regex CeldaRegex = new(
        "<c r=\"(?<ref>[A-Z]+\\d+)\"(?<attrs>[^>]*?)(?:/>|>(?<inner>.*?)</c>)",
        RegexOptions.Singleline | RegexOptions.Compiled);

    private static readonly Regex ValorRegex = new("<v>(?<v>[^<]*)</v>", RegexOptions.Compiled);
    private static readonly Regex SiRegex = new("<si>(?<cuerpo>.*?)</si>", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex TextoRegex = new("<t[^>]*>(?<t>[^<]*)</t>", RegexOptions.Singleline | RegexOptions.Compiled);

    [Fact]
    public void BaseAgosto2026082_EsInternamenteConsistenteConElManual()
    {
        var baseRuta = Insumos.PlantillaAgosto2026082;
        var manualRuta = Insumos.ManualAgosto2026082;
        Assert.True(File.Exists(baseRuta), $"Falta la base canónica: {baseRuta}");
        Assert.True(File.Exists(manualRuta), $"Falta el manual de agosto: {manualRuta}");

        // (1) Hojas nuevas presentes y viejas ausentes (workbook.xml + app.xml).
        var workbook = LeerParte(baseRuta, "xl/workbook.xml");
        foreach (var nombre in new[] { "DetRetri2026082", "DetValiRetri2026082", InformeNuevo })
        {
            Assert.Contains(nombre, workbook, StringComparison.Ordinal);
        }

        Assert.DoesNotContain("DetRetri2026072", workbook, StringComparison.Ordinal);
        Assert.DoesNotContain("DetValiRetri2026072", workbook, StringComparison.Ordinal);
        Assert.DoesNotContain(InformeViejo, workbook, StringComparison.Ordinal);

        var app = LeerParte(baseRuta, "docProps/app.xml");
        foreach (var nombre in new[] { "DetRetri2026082", "DetValiRetri2026082", InformeNuevo })
        {
            Assert.Contains(nombre, app, StringComparison.Ordinal);
        }

        // (2) 0 apariciones del token viejo en worksheets + workbook + app.
        Assert.DoesNotContain(TokenViejo, workbook, StringComparison.Ordinal);
        Assert.DoesNotContain(TokenViejo, app, StringComparison.Ordinal);
        foreach (var (_, xml) in PartesWorksheets(baseRuta))
        {
            Assert.DoesNotContain(TokenViejo, xml, StringComparison.Ordinal);
            Assert.DoesNotContain(InformeViejo, xml, StringComparison.Ordinal);
        }

        // (3) Exactamente 11 <f> con el token nuevo y 0 con el viejo.
        var formulasNuevas = 0;
        var formulasViejas = 0;
        foreach (var (_, xml) in PartesWorksheets(baseRuta))
        {
            formulasNuevas += ContarFormulasCon(xml, $"DetRetri{TokenNuevo}");
            formulasViejas += ContarFormulasCon(xml, $"DetRetri{TokenViejo}");
        }

        Assert.Equal(11, formulasNuevas);
        Assert.Equal(0, formulasViejas);

        // (4) definedNames limpios.
        var definedNames = Regex.Match(workbook, "<definedNames>.*?</definedNames>", RegexOptions.Singleline);
        Assert.False(definedNames.Success && definedNames.Value.Contains(TokenViejo, StringComparison.Ordinal),
            "definedNames conserva el token viejo.");

        // (5) Metadatos de período == manual (N3/D6/G7; DetRetri G7/J7).
        AssertMetadato(baseRuta, manualRuta, HojaConsolidado, "N3", esTexto: false);
        AssertMetadato(baseRuta, manualRuta, HojaConsolidado, "D6", esTexto: false);
        AssertMetadato(baseRuta, manualRuta, HojaConsolidado, "G7", esTexto: false);
        AssertMetadato(baseRuta, manualRuta, HojaDetRetri, "G7", esTexto: true);
        AssertMetadato(baseRuta, manualRuta, HojaDetRetri, "J7", esTexto: true);

        // (6) Columnas Q1 (D:E): toda celda con dato en el manual y sin fórmula en la base coincide.
        AssertColumnasQ1(baseRuta, manualRuta);
    }

    [Fact]
    public void BaseAgosto2026082_MetadatosCanonicosEsperados()
    {
        var baseRuta = Insumos.PlantillaAgosto2026082;
        Assert.True(File.Exists(baseRuta), $"Falta la base canónica: {baseRuta}");

        var consolidado = ParsearCeldas(LeerParte(baseRuta, HojaConsolidado), LeerSharedStrings(baseRuta));
        Assert.Equal("2026082", consolidado["N3"].Valor);
        Assert.Equal("46267", consolidado["D6"].Valor);
        Assert.Equal("46250", consolidado["G7"].Valor);
        Assert.Equal("46265", consolidado["K7"].Valor);

        var detRetri = ParsearCeldas(LeerParte(baseRuta, HojaDetRetri), LeerSharedStrings(baseRuta));
        Assert.Equal("16/08/2026", detRetri["G7"].Texto);
        Assert.Equal("31/08/2026", detRetri["J7"].Texto);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────────────────

    private sealed record InfoCelda(bool TieneFormula, bool EsTexto, string? Valor, string? Texto);

    private static void AssertMetadato(string baseRuta, string manualRuta, string parte, string referencia, bool esTexto)
    {
        var bases = ParsearCeldas(LeerParte(baseRuta, parte), LeerSharedStrings(baseRuta));
        var manuales = ParsearCeldas(LeerParte(manualRuta, parte), LeerSharedStrings(manualRuta));
        Assert.True(bases.ContainsKey(referencia), $"La base no tiene {parte}!{referencia}.");
        Assert.True(manuales.ContainsKey(referencia), $"El manual no tiene {parte}!{referencia}.");

        var valorBase = esTexto ? bases[referencia].Texto : bases[referencia].Valor;
        var valorManual = esTexto ? manuales[referencia].Texto : manuales[referencia].Valor;
        Assert.Equal(valorManual, valorBase);
    }

    private static void AssertColumnasQ1(string baseRuta, string manualRuta)
    {
        foreach (var parte in HojasRecaudoQ1)
        {
            var bases = ParsearCeldas(LeerParte(baseRuta, parte), LeerSharedStrings(baseRuta));
            var manuales = ParsearCeldas(LeerParte(manualRuta, parte), LeerSharedStrings(manualRuta));
            foreach (var (referencia, manual) in manuales)
            {
                if (referencia.Length < 2
                    || (referencia[0] != 'D' && referencia[0] != 'E')
                    || manual.TieneFormula
                    || manual.Valor is null
                    || !bases.TryGetValue(referencia, out var celdaBase)
                    || celdaBase.TieneFormula)
                {
                    continue;
                }

                Assert.Equal(manual.Valor, celdaBase.Valor);
            }
        }
    }

    private static int ContarFormulasCon(string xml, string token)
    {
        var total = 0;
        foreach (Match match in Regex.Matches(xml, "<f(?<attrs>[^>]*?)>(?<texto>[^<]*)</f>", RegexOptions.Singleline))
        {
            if (match.Groups["texto"].Value.Contains(token, StringComparison.Ordinal))
            {
                total++;
            }
        }

        return total;
    }

    private static Dictionary<string, InfoCelda> ParsearCeldas(string xml, IReadOnlyList<string> shared)
    {
        var celdas = new Dictionary<string, InfoCelda>(StringComparer.Ordinal);
        foreach (Match match in CeldaRegex.Matches(xml))
        {
            var referencia = match.Groups["ref"].Value;
            var attrs = match.Groups["attrs"].Value;
            var inner = match.Groups["inner"].Success ? match.Groups["inner"].Value : null;
            var tieneFormula = inner is not null && inner.Contains("<f", StringComparison.Ordinal);
            var esShared = attrs.Contains("t=\"s\"", StringComparison.Ordinal);
            var esTexto = esShared
                || attrs.Contains("t=\"inlineStr\"", StringComparison.Ordinal)
                || attrs.Contains("t=\"str\"", StringComparison.Ordinal);

            string? valor = null;
            string? texto = null;
            if (inner is not null)
            {
                var valorMatch = ValorRegex.Match(inner);
                if (valorMatch.Success)
                {
                    valor = valorMatch.Groups["v"].Value;
                }

                if (esShared && int.TryParse(valor, NumberStyles.None, CultureInfo.InvariantCulture, out var indice)
                    && indice >= 0 && indice < shared.Count)
                {
                    texto = shared[indice];
                }
                else if (attrs.Contains("t=\"inlineStr\"", StringComparison.Ordinal))
                {
                    texto = string.Concat(TextoRegex.Matches(inner).Select(m => m.Groups["t"].Value));
                }
            }

            texto ??= valor;
            celdas[referencia] = new InfoCelda(tieneFormula, esTexto, valor, texto);
        }

        return celdas;
    }

    private static IReadOnlyList<string> LeerSharedStrings(string ruta)
    {
        var xml = LeerParte(ruta, "xl/sharedStrings.xml");
        var lista = new List<string>();
        foreach (Match match in SiRegex.Matches(xml))
        {
            lista.Add(string.Concat(TextoRegex.Matches(match.Groups["cuerpo"].Value).Select(m => m.Groups["t"].Value)));
        }

        return lista;
    }

    private static string LeerParte(string ruta, string entrada)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var entradaZip = zip.GetEntry(entrada)
            ?? throw new InvalidOperationException($"Falta {entrada} en {Path.GetFileName(ruta)}.");
        using var lector = new StreamReader(entradaZip.Open(), Encoding.UTF8);
        return lector.ReadToEnd();
    }

    private static IReadOnlyList<(string Nombre, string Xml)> PartesWorksheets(string ruta)
    {
        var lista = new List<(string, string)>();
        using var zip = ZipFile.OpenRead(ruta);
        foreach (var entrada in zip.Entries)
        {
            if (!Regex.IsMatch(entrada.FullName, "^xl/worksheets/sheet\\d+\\.xml$"))
            {
                continue;
            }

            using var lector = new StreamReader(entrada.Open(), Encoding.UTF8);
            lista.Add((entrada.FullName, lector.ReadToEnd()));
        }

        lista.Sort(static (a, b) => string.CompareOrdinal(a.Item1, b.Item1));
        return lista;
    }
}
