using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// T0 residual (2026-10-15): congela desde bytes la diferencia de <b>24.329,38</b> de OCCIDENTE ASE5 en el
/// sub-bloque <c>VALIDACION_TOTAL!O25 = (O24 = N9)</c>.
///
/// CERO código de producción: este fixture + sus asserts SON la evidencia. Lectura SOLO con BCL
/// (<c>System.IO.Compression.ZipArchive</c> + <c>System.Xml.Linq</c>): sin Excel/COM, sin OpenXml SDK.
///
/// Tesis congelada: la fila Oportuno#1 de OCCIDENTE ASE5 (fuente R1 fila 39, <c>24.329,38</c>) SÍ se coloca
/// (corrida actual <c>Reporte Componentes R1!F628 = 24.329,38</c>), pero en el OFICIAL esa celda NO tiene
/// consumidor (ninguna fórmula la referencia) mientras que en el MANUAL la fila equivalente (<c>F537</c>/
/// <c>F538</c>) SÍ es consumida por las anclas EXTEMPORANEO (<c>F580</c>/<c>F560</c>). Es una diferencia
/// ESTRUCTURAL de ancla/consolidación, no un hueco de colocación.
///
/// Ver <c>Docs/T0-Residual-24329-OccidenteAse5.md</c>.
/// </summary>
public sealed class T0ResidualOccidenteAse5Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string HojaR1 = "Reporte Componentes R1";
    private const string HojaValidacionTotal = "VALIDACION_TOTAL";

    #region DEUDA-DECLARADA Plan 49 (T0-EXTEMP-Ase5-Plan49 §4.3/§8.1)
    // La celda F682/F702 es ancla consumidora (CONSOLIDADO!D51). El assert congela el
    // defecto de la plantilla v2 (literal 0), NO un comportamiento esperado.
    // Condición de retiro: v3 de la propietaria con la fórmula restaurada + T0 nuevo.
    [Fact]
    public void Oficial_Ase5_ExtempAnchorsSonValor_NoFormula()
    {
        // F682 (CONSOLIDADO!D51) y F702 (REMUNERACION_OCCIDENTE!D51) son VALORES 0 SIN <f>.
        // Por contraste, ASE1-4 traen fórmula EXTEMPORANEO que SÍ consume las filas Oportuno.
        using var zip = AbrirOficial();
        var hoja = Hoja(zip, HojaR1);

        Assert.Null(FormulaDe(hoja, "F682"));
        Assert.Null(FormulaDe(hoja, "F702"));
        Assert.Equal("0", ValorDe(hoja, "F682"));
        Assert.Equal("0", ValorDe(hoja, "F702"));

        // Patrón ASE1-4 (fórmula EXTEMPORANEO que consume Oportuno#1+Oportuno#2).
        Assert.Equal("F58+F31-L31", FormulaDe(hoja, "F97"));
        Assert.Equal("F189+F145-L145", FormulaDe(hoja, "F276"));
        Assert.Equal("F341+F314-L314", FormulaDe(hoja, "F437"));
        Assert.Equal("F529+F485-L485", FormulaDe(hoja, "F590"));
    }
    #endregion

    [Fact]
    public void Oficial_Ase5_AnclasNoConsumenLasFilasOportuno()
    {
        // OPORTUNO usa 1.ª + Aplic#1 + Aplic#2; TOT_OPT usa las filas Mes. Ninguno usa F628/F659.
        using var zip = AbrirOficial();
        var hoja = Hoja(zip, HojaR1);

        Assert.Equal("F647+F673+F612-L612-L647", FormulaDe(hoja, "F700"));
        Assert.Equal("F674+F648+F613-L613-L648", FormulaDe(hoja, "F680"));
    }

    [Fact]
    public void Oficial_F628YF659_SonHuerfanos_EnTodoElWorkbook()
    {
        // La fila Oportuno de OCCIDENTE ASE5 (F628) y la de Oportuno#2 (F659) no son referenciadas por
        // NINGUNA fórmula de NINGUNA hoja (los followers shared no traen texto; un consumo habría
        // aparecido en su maestro).
        using var zip = AbrirOficial();

        var consumidores = new List<string>();
        foreach (var nombre in NombresHojas(zip))
        {
            var hoja = Hoja(zip, nombre);
            foreach (var (referencia, formula) in Formulas(hoja))
            {
                if (ContieneReferencia(formula, "F628") || ContieneReferencia(formula, "F659")
                    || ContieneReferencia(formula, "L628") || ContieneReferencia(formula, "L659"))
                {
                    consumidores.Add($"{nombre}!{referencia}");
                }
            }
        }

        Assert.Empty(consumidores);
    }

    [Fact]
    public void Manual_ExtempAnchorsConsumenLaFilaOportuno_YValen441849()
    {
        using var zip = ZipFile.OpenRead(Ruta(
            "Docs", "Prueba Julio-2", "Resultado", "Remuneracion 202607-2 Total Administrativo.xlsx"));
        var hoja = Hoja(zip, HojaR1);

        // F537 = Oportuno#1 (familia REMUNERACION); F538 = Oportuno#1 (familia CONSOLIDADO).
        Assert.Equal("F537+F511-L511", FormulaDe(hoja, "F580"));
        Assert.Equal("F538+F512-L512", FormulaDe(hoja, "F560"));
        Assert.Equal("24329.38", ValorDe(hoja, "F537"));
        Assert.Equal("24329.38", ValorDe(hoja, "F538"));
        Assert.Equal("441849", ValorDe(hoja, "F580"));
        Assert.Equal("441849", ValorDe(hoja, "F560"));
    }

    [Fact]
    public void Manual_O24EsIgualAN9_ElModoComunCancela()
    {
        using var zip = ZipFile.OpenRead(Ruta(
            "Docs", "Prueba Julio-2", "Resultado", "Remuneracion 202607-2 Total Administrativo.xlsx"));
        var hoja = Hoja(zip, HojaValidacionTotal);

        var o24 = decimal.Parse(ValorDe(hoja, "O24"), CultureInfo.InvariantCulture);
        var n9 = decimal.Parse(ValorDe(hoja, "N9"), CultureInfo.InvariantCulture);
        Assert.Equal(o24, n9, 2);
        Assert.Equal("1", ValorDe(hoja, "O25")); // TRUE: el 24.329,38 entra en AMBOS lados y cancela.
    }

    [Fact]
    public void AppCorridaActual_ColocaLaFilaPorSeccion()
    {
        // Corrida actual (post-Plan 41): la fila Oportuno#1 va a F628; Aplic#1 a F647; Aplic#2 a F673.
        using var zip = ZipFile.OpenRead(Ruta(
            "Docs", "Prueba Julio-2", "Resultado", "6", "Remuneración 202607-2 Total.xlsx"));
        var hoja = Hoja(zip, HojaR1);

        Assert.Equal("417519.62", ValorDe(hoja, "F612"));
        Assert.Equal("24329.38", ValorDe(hoja, "F628"));
        Assert.Equal("229202451.62", ValorDe(hoja, "F647"));
        Assert.Equal("5904652.23", ValorDe(hoja, "L647"));
        Assert.Equal("81691821.42", ValorDe(hoja, "F673"));
    }

    // ── Lectura BCL (ZIP + XML) ─────────────────────────────────────────────────────────────────

    private static bool ContieneReferencia(string formula, string referencia)
    {
        // Evita falsos positivos de celda (p. ej. F6280): exige frontera no-dígito tras la referencia.
        var indice = 0;
        while ((indice = formula.IndexOf(referencia, indice, StringComparison.OrdinalIgnoreCase)) >= 0)
        {
            var fin = indice + referencia.Length;
            if (fin >= formula.Length || !char.IsAsciiDigit(formula[fin]))
            {
                return true;
            }

            indice = fin;
        }

        return false;
    }

    private static IEnumerable<(string Referencia, string Formula)> Formulas(XDocument hoja)
    {
        foreach (var row in hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row") ?? [])
        {
            foreach (var c in row.Elements(Ns + "c"))
            {
                var f = c.Element(Ns + "f");
                if (f is not null)
                {
                    yield return (c.Attribute("r")?.Value ?? string.Empty, f.Value);
                }
            }
        }
    }

    private static ZipArchive AbrirOficial() =>
        ZipFile.OpenRead(Ruta("Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx"));

    private static string Ruta(params string[] partes)
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx")))
            {
                return Path.Combine([dir.FullName, .. partes]);
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio.");
    }

    private static List<string> NombresHojas(ZipArchive zip)
    {
        var libro = LeerXml(zip, "xl/workbook.xml");
        return libro.Root?.Element(Ns + "sheets")?.Elements(Ns + "sheet")
            .Select(s => s.Attribute("name")!.Value).ToList() ?? [];
    }

    private static XDocument Hoja(ZipArchive zip, string nombre)
    {
        var libro = LeerXml(zip, "xl/workbook.xml");
        var sheet = libro.Root?.Element(Ns + "sheets")?.Elements(Ns + "sheet")
            .FirstOrDefault(s => string.Equals(s.Attribute("name")?.Value, nombre, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"El xlsx no trae la hoja '{nombre}'.");
        var id = sheet.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))!.Value;
        var rels = LeerXml(zip, "xl/_rels/workbook.xml.rels");
        var destino = rels.Root?.Elements().FirstOrDefault(r => r.Attribute("Id")?.Value == id)?.Attribute("Target")?.Value
            ?? throw new InvalidOperationException($"Sin rel para la hoja '{nombre}'.");
        return LeerXml(zip, "xl/" + destino);
    }

    private static XDocument LeerXml(ZipArchive zip, string entrada)
    {
        var entry = zip.GetEntry(entrada) ?? throw new InvalidOperationException($"El xlsx no trae '{entrada}'.");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static XElement? Celda(XDocument hoja, string referencia)
    {
        var numero = new string(referencia.Where(char.IsAsciiDigit).ToArray());
        return hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row")
            .FirstOrDefault(r => r.Attribute("r")?.Value == numero)?
            .Elements(Ns + "c")
            .FirstOrDefault(c => string.Equals(c.Attribute("r")?.Value, referencia, StringComparison.OrdinalIgnoreCase));
    }

    private static string? FormulaDe(XDocument hoja, string referencia) =>
        Celda(hoja, referencia)?.Element(Ns + "f")?.Value;

    private static string ValorDe(XDocument hoja, string referencia) =>
        Celda(hoja, referencia)?.Element(Ns + "v")?.Value ?? string.Empty;
}
