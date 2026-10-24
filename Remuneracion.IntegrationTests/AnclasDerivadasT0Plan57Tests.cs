using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 57 (H0/T0, S1.1-S1.4 — D1): contrato por BYTES de las anclas EXTEMPORANEO que el oráculo de
/// validaciones cruzadas debe derivar. CERO código de producción: este fixture y sus asserts SON la
/// evidencia que habilita H1.
///
/// Lectura SOLO con BCL (<see cref="ZipArchive"/> + <see cref="XDocument"/>): sin Excel/COM, sin
/// DocumentFormat.OpenXml, sin ExcelDataReader, sin <c>Insumos.Raiz()</c>. Rutas EXPLÍCITAS a la
/// plantilla v2 y a los tres períodos reales (2026071 / 2026072 / 2026101). Cero escritura.
///
/// <para>
/// <b>Lo que el T0 mide (y por qué importa).</b> <c>AnclasExtempR1</c> (Plan 49) congeló
/// <c>F682</c>/<c>F702</c> como LITERALES; los bytes de la v2 vigente traen
/// <c>F682 = &lt;f&gt;F660+F629-L629&lt;/f&gt;</c>. El inventario de abajo prueba, por bloque y por
/// bytes, que (a) la celda de publicación del CONSOLIDADO declara UNA referencia a
/// <c>Reporte Componentes R1</c> → ésa es el ancla (posición LEÍDA, §2-D3.1); (b) la naturaleza del
/// ancla es <c>&lt;f&gt;</c> en los 5 bloques (patrón homogéneo); (c) la fila de publicación de cada
/// bloque declara además las anclas por empresa que consumen las hojas <c>REMUNERACION_*</c>
/// (§2-D3.2), también por referencia leída; y (d) el cierre aritmético contra el R10 de los tres
/// períodos, al peso.
/// </para>
///
/// <para>
/// El T0 NO congela la expresión de las fórmulas (eso sería re-congelar la foto que este plan
/// retira): verifica que el ancla ES fórmula y que su texto sólo usa referencias y signos, que es
/// exactamente lo que el resolutor acotado (<c>ResolutorValorValidacion</c>) sabe evaluar.
/// </para>
/// </summary>
public sealed class AnclasDerivadasT0Plan57Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string ShaPlantillaV2 = "0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB";
    private const string ShaPlantillaV1 = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";

    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";
    private const string HojaR1 = "Reporte Componentes R1";

    /// <summary>Primera fila de la tabla de publicación del EXTEMP en el CONSOLIDADO (D47..D51).</summary>
    private const int PrimeraFilaPublicacionExtemp = 47;

    /// <summary>
    /// Fórmula admitida por el resolutor acotado: referencias de celda (relativas/absolutas) unidas
    /// por <c>+</c>/<c>-</c>, sin funciones, sin paréntesis, sin comas, sin literales de texto.
    /// </summary>
    private static readonly Regex FormulaDeReferencias = new(@"^[-+]?\$?[A-Z]{1,3}\$?\d+([-+]\$?[A-Z]{1,3}\$?\d+)*$", RegexOptions.Compiled);

    /// <summary>Celda única referenciada con hoja explícita (<c>'Reporte Componentes R1'!F682</c>).</summary>
    private static readonly Regex ReferenciaUnicaConHoja = new(@"^'?(?<hoja>[^'!]+)'?!(?<celda>\$?[A-Z]{1,3}\$?\d+)$", RegexOptions.Compiled);

    /// <summary>
    /// DetRetri (columna D de la hoja <c>DetRetri{período}</c> del R10) del plan §3 CA-A3 — el cierre
    /// que el oráculo NO puede mover (R-DETRI).
    /// </summary>
    private static readonly Dictionary<string, decimal[]> DetRetriR10PorPeriodo = new()
    {
        ["2026071"] = [16758167585m, 12227292053m, 10127195599m, 10553776376m, 8543663209m],
        ["2026072"] = [17450228673m, 20516143970m, 15221896467m, 7179595396m, 12073344662m],
        ["2026101"] = [14842425136m, 11302637605m, 8823787300m, 10483172762m, 7550668665m]
    };

    // ════════════════════════════════════════════════════════════════════════════════════════
    // S1.1 — Inventario + naturaleza real en la v2, en los CINCO bloques
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void T0_1_AnclaExtempDelBloque_EsFormulaEnLosCincoBloques()
    {
        var libro = CargarLibro(PlantillaV2);
        var anclas = AnclasDelBloque(libro);

        Assert.Equal(5, anclas.Count);
        Assert.Equal(HojaR1, anclas[0].Hoja);
        Assert.Equal("F77", anclas[0].Celda);
        Assert.Equal("F256", anclas[1].Celda);
        Assert.Equal("F417", anclas[2].Celda);
        Assert.Equal("F570", anclas[3].Celda);
        Assert.Equal("F682", anclas[4].Celda);

        foreach (var ancla in anclas)
        {
            Assert.True(
                ancla.Formula is not null,
                $"El ancla '{ancla.Hoja}!{ancla.Celda}' (EXTEMP del bloque) debe ser fórmula en la v2; "
                + "un literal aquí es la foto que el Plan 57 retira, no el contrato.");

            AssertFormulaDeReferencias(ancla.Formula!, (ancla.Hoja, ancla.Celda));
        }
    }

    [Fact]
    public void T0_2_AnclasPorEmpresaDeLaFilaDePublicacion_SeDeclaranPorReferencia()
    {
        var libro = CargarLibro(PlantillaV2);

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var anclas = AnclasDelBloque(libro);

            // El bloque del ASE publica su EXTEMP en D{46+aseId} del CONSOLIDADO y en D{46+aseId} de
            // cada hoja REMUNERACION_* (6 destinos distintos en la v2: consolidado + 5 empresas).
            var delBloque = AnclasPublicadasEnLaFila(libro, PrimeraFilaPublicacionExtemp + aseId - 1);
            Assert.Equal(6, delBloque.Count);
            Assert.Contains(anclas[aseId - 1].Celda, delBloque);

            foreach (var celda in delBloque)
            {
                Assert.True(
                    libro.Hojas.ContainsKey(HojaR1),
                    $"La referencia declarada apunta a la hoja inexistente '{HojaR1}' (bloque {aseId}).");
                Assert.True(
                    libro.Hojas[HojaR1].ContainsKey(celda),
                    $"La referencia declarada apunta a la celda inexistente '{HojaR1}!{celda}' (bloque {aseId}).");

                var destino = libro.Hojas[HojaR1][celda];
                if (destino.Formula is { } formula)
                {
                    AssertFormulaDeReferencias(formula, (HojaR1, celda));
                }
                else
                {
                    // Ancla literal: el oráculo la LEE (no hay descomposición declarada que evaluar).
                    Assert.True(
                        destino.Valor is not null,
                        $"El ancla literal '{HojaR1}!{celda}' (bloque {aseId}) no trae ni fórmula ni valor.");
                }
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // S1.2 — Referencia declarante por bloque (via §2-D3.1: seguir la referencia de la plantilla)
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void T0_3_LaCeldaDePublicacionDelConsolidado_DeclaraUnaSolaReferenciaALaHojaR1()
    {
        var libro = CargarLibro(PlantillaV2);
        var hoja = libro.Hojas[HojaConsolidado];

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var celda = $"D{PrimeraFilaPublicacionExtemp + aseId - 1}";
            Assert.True(hoja.ContainsKey(celda), $"Falta la celda de publicación '{HojaConsolidado}!{celda}'.");
            var formula = hoja[celda].Formula;
            Assert.True(formula is not null, $"'{HojaConsolidado}!{celda}' debe declarar la referencia del EXTEMP del bloque {aseId}.");

            var referencia = ReferenciaUnicaConHoja.Match(formula!);
            Assert.True(referencia.Success, $"'{HojaConsolidado}!{celda}' no es una referencia única: '{formula}'.");
            Assert.Equal(HojaR1, referencia.Groups["hoja"].Value);
        }

        // §2-D3.2: el consumidor declarado del ancla secundaria (familia consumida por
        // REMUNERACION_*) existe como <f> leído y publica en la MISMA fila del bloque.
        var occidente = libro.Hojas["REMUNERACION_OCCIDENTE_ Directa"]["D51"].Formula;
        Assert.Equal("'Reporte Componentes R1'!F702", occidente);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // S1.3 — Cierre aritmético contra el R10 de los tres períodos, al peso
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Theory]
    [InlineData("2026071")]
    [InlineData("2026072")]
    [InlineData("2026101")]
    public void T0_4_DetRetriDelR10_CierraAlPesoContraLaTablaDelPlan(string periodo)
    {
        var detRetri = LeerDetRetriR10(RutaR10(periodo));
        var esperado = DetRetriR10PorPeriodo[periodo];

        Assert.Equal(5, detRetri.Count);
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            Assert.Equal(esperado[aseId - 1], detRetri[aseId]);
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // S1.4 — SHAs de las plantillas (G-SHA): el T0 es read-only, nada cambió
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void T0_5_PlantillasV1YV2_ByteIdenticas()
    {
        Assert.Equal(ShaPlantillaV2, Sha256(PlantillaV2));
        Assert.Equal(ShaPlantillaV1, Sha256(PlantillaV1));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // Derivación (la MISMA regla que implementará H1, sin tabla y sin direcciones congeladas)
    // ════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Ancla EXTEMP de cada bloque = la celda que declara el <c>&lt;f&gt;</c> de
    /// <c>CONSOLIDADO_TOTAL RECAUDO!D{46+aseId}</c> (§2-D3.1). La posición se LEE; el bloque sólo
    /// aporta su fila de publicación.
    /// </summary>
    private static List<(string Hoja, string Celda, string? Formula)> AnclasDelBloque(Libro libro)
    {
        var resultado = new List<(string, string, string?)>(5);
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var publicacion = libro.Hojas[HojaConsolidado][$"D{PrimeraFilaPublicacionExtemp + aseId - 1}"].Formula
                ?? throw new InvalidOperationException($"'{HojaConsolidado}!D{PrimeraFilaPublicacionExtemp + aseId - 1}' sin <f>.");
            var referencia = ReferenciaUnicaConHoja.Match(publicacion);
            Assert.True(referencia.Success, $"Publicación del bloque {aseId} sin referencia única: '{publicacion}'.");

            var hoja = referencia.Groups["hoja"].Value;
            var celda = referencia.Groups["celda"].Value;
            var destino = libro.Hojas[hoja][celda];
            resultado.Add((hoja, celda, destino.Formula));
        }

        return resultado;
    }

    /// <summary>
    /// Destinos R1 declarados en la fila de publicación <paramref name="fila"/> por CUALQUIER hoja
    /// del workbook (§2-D3.2: la familia secundaria se localiza siguiendo su consumidor leído).
    /// Excluye la propia hoja R1 (ahí viven las anclas, no sus consumidores).
    /// </summary>
    private static HashSet<string> AnclasPublicadasEnLaFila(Libro libro, int fila)
    {
        var celdas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var celda = $"D{fila}";

        foreach (var (nombre, contenido) in libro.Hojas)
        {
            if (string.Equals(nombre, HojaR1, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!contenido.TryGetValue(celda, out var publicacion) || publicacion.Formula is not { } formula)
            {
                continue;
            }

            var referencia = ReferenciaUnicaConHoja.Match(formula);
            if (referencia.Success && string.Equals(referencia.Groups["hoja"].Value, HojaR1, StringComparison.OrdinalIgnoreCase))
            {
                celdas.Add(referencia.Groups["celda"].Value);
            }
        }

        return celdas;
    }

    private static void AssertFormulaDeReferencias(string formula, (string Hoja, string Celda) ancla)
    {
        Assert.True(
            FormulaDeReferencias.IsMatch(formula.Replace(" ", string.Empty, StringComparison.Ordinal)),
            $"El ancla '{ancla.Hoja}!{ancla.Celda}' trae una forma que el resolutor acotado no evalúa: '{formula}'.");
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // Lectura por bytes (BCL puro)
    // ════════════════════════════════════════════════════════════════════════════════════════

    private sealed record Celda(string? Formula, string? Valor);

    private sealed class Libro
    {
        public Dictionary<string, Dictionary<string, Celda>> Hojas { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private static Libro CargarLibro(string ruta)
    {
        Assert.True(File.Exists(ruta), $"Falta el archivo: {ruta}");

        using var zip = ZipFile.OpenRead(ruta);
        var workbook = LeerXml(zip, "xl/workbook.xml");
        var relaciones = LeerXml(zip, "xl/_rels/workbook.xml.rels");
        var compartidas = zip.GetEntry("xl/sharedStrings.xml") is { } entradaCompartidas
            ? LeerXml(entradaCompartidas).Descendants(Ns + "si").Select(si => string.Concat(si.Descendants(Ns + "t").Select(t => t.Value))).ToList()
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

            var celdas = new Dictionary<string, Celda>(StringComparer.OrdinalIgnoreCase);
            foreach (var celda in LeerXml(entrada).Descendants(Ns + "c"))
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

                celdas[referencia] = new Celda(formula is null ? null : formula.Value, valor);
            }

            libro.Hojas[nombre] = celdas;
        }

        return libro;
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

    private static Dictionary<int, decimal> LeerDetRetriR10(string ruta)
    {
        var libro = CargarLibro(ruta);
        var hoja = libro.Hojas.Single(h => h.Key.StartsWith("DetRetri", StringComparison.OrdinalIgnoreCase)).Value;

        var filas = hoja
            .Select(par => (Referencia: par.Key, Fila: FilaDe(par.Key), Celda: par.Value))
            .GroupBy(t => t.Fila)
            .ToDictionary(g => g.Key, g => g.ToDictionary(t => ColumnaDe(t.Referencia), t => t.Celda, StringComparer.OrdinalIgnoreCase));

        var indiceHeader = filas
            .Where(par => Texto(par.Value, "B").Equals("ASE", StringComparison.OrdinalIgnoreCase)
                && Texto(par.Value, "D").StartsWith("RECAUDO TOTAL", StringComparison.OrdinalIgnoreCase))
            .Select(par => par.Key)
            .DefaultIfEmpty(-1)
            .Max();
        Assert.True(indiceHeader > 0, $"El R10 '{Path.GetFileName(ruta)}' no declara el encabezado ASE / RECAUDO TOTAL.");

        var porAse = new Dictionary<int, decimal>();
        foreach (var celdas in filas.Where(par => par.Key > indiceHeader).OrderBy(par => par.Key).Select(par => par.Value))
        {
            if (int.TryParse(Texto(celdas, "B"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var aseId)
                && aseId is >= 1 and <= 5)
            {
                porAse[aseId] = ParsearMonetario(Texto(celdas, "D"));
                continue;
            }

            if (Texto(celdas, "C").Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                break;
            }
        }

        return porAse;
    }

    private static string Texto(Dictionary<string, Celda> celdas, string columna) =>
        celdas.TryGetValue(columna, out var celda) ? (celda.Valor ?? string.Empty) : string.Empty;

    private static int FilaDe(string referencia) =>
        int.Parse(Regex.Match(referencia, @"\d+$").Value, CultureInfo.InvariantCulture);

    private static string ColumnaDe(string referencia) =>
        Regex.Match(referencia, @"^[A-Z]{1,3}").Value;

    /// <summary>Montos del R10 en texto es-CO (<c>"$14.842.425.136"</c>, negativo <c>"-$4.217.145.527"</c>).</summary>
    private static decimal ParsearMonetario(string texto)
    {
        var negativo = texto.Contains('-', StringComparison.Ordinal);
        var digitos = new string(texto.Where(char.IsDigit).ToArray());
        if (digitos.Length == 0)
        {
            return 0m;
        }

        var valor = decimal.Parse(digitos, NumberStyles.Integer, CultureInfo.InvariantCulture);
        return negativo ? -valor : valor;
    }

    private static XDocument LeerXml(ZipArchive zip, string entrada) =>
        LeerXml(zip.GetEntry(entrada) ?? throw new InvalidOperationException($"El paquete no trae '{entrada}'."));

    private static XDocument LeerXml(ZipArchiveEntry entrada)
    {
        using var stream = entrada.Open();
        return XDocument.Load(stream);
    }

    private static string Sha256(string ruta)
    {
        using var stream = File.OpenRead(ruta);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    // ── rutas explícitas (prohibido Insumos.Raiz() en el T0 del Plan 57) ───────────────────────

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

    private static string PlantillaV2 => Path.Combine(
        RaizRepo(), "Docs", "Plantilla estandarizada", "v2", "Formato Remuneracion 2026xx-x Total vf.xlsx");

    private static string PlantillaV1 => Path.Combine(
        RaizRepo(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string RutaR10(string periodo)
    {
        var carpeta = periodo switch
        {
            "2026071" => Path.Combine("Docs", "Prueba Julio-1", "Insumos"),
            "2026072" => Path.Combine("Docs", "Prueba Julio-2", "Insumos"),
            "2026101" => Path.Combine("Docs", "Prueba Octubre-1", "Insumos"),
            _ => throw new ArgumentOutOfRangeException(nameof(periodo), periodo, "Período sin R10 declarado en el T0.")
        };

        return Path.Combine(RaizRepo(), carpeta, $"R10_Remuneracion_{periodo}.xlsx");
    }
}
