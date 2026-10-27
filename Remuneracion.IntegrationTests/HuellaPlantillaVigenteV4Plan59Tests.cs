using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (§0.3/§0.6, rebase 2026-10-27): la plantilla <b>v4</b> pasa a ser la VIGENTE y su
/// <b>huella semántica</b> es el gate de entrada. La v3 queda en disco como histórica.
///
/// Este fixture fija tres cosas que el resto de la suite no puede dar por supuestas:
/// <list type="number">
/// <item>La huella de la v4 es <see cref="HuellaSemanticaPlantilla.V4"/> (no el SHA del archivo: el
/// empaquetado cambia cuando un humano guarda en Excel, §0.6).</item>
/// <item>La v4 es la v3 <b>más exactamente dos celdas de fórmula</b> (<c>R1!F328</c>, <c>R1!F333</c>):
/// mismo conjunto de hojas, mismas dimensiones, mismo número de celdas. Se mide con el helper
/// canónico, no con una lista de constantes.</item>
/// <item><b>Prueba de inyección:</b> revertir <c>F328</c> al typo de la v3 en una copia enrojece el
/// gate (la huella deja de ser <see cref="HuellaSemanticaPlantilla.V4"/>). Un gate que no enrojece
/// ante un cambio de contenido no es un gate.</item>
/// </list>
/// </summary>
public sealed class HuellaPlantillaVigenteV4Plan59Tests
{
    private const string HojaR1 = "Reporte Componentes R1";

    /// <summary>Fórmula correcta de la v4 en <c>R1!F328</c> (la v3 traía <c>-M213</c>).</summary>
    private const string FormulaF328V4 = "F269+F213+F141-L141-L213";

    /// <summary>Fórmula correcta de la v4 en <c>R1!F333</c> (la v3 traía <c>--L126</c>).</summary>
    private const string FormulaF333V4 = "F254+F191+F126-L126-L191";

    /// <summary>Typo de la v3 en <c>R1!F328</c>: el que impedía cerrar julio-1 y julio-2.</summary>
    private const string FormulaF328V3 = "F269+F213+F141-L141-M213";

    /// <summary>SHA-256 del ARCHIVO v4: dato INFORMATIVO, nunca gate (§0.6).</summary>
    private const string Sha256ArchivoV4 = "5FEDFD316980276945DF2AF29784EFFC5FE178C5CF1AA16F3C5D5F6F0A4E38E5";

    private const long LongitudArchivoV4 = 705_805;

    [Fact]
    public void V4_1_ElArchivoVigenteTieneElShaInformativoyElTamanoDeclarados()
    {
        Assert.Equal(Sha256ArchivoV4, HuellaSemanticaPlantilla.Sha256ArchivoMayuscula(PlantillaOficialV4));
        Assert.Equal(LongitudArchivoV4, new FileInfo(PlantillaOficialV4).Length);
    }

    [Fact]
    public void V4_2_LaHuellaSemanticaDeLaPlantillaVigenteEsLaDelGate()
    {
        Assert.Equal(HuellaSemanticaPlantilla.V4, HuellaSemanticaPlantilla.Calcular(PlantillaOficialV4));
    }

    [Fact]
    public void V4_3_LaV3SigueEnDiscoComoHistoricaConSuHuellaPropia()
    {
        Assert.True(File.Exists(PlantillaHistoricaV3), $"La v3 debe quedar en disco (histórica): {PlantillaHistoricaV3}");
        Assert.Equal(HuellaSemanticaPlantilla.V3, HuellaSemanticaPlantilla.Calcular(PlantillaHistoricaV3));
        Assert.NotEqual(HuellaSemanticaPlantilla.V3, HuellaSemanticaPlantilla.V4);
    }

    [Fact]
    public void V4_4_LaV4EsLaV3MasExactamenteLasDosFormulasCorregidas()
    {
        var v3 = Leer(PlantillaHistoricaV3);
        var v4 = Leer(PlantillaOficialV4);

        Assert.Equal(v3.Hojas.Select(h => $"{h.Nombre}|{h.Dimension}"), v4.Hojas.Select(h => $"{h.Nombre}|{h.Dimension}"));
        Assert.Equal(v3.Celdas.Count, v4.Celdas.Count);

        var diferencias = v3.Celdas.Keys.Union(v4.Celdas.Keys, StringComparer.Ordinal)
            .Where(k => v3.Celdas.GetValueOrDefault(k) != v4.Celdas.GetValueOrDefault(k))
            .OrderBy(k => k, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(
            [$"{HojaR1}!F328", $"{HojaR1}!F333"],
            diferencias);

        Assert.Equal($"f|{FormulaF328V3}|v|0|si||ref|", v3.Celdas[$"{HojaR1}!F328"]);
        Assert.Equal($"f|{FormulaF328V4}|v|0|si||ref|", v4.Celdas[$"{HojaR1}!F328"]);
        Assert.Equal($"f|{FormulaF333V4}|v|0|si||ref|", v4.Celdas[$"{HojaR1}!F333"]);
    }

    /// <summary>
    /// Prueba de INYECCIÓN del gate: sobre una COPIA temporal de la v4 se revierte <c>F328</c> al typo
    /// de la v3 (<c>-M213</c>). La huella de la copia debe dejar de ser
    /// <see cref="HuellaSemanticaPlantilla.V4"/> — es decir, el gate enrojece ante un cambio de
    /// CONTENIDO aunque el empaquetado sea el mismo.
    /// </summary>
    [Fact]
    public void V4_5_INYECCION_RevertirF328AlTypoDeLaV3EnrojeceElGateDeHuella()
    {
        var copia = Path.Combine(Path.GetTempPath(), $"p59-v4-inyeccion-{Guid.NewGuid():N}.xlsx");
        try
        {
            File.Copy(PlantillaOficialV4, copia, overwrite: false);
            Assert.Equal(HuellaSemanticaPlantilla.V4, HuellaSemanticaPlantilla.Calcular(copia));

            SustituirFormula(copia, HojaR1, "F328", FormulaF328V3);

            var huellaInyectada = HuellaSemanticaPlantilla.Calcular(copia);
            Assert.NotEqual(HuellaSemanticaPlantilla.V4, huellaInyectada);
            Assert.NotEqual(HuellaSemanticaPlantilla.V3, huellaInyectada);
        }
        finally
        {
            if (File.Exists(copia))
            {
                File.Delete(copia);
            }
        }
    }

    private static string PlantillaOficialV4 => Insumos.PlantillaOficialV4();

    /// <summary>Ruta de la v3 histórica: solo este fixture la abre, como evidencia congelada (§0.3).</summary>
    private static string PlantillaHistoricaV3 => Path.Combine(
        Insumos.Raiz(), "Docs", "Plantilla estandarizada", "v3", "Formato Remuneracion 2026xx-x Total v3.xlsx");

    // ── Lectura por bytes (BCL: ZIP + XML), sin Excel ni OpenXML SDK ─────────────────────────────

    private sealed record Libro(IReadOnlyList<(string Nombre, string Dimension)> Hojas, IReadOnlyDictionary<string, string> Celdas);

    private static Libro Leer(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace prns = "http://schemas.openxmlformats.org/package/2006/relationships";

        var workbook = Cargar(zip, "xl/workbook.xml");
        var relaciones = Cargar(zip, "xl/_rels/workbook.xml.rels");
        var destinos = relaciones.Root!.Elements(prns + "Relationship")
            .ToDictionary(r => r.Attribute("Id")!.Value, r => r.Attribute("Target")!.Value);

        var hojas = new List<(string, string)>();
        var celdas = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (var hoja in workbook.Root!.Element(ns + "sheets")!.Elements(ns + "sheet"))
        {
            var nombre = hoja.Attribute("name")!.Value;
            var destino = destinos[hoja.Attribute(rns + "id")!.Value].Replace("worksheets/", "xl/worksheets/", StringComparison.Ordinal);
            if (!destino.StartsWith("xl/", StringComparison.Ordinal))
            {
                destino = "xl/" + destino;
            }

            var xml = Cargar(zip, destino);
            var dimension = xml.Descendants(ns + "dimension").FirstOrDefault()?.Attribute("ref")?.Value ?? string.Empty;
            hojas.Add((nombre, dimension));

            foreach (var celda in xml.Descendants(ns + "c"))
            {
                var formula = celda.Element(ns + "f");
                var valor = celda.Element(ns + "v");
                if (formula is null && valor is null)
                {
                    continue;
                }

                celdas[$"{nombre}!{celda.Attribute("r")?.Value}"] =
                    $"f|{formula?.Value}|v|{valor?.Value}|si|{formula?.Attribute("si")?.Value}|ref|{formula?.Attribute("ref")?.Value}";
            }
        }

        return new Libro(hojas, celdas);
    }

    private static XDocument Cargar(ZipArchive zip, string entrada)
    {
        using var stream = zip.GetEntry(entrada)!.Open();
        return XDocument.Load(stream);
    }

    /// <summary>Reescribe el <c>&lt;f&gt;</c> de una celda en una copia (solo para la prueba de inyección).</summary>
    private static void SustituirFormula(string ruta, string nombreHoja, string referencia, string formula)
    {
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace prns = "http://schemas.openxmlformats.org/package/2006/relationships";

        using var zip = ZipFile.Open(ruta, ZipArchiveMode.Update);

        var workbookDoc = Cargar(zip, "xl/workbook.xml");
        var relaciones = Cargar(zip, "xl/_rels/workbook.xml.rels");
        var destinos = relaciones.Root!.Elements(prns + "Relationship")
            .ToDictionary(r => r.Attribute("Id")!.Value, r => r.Attribute("Target")!.Value);

        var hoja = workbookDoc.Root!.Element(ns + "sheets")!.Elements(ns + "sheet")
            .Single(s => s.Attribute("name")!.Value == nombreHoja);
        var destino = destinos[hoja.Attribute(rns + "id")!.Value].Replace("worksheets/", "xl/worksheets/", StringComparison.Ordinal);
        if (!destino.StartsWith("xl/", StringComparison.Ordinal))
        {
            destino = "xl/" + destino;
        }

        var entrada = zip.GetEntry(destino)!;
        XDocument xml;
        using (var stream = entrada.Open())
        {
            xml = XDocument.Load(stream);
        }

        var celda = xml.Descendants(ns + "c").Single(c => c.Attribute("r")?.Value == referencia);
        var f = celda.Element(ns + "f") ?? throw new InvalidOperationException($"{nombreHoja}!{referencia} no declara <f>.");
        f.Value = formula;

        entrada.Delete();
        var nueva = zip.CreateEntry(destino);
        using var salida = nueva.Open();
        xml.Save(salida);
    }
}
