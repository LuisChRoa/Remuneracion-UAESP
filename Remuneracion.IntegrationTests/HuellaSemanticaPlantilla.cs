using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (§0.6 G-HUELLA): HUELLA SEMÁNTICA de una plantilla Excel.
///
/// El SHA-256 del ARCHIVO <c>.xlsx</c> NO sirve como gate: un humano abre la plantilla en Excel, la
/// guarda y el paquete cambia (orden de partes, compresión) aunque el CONTENIDO no cambie ni un byte.
/// Eso produjo los seis falsos rojos de <c>G-SHA</c>. La huella se calcula sobre el CONTENIDO
/// normalizado y ORDENADO de las hojas:
/// <list type="bullet">
/// <item><c>hoja|&lt;nombre&gt;|&lt;dimension&gt;</c> por cada hoja, en el orden del workbook;</item>
/// <item><c>celda|&lt;hoja&gt;|&lt;ref&gt;|f|&lt;texto de &lt;f&gt;&gt;|v|&lt;valor de &lt;v&gt;&gt;|si|&lt;índice compartido&gt;|ref|&lt;rango compartido&gt;</c>
/// por cada celda con <c>&lt;v&gt;</c> (las fórmulas compartidas viajan con su <c>si</c>/<c>ref</c>, así
/// que un maestro o una seguidora que cambie de rol cambia la huella).</item>
/// </list>
/// Las entradas se ordenan con comparación ordinal y se unen con salto de línea; el hash es SHA-256 en
/// hexadecimal minúscula. El SHA del archivo queda como dato INFORMATIVO (<see cref="Sha256Archivo"/>).
/// </summary>
internal static class HuellaSemanticaPlantilla
{
    /// <summary>Huella semántica de la plantilla v3 (vigente).</summary>
    internal const string V3 = "92935014f3611055c15f12db92c71020f741d009634e8ec59f1e4448235d72ff";

    /// <summary>Huella semántica de la plantilla v2 (histórica, no se consume).</summary>
    internal const string V2 = "a4ea462b57be95938d0a25a9369bfd1e19d2d1229b75de7d89fbf71fb5991879";

    /// <summary>Huella semántica de la plantilla v1 (histórica, no se consume).</summary>
    internal const string V1 = "cf9aa03604ed72b2252876df83abd573eb5609b14a97026ef7c1b23774490d17";

    /// <summary>SHA-256 del archivo v1: la v1 no fue re-guardada, así que su ancla de archivo sigue viva.</summary>
    internal const string V1Sha256Archivo = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";

    /// <summary>Huella semántica de la plantilla (contenido, no empaquetado).</summary>
    internal static string Calcular(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace prns = "http://schemas.openxmlformats.org/package/2006/relationships";

        XDocument workbook;
        using (var stream = zip.GetEntry("xl/workbook.xml")!.Open())
        {
            workbook = XDocument.Load(stream);
        }

        XDocument relaciones;
        using (var stream = zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open())
        {
            relaciones = XDocument.Load(stream);
        }

        var destinos = relaciones.Root!.Elements(prns + "Relationship")
            .ToDictionary(r => r.Attribute("Id")!.Value, r => r.Attribute("Target")!.Value);

        var entradas = new List<string>();
        foreach (var hoja in workbook.Root!.Element(ns + "sheets")!.Elements(ns + "sheet"))
        {
            var nombre = hoja.Attribute("name")!.Value;
            var destino = destinos[hoja.Attribute(rns + "id")!.Value].Replace("worksheets/", "xl/worksheets/", StringComparison.Ordinal);
            if (!destino.StartsWith("xl/", StringComparison.Ordinal))
            {
                destino = "xl/" + destino;
            }

            XDocument hojaXml;
            using (var stream = zip.GetEntry(destino)!.Open())
            {
                hojaXml = XDocument.Load(stream);
            }

            var dimension = hojaXml.Descendants(ns + "dimension").FirstOrDefault()?.Attribute("ref")?.Value ?? string.Empty;
            entradas.Add($"hoja|{nombre}|{dimension}");

            foreach (var celda in hojaXml.Descendants(ns + "c"))
            {
                var valor = celda.Element(ns + "v");
                if (valor is null)
                {
                    continue;
                }

                var formula = celda.Element(ns + "f");
                entradas.Add(
                    $"celda|{nombre}|{celda.Attribute("r")?.Value}|f|{formula?.Value}|v|{valor.Value}" +
                    $"|si|{formula?.Attribute("si")?.Value}|ref|{formula?.Attribute("ref")?.Value}");
            }
        }

        entradas.Sort(StringComparer.Ordinal);
        var bytes = Encoding.UTF8.GetBytes(string.Join("\n", entradas));
        return Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    }

    /// <summary>SHA-256 del ARCHIVO: dato informativo, nunca gate (ver doctrina §0.6).</summary>
    internal static string Sha256Archivo(string ruta) => Sha256Archivo(ruta, upper: false);

    /// <summary>SHA-256 del archivo en mayúscula (formato de las anclas históricas de la suite).</summary>
    internal static string Sha256ArchivoMayuscula(string ruta) => Sha256Archivo(ruta, upper: true);

    private static string Sha256Archivo(string ruta, bool upper)
    {
        using var stream = File.OpenRead(ruta);
        var hash = Convert.ToHexString(SHA256.HashData(stream));
        return upper ? hash : hash.ToLowerInvariant();
    }
}
