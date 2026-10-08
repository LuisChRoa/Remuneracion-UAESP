using System.Globalization;
using System.IO.Compression;
using System.Xml;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 33 (T1, R-G-1/R-G-2): gate <b>workbook-wide</b> del invariante shared de
/// <c>Reporte Componentes R1</c> (<c>sheet10.xml</c>). Lee el XML crudo (BCL <c>System.IO.Compression</c>
/// + <c>System.Xml</c>; sin Excel/COM — mismo método del comparador Plan 29 y los T0 de Planes 31/32)
/// y expone:
///   - <see cref="Leer"/>: masters (<c>t="shared"</c> con <c>ref</c>+<c>si</c>) y seguidoras
///     (<c>t="shared"</c> con <c>si</c> y sin <c>ref</c>) de la hoja.
///   - <see cref="ExigirInvariante"/>: todo <c>si</c> que aparece en seguidoras tiene un master con
///     ese <c>si</c> en la MISMA hoja; fail-fast nombrando hoja + celda + <c>si</c> (todos los huérfanos).
///
/// PROHIBIDO usar cachés <c>&lt;v&gt;</c> como oráculo (doctrina Plan 29/28): el gate solo lee
/// atributos (<c>t</c>/<c>ref</c>/<c>si</c>) y el texto de <c>&lt;f&gt;</c>. El fix de T2 preserva esos
/// atributos al recomponer; este gate es la red que hoy fallaría en rojo con la regen corrupta (E1–E4).
/// </summary>
internal static class SharedFormulaR1Gate
{
    /// <summary>Master shared: celda + <c>si</c> + rango <c>ref</c> + texto de <c>&lt;f&gt;</c>.</summary>
    public sealed record MaestroShared(string Celda, uint Si, string Rango, string Texto);

    /// <summary>Lectura shared de una hoja: masters por celda + seguidoras (celda, si).</summary>
    public sealed record LecturaShared(
        IReadOnlyDictionary<string, MaestroShared> Maestros,
        IReadOnlyList<(string Celda, uint Si)> Seguidoras);

    /// <summary>
    /// Lee masters y seguidoras shared de la hoja indicada. Resuelve el nombre de hoja →
    /// <c>xl/worksheets/sheetN.xml</c> vía <c>workbook.xml</c> + <c>workbook.xml.rels</c> (no asume el
    /// número de sheet).
    /// </summary>
    public static LecturaShared Leer(string rutaWorkbook, string hoja)
    {
        if (string.IsNullOrWhiteSpace(rutaWorkbook))
        {
            throw new ArgumentException("La ruta del workbook es requerida.", nameof(rutaWorkbook));
        }

        if (!File.Exists(rutaWorkbook))
        {
            throw new FileNotFoundException($"No existe el workbook a gatear: '{rutaWorkbook}'.", rutaWorkbook);
        }

        using var zip = ZipFile.OpenRead(rutaWorkbook);
        var rutaHoja = ResolverHoja(zip, hoja)
            ?? throw new InvalidOperationException($"El workbook '{Path.GetFileName(rutaWorkbook)}' no trae la hoja '{hoja}'.");
        var entrada = zip.GetEntry(rutaHoja)
            ?? throw new InvalidOperationException($"El workbook '{Path.GetFileName(rutaWorkbook)}' no tiene el XML de la hoja '{hoja}' ({rutaHoja}).");

        var doc = new XmlDocument { XmlResolver = null };
        using (var stream = entrada.Open())
        {
            doc.Load(stream);
        }

        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        var maestros = new Dictionary<string, MaestroShared>(StringComparer.OrdinalIgnoreCase);
        var seguidoras = new List<(string Celda, uint Si)>();

        foreach (XmlNode c in doc.SelectNodes("//m:sheetData/m:row/m:c", ns)!)
        {
            var referencia = c.Attributes?["r"]?.Value;
            var f = c.SelectSingleNode("m:f", ns);
            if (string.IsNullOrWhiteSpace(referencia) || f is null)
            {
                continue;
            }

            if (!string.Equals(f.Attributes?["t"]?.Value, "shared", StringComparison.Ordinal))
            {
                continue;
            }

            if (!uint.TryParse(f.Attributes?["si"]?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var si))
            {
                continue;
            }

            var rango = f.Attributes?["ref"]?.Value;
            if (!string.IsNullOrWhiteSpace(rango))
            {
                maestros[referencia!] = new MaestroShared(referencia!, si, rango!, f.InnerText);
            }
            else
            {
                seguidoras.Add((referencia!, si));
            }
        }

        return new LecturaShared(maestros, seguidoras);
    }

    /// <summary>
    /// Invariante workbook-wide: todo <c>si</c> de una seguidora tiene un master con ese <c>si</c> en
    /// la misma hoja. Fail-fast nombrando hoja + celda + <c>si</c> (lista completa de huérfanos).
    /// </summary>
    public static void ExigirInvariante(string rutaWorkbook, string hoja)
    {
        var lectura = Leer(rutaWorkbook, hoja);
        var masters = new HashSet<uint>(lectura.Maestros.Values.Select(m => m.Si));
        var huerfanos = lectura.Seguidoras
            .Where(s => !masters.Contains(s.Si))
            .Select(s => $"{hoja}!{s.Celda} [si={s.Si}]: sin master con ese si en la hoja")
            .ToList();

        if (huerfanos.Count > 0)
        {
            throw new InvalidOperationException(
                $"Invariante shared R1 rota ({huerfanos.Count} seguidora(s) huérfana(s)):\n" + string.Join("\n", huerfanos));
        }
    }

    private static string? ResolverHoja(ZipArchive zip, string hoja)
    {
        var workbook = Cargar(zip, "xl/workbook.xml");
        if (workbook is null)
        {
            return null;
        }

        var relaciones = LeerRelaciones(zip);
        var ns = new XmlNamespaceManager(workbook.NameTable);
        ns.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        foreach (XmlNode sheet in workbook.SelectNodes("//m:sheets/m:sheet", ns)!)
        {
            var nombre = sheet.Attributes?["name"]?.Value;
            var rid = sheet.Attributes?["r:id"]?.Value;
            if (!string.Equals(nombre, hoja, StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrWhiteSpace(rid)
                || !relaciones.TryGetValue(rid!, out var target))
            {
                continue;
            }

            return target.StartsWith("/", StringComparison.Ordinal)
                ? target.TrimStart('/')
                : "xl/" + target;
        }

        return null;
    }

    private static Dictionary<string, string> LeerRelaciones(ZipArchive zip)
    {
        var mapa = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var doc = Cargar(zip, "xl/_rels/workbook.xml.rels");
        if (doc is null)
        {
            return mapa;
        }

        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("p", "http://schemas.openxmlformats.org/package/2006/relationships");
        foreach (XmlNode rel in doc.SelectNodes("//p:Relationship", ns)!)
        {
            var id = rel.Attributes?["Id"]?.Value;
            var target = rel.Attributes?["Target"]?.Value;
            if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
            {
                mapa[id!] = target!;
            }
        }

        return mapa;
    }

    private static XmlDocument? Cargar(ZipArchive zip, string path)
    {
        var entrada = zip.GetEntry(path);
        if (entrada is null)
        {
            return null;
        }

        using var stream = entrada.Open();
        var doc = new XmlDocument { XmlResolver = null };
        doc.Load(stream);
        return doc;
    }
}
