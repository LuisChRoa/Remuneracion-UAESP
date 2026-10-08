using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 29 (T5, Unidad P — R-P-2 / S6): comparador de paridad entre la salida generada por la app
/// y el archivo manual del administrativo, hoja por celda.
///
/// BCL PURO: solo <see cref="System.IO.Compression"/> (ZIP) + <see cref="System.Xml"/> (XML). NO usa
/// DocumentFormat.OpenXml, ExcelDataReader, Excel ni COM. Es una herramienta de verificación del
/// pipeline (tests/regresión), no un motor de cálculo.
///
/// Regla de comparación (D-E del plan 29: VERIFICADOR SIN RECÁLCULO):
///   (i)  literal-vs-literal: |Δ| ≤ 0.5 para numéricos o igualdad de texto normalizado;
///   (ii) fórmula-vs-fórmula: igualdad del TEXTO de la fórmula normalizado (espacios fuera) — la
///        fórmula ES la evidencia de que post-recálculo el resultado coincide (Plan 28 garantiza
///        <c>fullCalcOnLoad</c>);
///   (iii) literal-vs-fórmula, o presencia unilateral con valor sustantivo: divergencia.
///
/// PROHIBIDO leer cachés <c>&lt;v&gt;</c> donde hay <c>&lt;f&gt;</c> (V8/D-E: los cachés son stale
/// por diseño). Por eso una celda con fórmula se compara SOLO por texto de fórmula.
///
/// Ausencia tratada como 0 solo para NUMÉRICOS: la plantilla deja literales-0 de captura que el
/// manual borró (V8/T0e, no-brecha cosmética) → 0-vs-ausente no es divergencia; ausente-vs-texto o
/// ausente-vs-fórmula sí lo es.
///
/// EXCLUSIONES VERSIONADAS EN CÓDIGO (R-FALSO-POSITIVO del plan; <see cref="EsExclusionVersionada"/>):
///   - Columnas de otra quincena en las hojas <c>Recaudo *</c> (en una corrida Q2 la app escribe
///     F/G; D/E son 1°Q y el manual las acumula — V3).
///   - Metadatos de período en <c>CONSOLIDADO_TOTAL RECAUDO</c>: N3/C6/D6 (Plan 28 los deja fuera de
///     scope) y los NOMBRES de hoja con sufijo de período (el par se resuelve por nombre
///     normalizado sin el token de período; H1/T0g: la base hereda el sufijo).
///   - <c>DetValiRetri*!J</c> (AJUSTE A LA DECENA): SALE documentado de T0d (manual-externo, sin
///     origen workbook-interno 10/10).
/// </summary>
public sealed class ComparadorSalidaVsManual
{
    private static readonly Regex TokenPeriodo = new(@"\d{6}-?\d?", RegexOptions.Compiled);
    private static readonly Regex ReferenciaCelda = new(@"^(?<col>[A-Za-z]+)(?<fila>\d+)$", RegexOptions.Compiled);

    private readonly int _numeroQuincena;

    /// <summary>
    /// Crea el comparador para una corrida de la quincena indicada (por defecto Q2, la del plan 29).
    /// </summary>
    public ComparadorSalidaVsManual(int numeroQuincena = 2)
    {
        _numeroQuincena = numeroQuincena;
    }

    /// <summary>Divergencia de una celda entre la salida de la app y el manual.</summary>
    public sealed record Divergencia(string Hoja, string Celda, string? ValorApp, string? ValorManual, string Motivo);

    /// <summary>
    /// Compara dos workbooks <c>.xlsx</c> y devuelve las divergencias fuera de las exclusiones
    /// versionadas. Archivo inexistente → fail-fast nombrando la ruta.
    /// </summary>
    public IReadOnlyList<Divergencia> Comparar(string rutaApp, string rutaManual)
    {
        if (!File.Exists(rutaApp))
        {
            throw new FileNotFoundException($"No existe la salida de la app a comparar: '{rutaApp}'.", rutaApp);
        }

        if (!File.Exists(rutaManual))
        {
            throw new FileNotFoundException($"No existe el archivo manual a comparar: '{rutaManual}'.", rutaManual);
        }

        var app = LeerWorkbook(rutaApp);
        var manual = LeerWorkbook(rutaManual);
        var divergencias = new List<Divergencia>();

        var nombresApp = app.Keys.ToDictionary(K => NormalizarHoja(K), StringComparer.OrdinalIgnoreCase);
        var nombresManual = manual.Keys.ToDictionary(K => NormalizarHoja(K), StringComparer.OrdinalIgnoreCase);

        foreach (var (hojaBase, hojaApp) in nombresApp)
        {
            if (!nombresManual.TryGetValue(hojaBase, out var hojaManual))
            {
                // Hoja entera sin contraparte: divergencia por presencia unilateral (fuera de
                // exclusiones; los nombres con sufijo de período ya se normalizaron).
                foreach (var celda in app[hojaApp].Keys)
                {
                    if (EsExclusionVersionada(hojaApp, celda))
                    {
                        continue;
                    }

                    divergencias.Add(new Divergencia(hojaApp, celda, Describir(app[hojaApp][celda]), null, "hoja ausente en el manual"));
                }

                continue;
            }

            CompararHoja(hojaApp, app[hojaApp], hojaManual, manual[hojaManual], divergencias);
        }

        foreach (var (hojaBase, hojaManual) in nombresManual)
        {
            if (nombresApp.ContainsKey(hojaBase))
            {
                continue;
            }

            foreach (var celda in manual[hojaManual].Keys)
            {
                if (EsExclusionVersionada(hojaManual, celda))
                {
                    continue;
                }

                divergencias.Add(new Divergencia(hojaManual, celda, null, Describir(manual[hojaManual][celda]), "hoja ausente en la app"));
            }
        }

        return divergencias
            .OrderBy(d => d.Hoja, StringComparer.OrdinalIgnoreCase)
            .ThenBy(d => d.Celda, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private void CompararHoja(
        string hoja,
        IReadOnlyDictionary<string, Celda> app,
        string hojaManual,
        IReadOnlyDictionary<string, Celda> manual,
        List<Divergencia> divergencias)
    {
        var celdas = app.Keys.Concat(manual.Keys).Distinct(StringComparer.OrdinalIgnoreCase);
        foreach (var celda in celdas)
        {
            if (EsExclusionVersionada(hoja, celda))
            {
                continue;
            }

            app.TryGetValue(celda, out var a);
            manual.TryGetValue(celda, out var m);
            if (!SonEquivalentes(a, m))
            {
                divergencias.Add(new Divergencia(
                    hoja,
                    celda,
                    a is null ? null : Describir(a),
                    m is null ? null : Describir(m),
                    Motivo(a, m)));
            }
        }
    }

    /// <summary>
    /// Equivalencia D-E: fórmulas por texto normalizado; literales por ±0.5 (numérico) o igualdad de
    /// texto; ausencia = 0 solo en numérico.
    /// </summary>
    private static bool SonEquivalentes(Celda? a, Celda? b)
    {
        if (a is null && b is null)
        {
            return true;
        }

        var aFormula = a?.Formula;
        var bFormula = b?.Formula;
        if (aFormula is not null || bFormula is not null)
        {
            // Fórmula-vs-fórmula: mismo texto normalizado. Cualquier mezcla con literal/ausente diverge.
            return aFormula is not null
                && bFormula is not null
                && NormalizarFormula(aFormula) == NormalizarFormula(bFormula);
        }

        return SonLiteralesEquivalentes(a, b);
    }

    private static bool SonLiteralesEquivalentes(Celda? a, Celda? b)
    {
        var aNum = a?.EsNumero() == true;
        var bNum = b?.EsNumero() == true;

        if (aNum || bNum)
        {
            if (a is not null && !aNum)
            {
                return false; // texto vs número
            }

            if (b is not null && !bNum)
            {
                return false; // número vs texto
            }

            var va = aNum ? a!.Numero!.Value : 0m;
            var vb = bNum ? b!.Numero!.Value : 0m;
            return Math.Abs(va - vb) <= 0.5m;
        }

        // Texto vs texto (o vacío/ausente): igualdad normalizada (trim, mayúsculas, sin acentos).
        var ta = string.IsNullOrWhiteSpace(a?.Texto) ? string.Empty : NormalizarTexto(a!.Texto!);
        var tb = string.IsNullOrWhiteSpace(b?.Texto) ? string.Empty : NormalizarTexto(b!.Texto!);
        return string.Equals(ta, tb, StringComparison.Ordinal);
    }

    private static string Motivo(Celda? a, Celda? b)
    {
        if ((a?.Formula is not null) != (b?.Formula is not null))
        {
            return "fórmula vs literal/ausente";
        }

        if (a is null || b is null)
        {
            return "presencia unilateral";
        }

        if (a.Formula is not null && b.Formula is not null)
        {
            return "texto de fórmula distinto";
        }

        var aNum = a.EsNumero();
        var bNum = b.EsNumero();
        if (aNum && bNum)
        {
            return "literal numérico fuera de tolerancia (±0.5)";
        }

        if (aNum != bNum)
        {
            return "tipo distinto (texto vs número)";
        }

        return "texto distinto";
    }

    /// <summary>
    /// Exclusiones versionadas (R-FALSO-POSITIVO). Ver el resumen del tipo para el detalle y motivo.
    /// </summary>
    public bool EsExclusionVersionada(string hoja, string celda)
    {
        var baseHoja = NormalizarHoja(hoja);
        var columna = ColumnaDe(celda);

        // Columnas de otra quincena en las hojas Recaudo * (Q2: F/G; Q1: D/E).
        if (baseHoja.StartsWith("RECAUDO ", StringComparison.OrdinalIgnoreCase))
        {
            if (_numeroQuincena == 2 && columna is "D" or "E")
            {
                return true;
            }

            if (_numeroQuincena == 1 && columna is "F" or "G")
            {
                return true;
            }
        }

        // Metadatos de período (Plan 28: fuera de scope del sello).
        if (baseHoja.Equals("CONSOLIDADO_TOTAL RECAUDO", StringComparison.OrdinalIgnoreCase)
            && celda is "N3" or "C6" or "D6")
        {
            return true;
        }

        // SALE T0d: DetValiRetri J (AJUSTE A LA DECENA), manual-externo sin origen 10/10.
        if (baseHoja.StartsWith("DETVALIRETRI", StringComparison.OrdinalIgnoreCase) && columna == "J")
        {
            return true;
        }

        return false;
    }

    /// <summary>
    /// Normaliza el nombre de hoja quitando el token de período (H1/T0g: la base hereda el sufijo,
    /// p. ej. <c>DetRetri2026072</c> vs <c>DetRetri2026082</c>). Mayúsculas y sin diacríticos.
    /// </summary>
    public static string NormalizarHoja(string nombre)
    {
        var sinToken = TokenPeriodo.Replace(nombre, string.Empty);
        return NormalizarTexto(sinToken);
    }

    private static string NormalizarFormula(string formula) =>
        new(formula.Where(c => !char.IsWhiteSpace(c)).ToArray());

    private static string NormalizarTexto(string texto)
    {
        var descompuesto = texto.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder(descompuesto.Length);
        foreach (var ch in descompuesto)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
            {
                sb.Append(char.ToUpperInvariant(ch));
            }
        }

        return sb.ToString().Trim();
    }

    private static string ColumnaDe(string celda)
    {
        var match = ReferenciaCelda.Match(celda);
        return match.Success ? match.Groups["col"].Value.ToUpperInvariant() : string.Empty;
    }

    private static string Describir(Celda celda)
    {
        if (celda.Formula is not null)
        {
            return $"=f({celda.Formula})";
        }

        return celda.Texto ?? celda.Numero?.ToString(CultureInfo.InvariantCulture) ?? string.Empty;
    }

    // ── Lectura ZIP+XML (BCL puro) ────────────────────────────────────────────────────────────

    private static Dictionary<string, Dictionary<string, Celda>> LeerWorkbook(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var shared = LeerSharedStrings(zip);
        var hojas = new Dictionary<string, Dictionary<string, Celda>>(StringComparer.OrdinalIgnoreCase);

        var workbookXml = Cargar(zip, "xl/workbook.xml")
            ?? throw new InvalidDataException($"'{Path.GetFileName(ruta)}' no tiene xl/workbook.xml (¿es un xlsx?).");
        var relaciones = LeerRelaciones(zip);
        var ns = new XmlNamespaceManager(workbookXml.NameTable);
        ns.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        foreach (XmlNode hoja in workbookXml.SelectNodes("//m:sheets/m:sheet", ns)!)
        {
            var nombre = hoja.Attributes?["name"]?.Value ?? string.Empty;
            var rid = hoja.Attributes?["r:id"]?.Value;
            if (string.IsNullOrWhiteSpace(nombre) || string.IsNullOrWhiteSpace(rid)
                || !relaciones.TryGetValue(rid!, out var target))
            {
                continue;
            }

            var path = target.StartsWith("/", StringComparison.Ordinal)
                ? target.TrimStart('/')
                : "xl/" + target;
            var sheetXml = Cargar(zip, path);
            hojas[nombre] = sheetXml is null
                ? new Dictionary<string, Celda>(StringComparer.OrdinalIgnoreCase)
                : LeerCeldas(sheetXml, shared);
        }

        return hojas;
    }

    private static List<string> LeerSharedStrings(ZipArchive zip)
    {
        var doc = Cargar(zip, "xl/sharedStrings.xml");
        var lista = new List<string>();
        if (doc is null)
        {
            return lista;
        }

        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        foreach (XmlNode si in doc.SelectNodes("//m:si", ns)!)
        {
            lista.Add(si.InnerText);
        }

        return lista;
    }

    private static Dictionary<string, string> LeerRelaciones(ZipArchive zip)
    {
        var doc = Cargar(zip, "xl/_rels/workbook.xml.rels");
        var mapa = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
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

    private static Dictionary<string, Celda> LeerCeldas(XmlDocument sheetXml, IReadOnlyList<string> shared)
    {
        var celdas = new Dictionary<string, Celda>(StringComparer.OrdinalIgnoreCase);
        var ns = new XmlNamespaceManager(sheetXml.NameTable);
        ns.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

        foreach (XmlNode c in sheetXml.SelectNodes("//m:sheetData/m:row/m:c", ns)!)
        {
            var referencia = c.Attributes?["r"]?.Value;
            if (string.IsNullOrWhiteSpace(referencia))
            {
                continue;
            }

            var formulaNodo = c.SelectSingleNode("m:f", ns);
            var valorNodo = c.SelectSingleNode("m:v", ns);
            var tipo = c.Attributes?["t"]?.Value;

            string? formula = null;
            if (formulaNodo is not null)
            {
                formula = formulaNodo.InnerText;
                if (string.IsNullOrEmpty(formula)
                    && formulaNodo.Attributes?["t"]?.Value == "shared")
                {
                    formula = "SHARED"; // follower: sin texto propio; dos followers se consideran iguales
                }
            }

            string? texto = null;
            decimal? numero = null;
            if (tipo == "s")
            {
                if (valorNodo is not null
                    && int.TryParse(valorNodo.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx)
                    && idx >= 0 && idx < shared.Count)
                {
                    texto = shared[idx];
                }
            }
            else if (tipo is "inlineStr")
            {
                texto = c.SelectSingleNode("m:is", ns)?.InnerText ?? string.Empty;
            }
            else if (tipo is "str" or "b" or "e")
            {
                texto = valorNodo?.InnerText;
            }
            else if (valorNodo is not null
                && decimal.TryParse(valorNodo.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var n))
            {
                numero = n;
            }
            else if (valorNodo is not null && !string.IsNullOrWhiteSpace(valorNodo.InnerText))
            {
                texto = valorNodo.InnerText;
            }

            if (formula is null && texto is null && numero is null)
            {
                continue; // celda sin contenido (incluye <c r="A1"/> sin <v>)
            }

            celdas[referencia!] = new Celda(formula, texto, numero);
        }

        return celdas;
    }

    private static XmlDocument? Cargar(ZipArchive zip, string path)
    {
        var entry = zip.GetEntry(path);
        if (entry is null)
        {
            return null;
        }

        using var stream = entry.Open();
        var doc = new XmlDocument { XmlResolver = null };
        doc.Load(stream);
        return doc;
    }

    private sealed record Celda(string? Formula, string? Texto, decimal? Numero)
    {
        public bool EsNumero() => Formula is null && Numero.HasValue;
    }
}
