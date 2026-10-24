using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 58 (H1.2 / §2-D2) — <b>gate de integridad del paquete</b> por bytes (BCL: <see cref="ZipArchive"/>
/// + <see cref="XDocument"/>), sin Excel, sin COM, sin el SDK de OpenXml.
///
/// Existe porque <c>OpenXmlValidator</c> es CIEGO a esta clase de defecto: un grupo de fórmula
/// <c>shared</c> sin maestro no lo reporta, y Excel sí (al abrir borra las fórmulas al "reparar").
/// Ningún criterio de este gate depende de <c>OpenXmlValidator</c>.
///
/// Clases de defecto (todas abortan con <see cref="CodigoError.Integridad"/> ANTES de entregar):
/// <list type="number">
/// <item><c>GRUPO-SHARED</c>: grupo <c>si</c> sin maestro o con maestros múltiples, o cuyo <c>ref</c> no cubre el grupo.</item>
/// <item><c>FORMULA-PERDIDA</c> / <c>FORMULA-ATRIBUTO-ALTERADO</c> / <c>FORMULA-ALTERADA</c> / <c>FORMULA-CREADA</c>:
/// el conjunto de fórmulas (celda → texto + <c>t</c>/<c>ref</c>/<c>si</c>) no se preserva.</item>
/// <item><c>CELDA-DUPLICADA</c>: la misma referencia dos veces en una hoja.</item>
/// <item><c>FILA-DESORDENADA</c>: filas o celdas fuera del orden monótono que exige el formato.</item>
/// <item><c>SHAREDSTRING-FUERA-DE-RANGO</c>: <c>t="s"</c> con índice inexistente en sharedStrings.</item>
/// <item><c>HOJA-INEXISTENTE</c>: fórmula o nombre definido que referencia una hoja que no existe.</item>
/// </list>
///
/// Es <b>reutilizable</b>: no conoce períodos, ASE, empresas ni celdas concretas. Todo lo decide el
/// contenido de los paquetes comparados.
/// </summary>
internal static class ValidadorIntegridadPaquete
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string PartWorkbook = "xl/workbook.xml";
    private const string PartRels = "xl/_rels/workbook.xml.rels";
    private const string PartShared = "xl/sharedStrings.xml";

    /// <summary>Tope de defectos que viajan en el mensaje (el resto se cuenta).</summary>
    private const int MaxDefectosEnMensaje = 20;

    private static readonly Regex ReferenciaRegex = new("^(?<c>[A-Z]+)(?<f>\\d+)$", RegexOptions.Compiled);
    private static readonly Regex RangoRegex = new(
        "^(?<c1>[A-Z]+)(?<f1>\\d+)(?::(?<c2>[A-Z]+)(?<f2>\\d+))?$", RegexOptions.Compiled);

    /// <summary>
    /// Referencia hoja entrecomillada (<c>'Recaudo EAAB + Ciud Limp'!F21</c>) o simple
    /// (<c>CONSOLIDADO!D109</c>). Se descartan referencias externas (<c>[1]Hoja!</c>) y errores
    /// (<c>#REF!</c>), que no son hojas del paquete.
    /// </summary>
    private static readonly Regex HojaEnFormulaRegex = new(
        "'(?<q>[^']+)'!|(?<![A-Za-z0-9_#$!'\\[\\].])(?<u>[A-Za-z_][A-Za-z0-9_.]*)!",
        RegexOptions.Compiled);

    // ── API ─────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Gate del ESCRITOR: el paquete de salida debe conservar EXACTAMENTE el conjunto de fórmulas de
    /// la base que recibió (misma celda, mismo texto, mismos <c>t</c>/<c>ref</c>/<c>si</c>) y no puede
    /// crear ninguna <c>&lt;f&gt;</c>. Cero excepciones: el escritor es values-only.
    /// </summary>
    internal static void ValidarEscritura(string rutaBasePreparada, string rutaSalida)
    {
        ArgumentNullException.ThrowIfNull(rutaBasePreparada);
        ArgumentNullException.ThrowIfNull(rutaSalida);

        ValidarEscritura(File.ReadAllBytes(rutaBasePreparada), File.ReadAllBytes(rutaSalida));
    }

    /// <summary>Variante en memoria de <see cref="ValidarEscritura(string, string)"/>.</summary>
    internal static void ValidarEscritura(byte[] basePreparada, byte[] salida)
    {
        var esperado = Leer(basePreparada);
        var obtenido = Leer(salida);

        var defectos = new List<string>();
        CompararHojas(esperado, obtenido, compararTextoFormula: true, defectos);
        ValidarAutocontenido(obtenido, defectos);
        Lanzar(defectos);
    }

    /// <summary>
    /// Gate de la PREPARACIÓN de la base (Phase 0): puede cambiar el TEXTO de una fórmula (variantes
    /// de período/quincena derivadas del dominio), pero NO puede perder celdas con <c>&lt;f&gt;</c>,
    /// crear fórmulas nuevas ni alterar <c>t</c>/<c>ref</c>/<c>si</c>. Es el invariante que el P0 del
    /// Plan 58 rompía.
    /// </summary>
    internal static void ValidarPreparacion(string rutaPlantilla, string rutaBasePreparada)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(rutaBasePreparada);

        ValidarPreparacion(File.ReadAllBytes(rutaPlantilla), File.ReadAllBytes(rutaBasePreparada));
    }

    /// <summary>Variante en memoria de <see cref="ValidarPreparacion(string, string)"/>.</summary>
    internal static void ValidarPreparacion(byte[] plantilla, byte[] basePreparada)
    {
        var esperado = Leer(plantilla);
        var obtenido = Leer(basePreparada);

        var defectos = new List<string>();
        CompararHojas(esperado, obtenido, compararTextoFormula: false, defectos);
        ValidarAutocontenido(obtenido, defectos);
        Lanzar(defectos);
    }

    /// <summary>
    /// Gate autocontenido: valida un paquete sin referencia (grupos shared, duplicadas, orden,
    /// sharedStrings, hojas inexistentes). Útil para auditar cualquier workbook por bytes.
    /// </summary>
    internal static void ValidarPaquete(string ruta)
    {
        ArgumentNullException.ThrowIfNull(ruta);

        ValidarPaquete(File.ReadAllBytes(ruta));
    }

    /// <summary>Variante en memoria de <see cref="ValidarPaquete(string)"/>.</summary>
    internal static void ValidarPaquete(byte[] paquete)
    {
        var leido = Leer(paquete);
        var defectos = new List<string>();
        ValidarAutocontenido(leido, defectos);
        Lanzar(defectos);
    }

    // ── Comparación referencia ↔ salida ─────────────────────────────────────────────────────────

    private static void CompararHojas(
        Paquete referencia,
        Paquete salida,
        bool compararTextoFormula,
        List<string> defectos)
    {
        if (referencia.Hojas.Count != salida.Hojas.Count)
        {
            defectos.Add(
                $"CONJUNTO-HOJAS · el paquete tiene {salida.Hojas.Count} hoja(s) y la referencia {referencia.Hojas.Count}: " +
                "la escritura no puede crear ni eliminar hojas.");
            return;
        }

        for (var i = 0; i < referencia.Hojas.Count; i++)
        {
            CompararHoja(referencia.Hojas[i], salida.Hojas[i], compararTextoFormula, defectos);
        }
    }

    private static void CompararHoja(
        Hoja referencia,
        Hoja salida,
        bool compararTextoFormula,
        List<string> defectos)
    {
        // La hoja se identifica por POSICIÓN (el nombre puede cambiar de período por diseño); se
        // informa con el nombre de la SALIDA porque es el que el Ingeniero verá al abrir el archivo.
        foreach (var celdaRef in referencia.EnOrden)
        {
            if (!celdaRef.TieneFormula)
            {
                continue;
            }

            if (!salida.Celdas.TryGetValue(celdaRef.Referencia, out var celdaSalida))
            {
                defectos.Add(
                    $"FORMULA-PERDIDA · {salida.Nombre}!{celdaRef.Referencia} · la plantilla/base trae " +
                    $"<f>{celdaRef.Texto}</f> y la salida no tiene la celda. Acción: la escritura no puede eliminar celdas con fórmula.");
                continue;
            }

            if (!celdaSalida.TieneFormula)
            {
                defectos.Add(
                    $"FORMULA-PERDIDA · {salida.Nombre}!{celdaRef.Referencia} · la plantilla/base trae " +
                    $"<f>{celdaRef.Texto}</f> y la salida la tiene como valor. Acción: el escritor es values-only sobre celdas sin fórmula.");
                continue;
            }

            if (!string.Equals(celdaRef.Tipo, celdaSalida.Tipo, StringComparison.Ordinal)
                || !string.Equals(celdaRef.Rango, celdaSalida.Rango, StringComparison.Ordinal)
                || !string.Equals(celdaRef.Indice, celdaSalida.Indice, StringComparison.Ordinal))
            {
                defectos.Add(
                    $"FORMULA-ATRIBUTO-ALTERADO · {salida.Nombre}!{celdaRef.Referencia} · atributos de <f> " +
                    $"esperados [t='{celdaRef.Tipo}' ref='{celdaRef.Rango}' si='{celdaRef.Indice}'] y obtenidos " +
                    $"[t='{celdaSalida.Tipo}' ref='{celdaSalida.Rango}' si='{celdaSalida.Indice}']. " +
                    "Acción: un grupo shared sin maestro hace que Excel borre TODAS las fórmulas al reparar; se aborta sin entregar.");
                continue;
            }

            if (compararTextoFormula && !string.Equals(celdaRef.Texto, celdaSalida.Texto, StringComparison.Ordinal))
            {
                defectos.Add(
                    $"FORMULA-ALTERADA · {salida.Nombre}!{celdaRef.Referencia} · esperado '{celdaRef.Texto}' y " +
                    $"obtenido '{celdaSalida.Texto}'. Acción: la escritura no escribe <f> (Plan 58, G-FORMULA).");
            }
        }

        // Cero <f> escritas por nosotros: ninguna fórmula puede aparecer donde la referencia no la tenía.
        foreach (var celdaSalida in salida.EnOrden)
        {
            if (!celdaSalida.TieneFormula)
            {
                continue;
            }

            if (!referencia.Celdas.TryGetValue(celdaSalida.Referencia, out var celdaRef) || !celdaRef.TieneFormula)
            {
                defectos.Add(
                    $"FORMULA-CREADA · {salida.Nombre}!{celdaSalida.Referencia} · la salida trae <f>{celdaSalida.Texto}</f> " +
                    "y la referencia no. Acción: la App no escribe fórmulas; solo valores.");
            }
        }
    }

    // ── Gate autocontenido ──────────────────────────────────────────────────────────────────────

    private static void ValidarAutocontenido(Paquete paquete, List<string> defectos)
    {
        var nombresHoja = paquete.Hojas.Select(h => h.Nombre).ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var hoja in paquete.Hojas)
        {
            ValidarDuplicadas(hoja, defectos);
            ValidarOrden(hoja, defectos);
            ValidarGruposShared(hoja, defectos);
            ValidarSharedStrings(hoja, paquete.TotalSharedStrings, defectos);
            ValidarReferenciasHoja(hoja, nombresHoja, defectos);
        }

        foreach (var (donde, texto) in paquete.NombresDefinidos)
        {
            ValidarReferenciasHoja(donde, texto, nombresHoja, defectos);
        }
    }

    private static void ValidarDuplicadas(Hoja hoja, List<string> defectos)
    {
        foreach (var referencia in hoja.Duplicadas)
        {
            defectos.Add(
                $"CELDA-DUPLICADA · {hoja.Nombre}!{referencia} · la referencia aparece más de una vez en sheetData. " +
                "Acción: el paquete está corrupto; Excel la descartaría al abrir.");
        }
    }

    private static void ValidarOrden(Hoja hoja, List<string> defectos)
    {
        var filaPrevia = 0;
        var columnaPrevia = 0;
        foreach (var celda in hoja.EnOrden)
        {
            if (celda.Fila < filaPrevia || (celda.Fila == filaPrevia && celda.Columna <= columnaPrevia))
            {
                defectos.Add(
                    $"FILA-DESORDENADA · {hoja.Nombre}!{celda.Referencia} · las filas/celdas no están en orden " +
                    $"creciente (fila {celda.Fila}, columna {celda.Columna}). Acción: el formato exige orden monótono.");
                break;
            }

            filaPrevia = celda.Fila;
            columnaPrevia = celda.Columna;
        }
    }

    private static void ValidarGruposShared(Hoja hoja, List<string> defectos)
    {
        var grupos = hoja.EnOrden
            .Where(c => c.TieneFormula && !string.IsNullOrEmpty(c.Indice))
            .GroupBy(c => c.Indice!, StringComparer.Ordinal);

        foreach (var grupo in grupos)
        {
            var celdas = grupo.ToList();
            var maestros = celdas
                .Where(c => string.Equals(c.Tipo, "shared", StringComparison.Ordinal) && !string.IsNullOrEmpty(c.Rango))
                .ToList();
            var listado = string.Join(",", celdas.Select(c => c.Referencia).OrderBy(r => r, StringComparer.Ordinal));

            if (maestros.Count != 1)
            {
                defectos.Add(
                    $"GRUPO-SHARED · {hoja.Nombre}!si={grupo.Key} · el grupo tiene {maestros.Count} maestro(s) con " +
                    $"t=\"shared\" y ref (celdas: {listado}). Acción: un grupo shared sin exactamente un maestro hace que " +
                    "Excel elimine las fórmulas al reparar el archivo; la escritura no puede tocar los atributos de <f>.");
                continue;
            }

            var maestro = maestros[0];
            if (!RangoCubre(maestro.Rango!, celdas))
            {
                defectos.Add(
                    $"GRUPO-SHARED · {hoja.Nombre}!si={grupo.Key} · el ref del maestro {maestro.Referencia} " +
                    $"('{maestro.Rango}') no cubre todas las celdas del grupo ({listado}). " +
                    "Acción: el grupo quedaría inconsistente y Excel repararía el libro.");
            }
        }
    }

    private static bool RangoCubre(string rango, IReadOnlyList<Celda> celdas)
    {
        var m = RangoRegex.Match(rango);
        if (!m.Success)
        {
            return false;
        }

        var col1 = ColumnaAIndice(m.Groups["c1"].Value);
        var fil1 = int.Parse(m.Groups["f1"].Value, CultureInfo.InvariantCulture);
        var col2 = m.Groups["c2"].Success ? ColumnaAIndice(m.Groups["c2"].Value) : col1;
        var fil2 = m.Groups["f2"].Success ? int.Parse(m.Groups["f2"].Value, CultureInfo.InvariantCulture) : fil1;

        var minCol = Math.Min(col1, col2);
        var maxCol = Math.Max(col1, col2);
        var minFil = Math.Min(fil1, fil2);
        var maxFil = Math.Max(fil1, fil2);

        return celdas.All(c => c.Columna >= minCol && c.Columna <= maxCol && c.Fila >= minFil && c.Fila <= maxFil);
    }

    private static void ValidarSharedStrings(Hoja hoja, int total, List<string> defectos)
    {
        foreach (var celda in hoja.EnOrden)
        {
            if (!string.Equals(celda.TipoCelda, "s", StringComparison.Ordinal))
            {
                continue;
            }

            if (!int.TryParse(celda.Valor, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indice)
                || indice < 0
                || indice >= total)
            {
                defectos.Add(
                    $"SHAREDSTRING-FUERA-DE-RANGO · {hoja.Nombre}!{celda.Referencia} · índice '{celda.Valor}' con " +
                    $"{total} entrada(s) en sharedStrings. Acción: la tabla de textos quedó inconsistente; Excel la repararía.");
            }
        }
    }

    private static void ValidarReferenciasHoja(Hoja hoja, IReadOnlySet<string> nombresHoja, List<string> defectos)
    {
        foreach (var celda in hoja.EnOrden)
        {
            if (celda.TieneFormula && celda.Texto is not null)
            {
                ValidarReferenciasHoja($"{hoja.Nombre}!{celda.Referencia}", celda.Texto, nombresHoja, defectos);
            }
        }
    }

    private static void ValidarReferenciasHoja(
        string donde,
        string texto,
        IReadOnlySet<string> nombresHoja,
        List<string> defectos)
    {
        foreach (Match match in HojaEnFormulaRegex.Matches(texto))
        {
            var nombre = match.Groups["q"].Success ? match.Groups["q"].Value : match.Groups["u"].Value;
            if (string.IsNullOrWhiteSpace(nombre)
                || nombre.Contains('[', StringComparison.Ordinal)
                || nombre.Contains('#', StringComparison.Ordinal))
            {
                continue; // referencia externa ([1]Hoja) o error (#REF!): no es una hoja del paquete.
            }

            if (!nombresHoja.Contains(nombre))
            {
                defectos.Add(
                    $"HOJA-INEXISTENTE · {donde} · la referencia '{nombre}!' apunta a una hoja que no existe en el " +
                    "workbook. Acción: la fórmula quedaría como #REF!; se aborta antes de entregar.");
            }
        }
    }

    // ── Lectura por bytes ──────────────────────────────────────────────────────────────────────

    private static Paquete Leer(byte[] contenido)
    {
        using var memoria = new MemoryStream(contenido, writable: false);
        using var zip = new ZipArchive(memoria, ZipArchiveMode.Read);

        var entradas = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entrada in zip.Entries)
        {
            entradas[entrada.FullName] = entrada;
        }

        var paquete = new Paquete();
        var workbook = Cargar(entradas, PartWorkbook);
        var relaciones = Cargar(entradas, PartRels);
        var mapaRelaciones = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var rel in relaciones.Root?.Elements() ?? [])
        {
            var id = (string?)rel.Attribute("Id");
            var target = (string?)rel.Attribute("Target");
            if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
            {
                mapaRelaciones[id] = NormalizarParte(target);
            }
        }

        paquete.TotalSharedStrings = entradas.ContainsKey(PartShared)
            ? Cargar(entradas, PartShared).Root?.Elements(Ns + "si").Count() ?? 0
            : 0;

        foreach (var definido in workbook.Root?.Element(Ns + "definedNames")?.Elements(Ns + "definedName") ?? [])
        {
            paquete.NombresDefinidos.Add(($"definedName '{definido.Attribute("name")?.Value}'", definido.Value));
        }

        var hojas = workbook.Root?.Element(Ns + "sheets")?.Elements(Ns + "sheet") ?? [];
        foreach (var hoja in hojas)
        {
            var nombre = (string?)hoja.Attribute("name") ?? string.Empty;
            var rid = (string?)hoja.Attribute(RelNs + "id");
            var parte = rid is not null && mapaRelaciones.TryGetValue(rid, out var encontrada)
                ? encontrada
                : throw new CalculoInvalidoException(
                    CodigoError.Integridad,
                    $"ERR-INTEGRIDAD · CONJUNTO-HOJAS · '{nombre}' no declara una relación de hoja válida en {PartRels}.");

            paquete.Hojas.Add(LeerHoja(entradas, nombre, parte));
        }

        return paquete;
    }

    private static Hoja LeerHoja(
        IReadOnlyDictionary<string, ZipArchiveEntry> entradas,
        string nombre,
        string parte)
    {
        var hoja = new Hoja { Nombre = nombre, Parte = parte };
        var documento = Cargar(entradas, parte);
        var sheetData = documento.Root?.Element(Ns + "sheetData");
        if (sheetData is null)
        {
            return hoja;
        }

        foreach (var fila in sheetData.Elements(Ns + "row"))
        {
            foreach (var celda in fila.Elements(Ns + "c"))
            {
                var referencia = (string?)celda.Attribute("r") ?? string.Empty;
                var match = ReferenciaRegex.Match(referencia);
                if (!match.Success)
                {
                    hoja.Duplicadas.Add(referencia);
                    continue;
                }

                var formula = celda.Element(Ns + "f");
                var valor = celda.Element(Ns + "v");
                var snapshot = new Celda(
                    referencia,
                    ColumnaAIndice(match.Groups["c"].Value),
                    int.Parse(match.Groups["f"].Value, CultureInfo.InvariantCulture),
                    formula is not null,
                    formula?.Value,
                    (string?)formula?.Attribute("t"),
                    (string?)formula?.Attribute("ref"),
                    (string?)formula?.Attribute("si"),
                    valor?.Value,
                    (string?)celda.Attribute("t"));

                if (!hoja.Celdas.TryAdd(referencia, snapshot))
                {
                    hoja.Duplicadas.Add(referencia);
                }

                hoja.EnOrden.Add(snapshot);
            }
        }

        return hoja;
    }

    private static XDocument Cargar(IReadOnlyDictionary<string, ZipArchiveEntry> entradas, string parte)
    {
        if (!entradas.TryGetValue(parte, out var entrada))
        {
            throw new CalculoInvalidoException(
                CodigoError.Integridad,
                $"ERR-INTEGRIDAD · PAQUETE-INCOMPLETO · falta la parte '{parte}' en el zip. Acción: el paquete generado está incompleto.");
        }

        using var stream = entrada.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static string NormalizarParte(string target)
    {
        var limpio = target.Replace('\\', '/');
        if (limpio.StartsWith('/'))
        {
            return limpio.TrimStart('/');
        }

        return limpio.StartsWith("xl/", StringComparison.OrdinalIgnoreCase) ? limpio : "xl/" + limpio;
    }

    private static int ColumnaAIndice(string letras)
    {
        var indice = 0;
        foreach (var letra in letras)
        {
            indice = (indice * 26) + (letra - 'A' + 1);
        }

        return indice;
    }

    // ── Salida ─────────────────────────────────────────────────────────────────────────────────

    private static void Lanzar(List<string> defectos)
    {
        if (defectos.Count == 0)
        {
            return;
        }

        var visibles = defectos.Take(MaxDefectosEnMensaje).ToList();
        var resto = defectos.Count - visibles.Count;
        var detalle = string.Join(" | ", visibles);
        if (resto > 0)
        {
            detalle += $" | ... y {resto} defecto(s) más.";
        }

        throw new CalculoInvalidoException(
            CodigoError.Integridad,
            $"El paquete generado no es íntegro ({defectos.Count} defecto(s); el archivo NO se entrega porque Excel lo " +
            $"repararía borrando fórmulas): {detalle}");
    }

    private sealed record Celda(
        string Referencia,
        int Columna,
        int Fila,
        bool TieneFormula,
        string? Texto,
        string? Tipo,
        string? Rango,
        string? Indice,
        string? Valor,
        string? TipoCelda);

    private sealed class Hoja
    {
        public string Nombre { get; init; } = string.Empty;

        public string Parte { get; init; } = string.Empty;

        public Dictionary<string, Celda> Celdas { get; } = new(StringComparer.OrdinalIgnoreCase);

        public List<Celda> EnOrden { get; } = [];

        public List<string> Duplicadas { get; } = [];
    }

    private sealed class Paquete
    {
        public List<Hoja> Hojas { get; } = [];

        public int TotalSharedStrings { get; set; }

        public List<(string Donde, string Texto)> NombresDefinidos { get; } = [];
    }
}
