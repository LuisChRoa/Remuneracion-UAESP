using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace Herramientas.PreparadorBasePeriodo;

/// <summary>
/// Herramienta OFFLINE one-shot (Plan 30, T3, Decision D-B): convierte la base de agosto
/// <c>Docs/Prueba2/Plantilla_Remuneracion.xlsx</c> (copia de la plantilla Q2 de julio, sufijo
/// 2026072) en la base canónica <c>Docs/Prueba2/Plantilla_Remuneracion_2026082.xlsx</c>,
/// internamente consistente con el manual del administrativo.
///
/// Qué hace (siempre sobre una COPIA; el origen NUNCA se modifica):
///   (a) renombra los 3 &lt;sheet name&gt; de <c>xl/workbook.xml</c> (rIds/sheetIds intactos);
///   (b) reescribe el token de período en las 11 fórmulas (&lt;f&gt;) que lo referencian;
///   (c) actualiza <c>TitlesOfParts</c> en <c>docProps/app.xml</c>;
///   (d) sella metadatos de período + columnas Q1 (D:E) copiando del manual celda por celda
///       (allow-list explícita; solo destinos SIN fórmula; cero datos inventados);
///   (e) autoverificación zip+XML embebida (hojas nuevas, 0 token viejo, 11 &lt;f&gt; nuevas,
///       definedNames limpios, metadatos == manual).
///
/// Método: BCL puro (<see cref="ZipArchive"/> + <see cref="Regex"/>) — mismo método del T0 del
/// Plan 30. PROHIBIDO OpenXML/ExcelDataReader/Excel/COM. No forma parte del pipeline de runtime.
/// </summary>
internal static class Program
{
    private const string TokenPeriodoViejo = "2026072";
    private const string TokenPeriodoNuevo = "2026082";
    private const string TokenInformeViejo = "202607-2";
    private const string TokenInformeNuevo = "202608-2";

    private const string NombreDetRetriViejo = "DetRetri2026072";
    private const string NombreDetRetriNuevo = "DetRetri2026082";
    private const string NombreDetValiViejo = "DetValiRetri2026072";
    private const string NombreDetValiNuevo = "DetValiRetri2026082";
    private const string NombreInformeViejo = "Informe AFaseo Recaudo 202607-2";
    private const string NombreInformeNuevo = "Informe AFaseo Recaudo 202608-2";

    private const string PartWorkbook = "xl/workbook.xml";
    private const string PartApp = "docProps/app.xml";
    private const string PartShared = "xl/sharedStrings.xml";
    private const string PartConsolidado = "xl/worksheets/sheet18.xml";
    private const string PartDetRetri = "xl/worksheets/sheet36.xml";

    /// <summary>Hojas con columnas Q1 (D:E) que el flujo Q2 no escribe: se heredan del manual (D-E).</summary>
    private static readonly string[] HojasRecaudoQ1 =
    [
        "xl/worksheets/sheet1.xml", // Recaudo EAAB Reciprocidad
        "xl/worksheets/sheet2.xml", // Recaudo ENEL
        "xl/worksheets/sheet3.xml", // Recaudo ENERBIT
        "xl/worksheets/sheet5.xml"  // Recaudo Directa Occidente
    ];

    /// <summary>Allow-list de metadatos (código de período / fechas) a copiar del manual (destino sin fórmula).</summary>
    private static readonly string[] MetadataConsolidado = ["N3", "D6", "G7", "K7"];
    private static readonly string[] MetadataDetRetri = ["G7", "J7"];

    private static readonly Regex TokenPeriodo = new(@"(?<![0-9])2026072(?![0-9])", RegexOptions.Compiled);
    private static readonly Regex CeldaRegex = new(
        "<c r=\"(?<ref>[A-Z]+\\d+)\"(?<attrs>[^>]*?)(?:/>|>(?<inner>.*?)</c>)",
        RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex FormulaRegex = new(
        "<f(?<attrs>[^>]*?)>(?<texto>[^<]*)</f>",
        RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex ValorRegex = new("<v>(?<v>[^<]*)</v>", RegexOptions.Compiled);
    private static readonly Regex SiRegex = new("<si>(?<cuerpo>.*?)</si>", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex TextoRegex = new("<t[^>]*>(?<t>[^<]*)</t>", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex SstHeaderRegex = new("<sst[^>]*>", RegexOptions.Compiled);
    private static readonly Regex HojaRegex = new("^xl/worksheets/sheet\\d+\\.xml$", RegexOptions.Compiled);

    private static readonly UTF8Encoding Utf8SinBom = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: false);

    private static int Main(string[] args)
    {
        try
        {
            Ejecutar(args);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine("FALLO: " + ex.Message);
            return 1;
        }
    }

    private static void Ejecutar(string[] args)
    {
        var raiz = Raiz();
        var origenPath = Argumento(args, "--origen")
            ?? Path.Combine(raiz, "Docs", "Prueba2", "Plantilla_Remuneracion.xlsx");
        var manualPath = Argumento(args, "--manual")
            ?? Path.Combine(raiz, "Docs", "Prueba2", "Resultado", "Resultado Manual por el administrativo", "Remuneracion 202608-2 Total_7721.xlsx");
        var salidaPath = Argumento(args, "--salida")
            ?? Path.Combine(raiz, "Docs", "Prueba2", "Plantilla_Remuneracion_2026082.xlsx");

        Exigir(File.Exists(origenPath), $"No existe la base origen: {origenPath}");
        Exigir(File.Exists(manualPath), $"No existe el manual: {manualPath}");

        Console.WriteLine("=== Preparador de base canónica por período (Plan 30, T3) ===");
        Console.WriteLine($"Origen : {origenPath}");
        Console.WriteLine($"Manual : {manualPath}");
        Console.WriteLine($"Salida : {salidaPath}");
        Console.WriteLine();

        // Backup obligatorio del origen a %TEMP% antes de tocar nada.
        var backupDir = Path.Combine(Path.GetTempPath(), "RemuneracionBackupPlan30");
        Directory.CreateDirectory(backupDir);
        var backup = Path.Combine(backupDir, $"Plantilla_Remuneracion_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        File.Copy(origenPath, backup, overwrite: false);
        Console.WriteLine($"[backup] {backup}");

        using var origenZip = ZipFile.OpenRead(origenPath);
        using var manualZip = ZipFile.OpenRead(manualPath);

        var entradasOrigen = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var e in origenZip.Entries)
        {
            Exigir(!entradasOrigen.ContainsKey(e.FullName), $"Entrada ZIP duplicada: {e.FullName}");
            entradasOrigen[e.FullName] = e;
        }

        var sharedManual = ResolverSharedStrings(LeerTexto(manualZip.GetEntry(PartShared)
            ?? throw new InvalidOperationException("El manual no tiene xl/sharedStrings.xml.")));

        // (0) sharedStrings base: token de período reemplazado primero; luego append de textos nuevos.
        var sharedBase = TokenPeriodo.Replace(LeerTexto(entradasOrigen[PartShared]), TokenPeriodoNuevo);
        var editorShared = new SharedStringsEditor(sharedBase);

        var partes = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        var resumen = new List<string>();

        foreach (var entrada in origenZip.Entries)
        {
            var nombre = entrada.FullName;
            if (string.Equals(nombre, PartShared, StringComparison.OrdinalIgnoreCase))
            {
                continue; // se emite al final, con los textos agregados.
            }

            var bytes = LeerBytes(entrada);
            if (!EsParteTextoTransformable(nombre))
            {
                partes[nombre] = bytes;
                continue;
            }

            var texto = Encoding.UTF8.GetString(bytes);
            var original = texto;
            var nToken = TokenPeriodo.Matches(texto).Count;
            var nFormulaVieja = ContarFormulasCon(texto, NombreDetRetriViejo);
            var nFormulaTotal = FormulaRegex.Matches(texto).Count;

            texto = TokenPeriodo.Replace(texto, TokenPeriodoNuevo);

            if (string.Equals(nombre, PartWorkbook, StringComparison.OrdinalIgnoreCase)
                || string.Equals(nombre, PartApp, StringComparison.OrdinalIgnoreCase))
            {
                texto = texto.Replace(TokenInformeViejo, TokenInformeNuevo, StringComparison.Ordinal);
            }

            if (string.Equals(nombre, PartConsolidado, StringComparison.OrdinalIgnoreCase))
            {
                texto = SellarNumericosDesdeManual(texto, manualZip, PartConsolidado, MetadataConsolidado, sharedManual, resumen);
            }

            if (string.Equals(nombre, PartDetRetri, StringComparison.OrdinalIgnoreCase))
            {
                texto = SellarStringsDesdeManual(texto, manualZip, PartDetRetri, MetadataDetRetri, sharedManual, editorShared, sharedBase, resumen);
            }

            if (HojasRecaudoQ1.Contains(nombre, StringComparer.OrdinalIgnoreCase))
            {
                texto = CopiarColumnasQ1(texto, manualZip, nombre, sharedManual, editorShared, sharedBase, resumen);
            }

            partes[nombre] = Utf8SinBom.GetBytes(texto);

            if (!string.Equals(texto, original, StringComparison.Ordinal) || nToken > 0)
            {
                resumen.Add($"  {nombre}: token {TokenPeriodoViejo}->{TokenPeriodoNuevo} x{nToken}; <f> con token viejo x{nFormulaVieja} (total <f> {nFormulaTotal})");
            }
        }

        // sharedStrings final (token reemplazado + textos nuevos agregados).
        partes[PartShared] = Utf8SinBom.GetBytes(editorShared.Xml());

        Console.WriteLine("[cambios]");
        foreach (var linea in resumen)
        {
            Console.WriteLine(linea);
        }

        Console.WriteLine();
        Console.WriteLine("[autoverificación embebida]");
        Verificar(partes, entradasOrigen, manualZip, sharedManual, editorShared, out var lineasVerificacion);
        foreach (var linea in lineasVerificacion)
        {
            Console.WriteLine("  " + linea);
        }

        // Escritura atómica: zip temporal -> verificación de apertura -> mover a la ruta final.
        var tmpPath = salidaPath + ".tmp";
        var tmpDir = Path.GetDirectoryName(salidaPath);
        if (!string.IsNullOrEmpty(tmpDir))
        {
            Directory.CreateDirectory(tmpDir);
        }

        if (File.Exists(tmpPath))
        {
            File.Delete(tmpPath);
        }

        using (var fs = new FileStream(tmpPath, FileMode.Create, FileAccess.Write))
        using (var salidaZip = new ZipArchive(fs, ZipArchiveMode.Create))
        {
            // Se recorre el orden original para preservar [Content_Types].xml como primera entrada.
            foreach (var entrada in origenZip.Entries)
            {
                var nueva = salidaZip.CreateEntry(entrada.FullName, CompressionLevel.Optimal);
                using var stream = nueva.Open();
                stream.Write(partes[entrada.FullName]);
            }
        }

        // Verificación estructural final: el zip producido se abre y conserva las entradas esperadas.
        using (var relectura = ZipFile.OpenRead(tmpPath))
        {
            Exigir(relectura.Entries.Count == partes.Count,
                $"El zip producido tiene {relectura.Entries.Count} entradas; se esperaban {partes.Count}.");
            foreach (var nombre in new[] { PartWorkbook, PartApp, PartShared, PartConsolidado, PartDetRetri })
            {
                Exigir(relectura.GetEntry(nombre) is not null, $"El zip producido no contiene {nombre}.");
            }
        }

        if (File.Exists(salidaPath))
        {
            File.Delete(salidaPath);
        }

        File.Move(tmpPath, salidaPath);
        Console.WriteLine();
        Console.WriteLine($"OK: base canónica escrita en {salidaPath}");
    }

    // ── (d) sellado desde el manual ──────────────────────────────────────────────────────────

    private static string SellarNumericosDesdeManual(
        string xml,
        ZipArchive manualZip,
        string parteManual,
        IReadOnlyList<string> referencias,
        IReadOnlyList<string> sharedManual,
        List<string> resumen)
    {
        var manualXml = LeerTexto(manualZip.GetEntry(parteManual)
            ?? throw new InvalidOperationException($"El manual no tiene {parteManual}."));
        var celdasManual = ParsearCeldas(manualXml, sharedManual);
        foreach (var referencia in referencias)
        {
            if (!celdasManual.TryGetValue(referencia, out var manual) || manual.Valor is null)
            {
                resumen.Add($"  {parteManual}!{referencia}: sin dato en el manual -> omitido");
                continue;
            }

            xml = SetValorNumerico(xml, referencia, manual.Valor, out var aplicado, out var motivo);
            resumen.Add(aplicado
                ? $"  {parteManual}!{referencia}: {manual.Valor} (desde manual)"
                : $"  {parteManual}!{referencia}: omitido ({motivo})");
        }

        return xml;
    }

    private static string SellarStringsDesdeManual(
        string xml,
        ZipArchive manualZip,
        string parteManual,
        IReadOnlyList<string> referencias,
        IReadOnlyList<string> sharedManual,
        SharedStringsEditor editor,
        string sharedBase,
        List<string> resumen)
    {
        var manualXml = LeerTexto(manualZip.GetEntry(parteManual)
            ?? throw new InvalidOperationException($"El manual no tiene {parteManual}."));
        var celdasManual = ParsearCeldas(manualXml, sharedManual);
        var celdasBase = ParsearCeldas(xml, ResolverSharedStrings(sharedBase));
        foreach (var referencia in referencias)
        {
            if (!celdasManual.TryGetValue(referencia, out var manual) || manual.Texto is null)
            {
                resumen.Add($"  {parteManual}!{referencia}: sin dato en el manual -> omitido");
                continue;
            }

            if (celdasBase.TryGetValue(referencia, out var baseCelda)
                && string.Equals(baseCelda.Texto, manual.Texto, StringComparison.Ordinal))
            {
                resumen.Add($"  {parteManual}!{referencia}: ya coincide con el manual ('{manual.Texto}')");
                continue;
            }

            var indice = editor.Agregar(manual.Texto);
            xml = SetValorSharedString(xml, referencia, indice, out var aplicado, out var motivo);
            resumen.Add(aplicado
                ? $"  {parteManual}!{referencia}: '{manual.Texto}' (ss#{indice})"
                : $"  {parteManual}!{referencia}: omitido ({motivo})");
        }

        return xml;
    }

    private static string CopiarColumnasQ1(
        string xml,
        ZipArchive manualZip,
        string parteManual,
        IReadOnlyList<string> sharedManual,
        SharedStringsEditor editor,
        string sharedBase,
        List<string> resumen)
    {
        var manualXml = LeerTexto(manualZip.GetEntry(parteManual)
            ?? throw new InvalidOperationException($"El manual no tiene {parteManual}."));
        var celdasManual = ParsearCeldas(manualXml, sharedManual);
        var celdasBase = ParsearCeldas(xml, ResolverSharedStrings(sharedBase));

        var referencias = celdasManual.Keys
            .Where(r => r.StartsWith('D') || r.StartsWith('E'))
            .OrderBy(r => r, StringComparer.Ordinal)
            .ToList();

        var copiadas = 0;
        var omitidas = 0;
        foreach (var referencia in referencias)
        {
            var manual = celdasManual[referencia];
            if (manual.TieneFormula || manual.Valor is null)
            {
                continue; // solo celdas con dato (no fórmula) del manual.
            }

            if (!celdasBase.TryGetValue(referencia, out var baseCelda))
            {
                omitidas++;
                continue; // destino inexistente: no se crea estructura.
            }

            if (baseCelda.TieneFormula)
            {
                omitidas++; // destino con fórmula: NUNCA se toca.
                continue;
            }

            if (string.Equals(baseCelda.Valor, manual.Valor, StringComparison.Ordinal)
                && string.Equals(baseCelda.Texto, manual.Texto, StringComparison.Ordinal))
            {
                continue; // ya coincide.
            }

            if (manual.EsTexto)
            {
                var indice = editor.Agregar(manual.Texto ?? string.Empty);
                xml = SetValorSharedString(xml, referencia, indice, out var ok, out _);
                if (ok)
                {
                    copiadas++;
                }
                else
                {
                    omitidas++;
                }
            }
            else
            {
                xml = SetValorNumerico(xml, referencia, manual.Valor, out var ok, out _);
                if (ok)
                {
                    copiadas++;
                }
                else
                {
                    omitidas++;
                }
            }
        }

        resumen.Add($"  {parteManual}: columnas Q1 (D:E) copiadas={copiadas}; omitidas={omitidas}");
        return xml;
    }

    // ── cirugía de celdas (string-safe, preserva todo el resto del XML) ──────────────────────

    private static string SetValorNumerico(string xml, string referencia, string valor, out bool aplicado, out string motivo)
    {
        var match = BuscarCelda(xml, referencia);
        if (match is null)
        {
            aplicado = false;
            motivo = "celda inexistente en la base";
            return xml;
        }

        var (attrs, inner, autocierre, tieneFormula) = AnalizarCelda(match);
        if (tieneFormula)
        {
            aplicado = false;
            motivo = "destino con fórmula (NUNCA se sobrescribe)";
            return xml;
        }

        attrs = QuitarAtributoTipo(attrs);
        string nueva;
        if (autocierre)
        {
            nueva = $"<c r=\"{referencia}\"{attrs}><v>{valor}</v></c>";
        }
        else
        {
            var cuerpo = inner ?? string.Empty;
            cuerpo = ValorRegex.IsMatch(cuerpo)
                ? ValorRegex.Replace(cuerpo, $"<v>{valor}</v>")
                : cuerpo + $"<v>{valor}</v>";
            nueva = $"<c r=\"{referencia}\"{attrs}>{cuerpo}</c>";
        }

        aplicado = true;
        motivo = "ok";
        return xml[..match.Index] + nueva + xml[(match.Index + match.Length)..];
    }

    private static string SetValorSharedString(string xml, string referencia, int indice, out bool aplicado, out string motivo)
    {
        var match = BuscarCelda(xml, referencia);
        if (match is null)
        {
            aplicado = false;
            motivo = "celda inexistente en la base";
            return xml;
        }

        var (attrs, _, _, tieneFormula) = AnalizarCelda(match);
        if (tieneFormula)
        {
            aplicado = false;
            motivo = "destino con fórmula (NUNCA se sobrescribe)";
            return xml;
        }

        attrs = QuitarAtributoTipo(attrs);
        var nueva = $"<c r=\"{referencia}\"{attrs} t=\"s\"><v>{indice}</v></c>";
        aplicado = true;
        motivo = "ok";
        return xml[..match.Index] + nueva + xml[(match.Index + match.Length)..];
    }

    private static Match? BuscarCelda(string xml, string referencia)
    {
        foreach (Match match in CeldaRegex.Matches(xml))
        {
            if (string.Equals(match.Groups["ref"].Value, referencia, StringComparison.Ordinal))
            {
                return match;
            }
        }

        return null;
    }

    private static (string Attrs, string? Inner, bool Autocierre, bool TieneFormula) AnalizarCelda(Match match)
    {
        var attrs = match.Groups["attrs"].Value;
        var autocierre = !match.Groups["inner"].Success;
        var inner = match.Groups["inner"].Success ? match.Groups["inner"].Value : null;
        var tieneFormula = inner is not null && inner.Contains("<f", StringComparison.Ordinal);
        return (attrs, inner, autocierre, tieneFormula);
    }

    private static string QuitarAtributoTipo(string attrs) =>
        Regex.Replace(attrs, "\\s+t=\"[^\"]*\"", string.Empty, RegexOptions.Compiled);

    // ── parseo de celdas ─────────────────────────────────────────────────────────────────────

    private sealed record InfoCelda(bool TieneFormula, bool EsTexto, string? Valor, string? Texto);

    private static Dictionary<string, InfoCelda> ParsearCeldas(string xml, IReadOnlyList<string> sharedStrings)
    {
        var celdas = new Dictionary<string, InfoCelda>(StringComparer.Ordinal);
        foreach (Match match in CeldaRegex.Matches(xml))
        {
            var referencia = match.Groups["ref"].Value;
            var (attrs, inner, _, tieneFormula) = AnalizarCelda(match);
            string? valor = null;
            var esTexto = attrs.Contains("t=\"s\"", StringComparison.Ordinal)
                || attrs.Contains("t=\"inlineStr\"", StringComparison.Ordinal)
                || attrs.Contains("t=\"str\"", StringComparison.Ordinal);
            string? texto = null;

            if (inner is not null)
            {
                var valorMatch = ValorRegex.Match(inner);
                if (valorMatch.Success)
                {
                    valor = valorMatch.Groups["v"].Value;
                }

                if (attrs.Contains("t=\"s\"", StringComparison.Ordinal)
                    && int.TryParse(valor, NumberStyles.None, CultureInfo.InvariantCulture, out var indice)
                    && indice >= 0 && indice < sharedStrings.Count)
                {
                    texto = sharedStrings[indice];
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

    private static List<string> ResolverSharedStrings(string xml)
    {
        var lista = new List<string>();
        foreach (Match match in SiRegex.Matches(xml))
        {
            lista.Add(string.Concat(TextoRegex.Matches(match.Groups["cuerpo"].Value).Select(m => m.Groups["t"].Value)));
        }

        return lista;
    }

    /// <summary>
    /// Editor de sharedStrings para agregar (deduplicando) los textos heredados del manual.
    /// Mantiene el XML original intacto y solo inserta al final + ajusta count/uniqueCount.
    /// </summary>
    private sealed class SharedStringsEditor
    {
        private readonly string _xmlOriginal;
        private readonly List<string> _textos;
        private readonly StringBuilder _agregados = new();
        private int _count;
        private int _uniqueCount;

        public SharedStringsEditor(string xml)
        {
            _xmlOriginal = xml;
            _textos = ResolverSharedStrings(xml);
            _uniqueCount = _textos.Count;
            var header = SstHeaderRegex.Match(xml);
            _count = ExtraerAtributoEntero(header.Success ? header.Value : string.Empty, "count", _uniqueCount);
        }

        public int Agregar(string texto)
        {
            var existente = _textos.FindIndex(t => string.Equals(t, texto, StringComparison.Ordinal));
            if (existente >= 0)
            {
                return existente;
            }

            var indice = _textos.Count;
            _textos.Add(texto);
            _agregados.Append("<si><t xml:space=\"preserve\">").Append(Escapar(texto)).Append("</t></si>");
            _count++;
            _uniqueCount++;
            return indice;
        }

        public string Xml()
        {
            if (_agregados.Length == 0)
            {
                return _xmlOriginal;
            }

            var nuevo = _xmlOriginal.Replace("</sst>", _agregados + "</sst>", StringComparison.Ordinal);
            var header = SstHeaderRegex.Match(nuevo);
            if (header.Success)
            {
                var tag = Regex.Replace(header.Value, "count=\"\\d+\"", $"count=\"{_count}\"");
                tag = Regex.Replace(tag, "uniqueCount=\"\\d+\"", $"uniqueCount=\"{_uniqueCount}\"");
                nuevo = nuevo[..header.Index] + tag + nuevo[(header.Index + header.Length)..];
            }

            return nuevo;
        }

        private static int ExtraerAtributoEntero(string tag, string nombre, int porDefecto)
        {
            var match = Regex.Match(tag, nombre + "=\"(\\d+)\"");
            return match.Success && int.TryParse(match.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var valor)
                ? valor
                : porDefecto;
        }
    }

    private static string Escapar(string texto) => texto
        .Replace("&", "&amp;", StringComparison.Ordinal)
        .Replace("<", "&lt;", StringComparison.Ordinal)
        .Replace(">", "&gt;", StringComparison.Ordinal);

    // ── autoverificación (e) ─────────────────────────────────────────────────────────────────

    private static void Verificar(
        IReadOnlyDictionary<string, byte[]> partes,
        IReadOnlyDictionary<string, ZipArchiveEntry> entradasOrigen,
        ZipArchive manualZip,
        IReadOnlyList<string> sharedManual,
        SharedStringsEditor editor,
        out List<string> lineas)
    {
        lineas = [];

        string TextoParte(string nombre) => Encoding.UTF8.GetString(partes[nombre]);

        // (1) 0 apariciones del token viejo en worksheets + workbook (+ app + sharedStrings).
        var partesConToken = new List<string>();
        foreach (var par in partes)
        {
            var esSheet = HojaRegex.IsMatch(par.Key);
            if (!esSheet
                && !string.Equals(par.Key, PartWorkbook, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(par.Key, PartApp, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(par.Key, PartShared, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var texto = Encoding.UTF8.GetString(par.Value);
            if (texto.Contains(TokenPeriodoViejo, StringComparison.Ordinal)
                || texto.Contains(NombreInformeViejo, StringComparison.Ordinal))
            {
                partesConToken.Add(par.Key);
            }
        }

        Exigir(partesConToken.Count == 0,
            $"Quedaron apariciones del token viejo en: {string.Join(", ", partesConToken)}");
        lineas.Add($"0 apariciones de '{TokenPeriodoViejo}' en worksheets + workbook + app + sharedStrings: OK");

        // (2) exactamente 11 <f> con el nombre nuevo en worksheets, y 0 con el viejo.
        var formulasNuevas = 0;
        var formulasViejas = 0;
        foreach (var par in partes)
        {
            if (!HojaRegex.IsMatch(par.Key))
            {
                continue;
            }

            var texto = Encoding.UTF8.GetString(par.Value);
            formulasNuevas += ContarFormulasCon(texto, NombreDetRetriNuevo);
            formulasViejas += ContarFormulasCon(texto, NombreDetRetriViejo);
        }

        Exigir(formulasNuevas == 11, $"Se esperaban 11 <f> con '{NombreDetRetriNuevo}'; se hallaron {formulasNuevas}.");
        Exigir(formulasViejas == 0, $"Quedaron {formulasViejas} <f> con '{NombreDetRetriViejo}'.");
        lineas.Add($"11 <f> con '{NombreDetRetriNuevo}' y 0 con el token viejo: OK");

        // (3) nombres nuevos presentes en workbook.xml y app.xml; viejos ausentes.
        var workbook = TextoParte(PartWorkbook);
        var app = TextoParte(PartApp);
        foreach (var nombre in new[] { NombreDetRetriNuevo, NombreDetValiNuevo, NombreInformeNuevo })
        {
            Exigir(workbook.Contains(nombre, StringComparison.Ordinal), $"workbook.xml no contiene '{nombre}'.");
        }

        foreach (var nombre in new[] { NombreDetRetriNuevo, NombreDetValiNuevo, NombreInformeNuevo })
        {
            Exigir(app.Contains(nombre, StringComparison.Ordinal), $"app.xml (TitlesOfParts) no contiene '{nombre}'.");
        }

        lineas.Add("Hojas nuevas presentes en workbook.xml y app.xml: OK");

        // (4) rIds/sheetIds intactos.
        var hojasViejas = ExtraerHojas(LeerTexto(entradasOrigen[PartWorkbook]));
        var hojasNuevas = ExtraerHojas(workbook);
        Exigir(hojasViejas.Count == hojasNuevas.Count, "Cambió el número de hojas del workbook.");
        for (var i = 0; i < hojasViejas.Count; i++)
        {
            Exigir(string.Equals(hojasViejas[i].SheetId, hojasNuevas[i].SheetId, StringComparison.Ordinal)
                && string.Equals(hojasViejas[i].Rid, hojasNuevas[i].Rid, StringComparison.Ordinal),
                $"Cambió sheetId/rId en la posición {i}: {hojasViejas[i].SheetId}/{hojasViejas[i].Rid} vs {hojasNuevas[i].SheetId}/{hojasNuevas[i].Rid}");
        }

        lineas.Add("rIds/sheetIds intactos: OK");

        // (5) definedNames sin token viejo.
        if (workbook.Contains("<definedNames>", StringComparison.Ordinal))
        {
            var definidos = Regex.Match(workbook, "<definedNames>.*?</definedNames>", RegexOptions.Singleline);
            Exigir(!definidos.Value.Contains(TokenPeriodoViejo, StringComparison.Ordinal),
                "definedNames conserva el token viejo.");
        }

        lineas.Add("definedNames sin token viejo: OK");

        // (6) metadatos == manual.
        var manualShared = new List<string>(sharedManual);
        VerificarCeldaManual(partes, manualZip, editor, manualShared, PartConsolidado, "N3", esTexto: false, lineas);
        VerificarCeldaManual(partes, manualZip, editor, manualShared, PartConsolidado, "D6", esTexto: false, lineas);
        VerificarCeldaManual(partes, manualZip, editor, manualShared, PartConsolidado, "G7", esTexto: false, lineas);
        VerificarCeldaManual(partes, manualZip, editor, manualShared, PartDetRetri, "G7", esTexto: true, lineas);
        VerificarCeldaManual(partes, manualZip, editor, manualShared, PartDetRetri, "J7", esTexto: true, lineas);
    }

    private static void VerificarCeldaManual(
        IReadOnlyDictionary<string, byte[]> partes,
        ZipArchive manualZip,
        SharedStringsEditor editor,
        IReadOnlyList<string> sharedManual,
        string parte,
        string referencia,
        bool esTexto,
        List<string> lineas)
    {
        var salida = ParsearCeldas(Encoding.UTF8.GetString(partes[parte]), ResolverSharedStrings(editor.Xml()));
        var manualXml = LeerTexto(manualZip.GetEntry(parte)!);
        var manual = ParsearCeldas(manualXml, sharedManual);

        Exigir(salida.TryGetValue(referencia, out var celdaSalida),
            $"La base producida no tiene {parte}!{referencia}.");
        Exigir(manual.TryGetValue(referencia, out var celdaManual),
            $"El manual no tiene {parte}!{referencia}.");

        var valorSalida = esTexto ? celdaSalida!.Texto : celdaSalida!.Valor;
        var valorManual = esTexto ? celdaManual!.Texto : celdaManual!.Valor;
        Exigir(string.Equals(valorSalida, valorManual, StringComparison.Ordinal),
            $"{parte}!{referencia}: base='{valorSalida}' != manual='{valorManual}'.");
        lineas.Add($"{parte}!{referencia} == manual ('{valorManual}'): OK");
    }

    private sealed record HojaInfo(string Name, string SheetId, string Rid);

    private static List<HojaInfo> ExtraerHojas(string workbookXml)
    {
        var lista = new List<HojaInfo>();
        foreach (Match match in Regex.Matches(workbookXml, "<sheet name=\"(?<n>[^\"]*)\" sheetId=\"(?<s>\\d+)\" r:id=\"(?<r>[^\"]+)\""))
        {
            lista.Add(new HojaInfo(match.Groups["n"].Value, match.Groups["s"].Value, match.Groups["r"].Value));
        }

        return lista;
    }

    // ── utilidades ───────────────────────────────────────────────────────────────────────────

    private static bool EsParteTextoTransformable(string nombre) =>
        string.Equals(nombre, PartWorkbook, StringComparison.OrdinalIgnoreCase)
        || string.Equals(nombre, PartApp, StringComparison.OrdinalIgnoreCase)
        || string.Equals(nombre, PartShared, StringComparison.OrdinalIgnoreCase)
        || HojaRegex.IsMatch(nombre);

    private static int ContarFormulasCon(string xml, string token)
    {
        var total = 0;
        foreach (Match match in FormulaRegex.Matches(xml))
        {
            if (match.Groups["texto"].Value.Contains(token, StringComparison.Ordinal))
            {
                total++;
            }
        }

        return total;
    }

    private static byte[] LeerBytes(ZipArchiveEntry entrada)
    {
        using var stream = entrada.Open();
        using var memoria = new MemoryStream();
        stream.CopyTo(memoria);
        return memoria.ToArray();
    }

    private static string LeerTexto(ZipArchiveEntry entrada) => Encoding.UTF8.GetString(LeerBytes(entrada));

    private static string? Argumento(string[] args, string nombre)
    {
        for (var i = 0; i < args.Length - 1; i++)
        {
            if (string.Equals(args[i], nombre, StringComparison.OrdinalIgnoreCase))
            {
                return args[i + 1];
            }
        }

        return null;
    }

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md"))
                && Directory.Exists(Path.Combine(dir.FullName, "Docs")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new InvalidOperationException("No se encontró la raíz del repositorio (AGENTS.md + Docs/).");
    }

    private static void Exigir(bool condicion, string mensaje)
    {
        if (!condicion)
        {
            throw new InvalidOperationException(mensaje);
        }
    }
}
