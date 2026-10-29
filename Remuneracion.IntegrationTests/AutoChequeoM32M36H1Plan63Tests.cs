using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 63 (H1 — T1.1/T1.2/T1.3, D-H1 / D-REEMPLAZO-DIRIGIDO / R5-GATE): los auto-chequeos
/// <c>DetRetri{per}!M32:M36</c> restan la interventoría CRUDA de la quincena del run
/// (Q1 → <c>INTERVENTORIA!N26:N30</c>, Q2 → <c>M26:M30</c>) sustituyendo SOLO la referencia
/// <c>INTERVENTORIA!F15:F19</c> del template; el resto de la fórmula queda byte-idéntico.
///
/// <para>
/// <b>Hechos T0 que el fixture congela</b> (<c>Docs/T0-Plan63-M32-36.md</c>): M32→F15/26,
/// M33→F16/27, M34→F17/28, M35→F18/29, M36→F19/30; el término <c>+'AJUSTES - SF-T'!M47:M51</c> y
/// el comparando <c>=M9:M13</c> NO se tocan. La columna la decide el DOMINIO
/// (<see cref="Periodo.NumeroQuincena"/>).
/// </para>
///
/// <para>
/// <b>Fuente de las bases:</b> Phase 0 de PRODUCCIÓN (<see cref="PreparadorBasePipeline"/>) desde la
/// plantilla v4 vigente vía <see cref="Insumos.BaseDelPeriodoV4"/> (caché por proceso). Cero escritura
/// sobre la plantilla: las mutaciones (master/follower compartido, tercera forma) se hacen sobre
/// COPIAS temporales.
/// </para>
/// </summary>
public sealed class AutoChequeoM32M36H1Plan63Tests
{
    private const string PeriodoQ1 = "2026091";
    private const string PeriodoQ2 = "2026072";
    private const string HojaConsolidado = WorkbookLeafCellMap.HojaConsolidado;

    /// <summary>Base interna de la plantilla v4: sus hojas <c>DetRetri</c>/<c>DetValiRetri</c> son de este período.</summary>
    private const string PeriodoBasePlantilla = "2026092";

    /// <summary>Geometría T0-06: celda del auto-chequeo → fila CRUDA de INTERVENTORIA (M32→26 … M36→30).</summary>
    private static readonly (string Celda, int FilaCruda)[] Bloque =
    [
        ("M32", 26), ("M33", 27), ("M34", 28), ("M35", 29), ("M36", 30),
    ];

    /// <summary>Fila del término redondeado del template por celda (M32→F15 … M36→F19).</summary>
    private static readonly Dictionary<string, int> FilaRedondeada = new(StringComparer.Ordinal)
    {
        ["M32"] = 15, ["M33"] = 16, ["M34"] = 17, ["M35"] = 18, ["M36"] = 19,
    };

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T1.1 — Phase 0: M32:M36 restan la interventoría cruda de la quincena del run
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T1.1: en la base preparada, <c>M32:M36</c> restan <c>INTERVENTORIA!{N,M}{26..30}</c> y ya NO
    /// contienen la referencia redondeada <c>INTERVENTORIA!F15:F19</c>. Q1 → <c>N</c>; Q2 → <c>M</c>
    /// (la columna la decide el dominio, no el contenido).
    /// </summary>
    [Theory]
    [InlineData(PeriodoQ1, "N")]
    [InlineData(PeriodoQ2, "M")]
    public void T1_1_Phase0_M32M36_RestaLaInterventoriaCrudaDelRun(string codigo, string columna)
    {
        var periodo = Periodo.Parse(codigo);
        Assert.Equal(columna, periodo.NumeroQuincena == 1 ? "N" : "M");

        var xml = HojaXml(Insumos.BaseDelPeriodoV4(codigo), NombresHojaPeriodo.DetRetri(periodo));
        foreach (var (celda, fila) in Bloque)
        {
            var formula = FormulaTexto(xml, celda)
                ?? throw new InvalidOperationException($"{celda} no trae <f> en la base de {codigo}.");

            Assert.Equal(1, ContarReferencia(formula, $"INTERVENTORIA!{columna}{fila}"));
            Assert.Equal(0, ContarReferencia(formula, $"INTERVENTORIA!F{FilaRedondeada[celda]}"));

            // El término AJUSTES y el comparando del ASE se conservan (no se reconstruye la fórmula).
            Assert.Equal(1, ContarReferencia(formula, $"'AJUSTES - SF-T'!M{47 + (fila - 26)}"));
            Assert.EndsWith($"=M{9 + (fila - 26)}", formula, StringComparison.Ordinal);
        }
    }

    /// <summary>
    /// T1.2 (D-REEMPLAZO-DIRIGIDO): la fórmula producida es <b>exactamente</b> la del template con el
    /// token de interventoría sustituido — ni un carácter más. Y en TODA la hoja <c>DetRetri</c> las
    /// únicas celdas cuyo texto de <c>&lt;f&gt;</c> cambia son esas 5 (diff por celda contra la
    /// plantilla oficial): cero daño colateral en los otros 127 auto-chequeos/consumidoras.
    /// </summary>
    [Theory]
    [InlineData(PeriodoQ1, "N")]
    [InlineData(PeriodoQ2, "M")]
    public void T1_2_Phase0_RestoDeLaFormulaByteIdentico(string codigo, string columna)
    {
        var periodo = Periodo.Parse(codigo);
        var hojaOficial = NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoBasePlantilla));
        var hojaProducida = NombresHojaPeriodo.DetRetri(periodo);
        var oficial = HojaXml(Insumos.PlantillaOficialV4(), hojaOficial);
        var producido = HojaXml(Insumos.BaseDelPeriodoV4(codigo), hojaProducida);

        foreach (var (celda, fila) in Bloque)
        {
            var original = FormulaTexto(oficial, celda)!;
            var esperada = original.Replace(
                $"INTERVENTORIA!F{FilaRedondeada[celda]}",
                $"INTERVENTORIA!{columna}{fila}",
                StringComparison.Ordinal);

            Assert.Contains($"INTERVENTORIA!F{FilaRedondeada[celda]}", original, StringComparison.Ordinal);
            Assert.Equal(esperada, FormulaTexto(producido, celda));
        }

        // Diff de textos de <f> en TODA la hoja: las de este plan (M32:M36) y las de la familia
        // F-NEW-1 del Plan 40 que Phase 0 ya reescribía en la misma hoja (L32:L36 / O32:O36).
        // Cualquier OTRA celda con texto distinto es daño colateral.
        var permitidas = Bloque.Select(b => b.Celda).ToHashSet(StringComparer.Ordinal);
        foreach (var prefijo in new[] { "L32", "L33", "L34", "L35", "L36", "O32", "O33", "O34", "O35", "O36" })
        {
            permitidas.Add(prefijo);
        }
        var formulasOficial = FormulasPorCelda(oficial);
        var formulasProducido = FormulasPorCelda(producido);
        Assert.Equal(formulasOficial.Count, formulasProducido.Count);

        var distintas = formulasOficial
            .Where(kv => !permitidas.Contains(kv.Key)
                && !string.Equals(kv.Value, formulasProducido.GetValueOrDefault(kv.Key), StringComparison.Ordinal))
            .Select(kv => kv.Key)
            .ToList();

        Assert.True(distintas.Count == 0,
            $"Plan 63 tocó <f> fuera de M32:M36 en {hojaProducida}: {string.Join(", ", distintas)}");

        // Las 5 celdas conservan su estilo/atributos de celda (solo cambia el texto de <f>).
        foreach (var (celda, _) in Bloque)
        {
            Assert.Equal(AtributoCelda(oficial, celda, "s"), AtributoCelda(producido, celda, "s"));
            Assert.Equal(AtributoCelda(oficial, celda, "t"), AtributoCelda(producido, celda, "t"));
            Assert.Null(AtributoFormula(producido, celda, "t"));
        }
    }

    /// <summary>
    /// T1.2 (R5-GATE): si la región trajera una fórmula compartida, los atributos
    /// (<c>t</c>/<c>ref</c>/<c>si</c>) del maestro se PRESERVAN — la costura
    /// <c>SustituirTextoFormula</c> del Plan 58/62 cambia solo el texto del nodo. Se injerta un master
    /// de 1 celda en <c>M32</c> de una COPIA de la plantilla: si alguien volviera a
    /// <c>new CellFormula(texto)</c>, el grupo quedaría huérfano y este assert enrojecería.
    /// </summary>
    [Fact]
    public void T1_2_R5Gate_MasterInjertado_ConservaTyRefYSi()
    {
        var periodo = Periodo.Parse(PeriodoQ1);
        using var copia = CopiaMutadaDeLaPlantilla();
        var hojaBase = NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoBasePlantilla));
        var hojaDestino = NombresHojaPeriodo.DetRetri(periodo);

        var original = FormulaTexto(HojaXml(copia.Ruta, hojaBase), "M32")!;
        InyectarAtributosFormula(copia.Ruta, hojaBase, "M32", " t=\"shared\" ref=\"M32:M32\" si=\"77\"");

        var salida = Path.Combine(copia.Carpeta, periodo.NombreArchivo);
        PreparadorBasePipeline.Ejecutar(copia.Ruta, periodo, salida);

        var producido = HojaXml(salida, hojaDestino);
        Assert.Equal("shared", AtributoFormula(producido, "M32", "t"));
        Assert.Equal("M32:M32", AtributoFormula(producido, "M32", "ref"));
        Assert.Equal("77", AtributoFormula(producido, "M32", "si"));
        Assert.Equal(
            original.Replace("INTERVENTORIA!F15", "INTERVENTORIA!N26", StringComparison.Ordinal),
            FormulaTexto(producido, "M32"));

        // El gate de integridad del paquete (Plan 58) sigue pasando: cero fórmulas nuevas ni celdas
        // con <f> desaparecidas.
        ValidadorIntegridadPaquete.ValidarPreparacion(File.ReadAllBytes(copia.Ruta), File.ReadAllBytes(salida));
    }

    /// <summary>
    /// T1.2 (R5-GATE, lado seguro): una SEGUIDORA compartida sin texto propio en la región NO se
    /// reescribe a ciegas — Phase 0 aborta con <c>ERR-PLANTILLA</c> nombrando la celda y no produce
    /// base. Reescribir el texto de una seguidora rompería el grupo (Excel borra fórmulas al reparar).
    /// </summary>
    [Fact]
    public void T1_2_R5Gate_SeguidoraCompartida_FailFastSinBase()
    {
        var periodo = Periodo.Parse(PeriodoQ1);
        using var copia = CopiaMutadaDeLaPlantilla();
        var hojaBase = NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoBasePlantilla));

        MutarFormula(copia.Ruta, hojaBase, "M33", string.Empty, " t=\"shared\" si=\"78\"");

        var salida = Path.Combine(copia.Carpeta, periodo.NombreArchivo);
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            PreparadorBasePipeline.Ejecutar(copia.Ruta, periodo, salida));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains("M33", ex.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(salida), "Phase 0 abortó y dejó una base escrita: se prohíbe la salida parcial.");
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T1.3 — ValidadorBasePeriodo: la variante del run, y la tercera forma es ERR-PLANTILLA
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T1.3 (control positivo): la base tal como sale de Phase 0 pasa el preflight de su quincena.
    /// Sin este control el rojo de abajo podría venir de cualquier otra familia de validación.
    /// </summary>
    [Theory]
    [InlineData(PeriodoQ1)]
    [InlineData(PeriodoQ2)]
    public void T1_3_BasePreparada_PasaElPreflightDeSuQuincena(string codigo)
    {
        using var documento = SpreadsheetDocument.Open(Insumos.BaseDelPeriodoV4(codigo), false);
        ValidadorBasePeriodo.Exigir(documento.WorkbookPart!, Periodo.Parse(codigo));
    }

    /// <summary>
    /// T1.3 (tercera forma): una base con la referencia redondeada del template (<c>F15</c>) donde se
    /// espera la cruda (<c>N26</c>) se rechaza ANTES de escribir, nombrando hoja + celda + variante
    /// esperada. Se perturba SOLO <c>M32</c> de una copia de la base Q1 (todo lo demás sigue correcto).
    /// </summary>
    [Fact]
    public void T1_3_TerceraForma_ValidadorBase_AbortaConErrPlantilla()
    {
        var periodo = Periodo.Parse(PeriodoQ1);
        using var copia = CopiaDeLaBase(PeriodoQ1);

        var hoja = NombresHojaPeriodo.DetRetri(periodo);
        var original = FormulaTexto(HojaXml(copia.Ruta, hoja), "M32")!;
        Assert.Contains("INTERVENTORIA!N26", original, StringComparison.Ordinal);
        MutarFormula(copia.Ruta, hoja, "M32", original.Replace("INTERVENTORIA!N26", "INTERVENTORIA!F15", StringComparison.Ordinal));

        using var documento = SpreadsheetDocument.Open(copia.Ruta, false);
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            ValidadorBasePeriodo.Exigir(documento.WorkbookPart!, periodo));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains(hoja, ex.Message, StringComparison.Ordinal);
        Assert.Contains("M32", ex.Message, StringComparison.Ordinal);
        Assert.Contains("INTERVENTORIA!N26", ex.Message, StringComparison.Ordinal);
        Assert.Contains("INTERVENTORIA!F15", ex.Message, StringComparison.Ordinal);
        Assert.Contains("Q1", ex.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// T1.3 (tercera forma mixta): una fórmula que contiene la variante del run Y la redondeada
    /// también se rechaza — el preflight no se conforma con "contiene lo esperado".
    /// </summary>
    [Fact]
    public void T1_3_FormaMixta_ValidadorBase_AbortaConErrPlantilla()
    {
        var periodo = Periodo.Parse(PeriodoQ1);
        using var copia = CopiaDeLaBase(PeriodoQ1);

        var hoja = NombresHojaPeriodo.DetRetri(periodo);
        var original = FormulaTexto(HojaXml(copia.Ruta, hoja), "M34")!;
        Assert.Contains("INTERVENTORIA!N28", original, StringComparison.Ordinal);
        MutarFormula(
            copia.Ruta,
            hoja,
            "M34",
            original.Replace("INTERVENTORIA!N28", "INTERVENTORIA!N28+INTERVENTORIA!F17", StringComparison.Ordinal));

        using var documento = SpreadsheetDocument.Open(copia.Ruta, false);
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            ValidadorBasePeriodo.Exigir(documento.WorkbookPart!, periodo));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains("M34", ex.Message, StringComparison.Ordinal);
        Assert.Contains("INTERVENTORIA!F17", ex.Message, StringComparison.Ordinal);
    }

    // ── Helpers ─────────────────────────────────────────────────────────────────────────────────

    private static int ContarReferencia(string formula, string referencia) =>
        Regex.Matches(formula, $"(?<![A-Za-z0-9_]){Regex.Escape(referencia)}(?![0-9])").Count;

    private static Dictionary<string, string> FormulasPorCelda(string xml)
    {
        var mapa = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (Match celda in Regex.Matches(
                     xml,
                     "<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c r=\"(?<ref>[A-Z]+\\d+)\"[^>]*(?:/>|>(?<inner>.*?)</(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c>)",
                     RegexOptions.Singleline))
        {
            if (!celda.Groups["inner"].Success)
            {
                continue;
            }

            var f = Regex.Match(celda.Groups["inner"].Value, "<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?f[^>]*>(?<texto>[^<]*)</");
            if (f.Success)
            {
                mapa[celda.Groups["ref"].Value] = f.Groups["texto"].Value;
            }
        }

        return mapa;
    }

    private static string HojaXml(string ruta, string nombreHoja)
    {
        using var zip = ZipFile.OpenRead(ruta);
        return LeerParte(zip, ParteDeHoja(LeerParte(zip, "xl/workbook.xml"), LeerParte(zip, "xl/_rels/workbook.xml.rels"), nombreHoja));
    }

    private static string ParteDeHoja(string workbook, string rels, string nombreHoja)
    {
        var sheet = Regex.Match(
            workbook,
            $"<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?sheet name=\"{Regex.Escape(nombreHoja)}\"[^>]*r:id=\"(?<rid>[^\"]+)\"");
        Assert.True(sheet.Success, $"El workbook no trae la hoja '{nombreHoja}'.");

        var rid = sheet.Groups["rid"].Value;
        foreach (Match relacion in Regex.Matches(rels, "<Relationship[^>]*>"))
        {
            var id = Regex.Match(relacion.Value, "\\bId=\"(?<v>[^\"]+)\"");
            if (!id.Success || !string.Equals(id.Groups["v"].Value, rid, StringComparison.Ordinal))
            {
                continue;
            }

            var target = Regex.Match(relacion.Value, "\\bTarget=\"(?<v>[^\"]+)\"").Groups["v"].Value;
            return target.StartsWith('/') ? target[1..] : "xl/" + target;
        }

        throw new InvalidOperationException($"No se resolvió la relación '{rid}' de '{nombreHoja}'.");
    }

    private static string LeerParte(ZipArchive zip, string parte)
    {
        var entrada = zip.GetEntry(parte) ?? throw new FileNotFoundException($"El paquete no trae '{parte}'.");
        using var stream = entrada.Open();
        using var lector = new StreamReader(stream);
        return lector.ReadToEnd();
    }

    private static string? CeldaCruda(string xml, string referencia)
    {
        var match = Regex.Match(
            xml,
            $"<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c r=\"{Regex.Escape(referencia)}\"[^>]*(?:/>|>(?<inner>.*?)</(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c>)",
            RegexOptions.Singleline);
        return match.Success ? match.Value : null;
    }

    private static string? FormulaTexto(string xml, string referencia)
    {
        var crudo = CeldaCruda(xml, referencia);
        if (crudo is null)
        {
            return null;
        }

        var match = Regex.Match(crudo, "<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?f(?<attrs>[^>]*)>(?<texto>[^<]*)</");
        return match.Success ? match.Groups["texto"].Value : null;
    }

    private static string? AtributoFormula(string xml, string referencia, string atributo)
    {
        var crudo = CeldaCruda(xml, referencia);
        if (crudo is null)
        {
            return null;
        }

        var match = Regex.Match(crudo, "<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?f(?<attrs>[^>]*)>", RegexOptions.Singleline);
        if (!match.Success)
        {
            return null;
        }

        var attr = Regex.Match(match.Groups["attrs"].Value, $"\\b{Regex.Escape(atributo)}=\"(?<v>[^\"]*)\"");
        return attr.Success ? attr.Groups["v"].Value : null;
    }

    private static string? AtributoCelda(string xml, string referencia, string atributo)
    {
        var crudo = CeldaCruda(xml, referencia);
        if (crudo is null)
        {
            return null;
        }

        var attr = Regex.Match(
            crudo,
            "^<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c r=\"[^\"]+\"(?<attrs>[^>]*?)(?:/>|>)",
            RegexOptions.Singleline);
        if (!attr.Success)
        {
            return null;
        }

        var m = Regex.Match(attr.Groups["attrs"].Value, $"\\b{Regex.Escape(atributo)}=\"(?<v>[^\"]*)\"");
        return m.Success ? m.Groups["v"].Value : null;
    }

    // ── Copias temporales y mutaciones por bytes (cero escritura sobre la plantilla) ────────────

    private static CopiaTemporal CopiaMutadaDeLaPlantilla()
    {
        var oficial = Insumos.PlantillaOficialV4();
        Assert.True(File.Exists(oficial), $"Falta la plantilla v4 vigente: {oficial}");

        var carpeta = Path.Combine(Path.GetTempPath(), "p63-h1-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(carpeta);
        var ruta = Path.Combine(carpeta, Path.GetFileName(oficial));
        File.Copy(oficial, ruta, overwrite: false);
        return new CopiaTemporal(carpeta, ruta);
    }

    private static CopiaTemporal CopiaDeLaBase(string codigo)
    {
        var basePreparada = Insumos.BaseDelPeriodoV4(codigo);
        Assert.True(File.Exists(basePreparada), $"Falta la base preparada de {codigo}: {basePreparada}");

        var carpeta = Path.Combine(Path.GetTempPath(), "p63-h1-base-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(carpeta);
        var ruta = Path.Combine(carpeta, $"Base_{codigo}_" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.Copy(basePreparada, ruta, overwrite: false);
        return new CopiaTemporal(carpeta, ruta);
    }

    /// <summary>
    /// Inyecta atributos en la <c>&lt;f&gt;</c> de una celda (BCL puro): sirve para simular un master
    /// shared sin arrastrar efectos de serialización ajenos al caso.
    /// </summary>
    private static void InyectarAtributosFormula(string ruta, string nombreHoja, string celda, string atributos)
    {
        using var zip = ZipFile.Open(ruta, ZipArchiveMode.Update);
        var parte = ParteDeHoja(LeerParte(zip, "xl/workbook.xml"), LeerParte(zip, "xl/_rels/workbook.xml.rels"), nombreHoja);
        var entrada = zip.GetEntry(parte) ?? throw new FileNotFoundException($"El paquete no trae '{parte}'.");

        string xml;
        using (var stream = entrada.Open())
        using (var lector = new StreamReader(stream))
        {
            xml = lector.ReadToEnd();
        }

        var crudo = CeldaCruda(xml, celda) ?? throw new InvalidOperationException($"Falta {celda} en '{parte}'.");
        var inicio = crudo.IndexOf("<f", StringComparison.Ordinal);
        Assert.True(inicio >= 0, $"{celda} no trae <f>.");
        var finAtributos = crudo.IndexOf('>', inicio);
        var nuevoCrudo = string.Concat(
            crudo.AsSpan(0, inicio),
            "<f",
            atributos,
            crudo.AsSpan(finAtributos));

        var indice = xml.IndexOf(crudo, StringComparison.Ordinal);
        var nuevoXml = string.Concat(xml.AsSpan(0, indice), nuevoCrudo, xml.AsSpan(indice + crudo.Length));
        ReescribirParte(zip, parte, nuevoXml);
    }

    /// <summary>
    /// Sustituye el TEXTO de la <c>&lt;f&gt;</c> de una celda por bytes. Con <paramref name="reemplazo"/>
    /// vacío produce una seguidora shared autocerrada (<c>&lt;f …/&gt;</c>).
    /// </summary>
    private static void MutarFormula(
        string ruta,
        string nombreHoja,
        string celda,
        string reemplazo,
        string atributos = "")
    {
        using var zip = ZipFile.Open(ruta, ZipArchiveMode.Update);
        var parte = ParteDeHoja(LeerParte(zip, "xl/workbook.xml"), LeerParte(zip, "xl/_rels/workbook.xml.rels"), nombreHoja);
        var entrada = zip.GetEntry(parte) ?? throw new FileNotFoundException($"El paquete no trae '{parte}'.");

        string xml;
        using (var stream = entrada.Open())
        using (var lector = new StreamReader(stream))
        {
            xml = lector.ReadToEnd();
        }

        var crudo = CeldaCruda(xml, celda) ?? throw new InvalidOperationException($"Falta {celda} en '{parte}'.");
        var nuevaFormula = string.IsNullOrEmpty(reemplazo)
            ? $"<f{atributos}/>"
            : $"<f{atributos}>{reemplazo}</f>";
        var viejaFormula = Regex.Match(crudo, "<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?f(?<attrs>[^>]*)>(?<texto>[^<]*)</.*?f>").Value;
        if (string.IsNullOrEmpty(viejaFormula))
        {
            viejaFormula = Regex.Match(crudo, "<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?f[^>]*/>").Value;
        }

        Assert.False(string.IsNullOrEmpty(viejaFormula), $"{celda} no trae <f> mutable.");
        var nuevoCrudo = crudo.Replace(viejaFormula, nuevaFormula, StringComparison.Ordinal);
        Assert.NotEqual(crudo, nuevoCrudo);

        var indice = xml.IndexOf(crudo, StringComparison.Ordinal);
        var nuevoXml = string.Concat(xml.AsSpan(0, indice), nuevoCrudo, xml.AsSpan(indice + crudo.Length));
        ReescribirParte(zip, parte, nuevoXml);
    }

    private static void ReescribirParte(ZipArchive zip, string parte, string contenido)
    {
        zip.GetEntry(parte)!.Delete();
        var nueva = zip.CreateEntry(parte, CompressionLevel.Optimal);
        using var escritura = nueva.Open();
        using var escritor = new StreamWriter(escritura);
        escritor.Write(contenido);
    }

    private sealed class CopiaTemporal : IDisposable
    {
        public CopiaTemporal(string carpeta, string ruta)
        {
            Carpeta = carpeta;
            Ruta = ruta;
        }

        public string Carpeta { get; }

        public string Ruta { get; }

        public void Dispose()
        {
            try
            {
                Directory.Delete(Carpeta, recursive: true);
            }
            catch
            {
                // best-effort
            }
        }
    }
}
