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
/// Plan 62 (H4 — T4.1/T4.2/T4.3, D-H4 / D-E1-4EXT): la familia E1-4 se extiende al bloque
/// <c>CONSOLIDADO_TOTAL RECAUDO!E142:E146</c> ("VALIDACION DEL RECAUDO VERSUS REMUNERACION").
///
/// <para><b>Hechos T0 que el fixture congela</b> (Plan 62 §T0, doc <c>Docs/T0-Plan62-Pegado-R10.md</c>):
/// la fila 142–146 es un ASE por fila; <c>E{k}</c> = Σ de los 5 canales <c>'Recaudo *'!{col}{21+i}</c>
/// con la columna FIJA en <c>F</c> (la de la 2.ª quincena) mientras el recaudo se escribe en <c>D</c>
/// en Q1 ⇒ bloque en 0 en Q1 (desalineación T0-12). La columna la decide el DOMINIO
/// (<c>Periodo.NumeroQuincena</c> vía <see cref="PreparadorBasePipeline.ColumnaRecaudoDe"/>).</para>
///
/// <para><b>R5-GATE</b> (ratificación §11 R-5): <c>L145</c> es el maestro de una fórmula compartida
/// (<c>si=68</c>, <c>ref=L145:L147</c>) y <c>M142</c> de otra (<c>si=67</c>, <c>ref=M142:M147</c>).
/// La reescritura de H4 toca SOLO el texto de <c>E142:E146</c>: el resto de la fila —incluidas las
/// dos compartidas— queda BYTE-idéntico (barrido XML completo del bloque 142–147).</para>
///
/// <para><b>Fuente de las bases:</b> Phase 0 de PRODUCCIÓN (<see cref="PreparadorBasePipeline"/>)
/// desde la plantilla v4 vigente, vía <see cref="Insumos.BaseDelPeriodoV4"/> (caché por proceso).
/// Cero escritura sobre plantillas o insumos: las mutaciones (tercera forma, saneo de calcChain) se
/// hacen sobre COPIAS temporales.</para>
/// </summary>
public sealed class RecaudoValidacionConsolidado142H4Plan62Tests
{
    private const string HojaConsolidado = WorkbookLeafCellMap.HojaConsolidado;

    /// <summary>Periodo Q1 del plan: el bloque debe quedar leyendo la columna <c>D</c>.</summary>
    private const string PeriodoQ1 = "2026091";

    /// <summary>Periodo Q2 del plan: el bloque debe quedar leyendo la columna <c>F</c> (identidad).</summary>
    private const string PeriodoQ2 = "2026072";

    private static readonly string[] CanalesRecaudo =
    [
        "Recaudo EAAB Reciprocidad",
        "Recaudo EAAB + Ciud Limp",
        "Recaudo Directa Occidente",
        "Recaudo ENEL",
        "Recaudo ENERBIT",
    ];

    /// <summary>Geometría T0-11: la fila del canal que suma cada celda del bloque (E142→21 … E146→25).</summary>
    private static readonly (string Celda, int FilaRecaudo)[] Bloque142 =
    [
        ("E142", 21), ("E143", 22), ("E144", 23), ("E145", 24), ("E146", 25),
    ];

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T4.1/T4.2 — Q1 → columna D (cierres la desalineación T0-12) · Q2 → columna F
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T4.2 (R-QUINCENA-FIJA): con la base del período preparada por Phase 0, el texto-<c>&lt;f&gt;</c>
    /// de <c>E142:E146</c> es Σ de los 5 canales <c>'Recaudo *'</c> en la columna de la quincena del
    /// run (Q1 → <c>D</c>, Q2 → <c>F</c>) y en la fila 21:25 que le toca. Se asertan los 5 términos
    /// por celda (no una subcadena): 5 refs, hojas del catálogo abierto, misma columna y misma fila.
    /// </summary>
    [Theory]
    [InlineData(PeriodoQ1, "D")]
    [InlineData(PeriodoQ2, "F")]
    public void T4_Celdas142_SumanElRecaudoEnLaColumnaDeLaQuincena(string codigo, string columna)
    {
        var periodo = Periodo.Parse(codigo);
        Assert.Equal(columna, PreparadorBasePipeline.ColumnaRecaudoDe(periodo));

        var xml = HojaXml(Insumos.BaseDelPeriodoV4(codigo), HojaConsolidado);
        foreach (var (celda, fila) in Bloque142)
        {
            var formula = FormulaTexto(xml, celda)
                ?? throw new InvalidOperationException($"{HojaConsolidado}!{celda} no trae <f> en la base de {codigo}.");

            var terminos = PreparadorBasePipeline.SepararTerminosSuma(formula);
            Assert.Equal(CanalesRecaudo.Length, terminos.Count);
            Assert.Equal(CanalesRecaudo, terminos.Select(HojaDeTermino).ToArray());

            for (var i = 0; i < terminos.Count; i++)
            {
                var m = Regex.Match(terminos[i], "^'(?<hoja>[^']+)'!(?<col>[A-Z]+)(?<fila>\\d+)$");
                Assert.True(m.Success, $"{HojaConsolidado}!{celda}: término fuera de forma '{terminos[i]}'.");
                Assert.Equal(columna, m.Groups["col"].Value);
                Assert.Equal(fila.ToString(CultureInfo.InvariantCulture), m.Groups["fila"].Value);
            }

            // La forma canónica del T0-11 es Σ con '+' y sin paréntesis: E142 termina en la fila
            // del ASE de su fila, no en la 21 de todos.
            Assert.DoesNotContain("(", formula, StringComparison.Ordinal);
            Assert.Equal(
                string.Join(
                    "+",
                    CanalesRecaudo.Select(h => $"'{h}'!{columna}{fila.ToString(CultureInfo.InvariantCulture)}")),
                formula);
        }
    }

    /// <summary>
    /// T4.1/T4.2 (R5-GATE): la extensión toca SOLO el texto de <c>E142:E146</c>. Todo el resto del
    /// bloque 142–147 del CONSOLIDADO (D, F..M, la fila de cierre 147 y el <c>E147 = SUM(E142:E146)</c>)
    /// queda BYTE-idéntico frente a la plantilla oficial, y el estilo (<c>s</c>) de las 5 celdas
    /// reescritas se preserva (solo cambia el texto de la <c>&lt;f&gt;</c>, nunca la celda).
    /// </summary>
    [Theory]
    [InlineData(PeriodoQ1)]
    [InlineData(PeriodoQ2)]
    public void T4_RestoDeLaFila142_147_ByteIdenticoYEstilosPreservados(string codigo)
    {
        var oficial = HojaXml(Insumos.PlantillaOficialV4(), HojaConsolidado);
        var producido = HojaXml(Insumos.BaseDelPeriodoV4(codigo), HojaConsolidado);

        var permitidas = Bloque142.Select(b => b.Celda).ToHashSet(StringComparer.Ordinal);
        var distintas = new List<string>();
        foreach (var referencia in CeldasDelBloque142A147(oficial))
        {
            if (permitidas.Contains(referencia))
            {
                continue; // E142:E146 es lo ÚNICO que H4 reescribe (texto de <f>).
            }

            if (!string.Equals(CeldaCruda(oficial, referencia), CeldaCruda(producido, referencia), StringComparison.Ordinal))
            {
                distintas.Add(referencia);
            }
        }

        Assert.True(distintas.Count == 0,
            $"Plan 62-H4 tocó celdas fuera de E142:E146 en el bloque 142–147: {string.Join(", ", distintas)}");

        // La fila de cierre NO se reescribe: E147 sigue siendo la suma del bloque.
        Assert.Equal("SUM(E142:E146)", FormulaTexto(producido, "E147"));

        // Solo el TEXTO de la <f> cambia: el estilo (s) de las 5 celdas es el del oficial.
        foreach (var (celda, _) in Bloque142)
        {
            Assert.Equal(AtributoCelda(oficial, celda, "s"), AtributoCelda(producido, celda, "s"));
        }
    }

    /// <summary>
    /// T4.2 (R5-GATE, ratificación §11 R-5): las fórmulas compartidas de la fila conservan su
    /// geometría. <c>L145</c> (<c>t="shared" ref="L145:L147" si=68</c>) y <c>M142</c>
    /// (<c>t="shared" ref="M142:M147" si=67</c>) mantienen <c>t</c>/<c>ref</c>/<c>si</c> y texto, y
    /// sus seguidoras NO ganan texto propio (si la reescritura hubiera extendido el rango compartido,
    /// el grupo quedaría huérfano y Excel borraría las fórmulas al reparar — Plan 33/58).
    /// </summary>
    [Theory]
    [InlineData(PeriodoQ1)]
    [InlineData(PeriodoQ2)]
    public void T4_FormulasCompartidas_L145_M142_ConservanSiYRef(string codigo)
    {
        var oficial = HojaXml(Insumos.PlantillaOficialV4(), HojaConsolidado);
        var producido = HojaXml(Insumos.BaseDelPeriodoV4(codigo), HojaConsolidado);

        foreach (var maestro in new[] { "L145", "M142" })
        {
            foreach (var atributo in new[] { "t", "ref", "si" })
            {
                Assert.Equal(AtributoFormula(oficial, maestro, atributo), AtributoFormula(producido, maestro, atributo));
            }

            Assert.Equal(FormulaTexto(oficial, maestro), FormulaTexto(producido, maestro));
            Assert.Equal(CeldaCruda(oficial, maestro), CeldaCruda(producido, maestro));
        }

        // Seguidoras del grupo: siguen sin texto propio (solo <f t="shared" si=".."/>).
        foreach (var seguidora in new[] { "L146", "L147", "M143", "M144", "M145", "M146", "M147" })
        {
            Assert.Equal(AtributoFormula(oficial, seguidora, "si"), AtributoFormula(producido, seguidora, "si"));
            Assert.Null(FormulaTexto(producido, seguidora));
            Assert.Equal(CeldaCruda(oficial, seguidora), CeldaCruda(producido, seguidora));
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T4.2 — Tercera forma ⇒ ERR-PLANTILLA (hoja + celda + quincena)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T4.2 (R-QUINCENA-FIJA): Phase 0 NO adivina. Si una de las 5 celdas trae una tercera forma
    /// (columna distinta de D/F, fila distinta de la del ASE, o algo que no es Σ de <c>'Recaudo *'</c>),
    /// aborta con <c>ERR-PLANTILLA</c> nombrando hoja + celda + quincena, sin escribir la base. La
    /// plantilla mutada es una COPIA temporal: el oficial nunca se abre en escritura.
    /// </summary>
    [Theory]
    [InlineData("E143", "G22")]                                    // columna fuera de {D,F}
    [InlineData("E144", "F24")]                                    // fila de otro ASE
    [InlineData("E145", "SUM('Recaudo ENEL'!D24:F24)")]            // no es Σ de referencias por canal
    public void T4_TerceraForma_Phase0_AbortaConErrPlantilla(string celda, string variante)
    {
        var periodo = Periodo.Parse(PeriodoQ1);
        using var copia = CopiaMutadaDeLaPlantilla();

        var original = FormulaTexto(HojaXml(copia.Ruta, HojaConsolidado), celda)!;
        var tercera = variante.StartsWith("SUM(", StringComparison.Ordinal)
            ? variante
            : string.Join("+", CanalesRecaudo.Select(h => $"'{h}'!{variante}"));
        MutarFormula(copia.Ruta, HojaConsolidado, celda, original, tercera);

        var salida = Path.Combine(copia.Carpeta, periodo.NombreArchivo);
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            PreparadorBasePipeline.Ejecutar(copia.Ruta, periodo, salida));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains(HojaConsolidado, ex.Message, StringComparison.Ordinal);
        Assert.Contains(celda, ex.Message, StringComparison.Ordinal);
        Assert.Contains("Q1", ex.Message, StringComparison.Ordinal);
        Assert.Contains("forma inesperada", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(salida), "Phase 0 abortó y dejó una base escrita: se prohibe la salida parcial.");
    }

    /// <summary>
    /// T4.2 (D-H4.1): la familia de VALIDACIÓN de base (<see cref="ValidadorBasePeriodo"/>) también
    /// exige la variante del run en las 5 celdas nuevas — una base con tercera forma se rechaza
    /// ANTES de escribir, nombre hoja + celda + quincena. Se perturba <c>E143</c> de una copia de la
    /// base Q1 (todo lo demás sigue siendo la variante Q1 correcta, así que el rechazo es de H4 y no
    /// del bloque L9:L13).
    /// </summary>
    [Fact]
    public void T4_TerceraForma_ValidadorBase_AbortaConErrPlantilla()
    {
        var periodo = Periodo.Parse(PeriodoQ1);
        using var copia = CopiaDeLaBase("2026091");

        var original = FormulaTexto(HojaXml(copia.Ruta, HojaConsolidado), "E143")!;
        var tercera = string.Join("+", CanalesRecaudo.Select(h => $"'{h}'!G22"));
        MutarFormula(copia.Ruta, HojaConsolidado, "E143", original, tercera);

        using var documento = SpreadsheetDocument.Open(copia.Ruta, false);
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            ValidadorBasePeriodo.Exigir(documento.WorkbookPart!, periodo));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains(HojaConsolidado, ex.Message, StringComparison.Ordinal);
        Assert.Contains("E143", ex.Message, StringComparison.Ordinal);
        Assert.Contains("Q1", ex.Message, StringComparison.Ordinal);
        Assert.Contains("no es Σ 'Recaudo *'!D22", ex.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// T4.2 (control positivo): la base del período, tal como sale de Phase 0, pasa el preflight de
    /// base en su propia quincena (Q1 y Q2). Sin este control, el rojo de arriba podría venir de
    /// cualquier otra familia de validación.
    /// </summary>
    [Theory]
    [InlineData(PeriodoQ1)]
    [InlineData(PeriodoQ2)]
    public void T4_BasePreparada_PasaElPreflightDeSuQuincena(string codigo)
    {
        using var documento = SpreadsheetDocument.Open(Insumos.BaseDelPeriodoV4(codigo), false);
        ValidadorBasePeriodo.Exigir(documento.WorkbookPart!, Periodo.Parse(codigo));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // CA-C4 — fullCalcOnLoad alcanzable tras la reescritura
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// CA-C4 (Plan 28 intacto): la base con <c>E142:E146</c> reescrito sigue siendo un paquete sano
    /// — <see cref="SaneadorCadenaCalculo.Sanear"/> pasa, no queda <c>xl/calcChain.xml</c>,
    /// <c>fullCalcOnLoad="1"</c> queda en <c>workbook.xml</c> y el gate de integridad autocontenido
    /// no ve grupos shared huérfanos: la salida abre sin reparación y Excel recalcula el bloque con
    /// la columna de la quincena. Todo sobre una COPIA temporal.
    /// </summary>
    [Fact]
    public void T4_FullCalcAlcanzable_SaneadorPasaTrasLaReescritura()
    {
        using var copia = CopiaDeLaBase("2026091");

        using (var documento = SpreadsheetDocument.Open(copia.Ruta, true))
        {
            SaneadorCadenaCalculo.Sanear(documento.WorkbookPart!);
            documento.WorkbookPart!.Workbook!.Save();
        }

        using (var zip = ZipFile.OpenRead(copia.Ruta))
        {
            Assert.Null(zip.GetEntry("xl/calcChain.xml"));
            Assert.DoesNotContain(zip.Entries, e => e.FullName.Contains("calcChain", StringComparison.OrdinalIgnoreCase));

            var workbook = LeerParte(zip, "xl/workbook.xml");
            Assert.Contains("fullCalcOnLoad=\"1\"", workbook, StringComparison.Ordinal);
        }

        ValidadorIntegridadPaquete.ValidarPaquete(copia.Ruta);

        // El bloque reescrito sobrevive al saneo (la <f> de E142:E146 no es calcChain).
        var xml = HojaXml(copia.Ruta, HojaConsolidado);
        foreach (var (celda, fila) in Bloque142)
        {
            Assert.Equal(
                string.Join(
                    "+",
                    CanalesRecaudo.Select(h => $"'{h}'!D{fila.ToString(CultureInfo.InvariantCulture)}")),
                FormulaTexto(xml, celda));
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T4.3 — Allow-list Phase 0: 68 → 73 celdas (la extensión es de 5 celdas exactas)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T4.3 (D-H4.3): la extensión declarada son EXACTAMENTE 5 celdas (68 → 73 en la convención del
    /// Plan 40). Se aserta contra la fuente única de la geometría —no contra una copia del número—
    /// y que cada celda existe como <c>&lt;f&gt;</c> en la plantilla oficial. El inventario por
    /// nombre de hoja del grep de cierre vive en <c>FNew1Q1BlockerT0Tests.EsCeldaDeAllowList</c>
    /// (Plan 62 le añade las mismas 5 celdas al bloque del CONSOLIDADO).
    /// </summary>
    [Fact]
    public void T4_3_AllowListFase0_CreceDe68A73ConLas5CeldasDelBloque142()
    {
        const int allowListPlan40 = 68;
        const int allowListConH4 = 73;

        var celdas = PreparadorBasePipeline.CeldasRecaudoDeConsolidado142;

        // Convención del Plan 40: la allow-list es un INVENTARIO de celdas declaradas; H4 suma
        // exactamente las 5 del bloque 142–147 (68 → 73).
        Assert.Equal(allowListConH4, allowListPlan40 + celdas.Count);

        var oficial = HojaXml(Insumos.PlantillaOficialV4(), HojaConsolidado);
        foreach (var (celda, fila) in celdas)
        {
            Assert.InRange(fila, 21, 25);
            Assert.NotNull(FormulaTexto(oficial, celda));
        }

        // Y las 5 celdas son las únicas que Phase 0 mueve en el bloque (ver T4_RestoDeLaFila142_147).
        Assert.Equal(new[] { "E142", "E143", "E144", "E145", "E146" }, celdas.Select(c => c.Celda).ToArray());
    }

    // ── Lectura BCL del workbook ────────────────────────────────────────────────────────────────

    private static string HojaXml(string ruta, string nombreHoja)
    {
        using var zip = ZipFile.OpenRead(ruta);
        return LeerParte(zip, ParteDeHoja(LeerParte(zip, "xl/workbook.xml"), LeerParte(zip, "xl/_rels/workbook.xml.rels"), nombreHoja));
    }

    /// <summary>
    /// Parte del ZIP de una hoja resuelta por NOMBRE (workbook.xml + rels). El ORDEN de los atributos
    /// del <c>&lt;Relationship&gt;</c> cambia según quién serializó el paquete (Excel vs SDK OpenXML),
    /// así que Id/Target se extraen por separado — nunca con un patrón que los ate al orden.
    /// </summary>
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
        var entrada = zip.GetEntry(parte)
            ?? throw new FileNotFoundException($"El paquete no trae '{parte}'.");
        using var stream = entrada.Open();
        using var lector = new StreamReader(stream);
        return lector.ReadToEnd();
    }

    private static string? CeldaCruda(string xml, string referencia)
    {
        var match = Regex.Match(
            xml,
            $"<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c r=\"{Regex.Escape(referencia)}\"[^>]*(?:/>|>.*?</(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c>)",
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

        var match = Regex.Match(
            crudo,
            $"<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?f(?<attrs>[^>]*)>",
            RegexOptions.Singleline);
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

        var attr = Regex.Match(crudo, "^<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c r=\"[^\"]+\"(?<attrs>[^>]*?)(?:/>|>)", RegexOptions.Singleline);
        if (!attr.Success)
        {
            return null;
        }

        var m = Regex.Match(attr.Groups["attrs"].Value, $"\\b{Regex.Escape(atributo)}=\"(?<v>[^\"]*)\"");
        return m.Success ? m.Groups["v"].Value : null;
    }

    private static string HojaDeTermino(string termino)
    {
        var m = Regex.Match(termino, "^'(?<hoja>[^']+)'!");
        return m.Success ? m.Groups["hoja"].Value : string.Empty;
    }

    /// <summary>Todas las referencias con celda presente en las filas 142–147 de la hoja.</summary>
    private static IEnumerable<string> CeldasDelBloque142A147(string xml)
    {
        foreach (Match match in Regex.Matches(
                     xml,
                     "<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c r=\"(?<ref>[A-Z]+\\d+)\"",
                     RegexOptions.Singleline))
        {
            var m = Regex.Match(match.Groups["ref"].Value, "^(?<col>[A-Z]+)(?<fila>\\d+)$");
            if (m.Success && int.Parse(m.Groups["fila"].Value, CultureInfo.InvariantCulture) is >= 142 and <= 147)
            {
                yield return match.Groups["ref"].Value;
            }
        }
    }

    // ── Copias temporales (cero escritura sobre plantilla/insumos) ──────────────────────────────

    /// <summary>Copia de la plantilla oficial v4 en un directorio temporal (mutable).</summary>
    private static CopiaTemporal CopiaMutadaDeLaPlantilla()
    {
        var oficial = Insumos.PlantillaOficialV4();
        Assert.True(File.Exists(oficial), $"Falta la plantilla v4 vigente: {oficial}");

        var carpeta = Path.Combine(Path.GetTempPath(), "p62-h4-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(carpeta);
        var ruta = Path.Combine(carpeta, Path.GetFileName(oficial));
        File.Copy(oficial, ruta, overwrite: false);
        return new CopiaTemporal(carpeta, ruta);
    }

    /// <summary>Copia de la base preparada por Phase 0 para el período (mutable).</summary>
    private static CopiaTemporal CopiaDeLaBase(string codigo)
    {
        var basePreparada = Insumos.BaseDelPeriodoV4(codigo);
        Assert.True(File.Exists(basePreparada), $"Falta la base preparada de {codigo}: {basePreparada}");

        var carpeta = Path.Combine(Path.GetTempPath(), "p62-h4-base-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(carpeta);
        var ruta = Path.Combine(carpeta, $"Base_{codigo}_" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.Copy(basePreparada, ruta, overwrite: false);
        return new CopiaTemporal(carpeta, ruta);
    }

    /// <summary>
    /// Sustituye el TEXTO de la <c>&lt;f&gt;</c> de una celda por bytes (BCL, sin OpenXML) para no
    /// arrastrar efectos de serialización ajenos al caso: solo cambia ese nodo.
    /// </summary>
    private static void MutarFormula(string ruta, string nombreHoja, string celda, string original, string reemplazo)
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

        var celdaCruda = CeldaCruda(xml, celda) ?? throw new InvalidOperationException($"Falta {celda} en '{parte}'.");
        var nuevaCelda = celdaCruda.Replace($">{original}</", $">{reemplazo}</", StringComparison.Ordinal);
        Assert.NotEqual(celdaCruda, nuevaCelda);

        var nuevoXml = xml[..xml.IndexOf(celdaCruda, StringComparison.Ordinal)]
            + nuevaCelda
            + xml[(xml.IndexOf(celdaCruda, StringComparison.Ordinal) + celdaCruda.Length)..];

        entrada.Delete();
        var nueva = zip.CreateEntry(parte, CompressionLevel.Optimal);
        using var escritura = nueva.Open();
        using var escritor = new StreamWriter(escritura);
        escritor.Write(nuevoXml);
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
                if (Directory.Exists(Carpeta))
                {
                    Directory.Delete(Carpeta, recursive: true);
                }
            }
            catch (IOException)
            {
                // Limpieza best-effort (temp).
            }
        }
    }
}
