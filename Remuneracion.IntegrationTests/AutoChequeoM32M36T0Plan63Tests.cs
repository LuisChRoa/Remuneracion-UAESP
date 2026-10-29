using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Remuneracion.Core.Models;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 63 (H0 — T0.1..T0.4): verificación read-only de los 8 hechos del §T0 que habilitan
/// reescribir los auto-chequeos <c>DetRetri{per}!M32:M36</c> para que resten la interventoría
/// CRUDA de la quincena del run (Q1 → <c>INTERVENTORIA!N26:N30</c>, Q2 → <c>M26:M30</c>).
///
/// <para>
/// <b>Método (D-H0):</b> BCL puro (<see cref="ZipArchive"/> + <see cref="XDocument"/>) sobre rutas
/// EXPLÍCITAS. Sin Excel/COM, sin DocumentFormat.OpenXml, sin ExcelDataReader y sin
/// <c>Insumos.Raiz()</c>. Los .xlsx se abren SOLO lectura. Cero código de producción: este fixture
/// y <c>Docs/T0-Plan63-M32-36.md</c> SON el entregable de H0.
/// </para>
///
/// <para>
/// <b>Fuentes medidas:</b>
/// <list type="bullet">
/// <item><c>Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx</c>
/// (plantilla VIGENTE; base interna 2026092 → hoja <c>DetRetri2026092</c> + <c>INTERVENTORIA</c>).</item>
/// <item><c>Docs/Prueba Septiembre 1/Resultados/7/Remuneración 202609-1 Total.xlsx</c> (salida Q1
/// del caso rojo).</item>
/// <item><c>Docs/Prueba Septiembre 1/Resultados/Administrativo/Remuneracion 202609-1 Total.xlsx</c>
/// (manual del administrativo Q1: oráculo <c>N26:N30</c>).</item>
/// <item><c>Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx</c>
/// (manual del administrativo Q2: oráculo <c>F15:F19</c> + AJUSTES).</item>
/// </list>
/// </para>
///
/// <para>
/// <b>Regla honesta:</b> el assert refleja el BYTE. Los cachés <c>&lt;v&gt;</c> se leen SOLO como
/// MEDICIÓN del estado del archivo (nunca como oráculo de la fórmula: doctrina Plan 28/29/40); la
/// equivalencia booleana se demuestra por aritmética de operandos (T0-02/T0-03), no por el caché.
/// </para>
/// </summary>
public sealed class AutoChequeoM32M36T0Plan63Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string PeriodoSalidaQ1 = "2026091";
    private const string PeriodoManualQ2 = "2026072";

    /// <summary>Base interna de la plantilla oficial v4 (N3 = 2026092): sus hojas Det son de este período.</summary>
    private const string PeriodoBasePlantilla = "2026092";

    private const string HojaInterventoria = "INTERVENTORIA";
    private const string PartWorkbook = "xl/workbook.xml";
    private const string PartRels = "xl/_rels/workbook.xml.rels";
    private const string PartCompartidas = "xl/sharedStrings.xml";

    /// <summary>Fila del término REDONDEADO en el template (F15..F19) por ASE 1..5.</summary>
    private static readonly int[] FilasRedondeadas = [15, 16, 17, 18, 19];

    /// <summary>Fila del término CRUDO de la quincena (26..30) por ASE 1..5.</summary>
    private static readonly int[] FilasCrudas = [26, 27, 28, 29, 30];

    /// <summary>Auto-chequeos protegidos del CONSOLIDADO por ASE: M32..M36 ↔ ASE 1..5.</summary>
    private static readonly int[] FilasAutoChequeo = [32, 33, 34, 35, 36];

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T0a — Salida 2026091 (Q1): el defecto (T0-01…T0-03)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T0-01 (forma actual de la salida Q1): <c>DetRetri2026091!M32:M36</c> restan
    /// <c>INTERVENTORIA!F15:F19</c> (columna REDONDEADA, fija) y conservan el término
    /// <c>+'AJUSTES - SF-T'!M47:M51</c>; el comparando es <c>M9:M13</c>. La referencia de
    /// interventoría aparece EXACTAMENTE UNA vez por celda (condición del reemplazo dirigido de H1).
    /// </summary>
    [Fact]
    public void T0_01_SalidaQ1_M32M36_FormaF15F19ConAjustes()
    {
        using var zip = Abrir(RutaSalidaQ1());
        var hoja = Hoja(zip, NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoSalidaQ1)));

        for (var i = 0; i < 5; i++)
        {
            var celda = $"M{FilasAutoChequeo[i]}";
            var formula = FormulaDe(hoja, celda);
            Assert.False(string.IsNullOrEmpty(formula), $"{celda} no trae <f> en la salida Q1.");

            var referencia = $"INTERVENTORIA!F{FilasRedondeadas[i]}";
            Assert.Equal(1, ContarReferencia(formula!, referencia));

            // El término AJUSTES y el comparando quedan en la fórmula (se PRESERVAN byte-idénticos
            // en H1: solo cambia el fragmento de interventoría).
            Assert.Equal(1, ContarReferencia(formula!, $"'AJUSTES - SF-T'!M{47 + i}"));
            Assert.EndsWith($"=M{9 + i}", formula, StringComparison.Ordinal);

            // Defecto: NO resta la columna cruda de la quincena (ni N ni M de interventoría).
            Assert.DoesNotContain("INTERVENTORIA!N", formula, StringComparison.Ordinal);
            Assert.DoesNotContain("INTERVENTORIA!M", formula, StringComparison.Ordinal);
        }
    }

    /// <summary>
    /// T0-02 (veredicto del defecto): la salida Q1 trae <c>M32:M36</c> en
    /// FALSO/FALSO/FALSO/VERDADERO/FALSO. Se demuestra por aritmética de operandos:
    /// <list type="bullet">
    /// <item>el cuerpo (R1/R2/R4/BCE + <c>M23:M27</c>) y el comparando <c>M9:M13</c> son los MISMOS
    /// del manual (T0-03);</item>
    /// <item>el término <c>AJUSTES - SF-T!M47:M51</c> vale 0 en Q1 (medido en la salida);</item>
    /// <item>el único operando distinto es la interventoría: <c>F15:F19</c> (redondeado, +1) contra
    /// <c>N26:N30</c> (crudo) ⇒ el chequeo solo puede cerrar donde <c>F = N</c>.</item>
    /// </list>
    /// ASE4 (<c>F18 = N29</c>) es el ÚNICO que cierra. El caché <c>&lt;v&gt;</c> se reporta como
    /// medición del archivo, no como oráculo de la fórmula.
    /// </summary>
    [Fact]
    public void T0_02_SalidaQ1_VeredictoFFFVF_PorEquivalenciaFvsN()
    {
        using var zip = Abrir(RutaSalidaQ1());
        var det = Hoja(zip, NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoSalidaQ1)));
        var interv = Hoja(zip, HojaInterventoria);
        var ajustes = Hoja(zip, "AJUSTES - SF-T");

        var redondeados = FilasRedondeadas.Select(f => ValorNumerico(interv, $"{'F'}{f}")).ToArray();
        var crudos = FilasCrudas.Select(f => ValorNumerico(interv, $"N{f}")).ToArray();

        // (a) Medición crudo-vs-redondeado por ASE: Δ = +1 (el redondeo sube el peso) salvo ASE4
        //     (Bogotá Limpia), donde F = N.
        for (var i = 0; i < 5; i++)
        {
            var delta = redondeados[i] - crudos[i];
            if (i == 3)
            {
                Assert.Equal(0m, delta);
            }
            else
            {
                Assert.Equal(1m, delta);
            }
        }

        Assert.Equal(
            new[]
            {
                189_185_988m, 266_580_256m, 154_788_536m, 120_391_083m, 128_990_446m,
            },
            redondeados);
        Assert.Equal(
            new[]
            {
                189_185_987m, 266_580_255m, 154_788_535m, 120_391_083m, 128_990_445m,
            },
            crudos);

        // (b) En Q1 el término AJUSTES vale 0 ⇒ no altera el veredicto (el defecto es SOLO la
        //     referencia de interventoría).
        for (var i = 0; i < 5; i++)
        {
            Assert.Equal(0m, ValorNumerico(ajustes, $"M{47 + i}"));
        }

        // (c) Veredicto esperado por ASE: cierra SOLO donde F = N.
        var veredictoEsperado = Enumerable.Range(0, 5).Select(i => redondeados[i] == crudos[i]).ToArray();
        Assert.Equal([false, false, false, true, false], veredictoEsperado);

        // (d) Medición del archivo (caché de la salida, post-recálculo de Excel): coincide con (c).
        //     Se documenta como LECTURA del estado, no como oráculo de la fórmula.
        var cache = FilasAutoChequeo.Select(f => ValorNumerico(det, $"M{f}")).ToArray();
        Assert.Equal(
            veredictoEsperado.Select(v => v ? 1m : 0m).ToArray(),
            cache);
    }

    /// <summary>
    /// T0-03 (invariante de operandos): el cuerpo de la fórmula de <c>M32:M36</c> de la salida Q1 es
    /// IDÉNTICO al del manual del administrativo <b>salvo las filas del layout</b> y la referencia de
    /// interventoría. Se normalizan las referencias (<c>!COLUMNA#</c>) y se excluyen los dos términos
    /// que el manual NO trae (interventoría) o que valen 0 (AJUSTES): el diff debe ser ∅.
    /// </summary>
    [Fact]
    public void T0_03_SalidaQ1_CuerpoDeOperandosIdenticoAlManual_SalvoFilas()
    {
        using var salida = Abrir(RutaSalidaQ1());
        using var manual = Abrir(RutaManualQ1());

        var hojaSalida = Hoja(salida, NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoSalidaQ1)));
        var hojaManual = Hoja(manual, NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoSalidaQ1)));

        for (var i = 0; i < 5; i++)
        {
            var celda = $"M{FilasAutoChequeo[i]}";
            var cuerpoSalida = NormalizarCuerpo(FormulaDe(hojaSalida, celda)!, i);
            var cuerpoManual = NormalizarCuerpo(FormulaDe(hojaManual, celda)!, i);

            Assert.Equal(cuerpoManual, cuerpoSalida);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T0b — Manuales: oráculos de paridad (T0-04…T0-05)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T0-04 (manual Q1 2026091): <c>M32:M36</c> restan <c>INTERVENTORIA!N26:N30</c> (crudo, por ASE),
    /// NO traen término AJUSTES y cierran todos VERDADERO (<c>&lt;v&gt;</c> = 1, medición).
    /// </summary>
    [Fact]
    public void T0_04_ManualQ1_RestaN26N30_SinAjustes_TodosVerdadero()
    {
        using var zip = Abrir(RutaManualQ1());
        var hoja = Hoja(zip, NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoSalidaQ1)));

        for (var i = 0; i < 5; i++)
        {
            var celda = $"M{FilasAutoChequeo[i]}";
            var formula = FormulaDe(hoja, celda);
            Assert.False(string.IsNullOrEmpty(formula), $"{celda} no trae <f> en el manual Q1.");

            Assert.Equal(1, ContarReferencia(formula!, $"INTERVENTORIA!N{FilasCrudas[i]}"));
            Assert.Equal(0, ContarReferencia(formula!, "'AJUSTES - SF-T'!"));

            Assert.Equal(1m, ValorNumerico(hoja, celda));
        }
    }

    /// <summary>
    /// T0-05 (manual Q2 2026072): <c>M32:M36</c> restan <c>INTERVENTORIA!F15:F19</c> (redondeado) +
    /// AJUSTES y cierran todos VERDADERO. En Q2 la plantilla canónica ya cierra porque
    /// <c>M26:M30 ≡ F15:F19</c> valor a valor ⇒ el mapeo pedido (Q2 → <c>M26:M30</c>) es SIN impacto
    /// numérico; solo cambia el tema de referencia. <c>N26:N31</c> es la serie cruda.
    /// </summary>
    [Fact]
    public void T0_05_ManualQ2_RestaF15F19ConAjustes_M26IgualF_N26Crudo()
    {
        using var zip = Abrir(RutaManualQ2());
        var hoja = Hoja(zip, NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoManualQ2)));
        var interv = Hoja(zip, HojaInterventoria);

        for (var i = 0; i < 5; i++)
        {
            var celda = $"M{FilasAutoChequeo[i]}";
            var formula = FormulaDe(hoja, celda);
            Assert.False(string.IsNullOrEmpty(formula), $"{celda} no trae <f> en el manual Q2.");

            Assert.Equal(1, ContarReferencia(formula!, $"INTERVENTORIA!F{FilasRedondeadas[i]}"));
            Assert.Equal(1, ContarReferencia(formula!, $"'AJUSTES - SF-T'!M{47 + i}"));
            Assert.Equal(1m, ValorNumerico(hoja, celda));
        }

        // M26:M30 ≡ F15:F19 valor a valor (la tabla de la 2.ª quincena es la redondeada).
        var f = FilasRedondeadas.Select(fila => ValorNumerico(interv, $"F{fila}")).ToArray();
        var m = FilasCrudas.Select(fila => ValorNumerico(interv, $"M{fila}")).ToArray();
        Assert.Equal(f, m);

        // M31 = SUM(M26:M30) y N26:N31 es la serie CRUDA (Δ = +1 salvo ASE4).
        Assert.Equal(859_936_309m, ValorNumerico(interv, "M31"));
        var n = FilasCrudas.Select(fila => ValorNumerico(interv, $"N{fila}")).ToArray();
        for (var i = 0; i < 5; i++)
        {
            Assert.Equal(i == 3 ? 0m : 1m, m[i] - n[i]);
        }

        Assert.Equal(859_936_305m, ValorNumerico(interv, "N31"));
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T0c — Plantilla v4 + familias vecinas (T0-06…T0-08)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T0-06 (plantilla v4): <c>DetRetri2026092!M32:M36</c> traen la forma
    /// <c>INTERVENTORIA!F15:F19</c> + término <c>'AJUSTES - SF-T'!M47:M51</c> con el comparando
    /// <c>M9:M13</c>: la geometría que H1 reescribe (M32→26 … M36→30, origen F15..F19).
    /// </summary>
    [Fact]
    public void T0_06_PlantillaV4_M32M36_FormaF15F19ConAjustes()
    {
        using var zip = Abrir(RutaPlantillaV4());
        var hoja = Hoja(zip, NombresHojaPeriodo.DetRetri(Periodo.Parse(PeriodoBasePlantilla)));

        for (var i = 0; i < 5; i++)
        {
            var celda = $"M{FilasAutoChequeo[i]}";
            var formula = FormulaDe(hoja, celda);
            Assert.False(string.IsNullOrEmpty(formula), $"{celda} no trae <f> en la plantilla v4.");

            Assert.Equal(1, ContarReferencia(formula!, $"INTERVENTORIA!F{FilasRedondeadas[i]}"));
            Assert.Equal(1, ContarReferencia(formula!, $"'AJUSTES - SF-T'!M{47 + i}"));
            Assert.EndsWith($"=M{9 + i}", formula, StringComparison.Ordinal);
        }
    }

    /// <summary>
    /// T0-07 (<c>DetValiRetri!I24:I29</c>, familia del Plan 40 — NO se toca): en la salida Q1 la
    /// columna de interventoría es la DEL RUN (<c>N26:N31</c>) y los 6 auto-cheques están
    /// VERDADERO (medido). Es la prueba de que la mecánica por quincena ya funciona en esa familia.
    /// </summary>
    [Fact]
    public void T0_07_SalidaQ1_DetValiRetriI24I29_ColumnaDelRunYVerdadero()
    {
        using var zip = Abrir(RutaSalidaQ1());
        var hoja = Hoja(zip, NombresHojaPeriodo.DetValiRetri(Periodo.Parse(PeriodoSalidaQ1)));

        for (var i = 0; i < 6; i++)
        {
            var celda = $"I{24 + i}";
            var formula = FormulaDe(hoja, celda);
            Assert.False(string.IsNullOrEmpty(formula), $"{celda} no trae <f> en la salida Q1.");

            // Q1 = N. La fila 31 es el total (N31 = SUM(N26:N30)).
            Assert.Equal(1, ContarReferencia(formula!, $"INTERVENTORIA!N{26 + i}"));
            Assert.DoesNotContain("INTERVENTORIA!M", formula, StringComparison.Ordinal);
            Assert.Equal(1m, ValorNumerico(hoja, celda));
        }
    }

    /// <summary>
    /// T0-08 (<c>INTERVENTORIA!S26:S31</c>, tercera familia F-NEW-1): se MIDE en la salida Q1 y se
    /// documenta el veredicto. <b>Medido: VERDE</b> (los 6 TRUE) porque Phase 0 (Plan 41/T3) reescribe
    /// el maestro <c>R26 = N26</c> en Q1 y <c>Q26 = CONSOLIDADO!L104</c> sigue la selección L.
    /// NO se amplía el alcance: este plan no toca <c>S26:S31</c>; si el veredicto hubiera sido FALSO,
    /// el doc lo declararía como pendiente.
    /// </summary>
    [Fact]
    public void T0_08_SalidaQ1_InterventoriaS26S31_Medido()
    {
        using var zip = Abrir(RutaSalidaQ1());
        var hoja = Hoja(zip, HojaInterventoria);

        // Forma: S26/S27 texto propio, S28 maestro shared (ref S28:S31, si=2), S29:S31 seguidoras.
        Assert.Equal("Q26=R26", FormulaDe(hoja, "S26"));
        Assert.Equal("Q27=R27", FormulaDe(hoja, "S27"));
        Assert.Equal("Q28=R28", FormulaDe(hoja, "S28"));
        Assert.Equal("shared", AtributoFormula(hoja, "S28", "t"));
        Assert.Equal("S28:S31", AtributoFormula(hoja, "S28", "ref"));

        // El maestro de la tabla J ya es la variante de la quincena del run (Q1 → N26).
        Assert.Equal("N26", FormulaDe(hoja, "R26"));

        // Veredicto medido: S26:S31 = VERDADERO×6 (no queda pendiente de la clase).
        for (var fila = 26; fila <= 31; fila++)
        {
            Assert.Equal(1m, ValorNumerico(hoja, $"S{fila}"));
        }
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // Lectura BCL (ZIP + XML) — rutas EXPLÍCITAS, sin Insumos.Raiz(), sin escritura
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>Abre el .xlsx SOLO lectura, tolerando que Excel lo tenga abierto en modo lectura.</summary>
    private static ZipArchive Abrir(string ruta)
    {
        var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return new ZipArchive(stream, ZipArchiveMode.Read);
    }

    private static string RutaPlantillaV4() => Path.Combine(
        Raiz(), "Docs", "Plantilla estandarizada", "v4", "Formato Remuneracion 2026xx-x Total v4.xlsx");

    private static string RutaSalidaQ1() => Path.Combine(
        Raiz(), "Docs", "Prueba Septiembre 1", "Resultados", "7", "Remuneración 202609-1 Total.xlsx");

    private static string RutaManualQ1() => Path.Combine(
        Raiz(), "Docs", "Prueba Septiembre 1", "Resultados", "Administrativo", "Remuneracion 202609-1 Total.xlsx");

    private static string RutaManualQ2() => Path.Combine(
        Raiz(), "Docs", "Prueba Julio-2", "Resultado", "Remuneracion 202607-2 Total Administrativo.xlsx");

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(
                    dir.FullName,
                    "Docs",
                    "Plantilla estandarizada",
                    "v4",
                    "Formato Remuneracion 2026xx-x Total v4.xlsx")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException(
            "No se encontró la raíz del repositorio (Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx).");
    }

    private static XDocument Hoja(ZipArchive zip, string nombre)
    {
        var libro = LeerXml(zip, PartWorkbook);
        var hoja = libro.Root?.Element(Ns + "sheets")?.Elements(Ns + "sheet")
            .FirstOrDefault(sheet => string.Equals(sheet.Attribute("name")?.Value, nombre, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"El workbook no trae la hoja '{nombre}'.");

        var id = hoja.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))!.Value;
        var rels = LeerXml(zip, PartRels);
        var destino = rels.Root?.Elements()
            .FirstOrDefault(rel => rel.Attribute("Id")?.Value == id)?
            .Attribute("Target")?.Value ?? throw new InvalidOperationException($"Sin relación para la hoja '{nombre}'.");

        var parte = destino.StartsWith('/') ? destino[1..] : "xl/" + destino;
        return LeerXml(zip, parte);
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
            .FirstOrDefault(row => row.Attribute("r")?.Value == numero)?
            .Elements(Ns + "c")
            .FirstOrDefault(celda => string.Equals(celda.Attribute("r")?.Value, referencia, StringComparison.OrdinalIgnoreCase));
    }

    private static string? FormulaDe(XDocument hoja, string referencia) =>
        Celda(hoja, referencia)?.Element(Ns + "f")?.Value;

    private static string? AtributoFormula(XDocument hoja, string referencia, string atributo) =>
        Celda(hoja, referencia)?.Element(Ns + "f")?.Attribute(atributo)?.Value;

    /// <summary>Valor numérico cacheado (<c>&lt;v&gt;</c>) de la celda. Solo LECTURA de estado.</summary>
    private static decimal ValorNumerico(XDocument hoja, string referencia)
    {
        var texto = Celda(hoja, referencia)?.Element(Ns + "v")?.Value;
        Assert.False(string.IsNullOrEmpty(texto), $"{referencia} no trae <v> (se esperaba un valor medible).");
        return decimal.Parse(texto!, NumberStyles.Float, CultureInfo.InvariantCulture);
    }

    /// <summary>Ocurrencias del fragmento en la fórmula, con frontera numérica (no `F15` dentro de `F150`).</summary>
    private static int ContarReferencia(string formula, string fragmento) =>
        Regex.Matches(formula, Regex.Escape(fragmento) + "(?![0-9])").Count;

    /// <summary>
    /// Cuerpo comparable de la fórmula del auto-chequeo: quita el comparando (<c>=M9</c>), el término
    /// AJUSTES (solo lo trae la geometría canónica; en Q1 vale 0) y la referencia de interventoría
    /// (es el ÚNICO operando que este plan cambia); normaliza el número de fila de cada referencia
    /// (el layout del manual usa sus propias anclas — R-TRAMPA-ANCLA).
    /// </summary>
    private static string NormalizarCuerpo(string formula, int ase)
    {
        var cuerpo = formula;
        var comparando = $"=M{9 + ase}";
        if (cuerpo.EndsWith(comparando, StringComparison.Ordinal))
        {
            cuerpo = cuerpo[..^comparando.Length];
        }

        cuerpo = cuerpo.Replace($"+'AJUSTES - SF-T'!M{47 + ase}", string.Empty, StringComparison.Ordinal);
        cuerpo = Regex.Replace(cuerpo, @"INTERVENTORIA![A-Z]+\d+", "INTERVENTORIA!#", RegexOptions.None);
        cuerpo = Regex.Replace(cuerpo, @"!(?<col>[A-Z]{1,3})\d+", "!${col}#", RegexOptions.None);
        return cuerpo;
    }
}
