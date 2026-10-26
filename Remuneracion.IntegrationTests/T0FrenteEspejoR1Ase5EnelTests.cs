using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 41 (T0 del frente espejo R1 — diagnóstico read-only, CERO código de producción): congela
/// desde bytes (BCL <c>ZipArchive</c> + <c>XDocument</c>; sin Excel/COM/OpenXml SDK) la evidencia del
/// defecto E1-2 del Plan 40: el espejo R1 <b>encuentra</b> las filas por empresa en la fuente pero
/// las coloca en slots equivocados del bloque fijo del oficial, dejando vacíos los slots que las
/// anclas <c>REMUNERACION_*!D9:D13</c> referencian.
///
/// Documento que manda: <c>Docs/T0-Plan41-Frente-EspejoR1-Ase5Enel.md</c>. Este fixture SOLO aserta el
/// estado congelado (el defecto vivo) para que un cambio futuro (el fix) rompa el test y exija
/// re-congelar. NO toca producción.
///
/// Prohibido: <c>Docs/Insumos</c> raíz; usar cachés <c>&lt;v&gt;</c> de celdas-fórmula como oráculo.
/// </summary>
public sealed class T0FrenteEspejoR1Ase5EnelTests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string HojaR1 = "Reporte Componentes R1";

    // ── Anclas por empresa del oficial (Reporte Componentes R1), leídas del disco. ───────────────
    // REMUNERACION_*!D9:D13 -> R1 F-cell (ASE1..ASE5). La familia REMUNERACION_* NO comparte anclas
    // con CONSOLIDADO_TOTAL RECAUDO (que usa F75/F254/F415/F568/F680 = TOT_OPT por ASE).
    private static readonly (string Hoja, string[] Anclas)[] AnclasPorEmpresa =
    [
        ("REMUNERACION Reciprocidad EAAB", ["F80", "F259", "F420", "F573", "F685"]),
        ("REMUNERACION_ENEL", ["F85", "F264", "F425", "F578", "F690"]),
        ("REMUNERACION_ENERBIT", ["F90", "F269", "F430", "F583", "F695"]),
        ("REMUNERACION_OCCIDENTE_ Directa", ["F95", "F274", "F435", "F588", "F700"]),
        ("REMUNERACION_EAAB-CL", ["F100", "F279", "F440", "F593", "F705"]),
    ];

    // ── Fórmulas de las anclas ENEL del oficial (texto-<f> byte-fiel; el oráculo es el TEXTO). ────
    private static readonly (string Celda, string Formula)[] EnelAnclasOficial =
    [
        ("F85", "F62+F36+F8-L8-L36"),
        ("F264", "F193+F150+F112--L112-L150"),
        ("F425", "F345+F319+F291-L291-L319"),
        ("F578", "F533+F490+F452-L452-L490"),
        ("F690", "F663+F633-L633"),
    ];

    [Fact]
    public void T0_Oficial_AnclasPorEmpresa_SonLasCincoFamilias()
    {
        using var zip = ZipFile.OpenRead(RutaOficial());
        var r1 = Hoja(zip, HojaR1);
        foreach (var (hojaEmpresa, anclas) in AnclasPorEmpresa)
        {
            var empresa = Hoja(zip, hojaEmpresa);
            for (var ase = 1; ase <= 5; ase++)
            {
                Assert.Equal($"'Reporte Componentes R1'!{anclas[ase - 1]}", FormulaDe(empresa, $"D{8 + ase}"));
            }
        }
    }

    [Fact]
    public void T0_Oficial_EnelAnclas_TextoFormulaCongelado()
    {
        using var zip = ZipFile.OpenRead(RutaOficial());
        var r1 = Hoja(zip, HojaR1);
        foreach (var (celda, formula) in EnelAnclasOficial)
        {
            Assert.Equal(formula, FormulaDe(r1, celda));
        }
    }

    [Fact]
    public void T0_App_Ase5Enel_DatoHalladoPeroSlotsVacios()
    {
        // El defecto E1-2: la app escribió las 2 filas ENEL de la fuente en los DOS PRIMEROS slots
        // (605/618) y dejó VACÍOS los slots que la ancla F690=F663+F633-L633 referencia (633/663).
        using var zip = ZipFile.OpenRead(RutaResultado5());
        var r1 = Hoja(zip, HojaR1);

        Assert.Equal("12823196468.26", ValorDe(r1, "F605"));
        Assert.Equal("-1113041546.26", ValorDe(r1, "F618"));
        Assert.Equal("0", ValorDe(r1, "F633"));
        Assert.Equal("0", ValorDe(r1, "L633"));
        Assert.Equal("0", ValorDe(r1, "F663"));

        // La ancla evalúa 0 (F663+F633-L633); el manual del administrativo da 11.701.157.844,90.
        Assert.Equal("F663+F633-L633", FormulaDe(r1, "F690"));
    }

    [Fact]
    public void T0_App_Ase3Enel_SlotsVacios()
    {
        // Mismo mecanismo en ASE3: la fuente trae 4 filas ENEL; la app llenó los 4 primeros slots
        // (291/304/319/334) y dejó vacío el 5.º (345), que la ancla F425 usa.
        using var zip = ZipFile.OpenRead(RutaResultado5());
        var r1 = Hoja(zip, HojaR1);
        Assert.Equal("680538.03", ValorDe(r1, "F291"));
        Assert.Equal("15107063589.67", ValorDe(r1, "F304"));
        Assert.Equal("627141.97", ValorDe(r1, "F319"));
        Assert.Equal("-1096861177.67", ValorDe(r1, "F334"));
        Assert.Equal("0", ValorDe(r1, "F345"));
    }

    [Fact]
    public void T0_Fuente_Ase5_TieneDosFilasEnel()
    {
        var ruta = RutaFuenteAse5();
        using var zip = ZipFile.OpenRead(ruta);
        var hoja = Hoja(zip, "Sheet1");
        var compartidas = Compartidas(zip);

        var filasEnel = new List<(int Fila, string Total)>();
        foreach (var row in hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row") ?? [])
        {
            var numero = int.Parse(row.Attribute("r")!.Value, CultureInfo.InvariantCulture);
            var c = row.Elements(Ns + "c").FirstOrDefault(x => x.Attribute("r")?.Value == $"C{numero}");
            if (c is not null && TextoCelda(c, compartidas) == "ENEL")
            {
                var f = row.Elements(Ns + "c").FirstOrDefault(x => x.Attribute("r")?.Value == $"F{numero}");
                filasEnel.Add((numero, f?.Element(Ns + "v")?.Value ?? string.Empty));
            }
        }

        Assert.Equal(2, filasEnel.Count);
        Assert.Equal((18, "12823196468.26"), filasEnel[0]);
        Assert.Equal((43, "-1113041546.26"), filasEnel[1]);
    }

    [Fact]
    public void T0_Manual_EnelAnclas_EsperadoDelAdministrativo()
    {
        // Oráculo: el manual del administrativo (geometría legacy). Correspondencia por (empresa, ASE):
        // manual F63/F216/F353/F478/F568 <-> oficial F85/F264/F425/F578/F690. NUNCA se comparan
        // direcciones entre plantillas distintas (R-TRAMPA-ANCLA).
        var ruta = Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Resultado",
            "Remuneracion 202607-2 Total Administrativo.xlsx");
        using var zip = ZipFile.OpenRead(ruta);
        var r1 = Hoja(zip, HojaR1);
        Assert.Equal("16560599150.5", ValorDe(r1, "F63"));
        Assert.Equal("19023525223.110004", ValorDe(r1, "F216"));
        Assert.Equal("13997194578.870001", ValorDe(r1, "F353"));
        Assert.Equal("6584787380.4700003", ValorDe(r1, "F478"));
        Assert.Equal("11701157844.9", ValorDe(r1, "F568"));
    }

    // ── Mini-T0 (Plan 41, R-T-1): (a) OCCIDENTE ASE5 F700 vs F578 (estructura) ────────────────────

    [Fact]
    public void MiniT0_Oficial_OccidenteAse5_F700_TresTerminos()
    {
        // El oficial F700 = F647+F673+F612-L612-L647 (3 términos F, 2 L). Operandos OCCIDENTE/Total
        // (en 0 en el template) en las secciones 1.ª (612), Aplic#1 (647) y Aplic#2 (673).
        using var zip = ZipFile.OpenRead(RutaOficial());
        var r1 = Hoja(zip, HojaR1);
        var ss = Compartidas(zip);
        Assert.Equal("F647+F673+F612-L612-L647", FormulaDe(r1, "F700"));
        foreach (var fila in new[] { 612, 647, 673 })
        {
            Assert.Equal("0", ValorDe(r1, $"F{fila}"));
            Assert.Equal("0", ValorDe(r1, $"L{fila}"));
            Assert.Equal("OCCIDENTE", TextoCelda(Celda(r1, $"C{fila}")!, ss));
            Assert.Equal("Total", TextoCelda(Celda(r1, $"D{fila}")!, ss));
        }
    }

    [Fact]
    public void MiniT0_Manual_OccidenteAse5_F578_DosTerminos_EsLaAnomalia()
    {
        // El manual F578 = F530+F551-L530 (2 términos). Sus ASE1-4 SÍ tienen 3 términos (misma
        // estructura que el oficial) → la ancla ASE5 del manual es la anomalía, no el oficial.
        using var zip = ZipFile.OpenRead(RutaManual());
        var r1 = Hoja(zip, HojaR1);
        Assert.Equal("F530+F551-L530", FormulaDe(r1, "F578"));
        Assert.Equal("229202451.62", ValorDe(r1, "F530"));
        Assert.Equal("5904652.2300000004", ValorDe(r1, "L530"));
        Assert.Equal("81691821.420000002", ValorDe(r1, "F551"));
        Assert.Equal("F47+F31+F11-L11-L31", FormulaDe(r1, "F73"));
        Assert.Equal("F280+F264+F243-L243-L264", FormulaDe(r1, "F363"));
    }

    // ── Mini-T0 (Plan 41, R-T-1): (b) geometría de SALDOS POR NOTA / RETRIBUCION NEGATIVA ─────────

    [Fact]
    public void MiniT0_Oficial_SaldosNota_SeisFilasConVlrIntereses_VisibleC8I8()
    {
        using var zip = ZipFile.OpenRead(RutaOficial());
        var hoja = Hoja(zip, "SALDOS POR NOTA");
        var ss = Compartidas(zip);
        Assert.Equal("Vlr Servicio", TextoCelda(Celda(hoja, "B3")!, ss));
        Assert.Equal("Vlr Intereses", TextoCelda(Celda(hoja, "B4")!, ss));
        Assert.Equal("Componente", TextoCelda(Celda(hoja, "A5")!, ss));
        Assert.Equal("Subsidio(-)/Contribucion(+)", TextoCelda(Celda(hoja, "B6")!, ss));
        Assert.Equal("Subs/Cont", TextoCelda(Celda(hoja, "A7")!, ss));
        Assert.Equal("Total", TextoCelda(Celda(hoja, "A8")!, ss));

        // El visible es Cn-In sobre la fila Total OFICIAL (8/21/34/47/60), no la legacy (7/20/…).
        Assert.Equal("C8-I8", FormulaDe(hoja, "C10"));
        Assert.Equal("C21-I21", FormulaDe(hoja, "C23"));
        Assert.Equal("C34-I34", FormulaDe(hoja, "C36"));
        Assert.Equal("C47-I47", FormulaDe(hoja, "C49"));
        Assert.Equal("C60-I60", FormulaDe(hoja, "C62"));

        // Los 5 bloques traen la fila 'Vlr Intereses' (Vlr Servicio + 1).
        foreach (var (inicioVlrServicio, inicioVlrIntereses) in new[] { (3, 4), (16, 17), (29, 30), (42, 43), (55, 56) })
        {
            Assert.Equal("Vlr Servicio", TextoCelda(Celda(hoja, $"B{inicioVlrServicio}")!, ss));
            Assert.Equal("Vlr Intereses", TextoCelda(Celda(hoja, $"B{inicioVlrIntereses}")!, ss));
        }
    }

    [Fact]
    public void MiniT0_Oficial_RetribucionNegativa_SeisFilasVlrInteresesPrimero()
    {
        using var zip = ZipFile.OpenRead(RutaOficial());
        var hoja = Hoja(zip, "RETRIBUCION NEGATIVA");
        var ss = Compartidas(zip);
        Assert.Equal("Vlr Intereses", TextoCelda(Celda(hoja, "B3")!, ss));
        Assert.Equal("Vlr Servicio", TextoCelda(Celda(hoja, "B4")!, ss));
        Assert.Equal("Total", TextoCelda(Celda(hoja, "A8")!, ss));
        Assert.Equal("C8-I8", FormulaDe(hoja, "C10"));
        Assert.Equal("C60-I60", FormulaDe(hoja, "C62"));
    }

    // ── Mini-T0 (Plan 41, R-T-1): (c) tabla J / R26:R31 (auto-chequeo S26:S31) ────────────────────

    [Fact]
    public void MiniT0_Oficial_Interventoria_TablaJYAutoChequeoS_NoEsInvariante()
    {
        using var zip = ZipFile.OpenRead(RutaOficial());
        var hoja = Hoja(zip, "INTERVENTORIA");

        // Tabla J (anual) = los valores de la 2.ª quincena (M26:M30); total J20.
        Assert.Equal("189185988", ValorDe(hoja, "J15"));
        Assert.Equal("266580256", ValorDe(hoja, "J16"));
        Assert.Equal("154788536", ValorDe(hoja, "J17"));
        Assert.Equal("120391083", ValorDe(hoja, "J18"));
        Assert.Equal("128990446", ValorDe(hoja, "J19"));
        Assert.Equal("859936309", ValorDe(hoja, "J20"));
        Assert.Equal("189185988", ValorDe(hoja, "M26"));
        Assert.Equal("189185987", ValorDe(hoja, "N26"));
        Assert.Equal("128990445", ValorDe(hoja, "N30"));

        // Auto-chequeo: S26 = Q26=R26; R26 (maestro shared R26:R31) = J15; Q26 = CONSOLIDADO!L104.
        Assert.Equal("Q26=R26", FormulaDe(hoja, "S26"));
        Assert.Equal("J15", FormulaDe(hoja, "R26"));
        Assert.Equal("'CONSOLIDADO_TOTAL RECAUDO'!L104", FormulaDe(hoja, "Q26"));

        // Q1: Q26 = N26 = 189185987 != R26 = J15 = 189185988 ⇒ S26 FALSE real (no es invariante).
        Assert.NotEqual(ValorDe(hoja, "N26"), ValorDe(hoja, "J15"));
    }

    // ── Lectura BCL (ZIP + XML) ──────────────────────────────────────────────────────────────────

    private static string RutaOficial() =>

        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    /// <summary>
    /// Fuente R1 de ASE5 de julio‑2 resuelta por PATRÓN (<c>5-*</c> + <c>Recaudoporcomponente*.xlsx</c>):
    /// el commit <c>88b3acb</c> reemplazó los insumos de julio‑2 por los vigentes y congelar el nombre
    /// del archivo es la misma clase de defecto que el Plan 59 elimina (cero nombres de archivo exactos).
    /// </summary>
    private static string RutaFuenteAse5()
    {
        var raizInsumos = Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Insumos");
        var carpetas = Directory.EnumerateDirectories(raizInsumos, "5-*", SearchOption.TopDirectoryOnly)
            .OrderBy(f => f, StringComparer.Ordinal)
            .ToList();
        Assert.Single(carpetas);

        var archivos = Directory.EnumerateFiles(carpetas[0], "Recaudoporcomponente*.xlsx", SearchOption.TopDirectoryOnly)
            .OrderBy(f => f, StringComparer.Ordinal)
            .ToList();
        Assert.Single(archivos);
        return archivos[0];
    }

    private static string RutaResultado5() =>
        Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Resultado", "Resultado5", "Remuneración 202607-2 Total.xlsx");

    private static string RutaManual() =>
        Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Resultado", "Remuneracion 202607-2 Total Administrativo.xlsx");

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (plantilla estandarizada única).");
    }

    private static XDocument Hoja(ZipArchive zip, string nombre)
    {
        var libro = LeerXml(zip, "xl/workbook.xml");
        var sheet = libro.Root?.Element(Ns + "sheets")?.Elements(Ns + "sheet")
            .FirstOrDefault(s => string.Equals(s.Attribute("name")?.Value, nombre, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"El workbook no trae la hoja '{nombre}'.");
        var id = sheet.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))!.Value;
        var rels = LeerXml(zip, "xl/_rels/workbook.xml.rels");
        var destino = rels.Root?.Elements().FirstOrDefault(r => r.Attribute("Id")?.Value == id)?.Attribute("Target")?.Value
            ?? throw new InvalidOperationException($"Sin rel para la hoja '{nombre}'.");
        return LeerXml(zip, "xl/" + destino);
    }

    private static XDocument LeerXml(ZipArchive zip, string entrada)
    {
        var entry = zip.GetEntry(entrada) ?? throw new InvalidOperationException($"El xlsx no trae '{entrada}'.");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static List<string> Compartidas(ZipArchive zip) =>
        LeerXml(zip, "xl/sharedStrings.xml").Root?.Elements(Ns + "si").Select(si => si.Value).ToList() ?? [];

    private static XElement? Celda(XDocument hoja, string referencia)
    {
        var numero = new string(referencia.Where(char.IsAsciiDigit).ToArray());
        return hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row")
            .FirstOrDefault(r => r.Attribute("r")?.Value == numero)?
            .Elements(Ns + "c")
            .FirstOrDefault(c => string.Equals(c.Attribute("r")?.Value, referencia, StringComparison.OrdinalIgnoreCase));
    }

    private static string? FormulaDe(XDocument hoja, string referencia) =>
        Celda(hoja, referencia)?.Element(Ns + "f")?.Value;

    private static string ValorDe(XDocument hoja, string referencia) =>
        Celda(hoja, referencia)?.Element(Ns + "v")?.Value ?? string.Empty;

    private static string TextoCelda(XElement celda, List<string> compartidas)
    {
        if (celda.Attribute("t")?.Value == "s" && int.TryParse(celda.Element(Ns + "v")?.Value, out var indice)
            && indice >= 0 && indice < compartidas.Count)
        {
            return compartidas[indice];
        }

        return celda.Element(Ns + "inlineStr")?.Value ?? celda.Element(Ns + "v")?.Value ?? string.Empty;
    }
}
