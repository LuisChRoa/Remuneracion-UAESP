using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 44 (T0, F0 — S1.1/S1.2/S1.3/S1.4): evidencia READ-ONLY congelada desde bytes para el
/// llenado R1 multicolumna (fuente x canal x concepto x componentes).
///
/// CERO código de producción: este fixture + sus asserts SON la evidencia. Lectura SOLO con BCL
/// (<see cref="ZipArchive"/> + <see cref="XDocument"/>): sin Excel/COM, sin OpenXml SDK, sin
/// ExcelDataReader. El OFICIAL nunca se modifica (SHA256 verificado antes/después).
///
/// Hallazgos congelados aquí (verificación 2026-10-16 por ZIP+XML directo):
/// E1: encabezados F..Z del OFICIAL (Total…Rec.Intereses) y F..U de la FUENTE (Total…Rec.Intereses).
/// E2: las 5 secciones-1.ª del OFICIAL (ASE1 6..15, ASE2 110..119, ASE3 289..298, ASE4 450..459,
///     ASE5 603..612) traen CERO filas <c>E='Vlr Intereses'</c>.
/// E3: la FUENTE ASE5 SI trae <c>Vlr Intereses</c> en la 1.ª sección (fila 8: F=402,40; fila 11:
///     F=7.697,23) con desglose G..U; la FUENTE ASE1 NO (1.ª = filas 5..10 sin Intereses).
/// E4: el MANUAL legacy ASE5 (bloque 501) SI registraba esas dos filas (R506 F=402,40; R509
///     F=7.697,23) — hueco de PLANTILLA (no se toca la plantilla; se reporta).
/// E5: anclas OCCIDENTE ASE5 del OFICIAL: F700=3 términos, F680=3 términos Mes, F682/F702 literal 0;
///     F628/F659 (Oportuno#1/#2 OCCIDENTE) SIN consumidor en ninguna fórmula del libro (huérfanas).
/// E6: micro-punto F654/F661: F654 (Oportuno#2, D=H) y F661 (Aplic#2, Subsidio) están en 0 en el
///     OFICIAL; la FUENTE trae DOS filas con -187.855,98 (R47 Subsidio / R48 D=H) → el duplicado es
///     un artefacto de ESCRITURA (cursor físico de filas sin C), no de emisión del lector.
/// E7: encabezado DUPLICADO en la plantilla (T y Y = 'Intereses Cte Financiac. Aseo'); el ancla de
///     empresa consume T+W+X+Y+Z+S → Y NO es espejo de T (no debe recibir el mismo valor fuente).
/// E8: la colocación por (firma + SECCIÓN ordinal) de las filas SIN C de ASE5 no tiene slot
///     (Componente/Subsidio de Oportuno#1 vs Oportuno#1 del oficial que es VS/VI) → STOP (ver doc).
/// </summary>
public sealed class R1FuenteCanalConceptoT0Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string HojaR1 = "Reporte Componentes R1";
    private const string Sha256Oficial = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    // Sección-1.ª del OFICIAL por ASE (inicio/fin inclusive; la primera frontera Mes/Aplicacion es fin+1).
    private static readonly (int Ase, int Inicio, int Fin)[] PrimeraSeccionOficial =
    [
        (1, 6, 15),
        (2, 110, 119),
        (3, 289, 298),
        (4, 450, 459),
        (5, 603, 612),
    ];

    // Encabezados F..Z del OFICIAL (fila 5 ASE1 / fila 602 ASE5) congelados desde bytes.
    private static readonly (string Celda, string Texto)[] EncabezadosOficial =
    [
        ("F", "Total"), ("G", "Componente TDF"), ("H", "Componente TTL"), ("I", "Componente TVIAT"),
        ("J", "Aprovechamiento"), ("K", "CCSA Prest.Aprov."), ("L", "SERVICIO ESPECIALES"),
        ("M", "Componente TCS"), ("N", "Componente TLU"), ("O", "Componente TBL"), ("P", "Componente TRT"),
        ("Q", "CCSA Prest. No Aprov."), ("R", "Ajuste Decena"), ("S", "Ajuste Recargo Aseo"),
        ("T", "Intereses Cte Financiac. Aseo"), ("U", "Exonera-Recargo Aseo"), ("V", "Exonera-Ajuste Decena Aseo"),
        ("W", "Exonera-Recargo Aseo"), ("X", "Exonera-Recargo Aseo"), ("Y", "Intereses Cte Financiac. Aseo"),
        ("Z", "Rec.Intereses"),
    ];

    // Encabezados F..U de la FUENTE R1 (fila 4, ASE1 y ASE5) — T=Intereses Cte, U=Rec.Intereses.
    private static readonly (string Celda, string Texto)[] EncabezadosFuente =
    [
        ("F", "Total"), ("G", "Componente TDF"), ("H", "Componente TTL"), ("I", "Componente TVIAT"),
        ("J", "Aprovechamiento"), ("K", "CCSA Prest.Aprov."), ("L", "SERVICIO ESPECIALES"),
        ("M", "Componente TCS"), ("N", "Componente TLU"), ("O", "Componente TBL"), ("P", "Componente TRT"),
        ("Q", "CCSA Prest. No Aprov."), ("R", "Ajuste Decena"), ("S", "Ajuste Recargo Aseo"),
        ("T", "Intereses Cte Financiac. Aseo"), ("U", "Rec.Intereses"),
    ];

    // ── E1: encabezados ─────────────────────────────────────────────────────────────────────────

    [Fact]
    public void T0_Sha256_OficialIntacto()
    {
        using var stream = File.OpenRead(Oficial());
        Assert.Equal(Sha256Oficial, Convert.ToHexString(SHA256.HashData(stream)));
    }

    [Fact]
    public void T0_Oficial_EncabezadosFZ_PorBloque()
    {
        using var zip = ZipFile.OpenRead(Oficial());
        var hoja = Hoja(zip, HojaR1);
        var ss = Compartidas(zip);
        foreach (var fila in new[] { 5, 602 })
        {
            foreach (var (letra, texto) in EncabezadosOficial)
            {
                Assert.Equal(texto, TextoCelda(hoja, ss, $"{letra}{fila}"));
            }
        }
    }

    [Fact]
    public void T0_Fuente_EncabezadosFU_Ase1YAse5()
    {
        foreach (var aseId in new[] { 1, 5 })
        {
            using var zip = ZipFile.OpenRead(FuenteR1(aseId));
            var hoja = Hoja(zip, "Sheet1");
            var ss = Compartidas(zip);
            foreach (var (letra, texto) in EncabezadosFuente)
            {
                Assert.Equal(texto, TextoCelda(hoja, ss, $"{letra}4"));
            }
        }
    }

    [Fact]
    public void T0_Oficial_TYY_EncabezadoDuplicado()
    {
        using var zip = ZipFile.OpenRead(Oficial());
        var hoja = Hoja(zip, HojaR1);
        var ss = Compartidas(zip);
        Assert.Equal("Intereses Cte Financiac. Aseo", TextoCelda(hoja, ss, "T5"));
        Assert.Equal("Intereses Cte Financiac. Aseo", TextoCelda(hoja, ss, "Y5"));
        // U/V/W/X del OFICIAL son Exonera-* (sin contraparte en la fuente).
        Assert.Equal("Exonera-Recargo Aseo", TextoCelda(hoja, ss, "U5"));
        Assert.Equal("Exonera-Ajuste Decena Aseo", TextoCelda(hoja, ss, "V5"));
        Assert.Equal("Exonera-Recargo Aseo", TextoCelda(hoja, ss, "W5"));
        Assert.Equal("Exonera-Recargo Aseo", TextoCelda(hoja, ss, "X5"));
    }

    // ── E2/E3: hueco de plantilla en la 1.ª sección ──────────────────────────────────────────────

    [Fact]
    public void T0_Oficial_PrimeraSeccion_SinVlrIntereses_CincoAse()
    {
        using var zip = ZipFile.OpenRead(Oficial());
        var hoja = Hoja(zip, HojaR1);
        var ss = Compartidas(zip);
        foreach (var (ase, inicio, fin) in PrimeraSeccionOficial)
        {
            var etiquetas = new List<string>();
            for (var fila = inicio; fila <= fin; fila++)
            {
                etiquetas.Add(TextoCelda(hoja, ss, $"E{fila}"));
            }

            Assert.DoesNotContain("Vlr Intereses", etiquetas);
            Assert.True(etiquetas.Any(e => e == "Vlr Servicio"), $"ASE {ase}: la 1.ª sección no trae 'Vlr Servicio'.");
        }
    }

    [Fact]
    public void T0_FuenteAse5_PrimeraSeccion_TraeVlrIntereses_ConDesglose()
    {
        using var zip = ZipFile.OpenRead(FuenteR1(5));
        var hoja = Hoja(zip, "Sheet1");
        var ss = Compartidas(zip);

        // Fila 8 (canal O) y fila 11 (canal T) de la 1.ª sección (antes de la frontera 'Aplicacion' fila 14).
        Assert.Equal("Vlr Intereses", TextoCelda(hoja, ss, "E8"));
        Assert.Equal("402.4", Celda(hoja, "F8")?.Element(Ns + "v")?.Value);
        Assert.Equal("32.54", Celda(hoja, "G8")?.Element(Ns + "v")?.Value);
        Assert.Equal("18.34", Celda(hoja, "H8")?.Element(Ns + "v")?.Value);

        Assert.Equal("Vlr Intereses", TextoCelda(hoja, ss, "E11"));
        Assert.Equal("7697.23", Celda(hoja, "F11")?.Element(Ns + "v")?.Value);
        Assert.Equal("484.57", Celda(hoja, "G11")?.Element(Ns + "v")?.Value);
        Assert.Equal("0.14", Celda(hoja, "R11")?.Element(Ns + "v")?.Value);
    }

    [Fact]
    public void T0_FuenteAse1_PrimeraSeccion_SinVlrIntereses()
    {
        using var zip = ZipFile.OpenRead(FuenteR1(1));
        var hoja = Hoja(zip, "Sheet1");
        var ss = Compartidas(zip);
        // 1.ª sección ASE1 = filas 5..10 (frontera 'Mes' en 11).
        for (var fila = 5; fila <= 10; fila++)
        {
            Assert.NotEqual("Vlr Intereses", TextoCelda(hoja, ss, $"E{fila}"));
        }
    }

    [Fact]
    public void T0_ManualLegacyAse5_SiRegistrabaPrimeraSeccionIntereses()
    {
        using var zip = ZipFile.OpenRead(Manual());
        var hoja = Hoja(zip, HojaR1);
        var ss = Compartidas(zip);
        Assert.Equal("Vlr Intereses", TextoCelda(hoja, ss, "E506"));
        Assert.Equal("402.4", Celda(hoja, "F506")?.Element(Ns + "v")?.Value);
        Assert.Equal("Vlr Intereses", TextoCelda(hoja, ss, "E509"));
        Assert.Equal("7697.23", Celda(hoja, "F509")?.Element(Ns + "v")?.Value);
    }

    // ── E5: anclas OCCIDENTE ASE5 + huérfanas ───────────────────────────────────────────────────

    [Fact]
    public void T0_Oficial_AnclasOccidenteAse5_YHuerfanas()
    {
        using var zip = ZipFile.OpenRead(Oficial());
        var hoja = Hoja(zip, HojaR1);
        var ss = Compartidas(zip);

        Assert.Equal("F647+F673+F612-L612-L647", FormulaDe(hoja, "F700"));
        Assert.Equal("F674+F648+F613-L613-L648", FormulaDe(hoja, "F680"));
        Assert.Null(FormulaDe(hoja, "F682"));
        Assert.Null(FormulaDe(hoja, "F702"));

        // F628 / F659 = filas Oportuno de OCCIDENTE (D=Total, C=OCCIDENTE) sin consumidor.
        Assert.Equal("OCCIDENTE", TextoCelda(hoja, ss, "C628"));
        Assert.Equal("Total", TextoCelda(hoja, ss, "D628"));
        Assert.Equal("OCCIDENTE", TextoCelda(hoja, ss, "C659"));
        Assert.Equal("Total", TextoCelda(hoja, ss, "D659"));

        // Grep exhaustivo de consumidores de F628/F659/L628/L659 en TODAS las hojas del libro.
        var consumidores = TodasLasFormulas(zip)
            .Where(f => ReferenciaAlguna(f, ["F628", "F659", "L628", "L659"]))
            .ToList();
        Assert.Empty(consumidores);
    }

    // ── E6: micro-punto F654/F661 (aislamiento lector vs escritor) ───────────────────────────────

    [Fact]
    public void T0_MicroPunto_F654_F661_EsArtefactoDeEscritura()
    {
        // Plantilla (OFICIAL): F654 (Oportuno#2, D=H) y F661 (Aplic#2, Subsidio) en 0 sin <f>.
        using (var zip = ZipFile.OpenRead(Oficial()))
        {
            var hoja = Hoja(zip, HojaR1);
            var ss = Compartidas(zip);
            Assert.Equal("H", TextoCelda(hoja, ss, "D654"));
            Assert.Equal("Total", TextoCelda(hoja, ss, "E654"));
            Assert.Equal("0", Celda(hoja, "F654")?.Element(Ns + "v")?.Value);
            Assert.Equal("Subsidio(-)/Contribucion(+)", TextoCelda(hoja, ss, "E661"));
            Assert.Equal("0", Celda(hoja, "F661")?.Element(Ns + "v")?.Value);
        }

        // Fuente ASE5: la 1.ª sección emite UNA fila Subsidio (R47) y UNA D=H (R48), ambas con el
        // MISMO F=-187.855,98; el lector las emite UNA sola vez cada una.
        using (var zip = ZipFile.OpenRead(FuenteR1(5)))
        {
            var hoja = Hoja(zip, "Sheet1");
            var ss = Compartidas(zip);
            Assert.Equal("Subsidio(-)/Contribucion(+)", TextoCelda(hoja, ss, "E47"));
            Assert.Equal("-187855.98", Celda(hoja, "F47")?.Element(Ns + "v")?.Value);
            Assert.Equal("H", TextoCelda(hoja, ss, "D48"));
            Assert.Equal("Total", TextoCelda(hoja, ss, "E48"));
            Assert.Equal("-187855.98", Celda(hoja, "F48")?.Element(Ns + "v")?.Value);
        }
    }

    // ── E7: el ancla de empresa consume T,W,X,Y,Z,S (Y no es espejo de T) ────────────────────────

    [Fact]
    public void T0_AnclaEmpresa_ConsumeTYWXZ_S()
    {
        using var zip = ZipFile.OpenRead(Oficial());
        var formulas = TodasLasFormulas(zip).ToList();
        Assert.Contains(formulas, f =>
            f.Contains("'Reporte Componentes R1'!T95", StringComparison.Ordinal)
            && f.Contains("'Reporte Componentes R1'!Y95", StringComparison.Ordinal)
            && f.Contains("'Reporte Componentes R1'!Z95", StringComparison.Ordinal)
            && f.Contains("'Reporte Componentes R1'!S95", StringComparison.Ordinal));
    }

    // ── E8: la colocación por (firma + sección ordinal) no tiene slot para filas sin C de ASE5 ──

    [Fact]
    public void T0_Ase5_Oportuno1_OficialEsVsVi_FuenteEsComponenteSubsidio()
    {
        // OFICIAL Oportuno#1 (frontera Mes 613; fin en Aplicacion 629) = filas VS/VI + códigos.
        using (var zip = ZipFile.OpenRead(Oficial()))
        {
            var hoja = Hoja(zip, HojaR1);
            var ss = Compartidas(zip);
            Assert.Equal("Mes", TextoCelda(hoja, ss, "B613"));
            Assert.Equal("Vlr Servicio", TextoCelda(hoja, ss, "E615"));
            Assert.Equal("Vlr Intereses", TextoCelda(hoja, ss, "E616"));
            Assert.Equal("Aplicacion nuevos x reversion", TextoCelda(hoja, ss, "B629"));
        }

        // FUENTE ASE5 Oportuno#1 (frontera Mes 33; fin en Aplicacion 40) = Componente/Subsidio.
        using (var zip = ZipFile.OpenRead(FuenteR1(5)))
        {
            var hoja = Hoja(zip, "Sheet1");
            var ss = Compartidas(zip);
            Assert.Equal("Mes", TextoCelda(hoja, ss, "B33"));
            Assert.Equal("Componente", TextoCelda(hoja, ss, "A34"));
            Assert.Equal("Subsidio(-)/Contribucion(+)", TextoCelda(hoja, ss, "E35"));
            Assert.Equal("H", TextoCelda(hoja, ss, "D36"));
            Assert.Equal("Aplicacion nuevos x reversion", TextoCelda(hoja, ss, "B40"));
        }
    }

    // ── Lectura BCL (ZIP + XML) ─────────────────────────────────────────────────────────────────

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

    private static string Oficial() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string Manual() =>
        Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Resultado", "Remuneracion 202607-2 Total Administrativo.xlsx");

    private static string FuenteR1(int aseId)
    {
        var carpeta = Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Insumos", CarpetasAse[aseId - 1]);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta explícita: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("Recaudoporcomponente", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Recaudoporcomponente del ASE{aseId} en {carpeta}");
        return archivo!;
    }

    private static XDocument LeerXml(ZipArchive zip, string entrada)
    {
        var entry = zip.GetEntry(entrada) ?? throw new InvalidOperationException($"El xlsx no trae '{entrada}'.");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static XDocument Hoja(ZipArchive zip, string nombre)
    {
        var libro = LeerXml(zip, "xl/workbook.xml");
        var sheet = libro.Root?.Element(Ns + "sheets")?.Elements(Ns + "sheet")
            .FirstOrDefault(s => string.Equals(s.Attribute("name")?.Value, nombre, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"El workbook no trae la hoja '{nombre}'.");
        var id = sheet.Attribute(RelNs + "id")!.Value;
        var rels = LeerXml(zip, "xl/_rels/workbook.xml.rels");
        var destino = rels.Root?.Elements()
            .FirstOrDefault(r => r.Attribute("Id")?.Value == id)?.Attribute("Target")?.Value
            ?? throw new InvalidOperationException($"Sin rel para la hoja '{nombre}'.");
        if (!destino.StartsWith("xl/", StringComparison.Ordinal))
        {
            destino = "xl/" + destino.TrimStart('/');
        }

        return LeerXml(zip, destino);
    }

    private static List<string> Compartidas(ZipArchive zip)
    {
        var doc = LeerXml(zip, "xl/sharedStrings.xml");
        return doc.Root?.Elements(Ns + "si").Select(si => si.Value).ToList() ?? [];
    }

    private static XElement? Fila(XDocument hoja, string numero) =>
        hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row")
            .FirstOrDefault(r => r.Attribute("r")?.Value == numero);

    private static XElement? Celda(XDocument hoja, string referencia)
    {
        var numero = new string(referencia.Where(char.IsAsciiDigit).ToArray());
        return Fila(hoja, numero)?.Elements(Ns + "c")
            .FirstOrDefault(c => string.Equals(c.Attribute("r")?.Value, referencia, StringComparison.OrdinalIgnoreCase));
    }

    private static string TextoCelda(XDocument hoja, List<string> compartidas, string referencia)
    {
        var celda = Celda(hoja, referencia);
        if (celda is null)
        {
            return string.Empty;
        }

        if (celda.Attribute("t")?.Value == "s" && int.TryParse(celda.Element(Ns + "v")?.Value, out var indice)
            && indice >= 0 && indice < compartidas.Count)
        {
            return compartidas[indice];
        }

        return celda.Element(Ns + "inlineStr")?.Value ?? celda.Element(Ns + "v")?.Value ?? string.Empty;
    }

    private static string? FormulaDe(XDocument hoja, string referencia) =>
        Celda(hoja, referencia)?.Element(Ns + "f")?.Value;

    private static IEnumerable<string> TodasLasFormulas(ZipArchive zip)
    {
        foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal)))
        {
            using var stream = entry.Open();
            var doc = XDocument.Load(stream);
            foreach (var f in doc.Descendants(Ns + "f"))
            {
                if (!string.IsNullOrEmpty(f.Value))
                {
                    yield return f.Value;
                }
            }
        }
    }

    private static bool ReferenciaAlguna(string formula, IReadOnlyList<string> referencias)
    {
        foreach (var referencia in referencias)
        {
            var i = 0;
            while ((i = formula.IndexOf(referencia, i, StringComparison.OrdinalIgnoreCase)) >= 0)
            {
                var antes = i == 0 ? '\0' : formula[i - 1];
                var despues = i + referencia.Length >= formula.Length ? '\0' : formula[i + referencia.Length];
                if (!char.IsLetterOrDigit(antes) && antes != '$' && !char.IsLetterOrDigit(despues) && despues != '$')
                {
                    return true;
                }

                i += referencia.Length;
            }
        }

        return false;
    }
}
