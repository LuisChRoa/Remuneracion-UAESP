using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 36 (T0, R-T-1/R-T-2 — D-A..D-G): contrato de slots congelado desde los bytes del disco de
/// la plantilla estandarizada única
/// (<c>Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx</c>).
///
/// CERO código de producción: este fixture + sus asserts SON el contrato. Lectura SOLO con BCL
/// (<c>System.IO.Compression.ZipArchive</c> + <c>System.Xml.Linq</c>): sin Excel/COM, sin
/// DocumentFormat.OpenXml, sin ExcelDataReader. Si el template cambia un byte contractual, este
/// test va a ROJO antes que cualquier writer (el fixture manda, no este documento).
///
/// Q1..Q4 (§0.2 del plan) están SIN RESPONDER: cada hallazgo se congela tal cual está en disco y
/// se cita como PENDING. No se asume ninguna respuesta.
///
/// Hallazgos congelados aquí (verificación 2026-10-13 por ZIP+XML directo):
/// E1: 40 hojas; R1 A1:AP712 con 111 filas hidden (213-251, 356-412, 552-566);
/// E2: labels-C estándar + slots NUEVO ESQUEMA ASE2 (F135/F160/F182/F200) y ASE4 (F475/F500/F522/F540);
/// E3: 37 masters shared si=0..36 + 1101 seguidoras; calcChain presente; base 2026092 (N3);
///     CONSOLIDADO!D111=D109+N109; sellos G7/K7/D6/D7; DetRetri2026092 editable-cero;
/// E4: AFaseo dual (slots A='AFaseo' B='Total' en F17/F121/F300/F461/F614) + resúmenes D=AFaseo
///     (F105→F17, F284→F120, F445→F299, F598→F460, F710→F614);
/// F5 (nuevo): R649 (ASE5 Componente/Total) trae valores reales F=13359789539.23/L=12154250.72,
///     única celda-dato no-cero del R1 (residuo de la base 2026092; el dato fuente manda).
/// </summary>
public sealed class ContratoSlotR1T0Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string HojaR1 = "Reporte Componentes R1";

    // Rango de filas de DATOS (slots) por ASE, 1-based inclusivo (ver §2.1 del plan).
    private static readonly (int Ase, int Inicio, int Fin)[] RangosDatos =
    [
        (1, 6, 72),
        (2, 110, 210),
        (3, 289, 355),
        (4, 450, 550),
        (5, 603, 676),
    ];

    // Cabeceras de bloque B por ASE (fila → rótulo).
    private static readonly (int Fila, string Rotulo)[] Cabeceras =
    [
        (4, "PROMOAMBIENTAL"),
        (108, "LIME"),
        (287, "CIUDAD LIMPIA"),
        (448, "BOGOTA LIMPIA"),
        (601, "AREA LIMPIA"),
    ];

    // Slots AFaseo variante-A por ASE (celda F del slot A='AFaseo' B='Total').
    private static readonly (int Ase, string Celda)[] SlotsAFaseo =
    [
        (1, "F17"),
        (2, "F121"),
        (3, "F300"),
        (4, "F461"),
        (5, "F614"),
    ];

    // Resúmenes interiores D=AFaseo (celda → texto de fórmula observado en disco).
    private static readonly (string Celda, string Formula)[] ResumenesAFaseo =
    [
        ("F105", "F17"),
        ("F284", "F120"),
        ("F445", "F299"),
        ("F598", "F460"),
        ("F710", "F614"),
    ];

    // Slots NUEVO ESQUEMA (celda F, literal 0 sin fórmula).
    private static readonly string[] SlotsNuevoEsquema =
        ["F135", "F160", "F182", "F200", "F475", "F500", "F522", "F540"];

    // Maestras shared del mirror G:AP (celda, si, ref, texto) — lista NO-TOUCH D-E.
    private static readonly (string Celda, uint Si, string Ref, string Texto)[] Maestras =
    [
        ("H75", 0, "H75:AP75", "H47+H16"),
        ("H77", 1, "H77:AP77", "H32"),
        ("H85", 2, "H85:AP85", "H36+H8"),
        ("H87", 3, "H87:AP87", "H21"),
        ("G95", 4, "G95:AP95", "G46+G15"),
        ("H97", 5, "H97:AP97", "H31"),
        ("F105", 6, "F105:AP105", "F17"),
        ("G254", 7, "G254:AP254", "G171+G120"),
        ("G256", 8, "G256:AP256", "G146"),
        ("G259", 9, "G259:AP259", "G160"),
        ("H261", 10, "H261:AP261", "H135"),
        ("G264", 11, "G264:AA264", "G150+G112"),
        ("G266", 12, "G266:AA266", "G125"),
        ("G274", 13, "G274:AP274", "G170+G119"),
        ("G276", 14, "G276:AP276", "G145"),
        ("F284", 15, "F284:AA284", "F120"),
        ("G415", 16, "G415:AO415", "G330+G299"),
        ("H417", 17, "H417:AO417", "H315"),
        ("G425", 18, "G425:AO425", "G319+G291"),
        ("H427", 19, "H427:AO427", "H304"),
        ("G435", 20, "G435:AO435", "G329+G298"),
        ("H437", 21, "H437:AO437", "H314"),
        ("F445", 22, "F445:Z445", "F299"),
        ("H568", 23, "H568:AO568", "H511+H460"),
        ("G570", 24, "G570:AA570", "G486"),
        ("G573", 25, "G573:AO573", "G500"),
        ("G575", 26, "G575:Z575", "G475"),
        ("G578", 27, "G578:Z578", "G490+G452"),
        ("G580", 28, "G580:Z580", "G465"),
        ("G588", 29, "G588:AO588", "G510+G459"),
        ("G590", 30, "G590:Z590", "G485"),
        ("G598", 31, "G598:Z598", "G460"),
        ("G680", 32, "G680:AO680", "G648+G613"),
        ("H690", 33, "H690:AO690", "H633"),
        ("H695", 34, "H695:AO695", "H637"),
        ("G700", 35, "G700:AO700", "G647+G612"),
        ("G710", 36, "G710:AA710", "G614"),
    ];

    // Filas hidden de reserva (Q4 PENDING): 39 + 57 + 15 = 111.
    private static readonly int[] FilasHidden =
    [
        .. Rango(213, 251),
        .. Rango(356, 412),
        .. Rango(552, 566),
    ];

    private static int[] Rango(int desde, int hasta)
    {
        var lista = new List<int>();
        for (var i = desde; i <= hasta; i++)
        {
            lista.Add(i);
        }

        return [.. lista];
    }

    [Fact]
    public void T0_PlantillaUnica_40HojasMismoOrden()
    {
        using var zip = AbrirPlantilla();
        var hojas = NombresHojas(zip);
        Assert.Equal(40, hojas.Count);
        Assert.Equal("Recaudo EAAB Reciprocidad", hojas[0]);
        Assert.Equal(HojaR1, hojas[9]);
        Assert.Equal("Informe AFaseo Recaudo 202609-2", hojas[16]);
        Assert.Equal("DetRetri2026092", hojas[35]);
        Assert.Equal("DetValiRetri2026092", hojas[36]);
        Assert.Equal("Valida - Control Recaudo", hojas[39]);
    }

    [Fact]
    public void T0_R1_DimensionYReservaHidden()
    {
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, HojaR1);
        Assert.Equal("A1:AP712", hoja.Root?.Element(Ns + "dimension")?.Attribute("ref")?.Value);

        var filas = hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row").ToList() ?? [];
        Assert.Equal(711, filas.Count);

        var hidden = filas
            .Where(r => string.Equals(r.Attribute("hidden")?.Value, "1", StringComparison.Ordinal))
            .Select(r => int.Parse(r.Attribute("r")!.Value))
            .OrderBy(n => n)
            .ToArray();
        Assert.Equal(FilasHidden, hidden);
    }

    [Fact]
    public void T0_BloquesAse_CabecerasYSlotsExistenSinFormula()
    {
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, HojaR1);
        var compartidas = Compartidas(zip);

        foreach (var (fila, rotulo) in Cabeceras)
        {
            Assert.Equal(rotulo, TextoCelda(hoja, compartidas, $"B{fila}"));
        }

        // Techo de capacidad por ASE: nº exacto de filas-dato del contrato (exceso → ERR-SIN-SLOT).
        Assert.Equal(67, ContarSlots(hoja, 6, 72));
        Assert.Equal(101, ContarSlots(hoja, 110, 210));
        Assert.Equal(67, ContarSlots(hoja, 289, 355));
        Assert.Equal(101, ContarSlots(hoja, 450, 550));
        Assert.Equal(74, ContarSlots(hoja, 603, 676));

        // Toda celda F/L de los rangos de datos existe y NO trae <f> (invariante values-only).
        foreach (var (_, inicio, fin) in RangosDatos)
        {
            for (var fila = inicio; fila <= fin; fila++)
            {
                foreach (var celda in new[] { $"F{fila}", $"L{fila}" })
                {
                    var elemento = Celda(hoja, celda);
                    Assert.True(elemento is not null, $"El slot {HojaR1}!{celda} no existe en la plantilla (contrato T0).");
                    Assert.True(elemento!.Element(Ns + "f") is null, $"El slot {HojaR1}!{celda} trae <f> (el contrato exige celda-dato sin fórmula).");
                }
            }
        }
    }

    [Fact]
    public void T0_AFaseoDualFirma_SlotsYCitas()
    {
        // R-T-2: variante A='AFaseo' B='Total' (slots F0 por ASE) + resúmenes D=AFaseo E=Total.
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, HojaR1);
        var compartidas = Compartidas(zip);

        foreach (var (ase, celda) in SlotsAFaseo)
        {
            var fila = celda[1..];
            Assert.Equal("AFaseo", TextoCelda(hoja, compartidas, $"A{fila}"));
            Assert.Equal("Total", TextoCelda(hoja, compartidas, $"B{fila}"));
            var elemento = Celda(hoja, celda);
            Assert.True(elemento?.Element(Ns + "f") is null, $"ASE{ase}: el slot AFaseo {celda} trae <f>.");
            Assert.Equal("0", elemento?.Element(Ns + "v")?.Value);
        }

        // Citas D=AFaseo E=Total con su fórmula observada (verbatum disco; ASE2-4 apuntan a la
        // fila Mes, ASE1/ASE5 a la fila AFaseo — hallazgo congelado, no juicio).
        Assert.Equal("AFaseo", TextoCelda(hoja, compartidas, "D105"));
        Assert.Equal("AFaseo", TextoCelda(hoja, compartidas, "D284"));
        Assert.Equal("AFaseo", TextoCelda(hoja, compartidas, "D445"));
        Assert.Equal("AFaseo", TextoCelda(hoja, compartidas, "D598"));
        Assert.Equal("AFaseo", TextoCelda(hoja, compartidas, "D710"));
        foreach (var (celda, formula) in ResumenesAFaseo)
        {
            Assert.Equal(formula, Celda(hoja, celda)?.Element(Ns + "f")?.Value);
        }
    }

    [Fact]
    public void T0_NuevoEsquemaYLabelsCEstandar()
    {
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, HojaR1);
        var compartidas = Compartidas(zip);

        foreach (var celda in SlotsNuevoEsquema)
        {
            var fila = celda[1..];
            Assert.Equal("NUEVO ESQUEMA", TextoCelda(hoja, compartidas, $"C{fila}"));
            Assert.Equal("Total", TextoCelda(hoja, compartidas, $"D{fila}"));
            Assert.True(Celda(hoja, celda)?.Element(Ns + "f") is null, $"El slot {celda} trae <f>.");
        }

        // Labels-C estándar (E2): rótulos de empresa en columna C tal cual los trae el disco.
        Assert.Equal("ENEL", TextoCelda(hoja, compartidas, "C8"));
        Assert.Equal("OCCIDENTE", TextoCelda(hoja, compartidas, "C15"));
        Assert.Equal("RECIPROCIDAD", TextoCelda(hoja, compartidas, "C80"));
        Assert.Equal("RECIPROCIDAD", TextoCelda(hoja, compartidas, "C259"));
        Assert.Equal("RECIPROCIDAD", TextoCelda(hoja, compartidas, "C573"));
    }

    [Fact]
    public void T0_AllowListL1_CerosExFormulaIntencionales()
    {
        // D-D: F209 es literal 0 (julio: F209=F142); R217/R219 hidden sin celdas (julio: F217=F145,
        // F219=F136). Q3 PENDING: el Ingeniero confirma que son defaults intencionales.
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, HojaR1);

        var f209 = Celda(hoja, "F209");
        Assert.True(f209?.Element(Ns + "f") is null, "F209 trae <f> (el contrato T0 exige literal 0).");
        Assert.Equal("0", f209?.Element(Ns + "v")?.Value);

        foreach (var fila in new[] { "217", "219" })
        {
            // Disco: filas hidden con celdas de SOLO-estilo (s=89/90, sin <v> ni <f>).
            var row = Fila(hoja, fila);
            Assert.True(row is not null, $"La fila hidden {fila} no existe (contrato T0).");
            foreach (var c in row!.Elements(Ns + "c"))
            {
                Assert.True(c.Element(Ns + "f") is null, $"La fila {fila} trae <f> en {c.Attribute("r")?.Value} (contrato T0: cero intencional).");
                var v = c.Element(Ns + "v")?.Value;
                Assert.True(string.IsNullOrEmpty(v), $"La fila {fila} trae <v>='{v}' en {c.Attribute("r")?.Value} (contrato T0: reserva vacía).");
            }
        }

        // Contrapeso julio: la base anterior SÍ traía esas fórmulas (diferencia congelada).
        using var julio = AbrirJulio();
        var hojaJulio = Hoja(julio, HojaR1);
        Assert.Equal("F142", FormulaDe(hojaJulio, "F209"));
        Assert.Equal("F145", FormulaDe(hojaJulio, "F217"));
        Assert.Equal("F136", FormulaDe(hojaJulio, "F219"));
    }

    [Fact]
    public void T0_NoTouchMirror_SharedByteIdentico()
    {
        // D-E: 37 maestras si=0..36 con t/ref/si/texto + 1101 seguidoras. Ni <v> ni <f> por este plan.
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, HojaR1);

        var maestras = new List<(string Celda, uint Si, string Ref, string Texto)>();
        var seguidoras = 0;
        foreach (var row in hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row") ?? [])
        {
            foreach (var c in row.Elements(Ns + "c"))
            {
                var f = c.Element(Ns + "f");
                if (f is null || f.Attribute("t")?.Value != "shared")
                {
                    continue;
                }

                if (f.Attribute("ref")?.Value is { Length: > 0 } rango)
                {
                    maestras.Add((c.Attribute("r")!.Value, uint.Parse(f.Attribute("si")!.Value), rango, f.Value));
                }
                else
                {
                    seguidoras++;
                }
            }
        }

        Assert.Equal(Maestras.Length, maestras.Count);
        Assert.Equal(1101, seguidoras);
        foreach (var esperada in Maestras)
        {
            Assert.Contains(maestras, m =>
                m.Celda == esperada.Celda && m.Si == esperada.Si && m.Ref == esperada.Ref && m.Texto == esperada.Texto);
        }
    }

    [Fact]
    public void T0_Q1_TypoReciprcidad_Pending()
    {
        // Q1 PENDING: ANTICIPOS USUARIOS!E1 = "RECIPRCIDAD" (typo, shared string, sin fórmula).
        // Se congela tal cual; si se corrige es cambio A LA PLANTILLA (re-congela T0).
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, "ANTICIPOS USUARIOS");
        var compartidas = Compartidas(zip);
        Assert.Equal("RECIPRCIDAD", TextoCelda(hoja, compartidas, "E1"));
        Assert.True(Celda(hoja, "E1")?.Element(Ns + "f") is null, "Q1: E1 trae <f> (disco: texto sin fórmula).");
    }

    [Fact]
    public void T0_Q2_PlaceholderD59_Pending()
    {
        // Q2 PENDING: REPORTE RECAUDO x BANCO!D59 = "OJO/// DILIGENCIAR" (sin fórmula).
        // Si es dato de la app entra como slot con su propio mini-T0; si no, D59 no se toca.
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, "REPORTE RECAUDO x BANCO");
        var compartidas = Compartidas(zip);
        Assert.Equal("OJO/// DILIGENCIAR", TextoCelda(hoja, compartidas, "D59"));
        Assert.True(Celda(hoja, "D59")?.Element(Ns + "f") is null, "Q2: D59 trae <f> (disco: texto sin fórmula).");
    }

    [Fact]
    public void T0_Q4_FilasHiddenReserva_Pending()
    {
        // Q4 PENDING: las 111 filas hidden son reserva vacía que la app NUNCA usa (ni desoculta,
        // ni dimensiona, ni escribe). Ninguna trae celdas con <v> ni <f>.
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, HojaR1);
        foreach (var numero in FilasHidden)
        {
            var row = Fila(hoja, numero.ToString());
            Assert.True(row is not null, $"Q4: falta la fila hidden {numero}.");
            Assert.True(string.Equals(row!.Attribute("hidden")?.Value, "1", StringComparison.Ordinal), $"Q4: la fila {numero} perdió hidden=1.");
            foreach (var c in row.Elements(Ns + "c"))
            {
                Assert.True(c.Element(Ns + "f") is null, $"Q4: la fila hidden {numero} trae <f> en {c.Attribute("r")?.Value}.");
                var v = c.Element(Ns + "v")?.Value;
                Assert.True(string.IsNullOrEmpty(v) || v == "0", $"Q4: la fila hidden {numero} trae <v>='{v}' en {c.Attribute("r")?.Value}.");
            }
        }
    }

    [Fact]
    public void T0_BaseCanonica2026092_SellosYFormulaD111()
    {
        // Base período interna 2026092 aunque el filename diga 2026xx-x; sellos y D111=D109+N109.
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, "CONSOLIDADO_TOTAL RECAUDO");
        var compartidas = Compartidas(zip);
        Assert.Equal("2026092", TextoCelda(hoja, compartidas, "N3"));
        Assert.Equal("46281", TextoCelda(hoja, compartidas, "G7"));
        Assert.Equal("46295", TextoCelda(hoja, compartidas, "K7"));
        Assert.Equal("46297", TextoCelda(hoja, compartidas, "D6"));
        Assert.Equal("D109+N109", FormulaDe(hoja, "D111"));
        Assert.Equal("SUM(D104:D108)", FormulaDe(hoja, "D109"));

        var calcChain = zip.GetEntry("xl/calcChain.xml");
        Assert.True(calcChain is not null, "La plantilla debe traer calcChain (doctrina Plan 28).");
    }

    [Fact]
    public void T0_R649_ValorResiduoBase_Congelado()
    {
        // F5: única celda-dato no-cero del R1 (residuo de la base 2026092). El dato fuente manda;
        // el writer puede sobrescribirla con el valor de la fuente (slot normal).
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, HojaR1);
        Assert.Equal("13359789539.23", Celda(hoja, "F649")?.Element(Ns + "v")?.Value);
        Assert.Equal("12154250.720000001", Celda(hoja, "L649")?.Element(Ns + "v")?.Value);
        Assert.True(Celda(hoja, "F649")?.Element(Ns + "f") is null, "F649 trae <f> (disco: valor sin fórmula).");
    }

    // ── Lectura BCL (ZIP + XML; prohibido OpenXML/Excel aquí) ──────────────────────────────

    private static ZipArchive AbrirPlantilla() =>
        ZipFile.OpenRead(Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx"));

    private static ZipArchive AbrirJulio() =>
        ZipFile.OpenRead(Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Plantilla_Remuneracion.xlsx"));

    private static string Raiz()
    {
        // Resolución explícita local (nunca Insumos.Raiz): se camina hacia arriba hasta la
        // plantilla estandarizada única.
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

    private static List<string> NombresHojas(ZipArchive zip)
    {
        var libro = LeerXml(zip, "xl/workbook.xml");
        return libro.Root?.Element(Ns + "sheets")?.Elements(Ns + "sheet")
            .Select(s => s.Attribute("name")!.Value).ToList() ?? [];
    }

    private static XDocument Hoja(ZipArchive zip, string nombre)
    {
        var libro = LeerXml(zip, "xl/workbook.xml");
        var sheet = libro.Root?.Element(Ns + "sheets")?.Elements(Ns + "sheet")
            .FirstOrDefault(s => string.Equals(s.Attribute("name")?.Value, nombre, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"La plantilla no trae la hoja '{nombre}'.");
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

    private static int ContarSlots(XDocument hoja, int inicio, int fin)
    {
        var filas = hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row")
            .Where(r => r.Attribute("r") is { } a && int.TryParse(a.Value, out var n) && n >= inicio && n <= fin)
            .ToList() ?? [];
        return filas.Count;
    }
}
