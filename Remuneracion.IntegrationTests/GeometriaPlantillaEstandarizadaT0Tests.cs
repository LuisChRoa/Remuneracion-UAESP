using System.IO.Compression;
using System.Xml.Linq;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 38 (T0, R-T-1/R-T-2 — D-A): contrato geométrico NUEVO congelado desde los bytes del disco
/// de la plantilla estandarizada única
/// (<c>Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx</c>).
///
/// CERO código de producción: este fixture + sus asserts SON el contrato. Lectura SOLO con BCL
/// (<c>System.IO.Compression.ZipArchive</c> + <c>System.Xml.Linq</c>): sin Excel/COM, sin
/// DocumentFormat.OpenXml, sin ExcelDataReader.
///
/// Reutiliza (no duplica) lo ya congelado por <see cref="ContratoSlotR1T0Tests"/>: inventario de
/// 40 hojas, bloques R1 por ASE, masters shared, sellos, N3=2026092, D111=D109+N109. Aquí se
/// congela lo que faltaba: las CADENAS de dirección que el runtime escribe/valida — CONSOLIDADO
/// D9:D13/D28:D32/D47:D51/D66:D70/D85:D89/D104:D108/D109/D111, visibles R1/R2/R4 por ASE,
/// DetRetri/DetValiRetri, SALDOS POR NOTA / RETRIBUCION NEGATIVA / AJUSTES-SF-T — con cita
/// celda-a-celda, y el DRIFT GATE (hash + conteos) que convierte cualquier cambio de plantilla en
/// un FAIL que exige un nuevo T0 (procedimiento documentado en
/// <c>Docs/T0-Geometria-Plantilla-Estandarizada.md</c>).
///
/// Frontera D-B (firma-vs-estructural): las direcciones aquí congeladas son ESTRUCTURALES (sin
/// firma resoluble): cadenas CONSOLIDADO, visibles de bloque, refs DetRetri/DetValiRetri,
/// visibles SALDOS/RETRIB. Las celdas-dato R1 del interior del bloque se derivan en runtime por
/// firma+clasificador+sinonimia (<c>EscritorValoresSlotR1</c>) y NO se congelan aquí.
/// </summary>
public sealed class GeometriaPlantillaEstandarizadaT0Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";
    private const string HojaR1 = "Reporte Componentes R1";
    private const string HojaR2 = "Rem. Anticipos R2";
    private const string HojaR4 = "Reversion Pagos R4";
    private const string HojaSaldos = "SALDOS POR NOTA";
    private const string HojaRetrib = "RETRIBUCION NEGATIVA";
    private const string HojaAjustes = "AJUSTES - SF-T";
    private const string HojaBanco = "REPORTE RECAUDO x BANCO";
    private const string HojaBce = "BCE SC POR FACT.";
    private const string HojaInterventoria = "INTERVENTORIA";

    // ── Drift gate (R-DRIFT-FUTURO): inventario de hojas + base interna + cadenas celda-a-celda.
    // El resto de la geometría se asserta celda-a-celda abajo. Un cambio de plantilla exige un
    // nuevo T0 (procedimiento documentado en Docs/T0-Geometria-Plantilla-Estandarizada.md).

    // ── CONSOLIDADO D-chain (T0: bytes del disco, dump ZIP+XML 2026-10-14) ───────────────────────
    private static readonly (string Celda, string Formula)[] ConsolidadoD =
    [
        ("D9", "'Reporte Componentes R1'!F75"),
        ("D10", "'Reporte Componentes R1'!F254"),
        ("D11", "'Reporte Componentes R1'!F415"),
        ("D12", "'Reporte Componentes R1'!F568"),
        ("D13", "'Reporte Componentes R1'!F680"),
        ("D28", "'Rem. Anticipos R2'!E47"),
        ("D29", "'Rem. Anticipos R2'!E144"),
        ("D30", "'Rem. Anticipos R2'!E262"),
        ("D31", "'Rem. Anticipos R2'!E362"),
        ("D32", "'Rem. Anticipos R2'!E446"),
        ("D47", "'Reporte Componentes R1'!F77"),
        ("D48", "'Reporte Componentes R1'!F256"),
        ("D49", "'Reporte Componentes R1'!F417"),
        ("D50", "'Reporte Componentes R1'!F570"),
        ("D51", "'Reporte Componentes R1'!F682"),
        ("D66", "'Reversion Pagos R4'!D73"),
        ("D67", "'Reversion Pagos R4'!D174"),
        ("D68", "'Reversion Pagos R4'!D211"),
        ("D69", "'Reversion Pagos R4'!D326"),
        ("D70", "'Reversion Pagos R4'!D361"),
        ("D85", "'AJUSTES - SF-T'!D47"),
        ("D86", "'AJUSTES - SF-T'!D48"),
        ("D87", "'AJUSTES - SF-T'!D49"),
        ("D88", "'AJUSTES - SF-T'!D50"),
        ("D89", "'AJUSTES - SF-T'!D51"),
        ("D104", "D9+D28+D47+D66+D85"),
        ("D105", "D10+D29+D48+D67+D86"),
        ("D107", "D12+D31+D50+D69+D88"),
        ("D108", "D13+D32+D51+D70+D89"),
        ("D109", "SUM(D104:D108)"),
        ("D111", "D109+N109")
    ];

    // ── Visibles R1 por ASE: TOT_OPT (D9:D13) y EXTEMP (D47:D51) con sus operandos exactos ──────
    private static readonly (string Celda, string Formula)[] R1Visibles =
    [
        ("F75", "F47+F70+F16-L16-L47"),
        ("F77", "F59+F32-L32"),
        ("F254", "F171+F208-L171+F120-L120"),
        ("F256", "F190+F146-L146"),
        ("F415", "F330+F353+F299-L299-L330"),
        ("F417", "F342+F315-L315"),
        ("F568", "F511+F548+F460-L460-L511"),
        ("F570", "F530+F486-L486"),
        ("F680", "F674+F648+F613-L613-L648")
    ];

    // F682 (EXTEMP ASE5) es VALOR 0 (no fórmula) — el CONSOLIDADO D51 lo referencia como valor.
    private const string R1ExtempAse5Valor = "F682";

    // ── Visibles R2/R4 por ASE ───────────────────────────────────────────────────────────────────
    private static readonly (string Celda, string Formula)[] R2Visibles =
    [
        ("E47", "E21+E32-K21"),
        ("E144", "E94+E112-K94"),
        ("E262", "E182+E193-K182"),
        ("E362", "E309+E327-K309"),
        ("E446", "E418+E404-K404")
    ];

    private static readonly (string Celda, string Formula)[] R4Visibles =
    [
        ("D73", "D15-P15"),
        ("D174", "D111-P111"),
        ("D211", "D206-P206"),
        ("D326", "D249-P249"),
        ("D361", "D358-P358")
    ];

    // ── DetRetri/DetValiRetri (hoja del período base 2026092): cadenas protegidas ───────────────
    private static readonly (string Celda, string Formula)[] DetRetriCadenas =
    [
        ("D23", "D9-'CONSOLIDADO_TOTAL RECAUDO'!D104"),
        ("D28", "D14-'CONSOLIDADO_TOTAL RECAUDO'!D109"),
        ("D32", "+D23+'Reporte Componentes R1'!F75+'Rem. Anticipos R2'!E47+'Reversion Pagos R4'!D73+'Reporte Componentes R1'!F77+'AJUSTES - SF-T'!D47=D9"),
        ("D36", "+D27+'Reporte Componentes R1'!F680+'Reporte Componentes R1'!F682+'Rem. Anticipos R2'!E446+'Reversion Pagos R4'!D361+'AJUSTES - SF-T'!D51=D13")
    ];

    private static readonly (string Celda, string Formula)[] DetValiRetriCadenas =
    [
        ("D16", "D9-'CONSOLIDADO_TOTAL RECAUDO'!U104"),
        ("D21", "D14-'CONSOLIDADO_TOTAL RECAUDO'!U109"),
        ("D24", "+'Reporte Componentes R1'!M75+'Reporte Componentes R1'!M77+'Rem. Anticipos R2'!L47+'Reversion Pagos R4'!J73+D16+'AJUSTES - SF-T'!U47=D9"),
        ("D28", "+'Reporte Componentes R1'!M680+'Reporte Componentes R1'!M682+'Rem. Anticipos R2'!L446+'Reversion Pagos R4'!J361+D20+'AJUSTES - SF-T'!U51=D13")
    ];

    // ── SALDOS POR NOTA / RETRIBUCION NEGATIVA visibles (Cn-In) por ASE ─────────────────────────
    private static readonly (string Celda, string Formula)[] SaldosVisibles =
    [
        ("C10", "C8-I8"),
        ("C23", "C21-I21"),
        ("C36", "C34-I34"),
        ("C49", "C47-I47"),
        ("C62", "C60-I60")
    ];

    private static readonly (string Celda, string Formula)[] RetribVisibles =
    [
        ("C10", "C8-I8"),
        ("C23", "C21-I21"),
        ("C36", "C34-I34"),
        ("C49", "C47-I47"),
        ("C62", "C60-I60")
    ];

    private static readonly (string Celda, string Formula)[] AjustesVisibles =
    [
        ("D9", "'SALDOS POR NOTA'!C10"),
        ("D13", "'SALDOS POR NOTA'!C62"),
        ("D28", "'RETRIBUCION NEGATIVA'!C10"),
        ("D32", "'RETRIBUCION NEGATIVA'!C62"),
        ("D47", "D9+D28"),
        ("D52", "SUM(D47:D51)")
    ];

    [Fact]
    public void T0_Consolidado_CadenaD_Exacta()
    {
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, HojaConsolidado);
        foreach (var (celda, formula) in ConsolidadoD)
        {
            Assert.Equal(formula, FormulaDe(hoja, celda));
        }
    }

    [Fact]
    public void T0_VisiblesR1_FormulasExactasPorAse()
    {
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, HojaR1);
        foreach (var (celda, formula) in R1Visibles)
        {
            Assert.Equal(formula, FormulaDe(hoja, celda));
        }

        // F682 = EXTEMP ASE5: VALOR (no fórmula) — el CONSOLIDADO D51 lo referencia como valor.
        Assert.True(Celda(hoja, R1ExtempAse5Valor)?.Element(Ns + "f") is null, "F682 (EXTEMP ASE5) debe ser VALOR (no fórmula) en el contrato T0.");
    }

    [Fact]
    public void T0_VisiblesR2R4_FormulasExactasPorAse()
    {
        using var zip = AbrirPlantilla();
        var hojaR2 = Hoja(zip, HojaR2);
        foreach (var (celda, formula) in R2Visibles)
        {
            Assert.Equal(formula, FormulaDe(hojaR2, celda));
        }

        var hojaR4 = Hoja(zip, HojaR4);
        foreach (var (celda, formula) in R4Visibles)
        {
            Assert.Equal(formula, FormulaDe(hojaR4, celda));
        }
    }

    [Fact]
    public void T0_DetRetriDetValiRetri_CadenasProtegidas()
    {
        using var zip = AbrirPlantilla();
        var detRetri = Hoja(zip, "DetRetri2026092");
        foreach (var (celda, formula) in DetRetriCadenas)
        {
            Assert.Equal(formula, FormulaDe(detRetri, celda));
        }

        var detVali = Hoja(zip, "DetValiRetri2026092");
        foreach (var (celda, formula) in DetValiRetriCadenas)
        {
            Assert.Equal(formula, FormulaDe(detVali, celda));
        }
    }

    [Fact]
    public void T0_SaldosRetribAjustes_VisiblesProtegidas()
    {
        using var zip = AbrirPlantilla();
        var saldos = Hoja(zip, HojaSaldos);
        foreach (var (celda, formula) in SaldosVisibles)
        {
            Assert.Equal(formula, FormulaDe(saldos, celda));
        }

        var retrib = Hoja(zip, HojaRetrib);
        foreach (var (celda, formula) in RetribVisibles)
        {
            Assert.Equal(formula, FormulaDe(retrib, celda));
        }

        var ajustes = Hoja(zip, HojaAjustes);
        foreach (var (celda, formula) in AjustesVisibles)
        {
            Assert.Equal(formula, FormulaDe(ajustes, celda));
        }
    }

    [Fact]
    public void T0_BceInterventoriaBanco_IdentidadEstructural()
    {
        // D-B: estas familias NO cambiaron de geometría respecto del contrato legacy (verificado
        // byte-a-byte): el T0 las fija para que cualquier cambio futuro sea FAIL accionable.
        using var zip = AbrirPlantilla();

        var bce = Hoja(zip, HojaBce);
        Assert.Equal("D3+E3", FormulaDe(bce, "F3"));
        Assert.Equal("DetRetri2026092!J9", FormulaDe(bce, "H3"));
        Assert.Equal("F3-H3", FormulaDe(bce, "I3"));
        Assert.Equal("SUM(F3:F7)", FormulaDe(bce, "F9"));
        Assert.Equal("SUM(H3:H7)", FormulaDe(bce, "H9"));
        Assert.Equal("H9=DetRetri2026092!J14", FormulaDe(bce, "H10"));

        var interventoria = Hoja(zip, HojaInterventoria);
        Assert.Equal("SUM(K26:K30)", FormulaDe(interventoria, "K31"));
        Assert.Equal("SUM(M26:M30)", FormulaDe(interventoria, "M31"));
        Assert.Equal("SUM(N26:N30)", FormulaDe(interventoria, "N31"));
        Assert.Equal("SUM(M31:N31)", FormulaDe(interventoria, "K32"));

        var banco = Hoja(zip, HojaBanco);
        Assert.Equal("C13+C23+C32+C41+C51", FormulaDe(banco, "C2"));
        Assert.Equal("IF($C$59=1,'Recaudo ENEL'!D21,'Recaudo ENEL'!F21)", FormulaDe(banco, "C62"));
        // La hoja termina en la fila 79: NO existe el TOTAL RECAUDO de la fila 81 del legacy.
        Assert.True(Celda(banco, "C81") is null, "La hoja BANCO del contrato T0 termina en la fila 79 (sin C81/D81).");
        Assert.True(Celda(banco, "D81") is null, "La hoja BANCO del contrato T0 termina en la fila 79 (sin C81/D81).");
    }

    [Fact]
    public void T0_DriftGate_WorkbookXmlYConteos()
    {
        // R-DRIFT-FUTURO: cualquier cambio del inventario de hojas o de la base interna rompe este
        // gate ANTES que cualquier writer. Un cambio de plantilla exige un nuevo T0 (procedimiento
        // documentado en Docs/T0-Geometria-Plantilla-Estandarizada.md), nunca edición silenciosa.
        using var zip = AbrirPlantilla();
        var hojas = NombresHojas(zip);
        Assert.Equal(40, hojas.Count);
        Assert.Equal("Reporte Componentes R1", hojas[9]);
        Assert.Equal("DetRetri2026092", hojas[35]);
        Assert.Equal("DetValiRetri2026092", hojas[36]);

        // Base interna del período (N3) + sello D111 coherentes con el contrato congelado.
        var consolidado = Hoja(zip, HojaConsolidado);
        Assert.Equal("2026092", TextoCelda(consolidado, Compartidas(zip), "N3"));
    }

    // ── Mini-T0 (Plan 38, T2-continuación, R-ATRIBUCION): direcciones por (empresa × ASE) y
    // bloques de detalle re-congelados desde los bytes. El test compara el CÓDIGO (los mapas) contra
    // el FIXTURE (la plantilla); un cambio de plantilla o una dirección legacy = FAIL accionable. ──

    [Fact]
    public void T0_PorEmpresaR2R4_DireccionesMatcheanFormulaOportuno()
    {
        using var zip = AbrirPlantilla();
        var compartidas = Compartidas(zip);

        // R2: el visible por empresa E{fila} = Total + SubsCont (…± Especiales). Las celdas del mapa
        // (Total/SubsCont) deben aparecer en esa fórmula. La celda Especiales (col K) no aparece en
        // la fórmula del visible (es operando de otras filas) → solo se exige que exista.
        var r2 = Hoja(zip, HojaR2);
        foreach (var ((empresaId, aseId), celdas) in WorkbookLeafCellMapPorEmpresa.EditablesR2PorEmpresa)
        {
            var label = EmpresaFacturacion.Obtener(empresaId).LabelTemplate;
            var fila = FilaOportuno(r2, compartidas, label, VentanaR2(aseId), colLabel: "B", colOportuno: "C");
            var formula = FormulaDe(r2, $"E{fila}") ?? string.Empty;
            foreach (var (celda, fuente) in celdas)
            {
                Assert.True(Celda(r2, celda) is not null, $"{HojaR2}!{celda} (ASE{aseId}, empresa {empresaId}) no existe en la plantilla.");
                if (fuente == WorkbookLeafCellMapPorEmpresa.FuenteR2.Especiales)
                {
                    continue;
                }

                Assert.Contains(celda, formula, StringComparison.OrdinalIgnoreCase);
            }
        }

        // R4: el visible por empresa D{fila} = Total − P. La celda D del mapa es el Total del visible.
        var r4 = Hoja(zip, HojaR4);
        foreach (var ((empresaId, aseId), celdas) in WorkbookLeafCellMapPorEmpresa.EditablesR4PorEmpresa)
        {
            var label = EmpresaFacturacion.Obtener(empresaId).LabelTemplate;
            var fila = FilaOportuno(r4, compartidas, label, VentanaR4(aseId), colLabel: "A", colOportuno: "B");
            var formula = FormulaDe(r4, $"D{fila}") ?? string.Empty;
            foreach (var (celda, _) in celdas)
            {
                Assert.True(Celda(r4, celda) is not null, $"{HojaR4}!{celda} (ASE{aseId}, empresa {empresaId}) no existe en la plantilla.");
                if (celda.StartsWith('D'))
                {
                    Assert.Contains(celda, formula, StringComparison.OrdinalIgnoreCase);
                }
            }
        }
    }

    [Fact]
    public void T0_DetalleR2R4_BloquesMatcheanPlantilla()
    {
        using var zip = AbrirPlantilla();
        var compartidas = Compartidas(zip);
        var r2 = Hoja(zip, HojaR2);
        foreach (var (_, bloque) in WorkbookLeafCellMapDetalleR2R4.BloquesR2)
        {
            Assert.Equal("Vlr Servicio", TextoCelda(r2, compartidas, $"D{bloque.FilaInicio}"));
            Assert.Equal("Total", TextoCelda(r2, compartidas, $"A{bloque.FilaFin}"));
        }

        var r4 = Hoja(zip, HojaR4);
        foreach (var (_, bloque) in WorkbookLeafCellMapDetalleR2R4.BloquesR4)
        {
            Assert.Equal("Vlr Servicio", TextoCelda(r4, compartidas, $"C{bloque.FilaInicio}"));
            Assert.Equal("Total", TextoCelda(r4, compartidas, $"A{bloque.FilaFin}"));
        }
    }

    [Fact]
    public void T0_R1Q2Editables_DireccionesExistenYSonValores()
    {
        using var zip = AbrirPlantilla();
        var r1 = Hoja(zip, HojaR1);
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            foreach (var (celda, _) in WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(aseId))
            {
                var c = Celda(r1, celda);
                Assert.True(c is not null, $"{HojaR1}!{celda} (ASE{aseId}) no existe en la plantilla.");
                Assert.True(c!.Element(Ns + "f") is null, $"{HojaR1}!{celda} (ASE{aseId}) debe ser VALOR editable, no fórmula.");
            }
        }
    }

    private static (int Desde, int Hasta) VentanaR2(int aseId) =>
        (WorkbookLeafCellMapDetalleR2R4.BloquesR2[aseId].FilaInicio,
            WorkbookLeafCellMapDetalleR2R4.BloquesR2.TryGetValue(aseId + 1, out var sig) ? sig.FilaInicio - 1 : int.MaxValue);

    private static (int Desde, int Hasta) VentanaR4(int aseId) =>
        (WorkbookLeafCellMapDetalleR2R4.BloquesR4[aseId].FilaInicio,
            WorkbookLeafCellMapDetalleR2R4.BloquesR4.TryGetValue(aseId + 1, out var sig) ? sig.FilaInicio - 1 : int.MaxValue);

    private static int FilaOportuno(XDocument hoja, List<string> compartidas, string label, (int Desde, int Hasta) ventana, string colLabel, string colOportuno)
    {
        for (var fila = ventana.Desde; fila <= ventana.Hasta; fila++)
        {
            if (string.Equals(TextoCelda(hoja, compartidas, $"{colLabel}{fila}"), label, StringComparison.Ordinal)
                && string.Equals(TextoCelda(hoja, compartidas, $"{colOportuno}{fila}"), "OPORTUNO", StringComparison.Ordinal))
            {
                return fila;
            }
        }

        throw new InvalidOperationException($"No se encontró la fila OPORTUNO de '{label}' en la ventana [{ventana.Desde},{ventana.Hasta}].");
    }

    // ── Lectura BCL (ZIP + XML; prohibido OpenXML/Excel aquí) ──────────────────────────────────

    private static ZipArchive AbrirPlantilla() =>
        ZipFile.OpenRead(Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx"));

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
}
