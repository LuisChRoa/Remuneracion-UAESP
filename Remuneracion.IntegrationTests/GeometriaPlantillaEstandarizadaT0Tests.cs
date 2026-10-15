using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Remuneracion.Core.Constants;
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
    private const string HojaValidacionTotal = "VALIDACION_TOTAL";
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

    // ── Plan 39 (T0, R-T-1/R-T-2 — R-INTERV-Q): variantes L9:L13 por quincena + allow-list Phase 0 ──
    // Congelado desde los bytes del oficial (dump BCL 2026-10-14). El TEXTO-<f> ES el oráculo; los
    // cachés <v> solo DOCUMENTAN la variante (nunca la demuestran). La quincena es DOMINIO
    // (Periodo.NumeroQuincena): Q1 -> N (Primera quincena), Q2 -> M (Segunda quincena). El oficial
    // trae la forma Q2 (base interna 2026092). ASE4 es idéntico en ambas (120391083): el gate NO
    // puede distinguir Q1/Q2 por ASE4 solo — la aserción es por texto-<f>, no por valor.
    private static readonly (string Celda, string FormulaQ2, string FormulaQ1, string ValorQ2, string ValorQ1)[] VariantesL9L13 =
    [
        ("L9", "INTERVENTORIA!M26", "INTERVENTORIA!N26", "189185988", "189185987"),
        ("L10", "INTERVENTORIA!M27", "INTERVENTORIA!N27", "266580256", "266580255"),
        ("L11", "INTERVENTORIA!M28", "INTERVENTORIA!N28", "154788536", "154788535"),
        ("L12", "INTERVENTORIA!M29", "INTERVENTORIA!N29", "120391083", "120391083"),
        ("L13", "INTERVENTORIA!M30", "INTERVENTORIA!N30", "128990446", "128990445"),
    ];

    // Las 11 <f> del período base (token 2026092): BCE H3:H7/H10 + DetValiRetri K24:K28.
    // (Dump BCL 2026-10-14; el parser de celdas respeta el autocierre `<c .../>`.)
    private static readonly (string Hoja, string Celda, string Formula)[] TokensPeriodoFase0 =
    [
        ("BCE SC POR FACT.", "H3", "DetRetri2026092!J9"),
        ("BCE SC POR FACT.", "H4", "DetRetri2026092!J10"),
        ("BCE SC POR FACT.", "H5", "DetRetri2026092!J11"),
        ("BCE SC POR FACT.", "H6", "DetRetri2026092!J12"),
        ("BCE SC POR FACT.", "H7", "DetRetri2026092!J13"),
        ("BCE SC POR FACT.", "H10", "H9=DetRetri2026092!J14"),
        ("DetValiRetri2026092", "K24", "+DetRetri2026092!N32"),
        ("DetValiRetri2026092", "K25", "+DetRetri2026092!N33"),
        ("DetValiRetri2026092", "K26", "+DetRetri2026092!N34"),
        ("DetValiRetri2026092", "K27", "+DetRetri2026092!N35"),
        ("DetValiRetri2026092", "K28", "+DetRetri2026092!N36"),
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

    // ── Plan 39 (T0, R-T-1/R-T-2): R-INTERV-Q + allow-list Phase 0 ──────────────────────────────

    [Fact]
    public void T0_L9L13_VarianteQ2EnElOficial_YValoresInterventoriaPorQuincena()
    {
        // El oficial trae la variante Q2 (INTERVENTORIA!M26..M30) como texto-<f>. La variante Q1
        // (INTERVENTORIA!N26..N30) se congela como la MISMA forma con la columna N; sus cachés son
        // los valores N26:N30 leídos del disco (documentan la variante, no la demuestran).
        using var zip = AbrirPlantilla();
        var consolidado = Hoja(zip, HojaConsolidado);
        foreach (var (celda, formulaQ2, formulaQ1, valorQ2, _) in VariantesL9L13)
        {
            Assert.Equal(formulaQ2, FormulaDe(consolidado, celda));
            Assert.Equal(valorQ2, ValorDe(consolidado, celda));
            Assert.Equal(formulaQ1, formulaQ2.Replace("!M", "!N", StringComparison.Ordinal));
        }

        // Columnas M/N de INTERVENTORIA = oráculo de cachés por variante (R-INTERV-Q §2.2 notas).
        var interventoria = Hoja(zip, HojaInterventoria);
        for (var i = 0; i < 5; i++)
        {
            var aseId = i + 1;
            var fila = 26 + i;
            Assert.Equal(
                InterventoriaDeclarada.SegundaQuincenaPorAse[aseId].ToString(CultureInfo.InvariantCulture),
                ValorDe(interventoria, $"M{fila}"));
            Assert.Equal(
                InterventoriaDeclarada.PrimeraQuincenaPorAse[aseId].ToString(CultureInfo.InvariantCulture),
                ValorDe(interventoria, $"N{fila}"));
        }
    }

    [Fact]
    public void T0_InventarioPeriodo_Y11TokensFase0()
    {
        // Inventario de hojas de período + las 11 <f> del token base (2026092): la superficie exacta
        // que Phase 0 reescribe por dominio (token -> periodo del run).
        using var zip = AbrirPlantilla();
        var hojas = NombresHojas(zip);
        Assert.Equal("Informe AFaseo Recaudo 202609-2", hojas[16]);
        Assert.Equal("DetRetri2026092", hojas[35]);
        Assert.Equal("DetValiRetri2026092", hojas[36]);

        Assert.Equal(11, TokensPeriodoFase0.Length);
        foreach (var (hoja, celda, formula) in TokensPeriodoFase0)
        {
            Assert.Equal(formula, FormulaDe(Hoja(zip, hoja), celda));
        }
    }

    [Fact]
    public void T0_AllowListFase0_16Celdas()
    {
        // R-T-2: Phase 0 SOLO puede tocar estas <f>: 11 tokens de período + 5 L9:L13. Cualquier otra
        // <f> tocada por cualquier fase = FAIL (grep de cierre por namespace/clase). La reescritura de
        // tokens también toca partes NO-<f> (workbook.xml/app.xml/sharedStrings/definedNames): eso es
        // la copia mecánica de período, fuera del alcance de esta allow-list de celdas-fórmula.
        using var zip = AbrirPlantilla();
        var allow = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (hoja, celda, _) in TokensPeriodoFase0)
        {
            allow.Add($"{hoja}!{celda}");
        }

        foreach (var (celda, _, _, _, _) in VariantesL9L13)
        {
            allow.Add($"{HojaConsolidado}!{celda}");
        }

        Assert.Equal(16, allow.Count);

        foreach (var (hoja, celda, _) in TokensPeriodoFase0)
        {
            Assert.True(Celda(Hoja(zip, hoja), celda)?.Element(Ns + "f") is not null, $"{hoja}!{celda} debe ser <f> en el oficial.");
        }

        var consolidado = Hoja(zip, HojaConsolidado);
        foreach (var (celda, _, _, _, _) in VariantesL9L13)
        {
            Assert.True(Celda(consolidado, celda)?.Element(Ns + "f") is not null, $"{HojaConsolidado}!{celda} debe ser <f> en el oficial.");
        }

        // El oficial SOLO trae 11 <f> con token de período: la allow-list de tokens es EXACTA.
        var conToken = new List<string>();
        foreach (var hoja in NombresHojas(zip))
        {
            foreach (var c in Hoja(zip, hoja).Descendants(Ns + "c"))
            {
                var texto = c.Element(Ns + "f")?.Value;
                if (texto is not null && Regex.IsMatch(texto, @"Det(Retri|ValiRetri)\d{7}"))
                {
                    conToken.Add($"{hoja}!{c.Attribute("r")?.Value}");
                }
            }
        }

        Assert.Equal(11, conToken.Count);
        Assert.All(conToken, celda => Assert.Contains(celda, allow));
    }

    // ── Plan 40 (T0, R-T-1): W-1 evalúa + sella — 5 sub-bloques + cadena C15 + diferencial de clase ──
    // Congelado desde bytes (BCL) del oficial + la salida real de hoy (Resultado5). Cero código de
    // producción en T0: estos asserts SON el contrato del gate evaluador (T1) y del sellado (T2).

    private static readonly (string Celda, string Formula)[] Plan40SubBloques =
    [
        ("C15", "C14=C9"),
        ("D25", "D24=C9"),
        ("O25", "O24=N9"),
        ("D34", "D33='CONSOLIDADO_TOTAL RECAUDO'!D109"),
        ("F34", "F33='CONSOLIDADO_TOTAL RECAUDO'!D109+'CONSOLIDADO_TOTAL RECAUDO'!N109"),
    ];

    private static readonly (string Hoja, string Celda, string Formula)[] Plan40CadenasC15 =
    [
        (HojaValidacionTotal, "C14", "'Valida - Control Recaudo'!F10"),
        (HojaValidacionTotal, "C9", "SUM(C3:C8)"),
        (HojaValidacionTotal, "C3", "'Recaudo EAAB Reciprocidad'!F21+'Recaudo EAAB + Ciud Limp'!F21+'Recaudo Directa Occidente'!F21+'Recaudo ENEL'!F21+'Recaudo ENERBIT'!F21"),
        (HojaValidacionTotal, "C8", "'Recaudo EAAB Reciprocidad'!F26+'Recaudo EAAB + Ciud Limp'!F26+'Recaudo Directa Occidente'!F26+'Recaudo ENEL'!F26+'Recaudo ENERBIT'!F26"),
        (HojaValidacionTotal, "D24", "SUM(D20:D23)"),
        (HojaValidacionTotal, "N9", "SUM(N3:N8)"),
        (HojaValidacionTotal, "D33", "SUM(D29:D32)"),
        (HojaValidacionTotal, "F33", "SUM(F29:F32)"),
        (HojaConsolidado, "D109", "SUM(D104:D108)"),
    ];

    [Fact]
    public void T0_Plan40_SubBloquesValidacionTotal_TextosYCadenas()
    {
        using var zip = AbrirPlantilla();
        var hoja = Hoja(zip, HojaValidacionTotal);

        // Los 5 sub-bloques con su <f> EXACTA; el fixture ES la fuente de FormasValidacionTotal.
        foreach (var (celda, formula) in Plan40SubBloques)
        {
            Assert.Equal(formula, FormulaDe(hoja, celda));
        }

        // Consistencia fixture <-> tabla del gate (una deriva del código = FAIL accionable).
        foreach (var (celda, textoEsperado) in FormasValidacionTotal.SubBloques)
        {
            Assert.Equal(textoEsperado, FormulaDe(hoja, celda));
        }

        // Cadenas de dependencia (recaudo + consolidado) que el resolutor del gate recorre.
        foreach (var (nombreHoja, celda, formula) in Plan40CadenasC15)
        {
            Assert.Equal(formula, FormulaDe(Hoja(zip, nombreHoja), celda));
        }

        // O24 (shared si=5, maestro H24:O24) y N109 (shared si=61, maestro D109:N109): el resolutor
        // los cubre por desplazamiento del maestro (no tienen texto propio).
        Assert.Equal("SUM(H20:H23)", FormulaDe(hoja, "H24"));
        Assert.Equal("SUM(D104:D108)", FormulaDe(Hoja(zip, HojaConsolidado), "D109"));
    }

    [Fact]
    public void T0_Plan40_EstadoF10C9C14C15_OficialYHoy()
    {
        // Oficial: F10 = literal residual 73981716599; C14 cache = 73981716599; C9 cache = 0;
        // C15 cache = 0 (FALSE heredado) — la cadena miente porque el caché nunca se recalculó.
        using (var oficial = AbrirPlantilla())
        {
            var vt = Hoja(oficial, HojaValidacionTotal);
            var control = Hoja(oficial, "Valida - Control Recaudo");
            Assert.Null(FormulaDe(control, "F10"));
            Assert.Equal("73981716599", ValorDe(control, "F10"));
            Assert.Equal("'Valida - Control Recaudo'!F10", FormulaDe(vt, "C14"));
            Assert.Equal("73981716599", ValorDe(vt, "C14"));
            Assert.Equal("SUM(C3:C8)", FormulaDe(vt, "C9"));
            Assert.Equal("0", ValorDe(vt, "C9"));
            Assert.Equal("C14=C9", FormulaDe(vt, "C15"));
            Assert.Equal("0", ValorDe(vt, "C15"));
        }

        // Hoy (Resultado5): la app ESCRIBIÓ F10 = 72497949967 (== Σ 'Recaudo *'!F21:F26), pero
        // C14/C9/C15 conservan los cachés heredados (73981716599 / 0 / 0 = FALSE).
        var rutaHoy = Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Resultado", "Resultado5", "Remuneración 202607-2 Total.xlsx");
        using (var hoy = ZipFile.OpenRead(rutaHoy))
        {
            var vt = Hoja(hoy, HojaValidacionTotal);
            var control = Hoja(hoy, "Valida - Control Recaudo");
            Assert.Equal("72497949967", ValorDe(control, "F10"));
            Assert.Equal("73981716599", ValorDe(vt, "C14")); // caché viejo
            Assert.Equal("0", ValorDe(vt, "C9"));            // caché viejo
            Assert.Equal("0", ValorDe(vt, "C15"));           // caché viejo (FALSE)
        }
    }

    [Fact]
    public void T0_Plan40_DiferencialClase_OficialVsHoy()
    {
        // Diferencial de TODAS las celdas de fórmula entre oficial y salida-hoy, alineando por
        // ÍNDICE de hoja (el período renombra 3 hojas). Se separa: (a) <f> distinta (reescritura
        // del token de período por Phase 0) y (b) <f> idéntica con <v> distinta (resello de N3).
        // La clase "caché stale" del incidente NO se ve por este diff: la app nunca recalcula, así
        // que C14/C9/C15 no cambian (0->0) — por eso el gate EVALÚA (no lee caché).
        using var oficial = AbrirPlantilla();
        var rutaHoy = Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Resultado", "Resultado5", "Remuneración 202607-2 Total.xlsx");
        using var hoy = ZipFile.OpenRead(rutaHoy);

        var nombresOficial = NombresHojas(oficial);
        var nombresHoy = NombresHojas(hoy);
        Assert.Equal(nombresOficial.Count, nombresHoy.Count);

        var formulaDiff = 0;
        var cacheDiff = 0;
        for (var i = 0; i < nombresOficial.Count; i++)
        {
            var celdasOficial = CeldasFormula(Hoja(oficial, nombresOficial[i]));
            var celdasHoy = CeldasFormula(Hoja(hoy, nombresHoy[i]));
            foreach (var referencia in celdasOficial.Keys.Union(celdasHoy.Keys))
            {
                celdasOficial.TryGetValue(referencia, out var a);
                celdasHoy.TryGetValue(referencia, out var b);
                if (a.Formula != b.Formula)
                {
                    formulaDiff++;
                }
                else if (a.Valor != b.Valor)
                {
                    cacheDiff++;
                }
            }
        }

        Assert.Equal(11, formulaDiff); // 11 <f> del token de período (BCE H3:H7/H10 + DetValiRetri K24:K28)
        Assert.Equal(103, cacheDiff);  // consumidores de N3 resellados por Phase 0 (REMUNERACION_*/GERENTES_*/AJUSTES/CONSOLIDADO/Valida -Remunera)

        // C15 NO cambia entre oficial y hoy: su caché stale no se detecta por diff (doctrina D-A).
        Assert.Equal(ValorDe(Hoja(oficial, HojaValidacionTotal), "C15"), ValorDe(Hoja(hoy, HojaValidacionTotal), "C15"));
    }

    [Fact]
    public void T0_Plan40_AllowListFase0_YPropuestaFNew1()
    {
        // Allow-list Phase 0 vigente = 16 celdas <f> (11 tokens de período + 5 L9:L13), ya
        // congelada por el T0 del Plan 39 (T0_AllowListFase0_16Celdas). Aquí se congela la
        // PROPUESTA de ampliación para F-NEW-1 (T3, fuera de esta corrida): las 4 familias de
        // auto-chequeo que hardcodean INTERVENTORIA!M26:M31 y quedarían FALSE en Q1.
        using var zip = AbrirPlantilla();

        var detRetri = Hoja(zip, "DetRetri2026092");
        for (var fila = 32; fila <= 36; fila++)
        {
            Assert.Contains("INTERVENTORIA!M", FormulaDe(detRetri, $"L{fila}"), StringComparison.Ordinal);
            Assert.Contains("INTERVENTORIA!M", FormulaDe(detRetri, $"O{fila}"), StringComparison.Ordinal);
        }

        var detVali = Hoja(zip, "DetValiRetri2026092");
        for (var fila = 24; fila <= 29; fila++)
        {
            Assert.Contains("INTERVENTORIA!M", FormulaDe(detVali, $"I{fila}"), StringComparison.Ordinal);
        }

        var interventoria = Hoja(zip, "INTERVENTORIA");
        // S26 es el maestro del bloque shared S26:S31 (Q26=R26); S27..S31 son followers sin texto.
        Assert.Equal("Q26=R26", FormulaDe(interventoria, "S26")?.Replace("$", string.Empty, StringComparison.Ordinal));
    }

    private static Dictionary<string, (string Formula, string Valor)> CeldasFormula(XDocument hoja)
    {
        var mapa = new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row") ?? [])
        {
            foreach (var c in row.Elements(Ns + "c"))
            {
                var f = c.Element(Ns + "f");
                if (f is null)
                {
                    continue;
                }

                var referencia = c.Attribute("r")?.Value ?? string.Empty;
                mapa[referencia] = (f.Value, c.Element(Ns + "v")?.Value ?? string.Empty);
            }
        }

        return mapa;
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

    private static string ValorDe(XDocument hoja, string referencia) =>
        Celda(hoja, referencia)?.Element(Ns + "v")?.Value ?? string.Empty;

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
