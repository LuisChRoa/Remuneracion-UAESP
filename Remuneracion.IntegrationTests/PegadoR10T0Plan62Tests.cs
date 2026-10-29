using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 62 (H0 — T0.1..T0.4): verificación read-only de los 12 hechos del §T0 que habilitan
/// (a) el pegado del bloque R10 en <c>DetRetri{AAAAMMQ}</c> / <c>DetValiRetri{AAAAMMQ}</c> y
/// (b) la extensión E1-4 al bloque <c>CONSOLIDADO_TOTAL RECAUDO!E142:E146</c>.
///
/// <para>
/// <b>Método (D-H0):</b> BCL puro (<see cref="ZipArchive"/> + <see cref="XDocument"/>) sobre rutas
/// EXPLÍCITAS. Sin Excel/COM, sin DocumentFormat.OpenXml, sin ExcelDataReader y sin
/// <c>Insumos.Raiz()</c>. Los .xlsx se abren SOLO lectura (la plantilla puede estar abierta en
/// Excel: <see cref="FileShare.ReadWrite"/> no se necesita porque se abre por stream de zip).
/// Cero código de producción: este fixture y <c>Docs/T0-Plan62-Pegado-R10.md</c> SON el entregable.
/// </para>
///
/// <para>
/// <b>Fuentes medidas:</b>
/// <list type="bullet">
/// <item><c>Docs/Prueba Septiembre 1/Insumos/R10_Remuneracion_2026091.xlsx</c> (oráculo del caso).</item>
/// <item><c>Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx</c> (oficial; base
/// interna 2026092 → <c>sheet36</c> DetRetri, <c>sheet37</c> DetValiRetri, <c>sheet18</c> CONSOLIDADO).</item>
/// <item><c>Docs/Prueba Septiembre 1/Resultados/Administrativo/Remuneracion 202609-1 Total.xlsx</c>
/// (manual del administrativo, base de paridad del pegado).</item>
/// </list>
/// </para>
///
/// <para>
/// <b>Regla honesta:</b> el assert refleja el BYTE. Donde el texto del plan y el archivo difieren
/// (detalle de columna del encabezado, literal de la fila 147, dígito de un monto citado) el fixture
/// congela lo medido y la desviación se reporta en <c>Docs/T0-Plan62-Pegado-R10.md</c> §6 — nunca se
/// "ajusta" el plan en silencio (§T0, puerta H0→H1).
/// </para>
/// </summary>
public sealed class PegadoR10T0Plan62Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly Regex MontoMonetarioRegex = new(@"^-?\$[0-9.]+$", RegexOptions.Compiled);

    private const string PeriodoR10 = "2026091";

    /// <summary>Base interna de la plantilla oficial (N3 = 2026092): las hojas Det son de este período.</summary>
    private const string PeriodoBasePlantilla = "2026092";

    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";
    private const string PartWorkbook = "xl/workbook.xml";
    private const string PartRels = "xl/_rels/workbook.xml.rels";
    private const string PartCompartidas = "xl/sharedStrings.xml";

    private static readonly string[] Columnas = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"];

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T0a — R10 2026091: exactamente 2 hojas con el MISMO layout que la plantilla (T0-01…T0-04)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T0-01 (hojas): el R10 trae EXACTAMENTE 2 hojas con los nombres que compone el dominio
    /// (<see cref="NombresHojaPeriodo"/>) para 2026091. Una tercera hoja es FAIL: el lector de bloque
    /// resuelve las 2 hojas por nombre (nunca por enumeración) y el contrato del §T0 fija el conteo.
    /// </summary>
    [Fact]
    public void T0_01_R10_TraeExactamenteDosHojasDelPeriodo()
    {
        var periodo = Periodo.Parse(PeriodoR10);
        using var zip = Abrir(RutaR10());

        var hojas = NombresHojas(zip);

        Assert.Equal(2, hojas.Count);
        Assert.Equal(NombresHojaPeriodo.DetRetri(periodo), hojas[0]);
        Assert.Equal(NombresHojaPeriodo.DetValiRetri(periodo), hojas[1]);
        Assert.Equal("DetRetri2026091", hojas[0]);
        Assert.Equal("DetValiRetri2026091", hojas[1]);
    }

    /// <summary>
    /// T0-02 (encabezado corporativo común): fila 3 razón social/NIT/código RMC 2026091, fila 6
    /// Fecha de Proceso 17/09/2026, fila 7 Hora 08:44 AM + rango 01/09–15/09, encabezado de tabla en
    /// fila 8 y ASE 1..5 en filas 9–13 con `Total` en la fila 14.
    ///
    /// <para>
    /// <b>Desviación declarada (§6 del doc):</b> el §T0 dice "literales de <c>A3:C8</c>"; en el byte
    /// los literales viven en <c>F3</c>/<c>H3</c>, <c>C6</c>/<c>D6</c>, <c>C7</c>/<c>D7</c>,
    /// <c>F7</c>/<c>G7</c>, <c>I7</c>/<c>J7</c> y el código de remuneración en <c>L3</c>/<c>M3</c>
    /// (DetRetri) o <c>J3</c>/<c>K3</c> (DetValiRetri). Las columnas A, B y E de las filas 3–7 están
    /// vacías. El assert congela la posición REAL (que es la de la plantilla de cada hoja, ver T0-04).
    /// </para>
    /// </summary>
    [Fact]
    public void T0_02_R10_EncabezadoCorporativo_SelloYRango()
    {
        var periodo = Periodo.Parse(PeriodoR10);
        using var zip = Abrir(RutaR10());
        var compartidas = Compartidas(zip);
        var det = Hoja(zip, NombresHojaPeriodo.DetRetri(periodo));
        var vali = Hoja(zip, NombresHojaPeriodo.DetValiRetri(periodo));

        // (a) Encabezado corporativo COMÚN a las 2 hojas (fila 3 a fila 7).
        foreach (var hoja in new[] { det, vali })
        {
            Assert.Equal("PROCERASEO S.A.S.", Texto(hoja, compartidas, "F3"));
            Assert.Equal("NIT: 901 153 639 - 0", Texto(hoja, compartidas, "H3"));
            Assert.Equal("3. ESQUEMA FINANCIERO DEL SERVICIO PUBLICO DE ASEO", Texto(hoja, compartidas, "F4"));
            Assert.Equal("Resolución UAESP Número 27 de 2018", Texto(hoja, compartidas, "F5"));

            // Fila 6: Fecha de Proceso. Fila 7: Hora + rango del período.
            Assert.Equal("Fecha de Proceso:", Texto(hoja, compartidas, "C6"));
            Assert.Equal("17/09/2026", Texto(hoja, compartidas, "D6"));
            Assert.Equal("R10 - Recaudo Consolidado", Texto(hoja, compartidas, "F6"));
            Assert.Equal("Hora:", Texto(hoja, compartidas, "C7"));
            Assert.Equal("08:44 AM", Texto(hoja, compartidas, "D7"));
            Assert.Equal("Fecha Desde:", Texto(hoja, compartidas, "F7"));
            Assert.Equal("01/09/2026", Texto(hoja, compartidas, "G7"));
            Assert.Equal("Fecha Hasta:", Texto(hoja, compartidas, "I7"));
            Assert.Equal("15/09/2026", Texto(hoja, compartidas, "J7"));
        }

        // (b) Código de remuneración del período: RMC 2026091 (columna distinta por hoja).
        Assert.Equal("Código Remuneración:", Texto(det, compartidas, "L3"));
        Assert.Equal("2026091", Texto(det, compartidas, "M3"));
        Assert.Equal("Código Remuneración:", Texto(vali, compartidas, "J3"));
        Assert.Equal("2026091", Texto(vali, compartidas, "K3"));

        // (c) Encabezado de tabla en la fila 8: 14 columnas B..O, con ASE / CONCESIONARIOS.
        foreach (var hoja in new[] { det, vali })
        {
            Assert.Equal(14, CeldasOcupadas(hoja, 8, 8, "B", "O").Count);
            Assert.Equal("ASE", Texto(hoja, compartidas, "B8"));
            Assert.Equal("CONCESIONARIOS", Texto(hoja, compartidas, "C8"));
        }

        // (d) ASE 1..5 en filas 9–13 y fila 14 = Total.
        foreach (var hoja in new[] { det, vali })
        {
            for (var i = 0; i < 5; i++)
            {
                Assert.Equal(
                    (i + 1).ToString(CultureInfo.InvariantCulture),
                    Texto(hoja, compartidas, $"B{9 + i}"));
            }

            Assert.Equal("Total", Texto(hoja, compartidas, "C14"));
        }
    }

    /// <summary>
    /// T0-03 (montos como TEXTO monetario): en el R10 los 72 montos de cada hoja
    /// (<c>D9:O14</c>) son <c>t="s"</c> con la convención es-CO del lector
    /// (<see cref="ExcelDataReaderDetRetriR10Reader"/>): positivo <c>"$17.450.250.543"</c>,
    /// negativo <c>"-$…"</c>. El assert exige ≥1 positivo y ≥1 negativo y verifica el parseo "al
    /// centavo" de los valores citados.
    ///
    /// <para>
    /// <b>Desviación declarada (§6 del doc):</b> el §T0 cita el negativo <c>-$5.093.257.797</c>; el
    /// byte dice <c>-$5.093.259.797</c> (ASE2, <c>DetRetri2026091!J10</c>). La CONVENCIÓN (texto
    /// monetario negativo con prefijo <c>-$</c>) se confirma; el dígito se re-mide y se reporta.
    /// </para>
    /// </summary>
    [Fact]
    public void T0_03_R10_MontosSonTextoMonetario_ParseoAlCentavo()
    {
        var periodo = Periodo.Parse(PeriodoR10);
        using var zip = Abrir(RutaR10());
        var compartidas = Compartidas(zip);
        var det = Hoja(zip, NombresHojaPeriodo.DetRetri(periodo));
        var vali = Hoja(zip, NombresHojaPeriodo.DetValiRetri(periodo));

        var positivos = 0;
        var negativos = 0;

        foreach (var hoja in new[] { det, vali })
        {
            for (var fila = 9; fila <= 14; fila++)
            {
                foreach (var columna in Columnas)
                {
                    var referencia = $"{columna}{fila}";
                    var celda = Celda(hoja, referencia);
                    Assert.True(celda is not null, $"{referencia} no existe en el R10 (slot de monto).");
                    Assert.Null(celda!.Element(Ns + "f"));

                    // El monto viaja como TEXTO (shared string), NO como número con formato moneda.
                    Assert.Equal("s", celda.Attribute("t")?.Value);

                    var texto = Texto(hoja, compartidas, referencia);
                    Assert.True(
                        MontoMonetarioRegex.IsMatch(texto),
                        $"{referencia} = '{texto}' no cumple la convención monetaria es-CO (^\\-?\\$[0-9.]+$).");

                    // "Al centavo": el parseo documentado no deja residuo fraccionario.
                    var valor = ParseoMonetarioT0(texto);
                    Assert.Equal(decimal.Truncate(valor), valor);

                    if (texto.StartsWith("-$", StringComparison.Ordinal))
                    {
                        negativos++;
                    }
                    else
                    {
                        positivos++;
                    }
                }
            }
        }

        Assert.True(positivos >= 1, "El R10 debe traer al menos un monto positivo con '$'.");
        Assert.True(negativos >= 1, "El R10 debe traer al menos un monto negativo con '-$'.");

        // Valores citados, medidos al centavo (bytes; ver §6 del doc para la desviación del dígito).
        Assert.Equal("$17.450.250.543", Texto(det, compartidas, "D9"));
        Assert.Equal(17_450_250_543m, ParseoMonetarioT0(Texto(det, compartidas, "D9")));
        Assert.Equal("-$5.093.259.797", Texto(det, compartidas, "J10"));
        Assert.Equal(-5_093_259_797m, ParseoMonetarioT0(Texto(det, compartidas, "J10")));

        // DetValiRetri trae sus propios negativos (COSTO DE LA INTERVENTORÍA) como texto monetario.
        Assert.Equal("-$189.185.987", Texto(vali, compartidas, "I9"));
        Assert.Equal(-189_185_987m, ParseoMonetarioT0(Texto(vali, compartidas, "I9")));
    }

    /// <summary>
    /// T0-04 (columna A vacía; B = ASE idéntico a la plantilla): en el R10 la columna A está vacía,
    /// la columna B es el id 1..5 y C..O son montos. El layout es el MISMO que la hoja homónima de la
    /// plantilla (base 2026092): mismos encabezados de la fila 8 (14 columnas), mismos rótulos
    /// <c>C9:C13</c> — incluido ASE3 sin punto final (H-59-4) — y misma fila 14 `Total`.
    /// </summary>
    [Fact]
    public void T0_04_R10_LayoutIdenticoALaPlantilla_ColumnaAVacia_BEsAse()
    {
        var periodo = Periodo.Parse(PeriodoR10);
        var periodoPlantilla = Periodo.Parse(PeriodoBasePlantilla);

        using var r10 = Abrir(RutaR10());
        using var plantilla = Abrir(RutaPlantilla());

        var sstR10 = Compartidas(r10);
        var sstPlantilla = Compartidas(plantilla);

        foreach (var (nombreR10, nombrePlantilla, filaFin) in new[]
                 {
                     (NombresHojaPeriodo.DetRetri(periodo), NombresHojaPeriodo.DetRetri(periodoPlantilla), 19),
                     (NombresHojaPeriodo.DetValiRetri(periodo), NombresHojaPeriodo.DetValiRetri(periodoPlantilla), 14),
                 })
        {
            var hojaR10 = Hoja(r10, nombreR10);
            var hojaPlantilla = Hoja(plantilla, nombrePlantilla);

            // (a) Columna A vacía en AMBOS archivos (el layout arranca en B).
            Assert.Empty(CeldasOcupadas(hojaR10, 3, filaFin, "A", "A"));
            Assert.Empty(CeldasOcupadas(hojaPlantilla, 3, filaFin, "A", "A"));

            // (b) B9:B13 = 1..5 en ambos.
            for (var i = 0; i < 5; i++)
            {
                var esperado = (i + 1).ToString(CultureInfo.InvariantCulture);
                Assert.Equal(esperado, Texto(hojaR10, sstR10, $"B{9 + i}"));
                Assert.Equal(esperado, Texto(hojaPlantilla, sstPlantilla, $"B{9 + i}"));
            }

            // (c) C9:C13 = mismos rótulos R10-vs-plantilla (texto exacto), incluido ASE3 sin punto.
            for (var i = 0; i < 5; i++)
            {
                Assert.Equal(
                    Texto(hojaPlantilla, sstPlantilla, $"C{9 + i}"),
                    Texto(hojaR10, sstR10, $"C{9 + i}"));
            }

            Assert.Equal("CIUDAD LIMPIA BOGOTA S.A. ESP", Texto(hojaR10, sstR10, "C11"));
            Assert.False(
                Texto(hojaR10, sstR10, "C11").EndsWith('.'),
                "H-59-4: ASE3 no lleva punto final en el rótulo (R10 ni plantilla).");

            // (d) Encabezado de tabla fila 8: 14 columnas con el MISMO texto R10-vs-plantilla.
            Assert.Equal(14, CeldasOcupadas(hojaR10, 8, 8, "B", "O").Count);
            foreach (var columna in "BCDEFGHIJKLMNO")
            {
                Assert.Equal(
                    Texto(hojaPlantilla, sstPlantilla, $"{columna}8"),
                    Texto(hojaR10, sstR10, $"{columna}8"));
            }

            // (e) Fila 14 = Total (mismo rótulo; D..O son montos en ambas hojas).
            Assert.Equal("Total", Texto(hojaR10, sstR10, "C14"));
            Assert.Equal("Total", Texto(hojaPlantilla, sstPlantilla, "C14"));

            // (f) El BLOQUE de encabezado (filas 3–7, A..O) ocupa exactamente las mismas celdas en
            //     ambos archivos: el pegado celda-a-celda no desplaza el encabezado.
            Assert.Equal(
                CeldasOcupadas(hojaPlantilla, 3, 7, "A", "O"),
                CeldasOcupadas(hojaR10, 3, 7, "A", "O"));
        }
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T0b — Plantilla DetRetri (sheet36): top-block values-only, fórmulas bajo la fila 19 (T0-05…06)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T0-05 (slots literales): en <c>DetRetri2026092</c> las filas 3–19 son literales+rótulos (sello
    /// D6/D7, headers fila 8, B9:C13, fila 14 `Total`) y los 72 slots de monto <c>D9:O14</c> son
    /// literales <c>&lt;v&gt;0&lt;/v&gt;</c> (destinos del pegado). Cero <c>&lt;f&gt;</c> en
    /// <c>A3:O19</c>.
    ///
    /// <para>
    /// <b>Desviación declarada (§6 del doc):</b> el §T0 rotula el bloque T2 como
    /// <c>REPORTE R10- RECAUDO CONSOLIDADO</c>; el byte de <c>F6</c> dice
    /// <c>R10 - Recaudo Consolidado</c>. Se congela el literal medido.
    /// </para>
    /// </summary>
    [Fact]
    public void T0_05_PlantillaDetRetri_TopBlockValuesOnly_CeroFormulasEnA3O19()
    {
        using var zip = Abrir(RutaPlantilla());
        var hoja = Hoja(zip, NombresHojaPeriodo.DetRetri(PeriodoBasePlantilla));
        var compartidas = Compartidas(zip);

        // (a) 0 celdas con <f> en el rango completo del pegado (A3:O19), barrido XML.
        var formulas = Formulas(hoja);
        Assert.DoesNotContain(
            formulas.Keys,
            referencia => EstaEnRango(referencia, "A3", "O19"));

        // (b) Los 72 slots de monto D9:O14 son literales numéricos en 0 (destinos de escritura).
        foreach (var columna in Columnas)
        {
            for (var fila = 9; fila <= 14; fila++)
            {
                var referencia = $"{columna}{fila}";
                var celda = Celda(hoja, referencia);
                Assert.True(celda is not null, $"{referencia} no existe en DetRetri2026092.");
                Assert.Null(celda!.Element(Ns + "f"));
                Assert.Equal("0", celda.Element(Ns + "v")?.Value);
                Assert.Null(celda.Attribute("t")); // numérico, sin shared string
            }
        }

        // (c) Rótulos y sello del top-block medidos por bytes.
        Assert.Equal("2026092", Texto(hoja, compartidas, "M3"));
        Assert.Equal("Fecha de Proceso:", Texto(hoja, compartidas, "C6"));
        Assert.Equal("02/10/2026", Texto(hoja, compartidas, "D6"));
        Assert.Equal("Hora:", Texto(hoja, compartidas, "C7"));
        Assert.Equal("10:43 AM", Texto(hoja, compartidas, "D7"));
        Assert.Equal("R10 - Recaudo Consolidado", Texto(hoja, compartidas, "F6"));
        Assert.Equal("ASE", Texto(hoja, compartidas, "B8"));
        Assert.Equal("Total", Texto(hoja, compartidas, "C14"));
    }

    /// <summary>
    /// T0-06 (fórmulas FUERA del pegado): las 132 fórmulas de <c>DetRetri2026092</c> viven
    /// EXACTAMENTE en <c>D23:O28</c> (consumidoras del CONSOLIDADO) y <c>D32:O36</c> (auto-chequeos,
    /// incluidos <c>L32:L36</c>/<c>O32:O36</c>): todas bajo la fila 19. El pegado <c>A3:O19</c> no
    /// toca ninguna <c>&lt;f&gt;</c> (diff antes/después = ∅ por construcción del rango).
    /// </summary>
    [Fact]
    public void T0_06_PlantillaDetRetri_FormulasSoloBajoLaFila19()
    {
        using var zip = Abrir(RutaPlantilla());
        var hoja = Hoja(zip, NombresHojaPeriodo.DetRetri(PeriodoBasePlantilla));

        var formulas = Formulas(hoja);
        Assert.Equal(132, formulas.Count);

        var esperadas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var fila = 23; fila <= 28; fila++)
        {
            foreach (var columna in Columnas)
            {
                esperadas.Add($"{columna}{fila}");
            }
        }

        for (var fila = 32; fila <= 36; fila++)
        {
            foreach (var columna in Columnas)
            {
                esperadas.Add($"{columna}{fila}");
            }
        }

        Assert.Equal(132, esperadas.Count);
        Assert.True(
            esperadas.SetEquals(formulas.Keys),
            $"Fórmulas fuera del contrato T0 en DetRetri: {string.Join(", ", formulas.Keys.Except(esperadas))}");

        // Todas BAJO la fila 19: el pegado A3:O19 nunca pisa una fórmula.
        Assert.Equal(23, formulas.Keys.Min(FilaDe));
        Assert.DoesNotContain(formulas.Keys, referencia => FilaDe(referencia) <= 19);

        // Forma de la cadena protegida (Plan 59 §3.4) que consume el CONSOLIDADO por ASE.
        for (var i = 0; i < 5; i++)
        {
            Assert.Equal($"D{9 + i}-'CONSOLIDADO_TOTAL RECAUDO'!D{104 + i}", formulas[$"D{23 + i}"]);
        }

        // Fila Total (28): D14 - D109 del CONSOLIDADO.
        Assert.Equal("D14-'CONSOLIDADO_TOTAL RECAUDO'!D109", formulas["D28"]);

        // Auto-chequeos (L) de interventoría y reconciliación (O): ambos bajo la fila 19.
        for (var i = 0; i < 5; i++)
        {
            Assert.StartsWith($"INTERVENTORIA!M{26 + i}+L{23 + i}=L{9 + i}", formulas[$"L{32 + i}"], StringComparison.Ordinal);
            Assert.Contains($"-INTERVENTORIA!M{26 + i}", formulas[$"O{32 + i}"], StringComparison.Ordinal);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T0c — Plantilla DetValiRetri (sheet37): idem + consumidoras de lo pegado (T0-07…T0-09)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T0-07 (slots literales): en <c>DetValiRetri2026092</c> el rango del pegado <c>A3:O14</c> tiene
    /// 0 celdas con <c>&lt;f&gt;</c> y los 72 slots <c>D9:O14</c> son literales numéricos en 0.
    /// </summary>
    [Fact]
    public void T0_07_PlantillaDetValiRetri_TopBlockValuesOnly_CeroFormulasEnA3O14()
    {
        using var zip = Abrir(RutaPlantilla());
        var hoja = Hoja(zip, NombresHojaPeriodo.DetValiRetri(PeriodoBasePlantilla));
        var compartidas = Compartidas(zip);

        var formulas = Formulas(hoja);
        Assert.DoesNotContain(formulas.Keys, referencia => EstaEnRango(referencia, "A3", "O14"));

        foreach (var columna in Columnas)
        {
            for (var fila = 9; fila <= 14; fila++)
            {
                var referencia = $"{columna}{fila}";
                var celda = Celda(hoja, referencia);
                Assert.True(celda is not null, $"{referencia} no existe en DetValiRetri2026092.");
                Assert.Null(celda!.Element(Ns + "f"));
                Assert.Equal("0", celda.Element(Ns + "v")?.Value);
                Assert.Null(celda.Attribute("t"));
            }
        }

        // Sello y rótulos del top-block (columna del código = J3/K3 en esta hoja).
        Assert.Equal("2026092", Texto(hoja, compartidas, "K3"));
        Assert.Equal("02/10/2026", Texto(hoja, compartidas, "D6"));
        Assert.Equal("10:43 AM", Texto(hoja, compartidas, "D7"));
        Assert.Equal("AJUSTE A LA DECENA", Texto(hoja, compartidas, "J8"));
        Assert.Equal("Total", Texto(hoja, compartidas, "C14"));

        // Todas las <f> de la hoja están BAJO el rango del pegado.
        Assert.Equal(16, formulas.Keys.Min(FilaDe));
    }

    /// <summary>
    /// T0-08 (fórmulas 16–20 intactas y CONSUMIDORAS): <c>D16:O20</c> = <c>D9:O13</c> menos
    /// <c>'CONSOLIDADO_TOTAL RECAUDO'!U104:AG104</c> por ASE (fila 16 ↔ 9 / 20 ↔ 13) queda FUERA del
    /// pegado y consume lo pegado. El assert congela el TEXTO-<c>&lt;f&gt;</c> exacto celda a celda
    /// (72 celdas), que es el contrato de "byte-idéntico antes/después" del H2.
    ///
    /// <para>
    /// Mapa de columnas medido: D→U, E→V, F→W, G→X, H→Y, I→Z, J→AA−AC (2 términos),
    /// K→AB, L→AE, M→AD, N→AF, O→AG. La fila <c>21</c> (<c>Total</c>) repite la forma contra
    /// <c>U109:AG109</c> y también queda fuera del pegado.
    /// </para>
    /// </summary>
    [Fact]
    public void T0_08_PlantillaDetValiRetri_D16O20ConsumidorasIntactas()
    {
        using var zip = Abrir(RutaPlantilla());
        var hoja = Hoja(zip, NombresHojaPeriodo.DetValiRetri(PeriodoBasePlantilla));

        var formulas = Formulas(hoja);
        var mapa = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["D"] = "U", ["E"] = "V", ["F"] = "W", ["G"] = "X", ["H"] = "Y", ["I"] = "Z",
            ["J"] = "AA", ["K"] = "AB", ["L"] = "AE", ["M"] = "AD", ["N"] = "AF", ["O"] = "AG",
        };

        for (var ase = 0; ase < 5; ase++)
        {
            var filaConsumo = 16 + ase;
            var filaOrigen = 9 + ase;
            var filaConsolidado = 104 + ase;

            foreach (var columna in Columnas)
            {
                var referencia = $"{columna}{filaConsumo}";
                var origen = $"{columna}{filaOrigen}";
                var destino = mapa[columna];

                var esperado = string.Equals(columna, "J", StringComparison.Ordinal)
                    ? $"J{filaOrigen}-'CONSOLIDADO_TOTAL RECAUDO'!AA{filaConsolidado}-'CONSOLIDADO_TOTAL RECAUDO'!AC{filaConsolidado}"
                    : $"{origen}-'CONSOLIDADO_TOTAL RECAUDO'!{destino}{filaConsolidado}";

                Assert.Equal(esperado, formulas[referencia]);

                // La fórmula CONSUME su fila 9–13 (lo que el pegado escribe).
                Assert.Contains(origen, formulas[referencia], StringComparison.Ordinal);
            }
        }

        // Ninguna <f> de la hoja cae en el rango del pegado (filas 3–14): el diff de <f> antes/
        // después del pegado es ∅ por construcción del rango.
        Assert.DoesNotContain(formulas.Keys, referencia => FilaDe(referencia) <= 14);

        // Fila Total (21) fuera del pegado y con la misma forma contra *109.
        Assert.Equal("D14-'CONSOLIDADO_TOTAL RECAUDO'!U109", formulas["D21"]);
        Assert.Equal("O14-'CONSOLIDADO_TOTAL RECAUDO'!AG109", formulas["O21"]);
    }

    /// <summary>
    /// T0-09 (<c>J9:J14</c> entra al pegado): la columna <c>J</c> (<c>AJUSTE A LA DECENA</c>) está
    /// dentro del rango <c>A3:O14</c>. El R10 trae DATO REAL (≠ 0) en <c>J9:J14</c> y el manual del
    /// administrativo trae el MISMO conjunto: la exclusión de comparador <c>DetValiRetri!J</c> (SALE,
    /// Plan 29 T0d) se retira en H5.
    /// </summary>
    [Fact]
    public void T0_09_J9J14TieneDatoRealEnR10YEnElManual()
    {
        var periodo = Periodo.Parse(PeriodoR10);
        using var r10 = Abrir(RutaR10());
        using var manual = Abrir(RutaManual());

        var sstR10 = Compartidas(r10);
        var hojaR10 = Hoja(r10, NombresHojaPeriodo.DetValiRetri(periodo));
        var hojaManual = Hoja(manual, NombresHojaPeriodo.DetValiRetri(periodo));

        Assert.Equal("AJUSTE A LA DECENA", Texto(hojaR10, sstR10, "J8"));

        var valoresR10 = new List<decimal>();
        for (var fila = 9; fila <= 14; fila++)
        {
            var referencia = $"J{fila}";

            // Está DENTRO del rango del pegado.
            Assert.True(EstaEnRango(referencia, "A3", "O14"), $"{referencia} debe caer en A3:O14.");

            var texto = Texto(hojaR10, sstR10, referencia);
            var valorR10 = ParseoMonetarioT0(texto);
            Assert.True(valorR10 != 0m, $"{referencia} del R10 no debe ser 0 (hay dato real que pegar).");
            valoresR10.Add(valorR10);

            // El manual (oráculo de paridad del pegado) trae el MISMO valor, como literal numérico.
            var celdaManual = Celda(hojaManual, referencia);
            Assert.True(celdaManual is not null, $"{referencia} no existe en el manual Administrativo.");
            Assert.Null(celdaManual!.Element(Ns + "f"));
            var valorManual = decimal.Parse(
                celdaManual.Element(Ns + "v")?.Value ?? "0",
                NumberStyles.Number,
                CultureInfo.InvariantCulture);
            Assert.Equal(valorManual, valorR10);
        }

        // Serie medida (bytes): AJUSTE A LA DECENA por ASE 1..5 + Total.
        Assert.Equal(
            new[] { -46_493m, 4_616m, 46_663m, -64_129m, -16_688m, -76_031m },
            valoresR10);
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T0d — CONSOLIDADO (sheet18) filas 141–147: estado actual y remapeo exigido (T0-10…T0-12)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T0-10 (geometría): fila 141 = headers, filas 142–146 = un ASE por fila (<c>B</c> = 1..5),
    /// fila 147 = Total. El assert congela la geometría y la <c>&lt;f&gt;</c> de la fila de cierre.
    ///
    /// <para>
    /// <b>Desviación declarada (§6 del doc):</b> el §T0 dice <c>C147 = Total</c> (y "comparte
    /// s-string con <c>Total</c>"); el byte dice <c>TOTAL</c> (mayúsculas) y usa un s-string DISTINTO
    /// del <c>Total</c> de <c>DetRetri!C14</c>. Se congela el literal medido.
    /// </para>
    /// </summary>
    [Fact]
    public void T0_10_Consolidado_GeometriaFilas141A147()
    {
        using var zip = Abrir(RutaPlantilla());
        var hoja = Hoja(zip, HojaConsolidado);
        var compartidas = Compartidas(zip);

        // (a) Fila 141: encabezado del bloque.
        Assert.Equal("ASE", Texto(hoja, compartidas, "B141"));
        Assert.Equal("Concesionarios", Texto(hoja, compartidas, "C141"));
        Assert.Equal("Recaudo REMUNERADO", Texto(hoja, compartidas, "D141"));
        Assert.Equal("RECAUDO RECIBIDO", Texto(hoja, compartidas, "E141"));

        // (b) Filas 142–146: B = 1..5 (literal numérico) y C = rótulo del ASE.
        for (var ase = 1; ase <= 5; ase++)
        {
            var fila = 141 + ase;
            var celdaB = Celda(hoja, $"B{fila}");
            Assert.True(celdaB is not null, $"B{fila} no existe en {HojaConsolidado}.");
            Assert.Null(celdaB!.Attribute("t"));
            Assert.Equal(
                ase.ToString(CultureInfo.InvariantCulture),
                celdaB.Element(Ns + "v")?.Value);
        }

        Assert.Equal("PROMOAMBIENTAL DISTRITO S.A.S. ESP.", Texto(hoja, compartidas, "C142"));
        Assert.Equal("AREA LIMPIA DISTRITO CAPITAL S.A.S. ESP.", Texto(hoja, compartidas, "C146"));

        // (c) Fila 147: Total del bloque (D..H con SUM; la <f> de E147 es el contrato del remapeo).
        Assert.Equal("TOTAL", Texto(hoja, compartidas, "C147"));
        Assert.Equal("SUM(D142:D146)", FormulaDe(hoja, "D147"));
        Assert.Equal("SUM(E142:E146)", FormulaDe(hoja, "E147"));
        Assert.Equal("SUM(F142:F146)", FormulaDe(hoja, "F147"));
        Assert.Equal("SUM(H142:H146)", FormulaDe(hoja, "H147"));
    }

    /// <summary>
    /// T0-11 (fórmulas por fila k, estado actual): <c>E{k}</c> = Σ de las 5 hojas <c>'Recaudo *'</c>
    /// en <b>columna F</b> y filas <c>21..25</c> (E142→21 … E146→25), en la forma exacta del §T0.
    /// <c>D{k}</c> internas sin quincena, <c>L{k}</c> = <c>D-E-F-G-J-H-I</c> y <c>M142</c> =
    /// <c>INT(ABS(L142))=0</c>.
    /// </summary>
    [Fact]
    public void T0_11_Consolidado_E142E146_SumaRecaudoColumnasF_Filas21a25()
    {
        using var zip = Abrir(RutaPlantilla());
        var hoja = Hoja(zip, HojaConsolidado);

        var hojasRecaudo = new[]
        {
            "Recaudo EAAB Reciprocidad",
            "Recaudo EAAB + Ciud Limp",
            "Recaudo Directa Occidente",
            "Recaudo ENEL",
            "Recaudo ENERBIT",
        };

        for (var i = 0; i < 5; i++)
        {
            var fila = 142 + i;
            var filaRecaudo = 21 + i;

            // (a) E{k} = forma EXACTA (5 términos, columna F, fila 21..25 de cada hoja de recaudo).
            var esperadoE = string.Join(
                "+",
                hojasRecaudo.Select(hojaRecaudo => $"'{hojaRecaudo}'!F{filaRecaudo}"));
            Assert.Equal(esperadoE, FormulaDe(hoja, $"E{fila}"));

            var terminos = SepararSuma(esperadoE);
            Assert.Equal(5, terminos.Count);
            foreach (var termino in terminos)
            {
                Assert.Matches(@"^'Recaudo [^']+'!F(2[1-5])$", termino);
                Assert.EndsWith($"F{filaRecaudo}", termino, StringComparison.Ordinal);
            }

            // (b) D{k} interna (sin quincena): D{104+i} + N{104+i} + D{125+i} + E{125+i}.
            Assert.Equal(
                $"D{104 + i}+N{104 + i}+D{125 + i}+E{125 + i}",
                FormulaDe(hoja, $"D{fila}"));

            // (c) F/G/H internas (R2/R4/R5) sin quincena.
            Assert.Equal($"D{28 + i}+N{28 + i}", FormulaDe(hoja, $"F{fila}"));
            Assert.Equal($"D{66 + i}+N{66 + i}", FormulaDe(hoja, $"G{fila}"));
            Assert.Equal($"D{85 + i}+N{85 + i}", FormulaDe(hoja, $"H{fila}"));
        }

        // (d) L{k} = D-E-F-G-J-H-I → el visible que arrastra E{k}. L145 es el MASTER de una
        //     shared-formula (ref L145:L147, si=68) y L146/L147 son seguidoras SIN texto propio:
        //     por eso el fixture no exige texto en 146/147 (doctrina Plan 33 — el pegado nunca toca L).
        for (var fila = 142; fila <= 145; fila++)
        {
            Assert.Equal(
                $"D{fila}-E{fila}-F{fila}-G{fila}-J{fila}-H{fila}-I{fila}",
                FormulaDe(hoja, $"L{fila}"));
        }

        Assert.Equal("shared", AtributoFormula(hoja, "L145", "t"));
        Assert.Equal("L145:L147", AtributoFormula(hoja, "L145", "ref"));
        var sharedL = AtributoFormula(hoja, "L145", "si");
        Assert.NotNull(sharedL);
        foreach (var fila in new[] { 146, 147 })
        {
            Assert.Equal(string.Empty, FormulaDe(hoja, $"L{fila}"));
            Assert.Equal("shared", AtributoFormula(hoja, $"L{fila}", "t"));
            Assert.Equal(sharedL, AtributoFormula(hoja, $"L{fila}", "si"));
        }

        // (e) M{k} = INT(ABS(L{k}))=0: M142 es el master (ref M142:M147, si=67) y M143:M147 son
        //     seguidoras booleanas (celda t="b") sin texto propio.
        Assert.Equal("INT(ABS(L142))=0", FormulaDe(hoja, "M142"));
        Assert.Equal("shared", AtributoFormula(hoja, "M142", "t"));
        Assert.Equal("M142:M147", AtributoFormula(hoja, "M142", "ref"));
        for (var fila = 143; fila <= 147; fila++)
        {
            Assert.Equal(string.Empty, FormulaDe(hoja, $"M{fila}"));
            Assert.Equal("shared", AtributoFormula(hoja, $"M{fila}", "t"));
            Assert.Equal("b", Celda(hoja, $"M{fila}")?.Attribute("t")?.Value);
        }
    }

    /// <summary>
    /// T0-12 (desalineación Q1): la columna E de esta fila NO pertenece a la familia E1-4, así que
    /// queda FIJA en la columna <b>F</b> de las hojas <c>'Recaudo *'</c>, mientras el writer de
    /// recaudo gobierna la columna por quincena (<see cref="PreparadorBasePipeline.ColumnaRecaudoDe"/>
    /// → <c>ColumnaRecaudoPorQuincena</c>: Q1 = D, Q2 = F) y <c>ValidadorBasePeriodo</c> valida con la
    /// misma fuente. En Q1 el recaudo aterriza en <c>D{21:25}</c> y <c>E142</c> sigue leyendo
    /// <c>F21</c> ⇒ bloque en 0.
    ///
    /// <para>
    /// <b>Demostración:</b> estática (texto-<c>&lt;f&gt;</c> medido + columna de dominio). La
    /// demostración DINÁMICA en una corrida Q1 real queda declarada para el Gate <b>G.3</b> del plan
    /// (este fixture NO ejecuta producción; §6 del doc).
    /// </para>
    /// </summary>
    [Fact]
    public void T0_12_Consolidado_DesalineacionQ1_E142LeeF_MientrasQ1EscribeD()
    {
        using var zip = Abrir(RutaPlantilla());
        var hoja = Hoja(zip, HojaConsolidado);

        // (a) Las 5 <f> del bloque fijan la columna F de 'Recaudo *' (ni una sola referencia a D).
        for (var i = 0; i < 5; i++)
        {
            var formula = FormulaDe(hoja, $"E{142 + i}");
            Assert.NotNull(formula);
            foreach (var termino in SepararSuma(formula!))
            {
                Assert.Matches(@"^'Recaudo [^']+'!F\d+$", termino);
            }

            Assert.DoesNotMatch(@"'Recaudo [^']+'!D\d+", formula);
        }

        // (b) La columna del recaudo la decide el DOMINIO (dominio, nunca contenido): Q1 → D, Q2 → F.
        var q1 = Periodo.Parse("2026091");
        var q2 = Periodo.Parse(PeriodoBasePlantilla);
        Assert.Equal(1, q1.NumeroQuincena);
        Assert.Equal(2, q2.NumeroQuincena);
        Assert.Equal("D", PreparadorBasePipeline.ColumnaRecaudoDe(q1));
        Assert.Equal("F", PreparadorBasePipeline.ColumnaRecaudoDe(q2));

        // (c) Desalineación Q1 demostrada: el bloque E142:E146 lee F{21:25} pero en Q1 el recaudo se
        //     escribe en D{21:25}. En Q2 la columna coincide (F) ⇒ el defecto es SOLO de Q1 (N2).
        const string columnaQueLeeElBloque = "F";
        var columnaQueEscribeQ1 = PreparadorBasePipeline.ColumnaRecaudoDe(q1);
        var columnaQueEscribeQ2 = PreparadorBasePipeline.ColumnaRecaudoDe(q2);

        Assert.NotEqual(columnaQueLeeElBloque, columnaQueEscribeQ1);
        Assert.Equal(columnaQueLeeElBloque, columnaQueEscribeQ2);
        Assert.Equal("'Recaudo EAAB Reciprocidad'!F21", SepararSuma(FormulaDe(hoja, "E142")!)[0]);
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // Lectura BCL (ZIP + XML) — rutas EXPLÍCITAS, sin Insumos.Raiz(), sin escritura
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    private static ZipArchive Abrir(string ruta) => ZipFile.OpenRead(ruta);

    private static string RutaR10() => Path.Combine(
        Raiz(), "Docs", "Prueba Septiembre 1", "Insumos", "R10_Remuneracion_2026091.xlsx");

    private static string RutaPlantilla() => Path.Combine(
        Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string RutaManual() => Path.Combine(
        Raiz(), "Docs", "Prueba Septiembre 1", "Resultados", "Administrativo", "Remuneracion 202609-1 Total.xlsx");

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(
                    dir.FullName,
                    "Docs",
                    "Plantilla estandarizada",
                    "Formato Remuneracion 2026xx-x Total.xlsx")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException(
            "No se encontró la raíz del repositorio (Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx).");
    }

    private static List<string> NombresHojas(ZipArchive zip)
    {
        var libro = LeerXml(zip, PartWorkbook);
        return libro.Root?.Element(Ns + "sheets")?.Elements(Ns + "sheet")
            .Select(hoja => hoja.Attribute("name")!.Value).ToList() ?? [];
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

    private static List<string> Compartidas(ZipArchive zip)
    {
        var doc = LeerXml(zip, PartCompartidas);
        return doc.Root?.Elements(Ns + "si").Select(si => si.Value).ToList() ?? [];
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

    private static string Texto(XDocument hoja, List<string> compartidas, string referencia)
    {
        var celda = Celda(hoja, referencia);
        if (celda is null)
        {
            return string.Empty;
        }

        if (celda.Attribute("t")?.Value == "s"
            && int.TryParse(celda.Element(Ns + "v")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indice)
            && indice >= 0 && indice < compartidas.Count)
        {
            return compartidas[indice];
        }

        return celda.Element(Ns + "inlineStr")?.Value
            ?? celda.Element(Ns + "v")?.Value
            ?? string.Empty;
    }

    /// <summary>Formulas de la hoja por referencia de celda (texto del <c>&lt;f&gt;</c>).</summary>
    private static Dictionary<string, string> Formulas(XDocument hoja)
    {
        var mapa = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var celda in hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row").Elements(Ns + "c") ?? [])
        {
            var formula = celda.Element(Ns + "f")?.Value;
            if (formula is null)
            {
                continue;
            }

            mapa[celda.Attribute("r")?.Value ?? string.Empty] = formula;
        }

        return mapa;
    }

    /// <summary>Referencias ocupadas (con contenido) del rectángulo indicado, ordenadas.</summary>
    private static List<string> CeldasOcupadas(XDocument hoja, int filaDesde, int filaHasta, string colDesde, string colHasta)
    {
        var desde = NumeroColumna(colDesde);
        var hasta = NumeroColumna(colHasta);
        var referencias = new List<string>();

        foreach (var fila in hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row") ?? [])
        {
            if (!int.TryParse(fila.Attribute("r")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var numero)
                || numero < filaDesde || numero > filaHasta)
            {
                continue;
            }

            foreach (var celda in fila.Elements(Ns + "c"))
            {
                var referencia = celda.Attribute("r")?.Value;
                if (string.IsNullOrEmpty(referencia))
                {
                    continue;
                }

                var columna = NumeroColumna(new string(referencia.Where(char.IsAsciiLetter).ToArray()));
                if (columna < desde || columna > hasta)
                {
                    continue;
                }

                // "Ocupada" = tiene valor, fórmula o texto en línea (un estilo suelto no cuenta).
                if (celda.Element(Ns + "v") is not null
                    || celda.Element(Ns + "f") is not null
                    || celda.Element(Ns + "is") is not null)
                {
                    referencias.Add(referencia);
                }
            }
        }

        referencias.Sort(StringComparer.Ordinal);
        return referencias;
    }

    private static bool EstaEnRango(string referencia, string desde, string hasta)
    {
        var columna = NumeroColumna(new string(referencia.Where(char.IsAsciiLetter).ToArray()));
        var fila = FilaDe(referencia);
        return columna >= NumeroColumna(new string(desde.Where(char.IsAsciiLetter).ToArray()))
            && columna <= NumeroColumna(new string(hasta.Where(char.IsAsciiLetter).ToArray()))
            && fila >= FilaDe(desde)
            && fila <= FilaDe(hasta);
    }

    private static int FilaDe(string referencia) =>
        int.Parse(new string(referencia.Where(char.IsAsciiDigit).ToArray()), CultureInfo.InvariantCulture);

    private static int NumeroColumna(string letras)
    {
        var numero = 0;
        foreach (var caracter in letras)
        {
            numero = (numero * 26) + (char.ToUpperInvariant(caracter) - 'A' + 1);
        }

        return numero;
    }

    /// <summary>
    /// Parte una suma por <c>'+'</c> respetando los nombres de hoja entrecomillados
    /// (<c>'Recaudo EAAB + Ciud Limp'</c> contiene un '+' literal). BCL puro.
    /// </summary>
    private static List<string> SepararSuma(string formula)
    {
        var terminos = new List<string>();
        var enComillas = false;
        var inicio = 0;

        for (var i = 0; i < formula.Length; i++)
        {
            if (formula[i] == '\'')
            {
                enComillas = !enComillas;
            }
            else if (formula[i] == '+' && !enComillas)
            {
                terminos.Add(formula[inicio..i]);
                inicio = i + 1;
            }
        }

        terminos.Add(formula[inicio..]);
        return terminos;
    }

    /// <summary>
    /// T0-03 (oráculo del fixture): delega en <see cref="ParseoMonetarioR10"/> — la fuente ÚNICA de
    /// la convención monetaria es-CO (<c>'.'</c> miles, <c>','</c> decimal) que el lector del R10
    /// consume. H1 cumplió lo que este helper declaraba: el conversor salió del lector a
    /// <c>ParseoMonetarioR10</c> y aquí ya no se replica (T0-03: "el parseo vive en UNA fuente").
    /// </summary>
    private static decimal ParseoMonetarioT0(string texto) => ParseoMonetarioR10.Convertir(texto);
}
