using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 62 (H1 — T1.1..T1.4): lector del BLOQUE del R10 (<c>DetRetri2026091</c> filas 3–19 y
/// <c>DetValiRetri2026091</c> filas 3–14, columnas A–O) que el pegado copiará a las hojas homónimas
/// de la plantilla (N1), y el conversor monetario ÚNICO <see cref="ParseoMonetarioR10"/>.
///
/// <para>
/// <b>Fuente medida:</b> <c>Docs/Prueba Septiembre 1/Insumos/R10_Remuneracion_2026091.xlsx</c> por
/// ruta EXPLÍCITA (Plan 62 §T0/D-H0). Cero escritura sobre los insumos: los casos negativos trabajan
/// sobre COPIAS temporales mutadas con OpenXML.
/// </para>
///
/// <para>
/// <b>Trazabilidad:</b> T1.1 (modelo + contrato: bloques completos 17×15 y 12×15) · T1.2 (monto
/// exacto y fix del decimal es-CO) · T1.3 (encabezados que VIAJAN, per-ASE/Total vs el oráculo,
/// celdas vacías explícitas, fail-fast sin bloque parcial) · T1.4 (este fixture).
/// </para>
/// </summary>
public sealed class BloqueR10H1Plan62Tests
{
    /// <summary>Tolerancia del proyecto entre valor leído y valor de referencia (redondeo).</summary>
    private const decimal Tolerancia = 0.5m;

    private const string CodigoPeriodo = "2026091";

    private static readonly string[] ColumnasDelBloque =
        ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"];

    private static ExcelDataReaderDetRetriR10Reader Reader() => new();

    private static Periodo PeriodoDelCaso() => Periodo.Parse(CodigoPeriodo);

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T1.1 — Contrato del bloque: modelo puro + rango completo (A3:O19 / A3:O14)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T1.1: el bloque de <c>DetRetri2026091</c> cubre EXACTAMENTE el rango del contrato T0
    /// (<c>A3:O19</c> = 17 filas × 15 columnas = 255 celdas), con las hojas y el período resueltos
    /// por el dominio. Cada celda existe con su coordenada real: el pegado y su read-back recorren
    /// las mismas coordenadas.
    /// </summary>
    [Fact]
    public void T1_1_BloqueDetRetri_CubreElContratoA3O19()
    {
        var bloque = Reader().LeerBloqueDetRetri(PeriodoDelCaso(), RutaR10());

        Assert.Equal("DetRetri2026091", bloque.Hoja);
        Assert.Equal(CodigoPeriodo, bloque.Periodo.CodigoCompleto);
        Assert.Equal(3, bloque.FilaDesde);
        Assert.Equal(19, bloque.FilaHasta);
        Assert.Equal(17 * 15, bloque.Celdas.Count);

        AfirmarRejillaCompleta(bloque, 3, 19);
        Assert.Equal("A3", bloque.Celdas[0].Referencia);
        Assert.Equal("O19", bloque.Celdas[^1].Referencia);
    }

    /// <summary>
    /// T1.1: el bloque de <c>DetValiRetri2026091</c> cubre <c>A3:O14</c> (12 filas × 15 columnas =
    /// 180 celdas) y se resuelve por <see cref="NombresHojaPeriodo"/>.
    /// </summary>
    [Fact]
    public void T1_1_BloqueDetValiRetri_CubreElContratoA3O14()
    {
        var bloque = Reader().LeerBloqueDetValiRetri(PeriodoDelCaso(), RutaR10());

        Assert.Equal("DetValiRetri2026091", bloque.Hoja);
        Assert.Equal(CodigoPeriodo, bloque.Periodo.CodigoCompleto);
        Assert.Equal(3, bloque.FilaDesde);
        Assert.Equal(14, bloque.FilaHasta);
        Assert.Equal(12 * 15, bloque.Celdas.Count);

        AfirmarRejillaCompleta(bloque, 3, 14);
        Assert.Equal("A3", bloque.Celdas[0].Referencia);
        Assert.Equal("O14", bloque.Celdas[^1].Referencia);
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T1.3 — Los encabezados de las filas 3–8 VIAJAN en el bloque (R-Aclareo ratificado)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T1.3 + §11 R-Aclareo: los encabezados corporativos del período (razón social, NIT, título,
    /// resolución, sello de proceso, código de remuneración y rango) viajan pegados en el bloque —
    /// no se filtran ni se reconstruyen. El sello se lee como TEXTO exacto (doctrina Plan 31 R-S-1:
    /// <c>dd/MM/yyyy</c> + <c>hh:mm AM/PM</c>), y las columnas que en el R10 están vacías viajan como
    /// celdas explícitas vacías (T0-04: la columna A y la E de las filas 3–7 no traen literales).
    /// </summary>
    [Fact]
    public void T1_3_LosEncabezadosCorporativosViajanEnElBloque()
    {
        var reader = Reader();
        var periodo = PeriodoDelCaso();
        var ruta = RutaR10();

        var det = reader.LeerBloqueDetRetri(periodo, ruta);
        var vali = reader.LeerBloqueDetValiRetri(periodo, ruta);

        foreach (var bloque in new[] { det, vali })
        {
            Assert.Equal("PROCERASEO S.A.S.", TextoDe(bloque, 3, "F"));
            Assert.Equal("NIT: 901 153 639 - 0", TextoDe(bloque, 3, "H"));
            Assert.Equal("3. ESQUEMA FINANCIERO DEL SERVICIO PUBLICO DE ASEO", TextoDe(bloque, 4, "F"));
            Assert.Equal("Resolución UAESP Número 27 de 2018", TextoDe(bloque, 5, "F"));

            // Sello de proceso del R10: fecha y hora EXACTAS (texto, no serial OADate).
            Assert.Equal("Fecha de Proceso:", TextoDe(bloque, 6, "C"));
            Assert.Equal("17/09/2026", TextoDe(bloque, 6, "D"));
            Assert.Equal("Hora:", TextoDe(bloque, 7, "C"));
            Assert.Equal("08:44 AM", TextoDe(bloque, 7, "D"));
            Assert.Equal("R10 - Recaudo Consolidado", TextoDe(bloque, 6, "F"));

            // Rango del período del R10 (no el del template).
            Assert.Equal("Fecha Desde:", TextoDe(bloque, 7, "F"));
            Assert.Equal("01/09/2026", TextoDe(bloque, 7, "G"));
            Assert.Equal("Fecha Hasta:", TextoDe(bloque, 7, "I"));
            Assert.Equal("15/09/2026", TextoDe(bloque, 7, "J"));

            // Encabezado de la tabla (14 columnas B..O) y rótulos por ASE.
            Assert.Equal("ASE", TextoDe(bloque, 8, "B"));
            Assert.Equal("CONCESIONARIOS", TextoDe(bloque, 8, "C"));
            for (var ase = 1; ase <= 5; ase++)
            {
                Assert.Equal(ase.ToString(CultureInfo.InvariantCulture), TextoDe(bloque, 8 + ase, "B"));
            }

            Assert.Equal("Total", TextoDe(bloque, 14, "C"));

            // R-2: la columna A (filas 3–8) y la columna E (filas 3–7) están VACÍAS en el R10 y
            // viajan como celdas explícitas vacías: no se omiten ni se inventan.
            for (var fila = 3; fila <= 8; fila++)
            {
                Assert.True(CeldaDe(bloque, fila, "A").Vacia, $"A{fila} debe viajar vacía en '{bloque.Hoja}'.");
            }

            for (var fila = 3; fila <= 7; fila++)
            {
                Assert.True(CeldaDe(bloque, fila, "E").Vacia, $"E{fila} debe viajar vacía en '{bloque.Hoja}'.");
            }
        }

        // Código de remuneración del período: columna distinta por hoja (T0-02: M3 / K3).
        Assert.Equal("Código Remuneración:", TextoDe(det, 3, "L"));
        Assert.Equal(CodigoPeriodo, TextoDe(det, 3, "M"));
        Assert.Equal("Código Remuneración:", TextoDe(vali, 3, "J"));
        Assert.Equal(CodigoPeriodo, TextoDe(vali, 3, "K"));
    }

    /// <summary>
    /// T1.3: el pie del bloque de <c>DetRetri</c> (filas 15–19) también viaja, por posición relativa
    /// al cierre <c>Total</c>: <c>C16</c> (costos de interventoría del período),
    /// <c>E18</c> y <c>F19</c> (firma). Es lo que el manual reemplaza del pie heredado del template
    /// (desplazamiento DECLARADO del pegado, Docs/T0-Plan62-Pegado-R10.md §1).
    /// </summary>
    [Fact]
    public void T1_3_ElPieDelBloqueDetRetriViajaConElPeriodo()
    {
        var bloque = Reader().LeerBloqueDetRetri(PeriodoDelCaso(), RutaR10());

        Assert.StartsWith("* Costos de Interventoría:", TextoDe(bloque, 16, "C"), StringComparison.Ordinal);
        Assert.Contains("202609-1", TextoDe(bloque, 16, "C"), StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(TextoDe(bloque, 18, "E")));
        Assert.Equal("Firma autorizada", TextoDe(bloque, 19, "F"));

        // Los montos del bloque terminan en la fila del Total: el pie no trae valores numéricos.
        Assert.True(CeldaDe(bloque, 15, "D").Vacia);
        Assert.True(CeldaDe(bloque, 17, "D").Vacia);
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T1.3/T1.4 — Per-ASE y Total: el bloque y el oráculo coinciden al centavo (identidad T1.2)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T1.4: por ASE y en el Total, el bloque <c>D9:D14</c> coincide con lo que devuelve el lector
    /// del oráculo (<see cref="IDetRetriR10Reader.LeerDetRetri"/>) dentro de ±0.5, y con los valores
    /// medidos en bytes del período (T0-03/T0-04). Esto es la identidad exigida por T1.2: extraer el
    /// conversor a <see cref="ParseoMonetarioR10"/> no movió un solo peso.
    /// </summary>
    [Fact]
    public void T1_4_PerAseYTotalDelBloqueCoincidenConElOraculo()
    {
        var reader = Reader();
        var periodo = PeriodoDelCaso();
        var ruta = RutaR10();

        var oraculo = reader.LeerDetRetri(periodo, ruta);
        var bloque = reader.LeerBloqueDetRetri(periodo, ruta);

        var esperadoPorAse = new Dictionary<int, decimal>
        {
            [1] = 17_450_250_543m,
            [2] = 14_147_357_755m,
            [3] = 10_891_338_739m,
            [4] = 11_508_947_207m,
            [5] = 9_147_179_378m,
        };

        foreach (var (ase, esperado) in esperadoPorAse)
        {
            var delBloque = NumeroDe(bloque, 8 + ase, "D");
            Assert.InRange(delBloque, esperado - Tolerancia, esperado + Tolerancia);
            Assert.InRange(oraculo.DetRetriPorAse[ase], delBloque - Tolerancia, delBloque + Tolerancia);
        }

        Assert.InRange(NumeroDe(bloque, 14, "D"), 63_145_073_622m - Tolerancia, 63_145_073_622m + Tolerancia);
        Assert.InRange(oraculo.Total, 63_145_073_622m - Tolerancia, 63_145_073_622m + Tolerancia);

        // Rango del período: el bloque lo trae como texto del R10 y el oráculo lo expone como fecha.
        Assert.Equal(new DateTime(2026, 9, 1), oraculo.FechaDesde);
        Assert.Equal(new DateTime(2026, 9, 15), oraculo.FechaHasta);
    }

    /// <summary>
    /// T1.3: los montos del bloque se representan como NUMÉRICO exacto y sin redondear (el redondeo
    /// a entero es regla de escritura del CONSOLIDADO, no de lectura). Se fija el negativo citado en
    /// §11 R-1 (<c>DetRetri2026091!J10 = -$5.093.259.797</c>) y la serie de <c>AJUSTE A LA DECENA</c>
    /// de <c>DetValiRetri</c> (T0-09), que entra al pegado con signo.
    /// </summary>
    [Fact]
    public void T1_3_MontosViajanComoNumericoExactoConSigno()
    {
        var reader = Reader();
        var periodo = PeriodoDelCaso();
        var ruta = RutaR10();

        var det = reader.LeerBloqueDetRetri(periodo, ruta);

        var j10 = CeldaDe(det, 10, "J");
        Assert.Null(j10.Texto);
        Assert.Equal(-5_093_259_797m, j10.Numerico);

        // Todo monto del bloque es numérico: nunca se transporta la representación "$…".
        for (var fila = 9; fila <= 14; fila++)
        {
            foreach (var columna in ColumnasDelBloque.Skip(3))
            {
                var celda = CeldaDe(det, fila, columna);
                Assert.Null(celda.Texto);
                Assert.NotNull(celda.Numerico);
                Assert.Equal(decimal.Truncate(celda.Numerico!.Value), celda.Numerico!.Value);
            }
        }

        var vali = reader.LeerBloqueDetValiRetri(periodo, ruta);
        var ajusteEsperado = new Dictionary<int, decimal>
        {
            [9] = -46_493m,
            [10] = 4_616m,
            [11] = 46_663m,
            [12] = -64_129m,
            [13] = -16_688m,
            [14] = -76_031m,
        };

        foreach (var (fila, esperado) in ajusteEsperado)
        {
            Assert.Equal(esperado, NumeroDe(vali, fila, "J"));
        }
    }

    /// <summary>
    /// T1.1: las celdas del rango que están vacías se representan de forma EXPLÍCITA (presentes,
    /// <see cref="CeldaBloqueR10.Vacia"/>) en vez de omitirse: el bloque cubre el rectángulo completo
    /// y la fila 15 del <c>DetRetri</c> (sin contenido en el R10) no desaparece.
    /// </summary>
    [Fact]
    public void T1_3_LasCeldasVaciasDelRangoSeRepresentanExplicitas()
    {
        var reader = Reader();
        var periodo = PeriodoDelCaso();
        var ruta = RutaR10();

        var det = reader.LeerBloqueDetRetri(periodo, ruta);
        for (var fila = 3; fila <= 19; fila++)
        {
            Assert.True(CeldaDe(det, fila, "A").Vacia, $"A{fila} debe existir vacía en el bloque de DetRetri.");
        }

        foreach (var columna in ColumnasDelBloque)
        {
            Assert.True(CeldaDe(det, 15, columna).Vacia, $"La fila 15 no trae contenido (columna {columna}).");
            Assert.True(CeldaDe(det, 17, columna).Vacia, $"La fila 17 no trae contenido (columna {columna}).");
        }

        var vali = reader.LeerBloqueDetValiRetri(periodo, ruta);
        for (var fila = 3; fila <= 14; fila++)
        {
            Assert.True(CeldaDe(vali, fila, "A").Vacia, $"A{fila} debe existir vacía en el bloque de DetValiRetri.");
        }

        // El bloque NO inventa celdas fuera del contrato.
        Assert.Null(det.Celda(20, "A"));
        Assert.Null(det.Celda(19, "P"));
        Assert.Null(vali.Celda(15, "A"));
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T1.2 — ParseoMonetarioR10: fuente única y fix del decimal es-CO
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T1.2 (fix del decimal, Plan 62 §11): la convención DOCUMENTADA (es-CO: <c>'.'</c> miles,
    /// <c>','</c> decimal) es la que se implementa. El código anterior
    /// (<c>texto.Where(char.IsDigit)</c>) leía <c>"$1.234,56"</c> como <c>123456</c>; aquí vale
    /// <c>1234.56</c>. Los montos del 2026091 son enteros, así que el fix no mueve los goldens — es
    /// precondición del retiro del gate con montos decimales de períodos futuros.
    /// </summary>
    [Theory]
    [InlineData("$1.234,56", "1234.56")]
    [InlineData("$17.450.250.543", "17450250543")]
    [InlineData("-$5.093.259.797", "-5093259797")]
    [InlineData("$9.147.179.378,75", "9147179378.75")]
    [InlineData("-$0", "0")]
    [InlineData("$0", "0")]
    [InlineData("0", "0")]
    [InlineData("1234", "1234")]
    [InlineData("1.234", "1234")]
    [InlineData("  $1.234,5  ", "1234.5")]
    public void T1_2_ParseoMonetarioR10_HonraElDecimalEsCo(string texto, string esperado)
    {
        var valorEsperado = decimal.Parse(esperado, NumberStyles.Number, CultureInfo.InvariantCulture);

        Assert.True(ParseoMonetarioR10.TryParseTexto(texto, out var numero), $"'{texto}' debería ser un monto legible.");
        Assert.Equal(valorEsperado, numero);
        Assert.Equal(valorEsperado, ParseoMonetarioR10.Convertir(texto));
    }

    /// <summary>
    /// T1.2: el conversor NO redondea (conserva el decimal) y NO adivina convenios ajenos: un texto
    /// con la coma antes que el punto (anglosajón) se rechaza en vez de producir un monto ×1000 en
    /// silencio; una celda vacía o sin dígitos es ilegible (modo estricto) y vale 0 en el modo
    /// histórico del oráculo.
    /// </summary>
    [Fact]
    public void T1_2_ParseoMonetarioR10_NoRedondeaNiAdivinaConvenios()
    {
        Assert.True(ParseoMonetarioR10.TryParseTexto("$9.147.179.378,75", out var conDecimales));
        Assert.Equal(9_147_179_378.75m, conDecimales);
        Assert.NotEqual(decimal.Truncate(conDecimales), conDecimales);

        Assert.False(ParseoMonetarioR10.TryParseTexto("1,234.56", out _));
        Assert.False(ParseoMonetarioR10.TryParseTexto("1,2,3", out _));
        Assert.False(ParseoMonetarioR10.TryParseTexto("abc", out _));
        Assert.False(ParseoMonetarioR10.TryParseTexto(string.Empty, out _));
        Assert.False(ParseoMonetarioR10.TryParseTexto(null, out _));

        // Semántica histórica del oráculo (celda vacía/ilegible = 0): identidad de los goldens.
        Assert.Equal(0m, ParseoMonetarioR10.Convertir(null));
        Assert.Equal(0m, ParseoMonetarioR10.Convertir(string.Empty));
        Assert.Equal(0m, ParseoMonetarioR10.Convertir("abc"));

        // Celda numérica: se toma el valor tal cual (sin redondear).
        Assert.Equal(1234.5m, ParseoMonetarioR10.Convertir(1234.5m));
        Assert.Equal(17_450_250_543m, ParseoMonetarioR10.Convertir(17_450_250_543d));
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // T1.3 — Fail-fast sin bloque parcial (copias temporales mutadas; los insumos NO se tocan)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// T1.3: si el R10 no trae la hoja del período, el lector de bloque falla antes de devolver nada
    /// y nombra el archivo y la hoja (nunca lee otra hoja del workbook ni cae a un período vecino).
    /// </summary>
    [Fact]
    public void T1_3_R10SinLaHojaDelPeriodo_FallaNombrandoArchivoYHoja()
    {
        var periodo = PeriodoDelCaso();
        using var copia = new CopiaTemporalDelR10();
        var hojaEsperada = NombresHojaPeriodo.DetValiRetri(periodo);
        RenombrarHoja(copia.Ruta, hojaEsperada, "DetValiRetri209901");

        var excepcion = Assert.Throws<ArchivoFuenteNoEncontradoException>(
            () => Reader().LeerBloqueDetValiRetri(periodo, copia.Ruta));

        Assert.Contains(hojaEsperada, excepcion.Message, StringComparison.Ordinal);
        Assert.Contains(Path.GetFileName(copia.Ruta), excepcion.Message, StringComparison.Ordinal);
        Assert.Equal(CodigoError.FuenteNoEncontrada, excepcion.Codigo);
    }

    /// <summary>
    /// T1.3: si un R10 existe pero no está en la ruta del período, el fallo nombra la ruta (nunca
    /// "bloque vacío").
    /// </summary>
    [Fact]
    public void T1_3_R10Inexistente_FallaNombrandoLaRuta()
    {
        var ruta = Path.Combine(Path.GetTempPath(), "r10-inexistente-" + Guid.NewGuid().ToString("N") + ".xlsx");

        var excepcion = Assert.Throws<ArchivoFuenteNoEncontradoException>(
            () => Reader().LeerBloqueDetRetri(PeriodoDelCaso(), ruta));

        Assert.Contains(ruta, excepcion.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(CodigoError.FuenteNoEncontrada, excepcion.Codigo);
    }

    /// <summary>
    /// T1.3: una celda del contrato faltante en el R10 (aquí el cierre <c>Total</c> de
    /// <c>DetRetri</c>) es fail-fast <c>ERR-FORMATO-FUENTE</c> que nombra período + archivo + hoja +
    /// celda, ANTES de devolver el bloque: el pegado nunca recibe un bloque parcial.
    /// </summary>
    [Fact]
    public void T1_3_R10SinElCierreTotal_FallaSinBloqueParcial()
    {
        var periodo = PeriodoDelCaso();
        using var copia = new CopiaTemporalDelR10();
        QuitarCelda(copia.Ruta, NombresHojaPeriodo.DetRetri(periodo), "C14");

        var excepcion = Assert.Throws<CalculoInvalidoException>(
            () => Reader().LeerBloqueDetRetri(periodo, copia.Ruta));

        Assert.Equal(CodigoError.FormatoFuente, excepcion.Codigo);
        Assert.Contains("C14", excepcion.Message, StringComparison.Ordinal);
        Assert.Contains(NombresHojaPeriodo.DetRetri(periodo), excepcion.Message, StringComparison.Ordinal);
        Assert.Contains(CodigoPeriodo, excepcion.Message, StringComparison.Ordinal);
        Assert.Contains(Path.GetFileName(copia.Ruta), excepcion.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// T1.3: un monto faltante (aquí <c>D9</c> de <c>DetValiRetri</c>) también es fail-fast
    /// <c>ERR-FORMATO-FUENTE</c> que nombra la celda exacta: el bloque no inventa un 0.
    /// </summary>
    [Fact]
    public void T1_3_R10SinUnMonto_FallaNombrandoLaCelda()
    {
        var periodo = PeriodoDelCaso();
        using var copia = new CopiaTemporalDelR10();
        QuitarCelda(copia.Ruta, NombresHojaPeriodo.DetValiRetri(periodo), "D9");

        var excepcion = Assert.Throws<CalculoInvalidoException>(
            () => Reader().LeerBloqueDetValiRetri(periodo, copia.Ruta));

        Assert.Equal(CodigoError.FormatoFuente, excepcion.Codigo);
        Assert.Contains("D9", excepcion.Message, StringComparison.Ordinal);
        Assert.Contains(NombresHojaPeriodo.DetValiRetri(periodo), excepcion.Message, StringComparison.Ordinal);
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // Utilidades del fixture (rutas EXPLÍCITAS; los insumos se abren SOLO lectura)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    private static void AfirmarRejillaCompleta(BloqueHojaR10 bloque, int filaDesde, int filaHasta)
    {
        var indice = 0;
        for (var fila = filaDesde; fila <= filaHasta; fila++)
        {
            foreach (var columna in ColumnasDelBloque)
            {
                var celda = bloque.Celdas[indice++];
                Assert.Equal(fila, celda.Fila);
                Assert.Equal(columna, celda.Columna);
                Assert.Equal($"{columna}{fila}", celda.Referencia);
            }
        }

        Assert.Equal(bloque.Celdas.Count, indice);
    }

    private static CeldaBloqueR10 CeldaDe(BloqueHojaR10 bloque, int fila, string columna) =>
        bloque.Celda(fila, columna)
        ?? throw new Xunit.Sdk.XunitException($"El bloque '{bloque.Hoja}' no trae la celda {columna}{fila}.");

    private static string TextoDe(BloqueHojaR10 bloque, int fila, string columna) =>
        CeldaDe(bloque, fila, columna).Texto
        ?? throw new Xunit.Sdk.XunitException($"La celda {columna}{fila} del bloque '{bloque.Hoja}' no es texto.");

    private static decimal NumeroDe(BloqueHojaR10 bloque, int fila, string columna) =>
        CeldaDe(bloque, fila, columna).Numerico
        ?? throw new Xunit.Sdk.XunitException($"La celda {columna}{fila} del bloque '{bloque.Hoja}' no es un monto.");

    /// <summary>
    /// Ruta EXPLÍCITA al R10 del caso (Plan 62 §T0). Prohibido resolverlo por enumeración de
    /// carpetas: la evidencia se ata al archivo medido.
    /// </summary>
    private static string RutaR10() => Path.Combine(
        Raiz(), "Docs", "Prueba Septiembre 1", "Insumos", "R10_Remuneracion_2026091.xlsx");

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (Directory.Exists(Path.Combine(dir.FullName, "Docs", "Prueba Septiembre 1", "Insumos")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException(
            "No se encontró la raíz del repositorio (Docs/Prueba Septiembre 1/Insumos).");
    }

    /// <summary>
    /// Copia de trabajo del R10 en temp: los casos negativos mutan la COPIA con OpenXML y jamás
    /// tocan el insumo (Excel cerrado ⇒ lectura/escritura de la copia, insumo intacto).
    /// </summary>
    private sealed class CopiaTemporalDelR10 : IDisposable
    {
        private readonly string _carpeta;

        public CopiaTemporalDelR10()
        {
            _carpeta = Path.Combine(
                Path.GetTempPath(), "remuneracion-p62-h1-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_carpeta);
            Ruta = Path.Combine(_carpeta, Path.GetFileName(RutaR10()));
            File.Copy(RutaR10(), Ruta, overwrite: false);
        }

        public string Ruta { get; }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_carpeta))
                {
                    Directory.Delete(_carpeta, recursive: true);
                }
            }
            catch (IOException)
            {
                // Limpieza best-effort: la copia vive en temp y no afecta el resultado del test.
            }
        }
    }

    private static void RenombrarHoja(string ruta, string nombreActual, string nombreNuevo)
    {
        using var documento = SpreadsheetDocument.Open(ruta, isEditable: true);
        var libro = documento.WorkbookPart
            ?? throw new InvalidOperationException($"'{Path.GetFileName(ruta)}' no es un workbook.");
        var libroXml = libro.Workbook
            ?? throw new InvalidOperationException($"'{Path.GetFileName(ruta)}' no es un workbook.");
        var hojas = libroXml.Sheets
            ?? throw new InvalidOperationException($"'{Path.GetFileName(ruta)}' no declara hojas.");

        var hoja = hojas.Elements<Sheet>()
            .FirstOrDefault(sheet => string.Equals(sheet.Name?.Value, nombreActual, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"El workbook no trae la hoja '{nombreActual}'.");

        hoja.Name = nombreNuevo;
        libroXml.Save();
    }

    private static void QuitarCelda(string ruta, string nombreHoja, string referencia)
    {
        using var documento = SpreadsheetDocument.Open(ruta, isEditable: true);
        var libro = documento.WorkbookPart
            ?? throw new InvalidOperationException($"'{Path.GetFileName(ruta)}' no es un workbook.");
        var libroXml = libro.Workbook
            ?? throw new InvalidOperationException($"'{Path.GetFileName(ruta)}' no es un workbook.");
        var hojas = libroXml.Sheets
            ?? throw new InvalidOperationException($"'{Path.GetFileName(ruta)}' no declara hojas.");

        var hoja = hojas.Elements<Sheet>()
            .FirstOrDefault(sheet => string.Equals(sheet.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"El workbook no trae la hoja '{nombreHoja}'.");

        var parte = libro.GetPartById(hoja.Id!.Value!) as WorksheetPart
            ?? throw new InvalidOperationException($"'{nombreHoja}' no es una hoja de cálculo.");
        var hojaXml = parte.Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombreHoja}' no se pudo abrir.");
        var celdas = hojaXml.Descendants<Cell>();
        var celda = celdas
            .FirstOrDefault(celda => string.Equals(celda.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"La hoja '{nombreHoja}' no trae la celda '{referencia}'.");

        celda.Remove();
        hojaXml.Save();
    }
}
