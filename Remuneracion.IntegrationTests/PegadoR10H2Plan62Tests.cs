using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 62 (H2 — T2.1..T2.5): pegado del bloque del R10 en <c>DetRetri{AAAAMMQ}</c> /
/// <c>DetValiRetri{AAAAMMQ}</c> (values-only, estilos del destino preservados, guard anti-fórmula,
/// geometría de fórmulas intacta) y read-back de integridad del pegado.
///
/// <para>
/// <b>Fuente medida:</b> <c>Docs/Prueba Septiembre 1/Insumos/R10_Remuneracion_2026091.xlsx</c> y la
/// plantilla oficial <c>Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx</c>
/// (gate de entrada del plan §gate_entrada), ambas por ruta EXPLÍCITA. Cero escritura sobre los
/// insumos: cada caso trabaja sobre una COPIA temporal preparada por la Phase 0 de producción
/// (<see cref="PreparadorBasePipeline"/>) — el mismo camino de la app.
/// </para>
///
/// <para>
/// <b>Trazabilidad:</b> T2.2 (pegado values-only + guard anti-fórmula + sin redondeo, N-5) · T2.3
/// (read-back con hoja+celda+esperado+escrito y forma-agnóstico en <c>B9:B13</c>, H3-SEMANTIC) ·
/// T2.4 (D6/D7 entregados ≡ R10, sin sello aparte) · T2.1 (el orquestador aporta los bloques antes
/// de escribir: el entregable de septiembre-1 trae el R10).
/// </para>
///
/// <para>
/// <b>Reclasificación (spectrum §7 del plan):</b> los asserts de <c>D9:D14</c> y del desglose
/// trazable de los fixtures del Plan 59 (<c>AsientoDetRetriEscritoPlan59Tests</c>,
/// <c>DesgloseTrazableAmbasQuincenasPlan59Tests</c>) se trasladan AQUÍ, al read-back contra el R10:
/// esas escrituras se retiraron en H2 (D-59H4) porque el pegado las cubre.
/// </para>
/// </summary>
public sealed class PegadoR10H2Plan62Tests
{
    /// <summary>Tolerancia del proyecto para montos (redondeo).</summary>
    private const decimal Tolerancia = 0.5m;

    private const string CodigoPeriodo = "2026091";

    private static readonly string[] ColumnasDelBloque =
        ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"];

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T2.2 / T2.3 — Septiembre-1: el pegado reproduce el R10 (read-back) y no toca una sola <f>
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T2.2: tras pegar el bloque del R10 en una copia preparada de la plantilla oficial, TODAS las
    /// celdas de los dos rangos del contrato (<c>DetRetri</c> A3:O19 y <c>DetValiRetri</c> A3:O14)
    /// coinciden con el R10 — montos ±0,5, textos exactos, vacías vacías — verificado por una
    /// lectura INDEPENDIENTE del fixture (no por el verificador de producción). Se fija además el
    /// negativo de §11 R-1 (<c>J10 = -5.093.259.797</c>) y el sello de proceso (T2.4).
    /// </summary>
    [Fact]
    public void T2_2_Septiembre1_ElPegadoReproduceElR10EnAmbasHojas()
    {
        using var caso = CasoPegado.Nuevo();
        caso.Pegar();

        using var documento = SpreadsheetDocument.Open(caso.Ruta, false);
        var workbookPart = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart nulo.");

        AfirmarBloquePegado(workbookPart, caso.BloqueDetRetri, 3, 19);
        AfirmarBloquePegado(workbookPart, caso.BloqueValiRetri, 3, 14);

        // §11 R-1: el negativo citado por el Ingeniero.
        Assert.Equal(-5_093_259_797m, NumeroDe(workbookPart, caso.BloqueDetRetri.Hoja, 10, "J"));

        // T0-09 (D-J9): la serie AJUSTE A LA DECENA entra pegada (antes quedaba en 0 y era exclusión
        // del comparador); se fija la fila del Total y la primera fila por ASE.
        var hojaVali = caso.BloqueValiRetri.Hoja;
        Assert.Equal(-46_493m, NumeroDe(workbookPart, hojaVali, 9, "J"));
        Assert.Equal(-76_031m, NumeroDe(workbookPart, hojaVali, 14, "J"));

        // T2.4 (D-31SELLO): el sello de proceso llega PEGADO del R10 en las dos hojas.
        foreach (var hoja in new[] { caso.BloqueDetRetri.Hoja, caso.BloqueValiRetri.Hoja })
        {
            Assert.Equal("17/09/2026", TextoDe(workbookPart, hoja, "D6"));
            Assert.Equal("08:44 AM", TextoDe(workbookPart, hoja, "D7"));
        }

        // §11 R-Aclareo: los encabezados corporativos del período VIAJAN (no se conserva el template).
        Assert.Equal("2026091", TextoDe(workbookPart, caso.BloqueDetRetri.Hoja, "M3"));
        Assert.Contains("202609-1", TextoDe(workbookPart, caso.BloqueDetRetri.Hoja, "C16"), StringComparison.Ordinal);
    }

    /// <summary>
    /// T2.2 (R5-GATE): el pegado NO altera ni una sola <c>&lt;f&gt;</c> — ni texto, ni <c>t</c>, ni
    /// <c>ref</c>, ni <c>si</c> — en las hojas tocadas, y las fórmulas que consumen lo pegado
    /// (<c>DetValiRetri!D16:O20</c>, T0-08) quedan byte-idénticas. Es el diff <c>&lt;f&gt;</c>
    /// antes/después = ∅ exigido por el plan.
    /// </summary>
    [Fact]
    public void T2_2_ElPegadoNoTocaNingunaFormula()
    {
        using var caso = CasoPegado.Nuevo();
        var antes = FormulasDe(caso.Ruta);

        caso.Pegar();

        var despues = FormulasDe(caso.Ruta);

        Assert.Equal(antes.Keys.OrderBy(k => k, StringComparer.Ordinal), despues.Keys.OrderBy(k => k, StringComparer.Ordinal));
        foreach (var (celda, texto) in antes)
        {
            Assert.Equal(texto, despues[celda]);
        }

        // T0-08, explícito: las 5 columnas × 5 filas que consumen lo pegado siguen siendo fórmula.
        using var documento = SpreadsheetDocument.Open(caso.Ruta, false);
        var hoja = Hoja(documento.WorkbookPart!, caso.BloqueValiRetri.Hoja);
        for (var fila = 16; fila <= 20; fila++)
        {
            foreach (var columna in new[] { "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O" })
            {
                var celda = Celda(hoja, $"{columna}{fila}")
                    ?? throw new InvalidOperationException($"Falta {columna}{fila} en DetValiRetri.");
                Assert.NotNull(celda.CellFormula);
            }
        }
    }

    /// <summary>
    /// T2.2: el pegado PRESERVA el estilo (<c>s</c>) del destino: es copia de valores, no de formato.
    /// Se mide el <c>s</c> de una celda de rótulo, una de monto y otra de texto antes y después.
    /// </summary>
    [Fact]
    public void T2_2_ElPegadoPreservaElEstiloDelDestino()
    {
        using var caso = CasoPegado.Nuevo();
        var estilosAntes = new Dictionary<string, uint?>(StringComparer.Ordinal);
        foreach (var (hoja, celda) in new[]
                 {
                     (caso.BloqueDetRetri.Hoja, "D9"),
                     (caso.BloqueDetRetri.Hoja, "C9"),
                     (caso.BloqueDetRetri.Hoja, "F3"),
                     (caso.BloqueValiRetri.Hoja, "J9"),
                 })
        {
            estilosAntes[$"{hoja}!{celda}"] = EstiloDe(caso.Ruta, hoja, celda);
        }

        caso.Pegar();

        foreach (var (clave, estilo) in estilosAntes)
        {
            var partes = clave.Split('!');
            Assert.Equal(estilo, EstiloDe(caso.Ruta, partes[0], partes[1]));
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T2.2 — Guard anti-fórmula (R-PEGADO-FORMULA) y fail-fast sin escritura parcial
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T2.2 / R-PEGADO-FORMULA: si el destino trae <c>&lt;f&gt;</c> en una celda del rango, el pegado
    /// aborta con <c>ERR-PLANTILLA</c> que NOMBRA hoja + celda, y no escribió NADA (plan-then-write):
    /// otra celda del mismo bloque conserva el valor del template.
    /// </summary>
    [Fact]
    public void T2_2_GuardAntiFormula_NombraHojaYCeldaYNoEscribeNada()
    {
        using var caso = CasoPegado.Nuevo();
        var valorPrevioF3 = TextoDeAlmacenado(caso.Ruta, caso.BloqueDetRetri.Hoja, "F3");
        InjertarFormula(caso.Ruta, caso.BloqueDetRetri.Hoja, "D9", "1+1");

        var excepcion = Assert.Throws<CalculoInvalidoException>(() => caso.Pegar());

        Assert.Equal(CodigoError.Plantilla, excepcion.Codigo);
        Assert.Contains(caso.BloqueDetRetri.Hoja, excepcion.Message, StringComparison.Ordinal);
        Assert.Contains("D9", excepcion.Message, StringComparison.Ordinal);

        // Sin escritura parcial: la primera celda del plan (F3) sigue con el literal del template.
        Assert.Equal(valorPrevioF3, TextoDeAlmacenado(caso.Ruta, caso.BloqueDetRetri.Hoja, "F3"));
        Assert.NotEqual("R10 - Recaudo Consolidado", TextoDeAlmacenado(caso.Ruta, caso.BloqueDetRetri.Hoja, "F3"));
    }

    /// <summary>
    /// T2.2 / D-H2.4: un R10 con un monto en blanco (celda vacía en un slot de monto) es fail-fast
    /// ANTES de escribir (<c>ERR-FORMATO-FUENTE</c> que nombra la celda): el pegado nunca inventa un
    /// <c>0</c> ni escribe medio bloque.
    /// </summary>
    [Fact]
    public void T2_2_R10ConMontoEnBlanco_FallaSinEscrituraParcial()
    {
        var periodo = Periodo.Parse(CodigoPeriodo);
        using var r10 = CopiaTemporalDelR10.Nueva();
        QuitarCelda(r10.Ruta, NombresHojaPeriodo.DetValiRetri(periodo), "D9");

        var excepcion = Assert.Throws<CalculoInvalidoException>(
            () => new ExcelDataReaderDetRetriR10Reader().LeerBloqueDetValiRetri(periodo, r10.Ruta));

        Assert.Equal(CodigoError.FormatoFuente, excepcion.Codigo);
        Assert.Contains("D9", excepcion.Message, StringComparison.Ordinal);
        Assert.Contains(NombresHojaPeriodo.DetValiRetri(periodo), excepcion.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// T2.2 / N-5: el monto del bloque se escribe SIN redondear (el entero es la regla de escritura
    /// del CONSOLIDADO, no del pegado). Un monto con decimales se lee de vuelta con sus centavos.
    /// </summary>
    [Fact]
    public void T2_2_ElPegadoNoRedondeaElMonto()
    {
        using var caso = CasoPegado.Nuevo();
        var bloqueConDecimales = ConCelda(caso.BloqueDetRetri, 9, "D", 39_148_067_546.64m);

        caso.Pegar(bloqueConDecimales, caso.BloqueValiRetri);

        using var documento = SpreadsheetDocument.Open(caso.Ruta, false);
        var valor = NumeroDe(documento.WorkbookPart!, caso.BloqueDetRetri.Hoja, 9, "D");

        Assert.Equal(39_148_067_546.64m, valor);
        Assert.NotEqual(39_148_067_547m, valor);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // N-1 (§12) — hardening del signo: el '-' SOLO vale en posición inicial
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// N-1 (Plan 62 §12, hallazgo #2 del review de H1): el signo <c>-</c> solo es válido en POSICIÓN
    /// INICIAL (al frente o detrás de <c>$</c>). Un texto con guiones laterales — separadores de
    /// fecha, prefijos de archivo, rangos — NO es un monto: se rechaza, jamás se adivina el negativo.
    /// Antes bastaba un <c>Contains('-')</c>, así que <c>"01-09"</c> se leía como <c>-109</c> (monto
    /// inventado) y <c>"01/09/2026"</c> como <c>-1092026</c>.
    /// </summary>
    [Theory]
    [InlineData("01/09/2026")]
    [InlineData("15/09/2026")]
    [InlineData("R10-2026091")]
    [InlineData("01-09")]
    [InlineData("01-09-2026")]
    [InlineData("-$ 5-093")]
    [InlineData("$5.093-257")]
    [InlineData("1.234,56-")]
    public void T2_N1_ParseoMonetarioR10_RechazaGuionesLateralesOFormatosNoMonetarios(string texto)
    {
        Assert.False(ParseoMonetarioR10.TryParseTexto(texto, out _), $"'{texto}' no es un monto y debe rechazarse.");
        Assert.Equal(0m, ParseoMonetarioR10.Convertir(texto)); // modo histórico del oráculo: 0, nunca un valor adivinado.
    }

    /// <summary>
    /// N-1: el signo inicial sigue funcionando en las dos formas del R10 (<c>"-$…"</c> y <c>"$-…"</c>)
    /// y la serie de los 10 casos del fixture de H1 queda intacta (aquí los 2 casos con signo, como
    /// anclaje de la no-regresión).
    /// </summary>
    [Theory]
    [InlineData("-$5.093.259.797", "-5093259797")]
    [InlineData("$-5.093.259.797", "-5093259797")]
    [InlineData("-$0", "0")]
    [InlineData("$9.147.179.378,75", "9147179378.75")]
    public void T2_N1_ParseoMonetarioR10_SignoInicialSigueValiendo(string texto, string esperado)
    {
        Assert.True(ParseoMonetarioR10.TryParseTexto(texto, out var numero), $"'{texto}' debería ser legible.");
        Assert.Equal(decimal.Parse(esperado, CultureInfo.InvariantCulture), numero);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T2.3 — Read-back: divergencia nombra hoja + celda + esperado + escrito; B9:B13 forma-agnóstico
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T2.3: el read-back de producción pasa con un pegado fiel, y ENROJECE nombrando hoja + celda +
    /// esperado + escrito cuando alguien altera una celda después del pegado (nunca se llega a los
    /// gates de validación con un pegado corrupto).
    /// </summary>
    [Fact]
    public void T2_3_ReadBack_EnrojeceNombrandoHojaCeldaEsperadoYEscrito()
    {
        using var caso = CasoPegado.Nuevo();
        caso.Pegar();

        // Control: el pegado fiel pasa.
        caso.VerificarIntegridad();

        var esperado = NumeroDe(caso.Ruta, caso.BloqueDetRetri.Hoja, 9, "D");
        CorromperNumero(caso.Ruta, caso.BloqueDetRetri.Hoja, "D9", esperado + 1m);

        var excepcion = Assert.Throws<CalculoInvalidoException>(() => caso.VerificarIntegridad());

        Assert.Equal(CodigoError.Escritura, excepcion.Codigo);
        Assert.Contains(caso.BloqueDetRetri.Hoja, excepcion.Message, StringComparison.Ordinal);
        Assert.Contains("D9", excepcion.Message, StringComparison.Ordinal);
        Assert.Contains(esperado.ToString("0.############################", CultureInfo.InvariantCulture), excepcion.Message, StringComparison.Ordinal);
        Assert.Contains((esperado + 1m).ToString("0.############################", CultureInfo.InvariantCulture), excepcion.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// T2.3 / H3-SEMANTIC: el read-back compara por VALOR semántico, no por forma de celda. Si un R10
    /// futuro trae el id del ASE como número, el pegado escribe número y la verificación sigue verde;
    /// el rótulo (<c>C9:C13</c>) sigue exigiéndose exacto.
    /// </summary>
    [Fact]
    public void T2_3_ReadBack_FormaAgnosticoEnLaColumnaDelAse()
    {
        using var caso = CasoPegado.Nuevo();
        caso.Pegar();

        // Se simula el R10 futuro: B9/B10 numéricos (el pegado los escribió como texto, porque hoy
        // el R10 los trae como texto).
        ConvertirEnNumero(caso.Ruta, caso.BloqueDetRetri.Hoja, "B9", 1m);
        ConvertirEnNumero(caso.Ruta, caso.BloqueDetRetri.Hoja, "B10", 2m);

        caso.VerificarIntegridad(); // sigue verde: mismo valor semántico (H3-SEMANTIC).
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T2.1 — El orquestador aporta los bloques: el entregable de septiembre-1 trae el R10
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T2.1 (D-ORDEN): la corrida real de septiembre-1 (ProcesadorPeriodo, 5 ASE) entrega las hojas
    /// DetRetri/DetValiRetri IGUALES al R10 — prueba de que el pegado corre dentro de la pasada del
    /// writer, después del CONSOLIDADO y de la Phase 0, y antes de los gates (que leen lo escrito).
    /// Se verifican los anchors por ASE y el sello del período; el caso del peso del ASE 3 (Δ = 1
    /// peso entre el cálculo del dominio y el R10, errata humana documentada) NO reaparece en la
    /// hoja: la hoja es el R10 por construcción (D-61R10).
    /// </summary>
    [Fact]
    public void T2_1_Septiembre1_ElEntregableTraeElR10Pegado()
    {
        var (resultado, _) = EspecialesR4H2Plan61Tests.Generacion2026091.Value;
        var ruta = resultado.RutaSalida;
        Assert.True(File.Exists(ruta), $"La corrida de septiembre-1 debía dejar el Excel: {ruta}");

        var periodo = Periodo.Parse(CodigoPeriodo);
        var lector = new ExcelDataReaderDetRetriR10Reader();
        var bloqueDet = lector.LeerBloqueDetRetri(periodo, RutaR10());
        var bloqueVali = lector.LeerBloqueDetValiRetri(periodo, RutaR10());

        using var documento = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart nulo.");

        AfirmarBloquePegado(workbookPart, bloqueDet, 3, 19);
        AfirmarBloquePegado(workbookPart, bloqueVali, 3, 14);

        // La hoja entregada es el R10 por construccion (D-61R10). Discriminadores medidos frente al
        // template: la serie AJUSTE A LA DECENA de DetValiRetri (J9:J14, template en 0) y el sello de
        // proceso (D6/D7), que en la base del periodo no coinciden con los del R10.
        Assert.Equal(-46_493m, NumeroDe(workbookPart, bloqueVali.Hoja, 9, "J"));
        Assert.Equal(10_891_338_739m, NumeroDe(workbookPart, bloqueDet.Hoja, 11, "D"));
        Assert.Equal("17/09/2026", TextoDe(workbookPart, bloqueDet.Hoja, "D6"));
        Assert.Equal("08:44 AM", TextoDe(workbookPart, bloqueDet.Hoja, "D7"));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Caso compartido: copia preparada de la plantilla oficial + bloques del R10 real
    // ════════════════════════════════════════════════════════════════════════════════════════════

    private sealed class CasoPegado : IDisposable
    {
        private readonly string _carpeta;
        private readonly Periodo _periodo;

        private CasoPegado(string carpeta, string ruta, Periodo periodo, BloqueHojaR10 bloqueDetRetri, BloqueHojaR10 bloqueValiRetri)
        {
            _carpeta = carpeta;
            _periodo = periodo;
            Ruta = ruta;
            BloqueDetRetri = bloqueDetRetri;
            BloqueValiRetri = bloqueValiRetri;
        }

        public string Ruta { get; }

        public BloqueHojaR10 BloqueDetRetri { get; }

        public BloqueHojaR10 BloqueValiRetri { get; }

        /// <summary>
        /// Copia FRESCA de la base preparada para el período (Phase 0 ejecutada una vez por proceso):
        /// cada caso muta SU copia y el insumo/plantilla quedan intactos.
        /// </summary>
        public static CasoPegado Nuevo()
        {
            var periodo = Periodo.Parse(CodigoPeriodo);
            var carpeta = Path.Combine(
                Path.GetTempPath(), "remuneracion-p62-h2-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(carpeta);
            var ruta = Path.Combine(carpeta, $"Remuneracion 202609-1 (pegado).xlsx");
            File.Copy(BasePreparadaDelCaso.Value, ruta, overwrite: false);

            var lector = new ExcelDataReaderDetRetriR10Reader();
            return new CasoPegado(
                carpeta,
                ruta,
                periodo,
                lector.LeerBloqueDetRetri(periodo, RutaR10()),
                lector.LeerBloqueDetValiRetri(periodo, RutaR10()));
        }

        /// <summary>Ejecuta el punto único de pegado del writer sobre la copia de trabajo.</summary>
        public void Pegar(BloqueHojaR10? bloqueDetRetri = null, BloqueHojaR10? bloqueValiRetri = null)
        {
            using var documento = SpreadsheetDocument.Open(Ruta, true);
            OpenXmlPlantillaWriter.PegarBloqueR10(
                documento.WorkbookPart!,
                _periodo,
                bloqueDetRetri ?? BloqueDetRetri,
                bloqueValiRetri ?? BloqueValiRetri);
            documento.WorkbookPart!.Workbook!.Save();
        }

        /// <summary>Ejecuta el read-back de integridad del writer (T2.3) sobre los bytes en disco.</summary>
        public void VerificarIntegridad()
        {
            using var documento = SpreadsheetDocument.Open(Ruta, false);
            OpenXmlPlantillaWriter.VerificarIntegridadPegadoR10(
                documento.WorkbookPart!, _periodo, BloqueDetRetri, BloqueValiRetri);
        }

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
                // Limpieza best-effort (temp).
            }
        }
    }

    /// <summary>
    /// Base del período preparada por la Phase 0 de PRODUCCIÓN desde la plantilla oficial (gate de
    /// entrada del plan), UNA vez por proceso. Los casos copian de aquí y nunca la mutan.
    /// </summary>
    private static readonly Lazy<string> BasePreparadaDelCaso = new(() =>
    {
        var carpeta = Path.Combine(Path.GetTempPath(), "remuneracion-p62-h2-base-" + Environment.ProcessId);
        Directory.CreateDirectory(carpeta);
        var salida = Path.Combine(carpeta, "Base_Remuneracion_2026091.xlsx");
        if (!File.Exists(salida))
        {
            PreparadorBasePipeline.Ejecutar(RutaPlantillaOficial(), Periodo.Parse(CodigoPeriodo), salida);
        }

        return salida;
    });

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Assert del read-back independiente (no usa el verificador de producción)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    private static void AfirmarBloquePegado(WorkbookPart workbookPart, BloqueHojaR10 bloque, int filaDesde, int filaHasta)
    {
        var hoja = Hoja(workbookPart, bloque.Hoja);

        for (var fila = filaDesde; fila <= filaHasta; fila++)
        {
            foreach (var columna in ColumnasDelBloque)
            {
                var origen = bloque.Celda(fila, columna)
                    ?? throw new InvalidOperationException($"El bloque '{bloque.Hoja}' no trae {columna}{fila}.");
                var destino = Celda(hoja, origen.Referencia);

                if (origen.Vacia)
                {
                    Assert.True(
                        destino is null || string.IsNullOrEmpty(destino.CellValue?.InnerText) && string.IsNullOrEmpty(destino.InlineString?.InnerText),
                        $"{bloque.Hoja}!{origen.Referencia} debía quedar vacía.");
                    continue;
                }

                Assert.NotNull(destino);
                Assert.Null(destino!.CellFormula);

                if (origen.Numerico is { } monto)
                {
                    Assert.InRange(NumeroDe(destino), monto - Tolerancia, monto + Tolerancia);
                }
                else
                {
                    Assert.Equal(origen.Texto, TextoDe(destino, workbookPart).Trim());
                }
            }
        }
    }

    private static decimal NumeroDe(WorkbookPart workbookPart, string hoja, int fila, string columna) =>
        NumeroDe(Celda(Hoja(workbookPart, hoja), $"{columna}{fila}")
            ?? throw new InvalidOperationException($"Falta {hoja}!{columna}{fila}."));

    private static decimal NumeroDe(string ruta, string hoja, int fila, string columna)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        return NumeroDe(documento.WorkbookPart!, hoja, fila, columna);
    }

    private static decimal NumeroDe(Cell celda) =>
        decimal.Parse(celda.CellValue?.InnerText ?? "0", NumberStyles.Any, CultureInfo.InvariantCulture);

    private static string TextoDe(WorkbookPart workbookPart, string hoja, string celda) =>
        TextoDe(Celda(Hoja(workbookPart, hoja), celda)
            ?? throw new InvalidOperationException($"Falta {hoja}!{celda}."), workbookPart);

    private static string TextoDe(Cell celda, WorkbookPart workbookPart)
    {
        if (celda.DataType?.Value == CellValues.InlineString)
        {
            return celda.InlineString?.InnerText ?? string.Empty;
        }

        if (celda.DataType?.Value == CellValues.SharedString)
        {
            var tabla = workbookPart.SharedStringTablePart?.SharedStringTable;
            var indice = int.Parse(celda.CellValue?.InnerText ?? "-1", CultureInfo.InvariantCulture);
            return tabla is not null && indice >= 0 && indice < tabla.Count()
                ? tabla.ElementAt(indice).InnerText
                : string.Empty;
        }

        return celda.CellValue?.InnerText ?? string.Empty;
    }

    /// <summary>Texto almacenado tal cual (sin resolver shared strings): sirve para comparar antes/después.</summary>
    private static string TextoDeAlmacenado(string ruta, string hoja, string celda)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var destino = Celda(Hoja(documento.WorkbookPart!, hoja), celda);
        return destino?.InlineString?.InnerText ?? destino?.CellValue?.InnerText ?? string.Empty;
    }

    /// <summary>Mapa (hoja!celda → texto + t/ref/si) de TODAS las fórmulas de las hojas del período.</summary>
    private static Dictionary<string, string> FormulasDe(string ruta)
    {
        var periodo = Periodo.Parse(CodigoPeriodo);
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = documento.WorkbookPart!;
        var mapa = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (var hoja in new[] { NombresHojaPeriodo.DetRetri(periodo), NombresHojaPeriodo.DetValiRetri(periodo) })
        {
            foreach (var celda in Hoja(workbookPart, hoja).Descendants<Cell>().Where(c => c.CellFormula is not null))
            {
                mapa[$"{hoja}!{celda.CellReference!.Value}"] =
                    $"{celda.CellFormula!.Text}|{celda.CellFormula.FormulaType?.Value}|{celda.CellFormula.Reference?.Value}|{celda.CellFormula.SharedIndex?.Value}";
            }
        }

        return mapa;
    }

    private static uint? EstiloDe(string ruta, string hoja, string celda)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        return Celda(Hoja(documento.WorkbookPart!, hoja), celda)?.StyleIndex?.Value;
    }

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"El workbook no trae la hoja '{nombre}'.");
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet.");
    }

    private static Cell? Celda(Worksheet worksheet, string referencia) =>
        worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // Mutaciones de fixture (siempre sobre COPIA temporal)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    /// <summary>Devuelve el bloque con UNA celda reemplazada por un monto exacto (para N-5).</summary>
    private static BloqueHojaR10 ConCelda(BloqueHojaR10 bloque, int fila, string columna, decimal monto)
    {
        var celdas = bloque.Celdas
            .Select(c => c.Fila == fila && string.Equals(c.Columna, columna, StringComparison.OrdinalIgnoreCase)
                ? new CeldaBloqueR10(c.Fila, c.Columna, null, monto)
                : c)
            .ToList();
        return new BloqueHojaR10(bloque.Hoja, bloque.Periodo, celdas);
    }

    private static void InjertarFormula(string ruta, string hoja, string referencia, string texto)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var destino = Celda(Hoja(documento.WorkbookPart!, hoja), referencia)
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no trae '{referencia}'.");
        destino.CellValue = null;
        destino.InlineString = null;
        destino.DataType = null;
        destino.CellFormula = new CellFormula(texto);
        documento.WorkbookPart!.Workbook!.Save();
    }

    private static void CorromperNumero(string ruta, string hoja, string referencia, decimal valor)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        ConvertirEnNumero(Celda(Hoja(documento.WorkbookPart!, hoja), referencia)
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no trae '{referencia}'."), valor);
        documento.WorkbookPart!.Workbook!.Save();
    }

    private static void ConvertirEnNumero(string ruta, string hoja, string referencia, decimal valor)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        ConvertirEnNumero(Celda(Hoja(documento.WorkbookPart!, hoja), referencia)
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no trae '{referencia}'."), valor);
        documento.WorkbookPart!.Workbook!.Save();
    }

    private static void ConvertirEnNumero(Cell celda, decimal valor)
    {
        celda.InlineString = null;
        celda.DataType = CellValues.Number;
        celda.CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture));
    }

    private static void QuitarCelda(string ruta, string nombreHoja, string referencia)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var celda = Celda(Hoja(documento.WorkbookPart!, nombreHoja), referencia)
            ?? throw new InvalidOperationException($"La hoja '{nombreHoja}' no trae '{referencia}'.");
        celda.Remove();
        documento.WorkbookPart!.Workbook!.Save();
    }

    /// <summary>Copia temporal del R10 real: los casos negativos mutan la copia, nunca el insumo.</summary>
    private sealed class CopiaTemporalDelR10 : IDisposable
    {
        private readonly string _carpeta;

        private CopiaTemporalDelR10(string carpeta, string ruta)
        {
            _carpeta = carpeta;
            Ruta = ruta;
        }

        public string Ruta { get; }

        public static CopiaTemporalDelR10 Nueva()
        {
            var carpeta = Path.Combine(Path.GetTempPath(), "remuneracion-p62-h2-r10-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(carpeta);
            var ruta = Path.Combine(carpeta, Path.GetFileName(RutaR10()));
            File.Copy(RutaR10(), ruta, overwrite: false);
            return new CopiaTemporalDelR10(carpeta, ruta);
        }

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
                // Limpieza best-effort (temp).
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────────────────────────
    // Rutas EXPLÍCITAS (Plan 62 §gate_entrada / §T0: nada se resuelve por enumeración)
    // ─────────────────────────────────────────────────────────────────────────────────────────────

    private static string RutaR10() => Path.Combine(
        Raiz(), "Docs", "Prueba Septiembre 1", "Insumos", "R10_Remuneracion_2026091.xlsx");

    private static string RutaPlantillaOficial() => Path.Combine(
        Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }
}
