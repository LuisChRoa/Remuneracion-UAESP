using System.Globalization;
using System.Text.RegularExpressions;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Services;

/// <summary>
/// Plan 61 (H3.2 / S3.2, §2-D-H3): traduce los gates de DATOS a <see cref="NovedadEntrega"/> en
/// lenguaje de NEGOCIO. La clasificación es por CLASE (toda divergencia de datos con workbook
/// escribible es novedad), nunca por caso: no hay ramas por ASE, período, quincena ni letra de
/// columna. Los textos visibles omiten la jerga del catálogo de errores (<c>ERR-*</c>,
/// <c>&lt;f&gt;</c>, nombres de clase).
///
/// <para>Los montos se extraen de los mensajes de los gates cuando el mensaje los expone en la forma
/// <c>(observado) … (esperado)</c>, que es la convención de todos los gates de coherencia. Si no hay
/// dos montos comparables se dejan en <c>null</c>: el modelo NO inventa ceros.</para>
/// </summary>
internal static class ClasificadorNovedadesEntrega
{
    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";

    private const string DuenoAplicacion = "Equipo de la aplicación (control interno del cálculo)";

    private const string DuenoAdministrativo = "Área administrativa (elabora el reporte oficial)";

    private static readonly Regex PatronAse = new(@"ASE\s+(\d+)", RegexOptions.CultureInvariant);

    private static readonly Regex PatronPrefijoError = new(@"^\[[A-Z]+(?:-[A-Z]+)*\]\s*", RegexOptions.CultureInvariant);

    private static readonly Regex PatronMonto = new(@"\(([-+]?\d[\d.,]*)\)", RegexOptions.CultureInvariant);

    /// <summary>
    /// Gates de coherencia del período (§2.5): Σ visibles-empresa vs bloque, cifras del ASE frente
    /// al consolidado, ajustes, banco, balance. Cada error del validador es UNA ficha.
    /// </summary>
    internal static IReadOnlyList<NovedadEntrega> DesdeCoherencia(IReadOnlyList<string> errores, Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(errores);
        ArgumentNullException.ThrowIfNull(periodo);

        var novedades = new List<NovedadEntrega>(errores.Count);
        foreach (var error in errores)
        {
            novedades.Add(DesdeMensaje(error, periodo, esValidacionCruzada: false));
        }

        return novedades;
    }

    /// <summary>
    /// Gate DetRetri calculado vs R10 (frente per-ASE). Es la ficha con más detalle útil, así que se
    /// construye con los números del dominio (no con el texto del gate): el administrativo ve el
    /// valor de la plantilla y el del reporte oficial del período.
    ///
    /// <para><b>D8 (ratificada).</b> Cuando el R10 y la cadena de fórmulas de la plantilla
    /// discrepan, manda la plantilla: la diferencia se reporta y JAMÁS se corrige el cálculo ni se
    /// amplía la tolerancia.</para>
    ///
    /// <para><b>Plan 62 (T3.2, D-61R10) — sin novedad de TOTAL.</b> La capa de TOTAL
    /// (Σ de los <c>Detalle</c> vs <c>r10.Total</c>) dejó de emitirse: la hoja DetRetri es el R10
    /// por construcción (pegado N1), así que esa ficha sería una tautología y su residuo de
    /// redondeo vive donde ya lo tiene el manual (<c>DetValiRetri!D21</c>, D6). Solo quedan las
    /// fichas per-ASE, que siguen probando que el cálculo de la plantilla cuadra con el reporte
    /// oficial de cada operador.</para>
    /// </summary>
    internal static IReadOnlyList<NovedadEntrega> DesdeR10(
        IReadOnlyList<WorkbookLeafInputs> leafs,
        DetRetriInputs r10,
        Periodo periodo,
        string nombreArchivoR10)
    {
        ArgumentNullException.ThrowIfNull(leafs);
        ArgumentNullException.ThrowIfNull(r10);
        ArgumentNullException.ThrowIfNull(periodo);

        var hoja = NombresHojaPeriodo.DetRetri(periodo.CodigoCompleto);
        var reporte = $"R10 del período ({nombreArchivoR10})";
        var novedades = new List<NovedadEntrega>();

        var conDetalle = leafs.Where(l => l.DetRetriQ2 is not null).OrderBy(l => l.Ase.Id).ToList();

        foreach (var leaf in conDetalle)
        {
            var calculado = leaf.DetRetriQ2!.Detalle;
            if (!r10.DetRetriPorAse.TryGetValue(leaf.Ase.Id, out var oraculo))
            {
                novedades.Add(new NovedadEntrega(
                    reporte,
                    hoja,
                    leaf.Ase.Id,
                    "Fila del ASE en el reporte oficial",
                    calculado,
                    null,
                    null,
                    DuenoAdministrativo,
                    "Verifique que el R10 del período incluya la fila de este ASE; el reporte entregado no la trae y por eso no se pudo comparar.",
                    $"El reporte oficial del período no trae el valor de este ASE; la plantilla calcula {calculado:0}."));
                continue;
            }

            var diferencia = Math.Abs(calculado - oraculo);
            if (diferencia <= 0.5m)
            {
                continue;
            }

            var origen = leaf.DetRetriQ2.TotalD104;
            novedades.Add(new NovedadEntrega(
                reporte,
                hoja,
                leaf.Ase.Id,
                "Retribución total del ASE (asiento del período)",
                calculado,
                oraculo,
                calculado - oraculo,
                DuenoAdministrativo,
                "Compare, por la etiqueta del ASE, el valor del asiento en el Excel entregado contra el del reporte oficial del período. La plantilla manda: si la diferencia es de transcripción del reporte oficial, corríjala allí y vuelva a ejecutar.",
                ExplicarDiferenciaDelAsiento(origen, calculado, oraculo)));
        }

        // Plan 62 (T3.2, D-61R10): NO se emite la ficha de "Total del período". Con el pegado (N1)
        // el total entregado es el del R10 por construcción; repetirlo sería un verde que no prueba
        // nada. El residuo de composición vive en DetValiRetri!D21 (divergencia D6 existente).
        return novedades;
    }

    /// <summary>
    /// Validación cruzada 2.7: sub-bloques FALSE no declarados (fichas estructuradas, con celda y
    /// montos) + los errores de esa validación que no son sub-bloques. Los errores del gate §2.5 se
    /// excluyen porque ya vienen de <see cref="DesdeCoherencia"/> (la clasificación 2.7 los incluye
    /// al reutilizar el validador completo).
    /// </summary>
    internal static IReadOnlyList<NovedadEntrega> DesdeValidacionCruzada(
        ClasificacionDivergenciasDeclaradas clasificacion,
        IReadOnlyList<string> erroresCoherencia,
        Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(clasificacion);
        ArgumentNullException.ThrowIfNull(erroresCoherencia);
        ArgumentNullException.ThrowIfNull(periodo);

        var novedades = new List<NovedadEntrega>();

        foreach (var veredicto in clasificacion.NoDeclaradas)
        {
            novedades.Add(new NovedadEntrega(
                "Validación cruzada de la salida (VALIDACION_TOTAL)",
                "VALIDACION_TOTAL",
                veredicto.AseId,
                $"Bloque {veredicto.Celda} de la hoja de validación",
                veredicto.MontoObservado,
                veredicto.MontoEsperado,
                veredicto.MontoObservado - veredicto.MontoEsperado,
                DuenoAplicacion,
                "Abra la hoja VALIDACION_TOTAL del Excel entregado, ubique el bloque indicado y compare los dos montos que lo componen; la diferencia debe explicarse antes de dar la corrida por buena.",
                "Un control interno de la salida quedó en falso: los dos montos que ese bloque iguala no coinciden."));
        }

        foreach (var error in clasificacion.ErroresNoDeclarados)
        {
            if (erroresCoherencia.Contains(error, StringComparer.Ordinal))
            {
                continue;
            }

            novedades.Add(DesdeMensaje(error, periodo, esValidacionCruzada: true));
        }

        return novedades;
    }

    /// <summary>
    /// Divergencias DECLARADAS pendientes de reconocimiento (flujo Plan 42): la corrida entrega el
    /// Excel y el reporte, pero no se certifica hasta que el usuario reconozca la diferencia.
    /// </summary>
    internal static IReadOnlyList<NovedadEntrega> DesdeDeclaradasPendientes(
        ClasificacionDivergenciasDeclaradas clasificacion)
    {
        ArgumentNullException.ThrowIfNull(clasificacion);

        var novedades = new List<NovedadEntrega>();
        foreach (var veredicto in clasificacion.Declaradas
                     .GroupBy(v => (v.EntradaId, v.Celda))
                     .Select(g => g.First()))
        {
            var entrada = BuscarEntrada(veredicto.EntradaId);
            novedades.Add(new NovedadEntrega(
                "Divergencia declarada de la validación cruzada",
                "VALIDACION_TOTAL",
                veredicto.AseId,
                $"Bloque {veredicto.Celda} de la hoja de validación",
                veredicto.MontoObservado,
                veredicto.MontoEsperado,
                veredicto.MontoObservado - veredicto.MontoEsperado,
                entrada?.Dueno ?? "Responsable externo por confirmar",
                "Esta diferencia ya está declarada y tiene responsable asignado, pero exige reconocimiento explícito en esta corrida para cerrar. Sin ese reconocimiento la corrida no se certifica.",
                "La diferencia está declarada en el registro de divergencias conocidas; no es un error nuevo del cálculo."));
        }

        return novedades;
    }

    private static NovedadEntrega DesdeMensaje(string error, Periodo periodo, bool esValidacionCruzada)
    {
        var texto = PatronPrefijoError.Replace(error, string.Empty).Trim();
        var aseId = 0;
        var coincidenciaAse = PatronAse.Match(texto);
        if (coincidenciaAse.Success)
        {
            _ = int.TryParse(coincidenciaAse.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out aseId);
        }

        var (reporte, hoja, concepto, dueno, queRevisar) = Familia(texto, periodo, esValidacionCruzada);
        var montos = ExtraerMontos(texto);

        decimal? observado = montos.Count >= 2 ? montos[0] : null;
        decimal? esperado = montos.Count >= 2 ? montos[1] : null;
        decimal? delta = observado is not null && esperado is not null ? observado - esperado : null;

        return new NovedadEntrega(
            reporte,
            hoja,
            aseId,
            concepto,
            observado,
            esperado,
            delta,
            dueno,
            queRevisar,
            texto);
    }

    private static (string Reporte, string Hoja, string Concepto, string Dueno, string QueRevisar) Familia(
        string texto, Periodo periodo, bool esValidacionCruzada)
    {
        if (esValidacionCruzada || texto.Contains("VALIDACION_TOTAL", StringComparison.OrdinalIgnoreCase))
        {
            return (
                "Validación cruzada de la salida (VALIDACION_TOTAL)",
                "VALIDACION_TOTAL",
                "Control de la hoja de validaciones cruzadas",
                DuenoAplicacion,
                "Revise la hoja VALIDACION_TOTAL del Excel entregado: el bloque señalado compara dos cifras que deben coincidir.");
        }

        if (texto.Contains("visibles de empresas", StringComparison.OrdinalIgnoreCase))
        {
            var familia = texto.Contains(" R4", StringComparison.Ordinal) || texto.Contains("en R4", StringComparison.Ordinal)
                ? "Reversión por componente (R4)"
                : texto.Contains(" R2", StringComparison.Ordinal) || texto.Contains("en R2", StringComparison.Ordinal)
                    ? "Reporte detalle saldos a favor (R2)"
                    : "Recaudo por componente (R1)";

            return (
                familia,
                HojaConsolidado,
                "Suma de las empresas frente al total del ASE",
                "Área administrativa y equipo de la aplicación (conciliación por empresa)",
                "Compare, empresa por empresa y por su etiqueta, la cifra del reporte contra el total del ASE que muestra el Excel entregado. El detalle de las empresas acompañante esta ficha.");
        }

        if (texto.Contains("AJUSTES-SF-T", StringComparison.OrdinalIgnoreCase) || texto.Contains("AjustesSfT", StringComparison.OrdinalIgnoreCase))
        {
            return (
                "Ajustes de saldos a favor (SALDOS POR NOTA + RETRIBUCIÓN NEGATIVA)",
                "AJUSTES - SF-T",
                "Ajustes del ASE frente a su consolidado",
                DuenoAplicacion,
                "Revise la hoja de ajustes del Excel entregado y compare sus dos componentes (saldos por nota y retribución negativa) con el consolidado del ASE.");
        }

        if (texto.Contains("DetRetri", StringComparison.OrdinalIgnoreCase))
        {
            return (
                "Asiento de retribución del período",
                NombresHojaPeriodo.DetRetri(periodo.CodigoCompleto),
                "Asiento de retribución del ASE",
                DuenoAplicacion,
                "Compare el asiento del ASE en el Excel entregado contra su origen (la suma de sus cinco conceptos).");
        }

        if (texto.Contains("BSC", StringComparison.OrdinalIgnoreCase) || texto.Contains("Balance", StringComparison.OrdinalIgnoreCase))
        {
            return (
                "Balance de subsidios y contribuciones",
                "BCE SC POR FACT.",
                "Balance del ASE frente al total de la fuente",
                DuenoAplicacion,
                "Compare, por la etiqueta del ASE, el subsidio y la contribución del balance contra el total de su fuente.");
        }

        if (texto.Contains("banco", StringComparison.OrdinalIgnoreCase))
        {
            return (
                "Reporte de recaudo por banco",
                "REPORTE RECAUDO x BANCO",
                "Recaudo por banco del ASE",
                DuenoAplicacion,
                "Compare, empresa por empresa, los conceptos del reporte de banco contra el resumen de su fuente.");
        }

        // Plan 64 (§2.9): el gate de coherencia de ANT EXT-REV es de DATOS (novedad), y su ficha debe
        // nombrar la hoja REAL que el administrativo tiene que abrir. Sin esta familia la ficha caía en
        // el texto genérico y mandaba a revisar la hoja del consolidado, donde ese dinero no está
        // (G-TRAZA: reporte + hoja + ASE + montos comparados).
        if (texto.Contains("ANT EXT-REV", StringComparison.Ordinal))
        {
            return (
                "Recaudo reversado de anticipos (ANT EXT-REV)",
                "ANT EXT-REV",
                "Cierre de los bloques de fuente del ASE en la hoja de anticipos y recaudos reversados",
                DuenoAdministrativo,
                "Compare, por la etiqueta de la fuente, los cuatro conceptos del insumo con el total que cierra cada grupo y el total leído con lo escrito en la hoja ANT EXT-REV del Excel entregado; si el insumo está mal, corríjalo allí y vuelva a ejecutar.");
        }

        if (texto.Contains(" R4", StringComparison.Ordinal) || texto.Contains("Reversion", StringComparison.OrdinalIgnoreCase))
        {
            return ("Reversión por componente (R4)", HojaConsolidado, "Cifra del ASE frente a su consolidado", DuenoAplicacion,
                "Compare la cifra del ASE en el Excel entregado contra el consolidado del mismo ASE.");
        }

        if (texto.Contains(" R2", StringComparison.Ordinal) || texto.Contains("R2TotalOportuno", StringComparison.Ordinal))
        {
            return ("Reporte detalle saldos a favor (R2)", HojaConsolidado, "Cifra del ASE frente a su consolidado", DuenoAplicacion,
                "Compare la cifra del ASE en el Excel entregado contra el consolidado del mismo ASE.");
        }

        return (
            "Validación de coherencia del período",
            HojaConsolidado,
            "Cifra del período que no cuadra",
            DuenoAplicacion,
            "Revise el Excel entregado en la hoja señalada y compare la cifra indicada contra su origen; el detalle de esta ficha explica qué dos cifras debían coincidir.");
    }

    /// <summary>
    /// Extrae los montos de la forma <c>(observado) … (esperado)</c> que usan los gates de
    /// coherencia. Tolera separador de miles con punto o coma y descarta lo que no sea número.
    /// </summary>
    private static List<decimal> ExtraerMontos(string texto)
    {
        var montos = new List<decimal>();
        foreach (Match coincidencia in PatronMonto.Matches(texto))
        {
            var crudo = coincidencia.Groups[1].Value;
            if (TryParsearMonto(crudo, out var monto))
            {
                montos.Add(monto);
            }
        }

        // Solo interesan el observado y el esperado (primer par comparado).
        if (montos.Count > 2)
        {
            montos = montos.Take(2).ToList();
        }

        return montos;
    }

    private static bool TryParsearMonto(string crudo, out decimal monto)
    {
        var texto = crudo.Trim();
        var ultimaComa = texto.LastIndexOf(',');
        var ultimoPunto = texto.LastIndexOf('.');

        string normalizado;
        if (ultimaComa >= 0 && ultimoPunto >= 0)
        {
            // El separador decimal es el que aparece MÁS A LA DERECHA; el otro es de miles.
            normalizado = ultimaComa > ultimoPunto
                ? texto.Replace(".", string.Empty, StringComparison.Ordinal).Replace(',', '.')
                : texto.Replace(",", string.Empty, StringComparison.Ordinal);
        }
        else if (ultimaComa >= 0)
        {
            // Un solo tipo de separador: es de miles si se repite o si agrupa exactamente 3 dígitos.
            normalizado = ContarSeparadores(texto, ',') > 1 || texto.Length - ultimaComa - 1 == 3
                ? texto.Replace(",", string.Empty, StringComparison.Ordinal)
                : texto.Replace(',', '.');
        }
        else if (ultimoPunto >= 0)
        {
            normalizado = ContarSeparadores(texto, '.') > 1 || texto.Length - ultimoPunto - 1 == 3
                ? texto.Replace(".", string.Empty, StringComparison.Ordinal)
                : texto;
        }
        else
        {
            normalizado = texto;
        }

        return decimal.TryParse(normalizado, NumberStyles.Number, CultureInfo.InvariantCulture, out monto);
    }

    private static int ContarSeparadores(string texto, char separador)
    {
        var total = 0;
        foreach (var caracter in texto)
        {
            if (caracter == separador)
            {
                total++;
            }
        }

        return total;
    }

    /// <summary>
    /// Explica la diferencia del asiento en lenguaje de negocio, mostrando la cadena completa: el
    /// valor de la plantilla con sus centavos, el asiento redondeado y el valor del reporte oficial.
    /// Es la lectura que el administrativo necesita para decidir de quién es la diferencia (D8).
    /// </summary>
    private static string ExplicarDiferenciaDelAsiento(decimal origen, decimal redondeado, decimal oraculo)
    {
        var cadena = $"La cadena de cálculo de la plantilla da {FormatearMonto(origen)} y el asiento se redondea a {FormatearMonto(redondeado)}";
        var reporte = $"el reporte oficial trae {FormatearMonto(oraculo)}";
        var diferencia = redondeado - oraculo;
        var lectura = Math.Abs(diferencia) < 0.0000001m
            ? "no hay diferencia material"
            : diferencia < 0
                ? $"la diferencia de {FormatearMonto(Math.Abs(diferencia))} está en el reporte oficial"
                : $"la diferencia de {FormatearMonto(Math.Abs(diferencia))} está en el asiento de la plantilla";

        return $"{cadena}; {reporte}: {lectura}.";
    }

    /// <summary>Monto de negocio: miles con punto y decimales con coma (10.891.338.739,48).</summary>
    private static string FormatearMonto(decimal valor) => valor.ToString("#,##0.####", FormatoNegocio);

    private static readonly NumberFormatInfo FormatoNegocio = new()
    {
        NumberGroupSeparator = ".",
        NumberDecimalSeparator = ",",
        NumberGroupSizes = [3],
        NumberDecimalDigits = 4,
    };

    private static DivergenciasDeclaradasCatalogo.Entrada? BuscarEntrada(string? id)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            return null;
        }

        foreach (var entrada in DivergenciasDeclaradasCatalogo.Entradas)
        {
            if (string.Equals(entrada.Id, id, StringComparison.OrdinalIgnoreCase))
            {
                return entrada;
            }
        }

        return null;
    }
}
