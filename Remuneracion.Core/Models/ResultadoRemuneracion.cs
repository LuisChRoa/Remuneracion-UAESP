using Remuneracion.Core.Models;

namespace Remuneracion.Core.Models;

/// <summary>
/// Resultado global de la liquidación de remuneración quincenal, con los consolidados
/// por ASE, el gran total y el estado del proceso.
/// </summary>
public class ResultadoRemuneracion
{
    /// <summary>
    /// Periodo quincenal al que corresponde la liquidación.
    /// </summary>
    public Periodo Periodo { get; set; } = new();

    /// <summary>
    /// Lista de consolidados, uno por ASE.
    /// </summary>
    public List<ConsolidadoAse> Consolidados { get; set; } = new();

    /// <summary>
    /// Gran total de la liquidación, suma de los totales de cada ASE.
    /// </summary>
    public decimal GranTotal => Consolidados.Sum(c => c.TotalAse);

    /// <summary>
    /// Indica si el proceso de liquidación finalizó sin errores bloqueantes.
    /// </summary>
    public bool Exitoso { get; set; }

    /// <summary>
    /// Mensajes de información, advertencia o error generados durante el proceso.
    /// </summary>
    public List<string> Mensajes { get; set; } = new();

    /// <summary>
    /// Marca de tiempo de la generación del resultado.
    /// </summary>
    public DateTime Timestamp { get; set; } = DateTime.Now;

    /// <summary>
    /// Plan 28 (Unidad F): fecha de inicio del rango del período, leída del R10 (celda G7 de
    /// <c>DetRetri{AAAAMMQ}</c>). La setea <c>ProcesadorPeriodo</c> tras leer el R10; el writer la
    /// sella en <c>CONSOLIDADO_TOTAL RECAUDO!G7</c>. <c>null</c> = no se sella (path single-ASE
    /// sin R10).
    /// </summary>
    public DateTime? FechaDesde { get; set; }

    /// <summary>
    /// Plan 28 (Unidad F): fecha de fin del rango del período, leída del R10 (celda J7). El writer
    /// la sella en <c>CONSOLIDADO_TOTAL RECAUDO!K7</c> (columna real fijada por T0; J7 de la
    /// plantilla es el rótulo). <c>null</c> = no se sella.
    /// </summary>
    public DateTime? FechaHasta { get; set; }

    /// <summary>
    /// Plan 31 (T3, R-S-1/R-S-2): "Fecha de Proceso" leída del encabezado del R10 (celda D6 de
    /// <c>DetRetri{AAAAMMQ}</c>), sello de la corrida administrativa. <c>null</c> = no se sella.
    ///
    /// <para>
    /// Plan 62 (T2.4, D-31SELLO): el sello <c>D6/D7</c> de <c>DetRetri*/DetValiRetri*</c> ya NO se
    /// escribe aparte — viaja pegado en el bloque del R10 (<see cref="BloqueDetRetriR10"/> /
    /// <see cref="BloqueValiRetriR10"/>). Este campo se conserva porque la LECTURA del R10 y su
    /// fail-fast <c>ERR-FORMATO-FUENTE</c> siguen vigentes (lo usa el orquestador para el sello de
    /// proceso del resultado), y porque la fecha de proceso identifica la corrida administrativa.
    /// </para>
    /// </summary>
    public DateTime? FechaProceso { get; set; }

    /// <summary>
    /// Plan 31 (T3, R-S-1/R-S-2): "Hora" del proceso leída del R10 (celda D7), hora-del-día del
    /// sello administrativo. <c>null</c> = no se sella. Misma doctrina que
    /// <see cref="FechaProceso"/> (Plan 62, T2.4).
    /// </summary>
    public TimeSpan? HoraProceso { get; set; }

    /// <summary>
    /// Plan 62 (N1, T2.1/T2.2): bloque COMPLETO del R10 para <c>DetRetri{AAAAMMQ}</c> (filas 3-19,
    /// columnas A-O) que el writer pega celda a celda, values-only, en la hoja homónima de la
    /// plantilla de procesamiento. Lo aporta el orquestador (<c>ProcesadorPeriodo</c>) desde el
    /// lector de bloque del R10; <c>null</c> = no se pega (path single-ASE sin R10).
    ///
    /// <para>
    /// El pegado corre DESPUÉS del cálculo del CONSOLIDADO y de la Phase 0 de la base, y ANTES de
    /// las lecturas-oráculo/gates, que así evalúan lo pegado (Plan 62 §2-D-H2.1 y §8 D-ORDEN).
    /// </para>
    /// </summary>
    public BloqueHojaR10? BloqueDetRetriR10 { get; set; }

    /// <summary>
    /// Plan 62 (N1, T2.1/T2.2): bloque COMPLETO del R10 para <c>DetValiRetri{AAAAMMQ}</c> (filas
    /// 3-14, columnas A-O). Misma doctrina que <see cref="BloqueDetRetriR10"/>.
    /// </summary>
    public BloqueHojaR10? BloqueValiRetriR10 { get; set; }
}
