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
    /// <c>DetRetri{AAAAMMQ}</c>), sello de la corrida administrativa. El writer la propaga a
    /// <c>DetRetri*/DetValiRetri*!D6</c> (texto <c>dd/MM/yyyy</c>, estilo de celda preservado).
    /// <c>null</c> = no se sella (path single-ASE sin R10).
    /// </summary>
    public DateTime? FechaProceso { get; set; }

    /// <summary>
    /// Plan 31 (T3, R-S-1/R-S-2): "Hora" del proceso leída del R10 (celda D7), hora-del-día del
    /// sello administrativo. El writer la propaga a <c>DetRetri*/DetValiRetri*!D7</c> (texto
    /// <c>hh:mm AM/PM</c>). <c>null</c> = no se sella.
    /// </summary>
    public TimeSpan? HoraProceso { get; set; }
}
