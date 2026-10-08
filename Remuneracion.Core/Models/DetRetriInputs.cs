namespace Remuneracion.Core.Models;

/// <summary>
/// HU-20 (G3): lectura del oráculo <c>R10</c> por período. Representa la hoja
/// <c>DetRetri{AAAAMMQ}</c> del archivo <c>R10_Remuneracion_{AAAAMMQ}.xlsx</c>:
/// el DetRetri (entero) por ASE (<c>D9:D13</c>) y su total (<c>D14</c>), más el rango del
/// período (<c>G7</c>/<c>J7</c>).
///
/// Es un <b>oráculo de VALIDACIÓN</b> (G3-D1): el valor escrito en el template es el DetRetri
/// calculado bottom-up (HU-12, <see cref="DetRetriQ2Inputs"/>) y este modelo solo sirve para
/// contrastarlo con tolerancia ±0.5 post-redondeo. El R10 NUNCA se escribe al template.
/// </summary>
public sealed class DetRetriInputs
{
    /// <summary>
    /// Código de remuneración del R10 (AAAAMMQ; p. ej. "2026071"), leído de la celda M3.
    /// </summary>
    public string CodigoRemuneracion { get; set; } = string.Empty;

    /// <summary>
    /// Fecha de inicio del rango del R10 ("Fecha Desde", G7).
    /// </summary>
    public DateTime FechaDesde { get; set; }

    /// <summary>
    /// Fecha de fin del rango del R10 ("Fecha Hasta", J7).
    /// </summary>
    public DateTime FechaHasta { get; set; }

    /// <summary>
    /// Plan 31 (T3, R-S-1): "Fecha de Proceso" (D6) del encabezado del R10 — es el sello de la
    /// corrida administrativa que produjo el R10, no el rango del período (G7/J7). <c>null</c> si la
    /// celda no existe o no parsea.
    /// </summary>
    public DateTime? FechaProceso { get; set; }

    /// <summary>
    /// Plan 31 (T3, R-S-1): "Hora" (D7) del encabezado del R10 — hora-del-día del sello de proceso
    /// (fracción de día, p. ej. 07:42 → 0.3208…). <c>null</c> si no existe o no parsea.
    /// </summary>
    public TimeSpan? HoraProceso { get; set; }

    /// <summary>
    /// DetRetri por ASE (clave = <see cref="Ase.Id"/> 1..5) leído de D9:D13.
    /// </summary>
    public IReadOnlyDictionary<int, decimal> DetRetriPorAse { get; set; } = new Dictionary<int, decimal>();

    /// <summary>
    /// DetRetri total del período (D14) leído del R10.
    /// </summary>
    public decimal Total { get; set; }
}
