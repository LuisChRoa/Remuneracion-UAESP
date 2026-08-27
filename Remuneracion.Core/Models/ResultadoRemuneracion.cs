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
}
