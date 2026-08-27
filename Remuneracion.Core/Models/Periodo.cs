namespace Remuneracion.Core.Models;

/// <summary>
/// Representa el periodo de liquidación quincenal de la remuneración UAESP.
/// </summary>
public class Periodo
{
    /// <summary>
    /// Código del mes en formato AAAAMM (por ejemplo "202605").
    /// </summary>
    public string CodigoAAAAMM { get; set; } = string.Empty;

    /// <summary>
    /// Número de quincena dentro del mes: 1 (primera) o 2 (segunda).
    /// </summary>
    public int NumeroQuincena { get; set; }

    /// <summary>
    /// Fecha de inicio del periodo quincenal.
    /// </summary>
    public DateTime FechaInicio { get; set; }

    /// <summary>
    /// Fecha de fin del periodo quincenal.
    /// </summary>
    public DateTime FechaFin { get; set; }

    /// <summary>
    /// Código completo del periodo, resultado de concatenar <see cref="CodigoAAAAMM"/> y
    /// <see cref="NumeroQuincena"/> (por ejemplo "2026051").
    /// </summary>
    public string CodigoCompleto => $"{CodigoAAAAMM}{NumeroQuincena}";

    /// <summary>
    /// Nombre sugerido para el archivo de salida de la remuneración total.
    /// </summary>
    public string NombreArchivo => $"Remuneración {CodigoAAAAMM}-{NumeroQuincena} Total.xlsx";
}
