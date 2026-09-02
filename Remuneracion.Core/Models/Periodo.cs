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

    /// <summary>
    /// Constructor útil para parsear el código de periodo del formulario: "2026071" → AAAAMM=202607, quincena=1.
    /// </summary>
    public static Periodo Parse(string codigoCompleto)
    {
        if (string.IsNullOrWhiteSpace(codigoCompleto) || codigoCompleto.Length < 7)
        {
            throw new ArgumentException("El código de periodo debe incluir AAAAMM + quincena.", nameof(codigoCompleto));
        }

        var codigoAaaamm = codigoCompleto.Substring(0, 6);
        var quincena = codigoCompleto[^1..];

        if (!int.TryParse(quincena, out var numeroQuincena) || (numeroQuincena is not 1 and not 2))
        {
            throw new ArgumentException("La quincena debe ser 1 o 2.", nameof(codigoCompleto));
        }

        return new Periodo
        {
            CodigoAAAAMM = codigoAaaamm,
            NumeroQuincena = numeroQuincena
        };
    }
}
