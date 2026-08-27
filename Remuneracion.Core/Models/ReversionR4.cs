namespace Remuneracion.Core.Models;

/// <summary>
/// Datos de reversión extraídos del reporte R4 (Reversión por componente).
/// </summary>
public class ReversionR4
{
    /// <summary>
    /// Nombre del ASE al que corresponden los datos.
    /// </summary>
    public string NombreAse { get; set; } = string.Empty;

    /// <summary>
    /// Total de reversores (última fila, columna 4). Se almacena como valor negativo.
    /// </summary>
    public decimal TotalReversiones { get; set; }
}
