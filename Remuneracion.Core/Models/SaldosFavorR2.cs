namespace Remuneracion.Core.Models;

/// <summary>
/// Datos de saldos a favor extraídos del reporte R2 (Reporte detalle saldos a favor).
/// </summary>
public class SaldosFavorR2
{
    /// <summary>
    /// Nombre del ASE al que corresponden los datos.
    /// </summary>
    public string NombreAse { get; set; } = string.Empty;

    /// <summary>
    /// Gran total del reporte (columna E, última fila).
    /// </summary>
    public decimal GrandTotal { get; set; }

    /// <summary>
    /// Indica si el reporte contiene la columna "Especiales" (encabezado en columna 11).
    /// </summary>
    public bool TieneColumnaEspeciales { get; set; }

    /// <summary>
    /// Valor de servicios especiales (columna 11) cuando existe; 0 en caso contrario.
    /// </summary>
    public decimal ServEspK { get; set; }

    /// <summary>
    /// Total oportuno calculado como <see cref="GrandTotal"/> menos <see cref="ServEspK"/>.
    /// </summary>
    public decimal TotalOportuno => GrandTotal - ServEspK;
}
