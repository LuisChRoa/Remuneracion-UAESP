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
    /// Total del cierre del reporte (fila <c>A="Total"</c> con <c>B</c> vacío, componente
    /// <c>Total</c>). Observación FIEL de la fuente: se almacena tal como viene, negativo.
    /// </summary>
    public decimal TotalReversiones { get; set; }

    /// <summary>
    /// Componente <c>SERVICIO ESPECIALES</c> (alias admitido <c>Especiales</c>) de la MISMA fila de
    /// cierre, resuelto por ENCABEZADO — nunca por letra de columna. Plan 61 (D-MODELO ratificada):
    /// se agrega de forma ADITIVA para no alterar la semántica de <see cref="TotalReversiones"/>.
    /// <c>0</c> cuando la fuente no trae ese encabezado en el período/ASE (columna intermitente).
    /// El visible que consume la plantilla y el consolidado es <c>TotalReversiones −
    /// EspecialesReversion</c>.
    /// </summary>
    public decimal EspecialesReversion { get; set; }
}
