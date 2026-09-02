namespace Remuneracion.Core.Models;

/// <summary>
/// Inputs editables reales de la hoja <c>Reversion Pagos R4</c>.
/// La celda visible <c>D67</c> es una fórmula del workbook: <c>D9 - P9</c>.
/// </summary>
public sealed class WorkbookLeafInputsR4
{
    /// <summary>
    /// Valor de la celda leaf <c>D9</c>.
    /// </summary>
    public decimal D9 { get; set; }

    /// <summary>
    /// Valor de la celda leaf <c>P9</c>.
    /// </summary>
    public decimal P9 { get; set; }

    /// <summary>
    /// Visible esperado de la reversión del R4 cuando se recalcula la fórmula del workbook:
    /// <c>D67 = D9 - P9</c>.
    /// </summary>
    public decimal TotalReversionEsperada => D9 - P9;
}
