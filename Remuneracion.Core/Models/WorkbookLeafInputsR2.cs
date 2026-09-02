namespace Remuneracion.Core.Models;

/// <summary>
/// Inputs editables reales de la hoja <c>Rem. Anticipos R2</c>.
/// La celda visible <c>E41</c> es una fórmula del workbook: <c>E15 + E26 - K15</c>.
/// </summary>
public sealed class WorkbookLeafInputsR2
{
    /// <summary>
    /// Valor en la celda leaf <c>E15</c>.
    /// </summary>
    public decimal E15 { get; set; }

    /// <summary>
    /// Valor en la celda leaf <c>E26</c>.
    /// </summary>
    public decimal E26 { get; set; }

    /// <summary>
    /// Valor en la celda leaf <c>K15</c>.
    /// </summary>
    public decimal K15 { get; set; }

    /// <summary>
    /// Visible esperado del total oportuno del R2 cuando se recalcula la fórmula del workbook:
    /// <c>E41 = E15 + E26 - K15</c>.
    /// </summary>
    public decimal TotalOportunoEsperado => E15 + E26 - K15;
}
