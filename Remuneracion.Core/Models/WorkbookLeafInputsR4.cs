namespace Remuneracion.Core.Models;

/// <summary>
/// Inputs editables reales de la hoja <c>Reversion Pagos R4</c>.
/// La celda visible <c>D67</c> es una fórmula del workbook: <c>D9 - P9</c>.
///
/// HU-07 (multi-ASE): el mapeo por labels es uniforme para los 5 bloques; solo cambian las
/// direcciones del template (D98/P98 para ASE2, D193/P193 para ASE3, etc.), expuestas como ROL de
/// destino en <see cref="ValoresPorRol"/> (Plan 59 H2.1: sin celdas en el transporte).
/// </summary>
public sealed class WorkbookLeafInputsR4
{
    /// <summary>
    /// Valor de la celda leaf <c>D9</c> (Total del bloque col D, negativo).
    /// </summary>
    public decimal D9 { get; set; }

    /// <summary>
    /// Valor de la celda leaf <c>P9</c> (0 en la plantilla de referencia).
    /// </summary>
    public decimal P9 { get; set; }

    /// <summary>
    /// Visible esperado de la reversión del R4 cuando se recalcula la fórmula del workbook:
    /// <c>D67 = D9 - P9</c>. Equivalente por bloque (D161, D198, D312, D347).
    /// </summary>
    public decimal TotalReversionEsperada => D9 - P9;

    /// <summary>
    /// Valores del bloque R4 por ASE, keyed por ROL de destino (concepto, no coordenada): el gran
    /// total del bloque y sus especiales. Plan 59 (H2.1): el escritor resuelve el rol contra la
    /// plantilla; el transporte no lleva celdas.
    /// </summary>
    public IReadOnlyDictionary<string, decimal> ValoresPorRol { get; set; } = new Dictionary<string, decimal>();
}
