namespace Remuneracion.Core.Models;

/// <summary>
/// Inputs editables reales de la hoja <c>Rem. Anticipos R2</c>.
/// La celda visible <c>E41</c> es una fórmula del workbook: <c>E15 + E26 - K15</c>.
///
/// HU-07 (multi-ASE): el mapeo por labels es uniforme para los 5 bloques; solo cambian las
/// direcciones del template (E85/E103/K85 para ASE2, E167/E178/K167 para ASE3, etc.), expuestas
/// como ROL de destino en <see cref="ValoresPorRol"/> (Plan 59 H2.1: sin celdas en el transporte).
/// </summary>
public sealed class WorkbookLeafInputsR2
{
    /// <summary>
    /// Valor en la celda leaf <c>E15</c> (Componente/Total col E del bloque).
    /// </summary>
    public decimal E15 { get; set; }

    /// <summary>
    /// Valor en la celda leaf <c>E26</c> (Subs/Cont/Total col E del bloque).
    /// </summary>
    public decimal E26 { get; set; }

    /// <summary>
    /// Valor en la celda leaf <c>K15</c> (Especiales del bloque; 0 si no hay columna).
    /// </summary>
    public decimal K15 { get; set; }

    /// <summary>
    /// Visible esperado del total oportuno del R2 cuando se recalcula la fórmula del workbook:
    /// <c>E41 = E15 + E26 - K15</c>. Equivalente por bloque (E135, E247, E343, E413).
    /// </summary>
    public decimal TotalOportunoEsperado => E15 + E26 - K15;

    /// <summary>
    /// Valores del bloque R2 por ASE, keyed por ROL de destino (concepto, no coordenada): el cierre
    /// de cada sección y los especiales. Plan 59 (H2.1): el escritor resuelve el rol contra la
    /// plantilla; el transporte no lleva celdas.
    /// </summary>
    public IReadOnlyDictionary<string, decimal> ValoresPorRol { get; set; } = new Dictionary<string, decimal>();
}
