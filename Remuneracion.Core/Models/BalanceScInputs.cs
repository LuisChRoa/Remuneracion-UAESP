namespace Remuneracion.Core.Models;

/// <summary>
/// HU-10 (2.4): una fila ASE (3–7) de la hoja <c>BCE SC POR FACT.</c>.
/// Los valores provienen de la fila "Total General" del <c>R4-BalanceSubsidioyContribuciones_*.xlsx</c>
/// con la asignación D/E del veredicto T0-0.2 (hipótesis líder PROBADA contra golden Q1):
/// template-D (CONTRIBUCION, positivo) ← columna F-fuente; template-E (SUBSIDIO, negativo) ←
/// columna E-fuente. El modelo usa <see cref="Contribucion"/>/<see cref="Subsidio"/> por
/// significado de dominio; el mapa (<c>WorkbookLeafCellMapBalanceSc</c>) fija qué propiedad va
/// a qué columna template. Si T0 hubiera probado la asignación inversa, solo cambia el mapa.
/// </summary>
public sealed class BalanceScAseInputs
{
    /// <summary>
    /// ASE al que pertenece la fila.
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Contribución de la quincena (columna F-fuente del "Total General"; template-D, positiva).
    /// </summary>
    public decimal Contribucion { get; set; }

    /// <summary>
    /// Subsidio de la quincena (columna E-fuente del "Total General"; template-E, negativa).
    /// </summary>
    public decimal Subsidio { get; set; }

    /// <summary>
    /// Total BSC de la fila = <see cref="Contribucion"/> + <see cref="Subsidio"/> (análogo a
    /// <see cref="ConsolidadoAse.TotalAse"/>). Es la columna F del template (F = D+E, fórmula
    /// protegida que recalcula sola) y la columna G-fuente (Valor) del Balance.
    /// </summary>
    public decimal TotalBsc => Contribucion + Subsidio;

    /// <summary>
    /// Valor de la columna G-fuente (Valor) de la fila "Total General". Es la contraparte del
    /// gate D5(i) "BCE = fuente" ±0.5: el validador compara <see cref="TotalBsc"/> contra este
    /// valor sin abrir ningún xlsx.
    /// </summary>
    public decimal TotalFuente { get; set; }

    /// <summary>
    /// Valor de la columna H (SISTEMA) del template. Veredicto T0-0.6 = D2(b): H es FÓRMULA
    /// (<c>DetRetri2026071!J9..J13</c>) en el template → este campo queda <c>null</c>, H entra
    /// al mapa de fórmulas protegidas y solo se verifica H≈F contra el caché golden en Capa A.
    /// Si un template futuro trajera H como valor (D2(a)), el mapa lo escribiría con esta regla:
    /// redondeo a entero (<see cref="Math.Round(decimal, MidpointRounding.AwayFromZero)"/>).
    /// </summary>
    public decimal? Sistema { get; set; }
}

/// <summary>
/// HU-10 (2.4): inputs de la hoja <c>BCE SC POR FACT.</c> para un ASE (dentro de
/// <see cref="WorkbookLeafInputs.BalanceSc"/>). En el flujo período cada leaf carga SU fila
/// (<see cref="Ases"/> con un solo elemento); la fila 9 (total), la fila 11 (sumas) y el bloque
/// 18–24 de validación son fórmulas protegidas del template (2.7) y nunca se escriben.
/// </summary>
public sealed class BalanceScInputs
{
    /// <summary>
    /// Filas ASE del balance. En el flujo por-ASE contiene la fila del leaf actual.
    /// </summary>
    public IReadOnlyList<BalanceScAseInputs> Ases { get; set; } = [];
}
