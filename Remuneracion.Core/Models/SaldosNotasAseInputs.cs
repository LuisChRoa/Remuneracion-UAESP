namespace Remuneracion.Core.Models;

/// <summary>
/// HU-11 (2.5): inputs de un ASE desde la fuente <c>SaldosaFavorAplicadosPorNotas_*.xlsx</c>
/// (SALDOS POR NOTA). Patrón R2 literal: columna "Especiales" opcional — en las 5 fuentes Q2
/// verificadas por T0-0.5 la columna NO existe (el header de esa posición es "Componente TCS"),
/// por lo que <see cref="TieneColumnaEspeciales"/> es <c>false</c> y <see cref="ServEspK"/> = 0.
///
/// Veredicto T0-0.3 (probado contra golden Q2, no asumido): la fila "Total" de cada bloque del
/// template es VALOR editable (C7/C20/C32/C45/C58) y el visible del bloque (C9/C22/C34/C47/C60)
/// es fórmula <c>Cn-In</c> (Total − Especiales). La aritmética de dominio
/// <see cref="TotalSaldosNotas"/> replica esa fórmula: Total (col C de la fila Total de la
/// fuente) − ServEspK (ausente = 0). Es análoga a <see cref="SaldosFavorR2.TotalOportuno"/>.
/// </summary>
public sealed class SaldosNotasAseInputs
{
    /// <summary>
    /// ASE al que pertenece el bloque.
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Indica si la fuente trae columna "Especiales" (patrón R2). T0-0.5: false en los 5 ASE Q2.
    /// </summary>
    public bool TieneColumnaEspeciales { get; set; }

    /// <summary>
    /// Total (col C) de la fila "Total" del bloque en la fuente (operando del visible Cn-In).
    /// </summary>
    public decimal Total { get; set; }

    /// <summary>
    /// Especiales de la fila "Total" (columna "Especiales" de la fuente; 0 si ausente).
    /// </summary>
    public decimal ServEspK { get; set; }

    /// <summary>
    /// Celdas editables del bloque SALDOS POR NOTA por ASE, keyed por referencia de celda del
    /// template (ej. "C3", "D3", …, "O7"). Congelado por T0-0.7; el writer escribe SOLO estas
    /// celdas. El valor de "Especiales" (columna I del template) es 0 (fuente sin esa columna).
    /// </summary>
    public IReadOnlyDictionary<string, decimal> Celdas { get; set; } = new Dictionary<string, decimal>();

    /// <summary>
    /// Total de saldos por nota del ASE = Total − ServEspK (aritmética T0-0.3, visible Cn-In).
    /// </summary>
    public decimal TotalSaldosNotas => Total - ServEspK;
}