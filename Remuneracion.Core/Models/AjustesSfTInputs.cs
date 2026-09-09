namespace Remuneracion.Core.Models;

/// <summary>
/// HU-11 (2.5): composición por ASE de la cadena AJUSTES-SF-T → CONSOLIDADO D85:D89.
///
/// Veredicto T0-0.3 (probado contra el golden Q2, NUNCA asumido): la composición exacta es
/// <c>CONSOLIDADO D85:D89 = 'AJUSTES - SF-T'!D47..D51</c> y en la hoja AJUSTES - SF-T
/// <c>D47 = D9 + D28</c> (sección SALDOS A FAVOR POR NOTA fila 9 + sección RETRIBUCION NEGATIVA
/// fila 28). A su vez <c>D9 = 'SALDOS POR NOTA'!C9</c> (visible Cn-In del bloque SALDOS) y
/// <c>D28 = 'RETRIBUCION NEGATIVA'!C10</c> (visible Cn-In del bloque RETRIBUCION).
///
/// La aritmética de dominio <see cref="TotalAjustes"/> replica esa composición:
/// <see cref="SaldosNotas.TotalSaldosNotas"/> + <see cref="RetribucionNegativa.TotalRetribucionNegativa"/>.
/// Los totales por ASE probados contra el golden: [973693.46, 216025.77, 104231.83, 35954.44, 0].
/// </summary>
public sealed class AjustesSfTInputs
{
    /// <summary>
    /// ASE al que pertenece la composición.
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Bloque SALDOS POR NOTA del ASE (fuente <c>SaldosaFavorAplicadosPorNotas_*</c>).
    /// </summary>
    public SaldosNotasAseInputs SaldosNotas { get; set; } = new();

    /// <summary>
    /// Bloque RETRIBUCION NEGATIVA del ASE (fuente <c>RetribuciónNegativa_*</c>).
    /// </summary>
    public RetribucionNegativaAseInputs RetribucionNegativa { get; set; } = new();

    /// <summary>
    /// Total de ajustes del ASE = saldos por nota + retribución negativa (composición T0-0.3).
    /// Es el valor que alimenta <see cref="ConsolidadoAse.AjustesSfT"/> en Q2.
    /// </summary>
    public decimal TotalAjustes => SaldosNotas.TotalSaldosNotas + RetribucionNegativa.TotalRetribucionNegativa;
}