using Remuneracion.Core.Rules;

namespace Remuneracion.Core.Models;

/// <summary>
/// HU-12 (2.6 ampliada, V0.4): DetRetri-Q2 por ASE con composición CONGELADA y probada contra el
/// golden 5/5: <c>DetRetri_D(ase) = ROUND(D104:D108(ase), 0)</c> donde D104:D108 es el total del
/// ASE en CONSOLIDADO (Σ visibles leaf Q2: R1 TOT_OPT + R2 + EXTEMP + R4 + AJUSTES-SF-T).
///
/// <see cref="TotalD104"/> replica el valor que Excel computa en D104:D108 (suma de los visibles
/// leaf; NUNCA agregados HU-02 — A5) y <see cref="Detalle"/> aplica la ÚNICA regla de redondeo
/// permitida (<see cref="DetRetriRounder"/>; prohibido duplicarla, V0.7).
/// </summary>
public sealed class DetRetriQ2Inputs
{
    /// <summary>
    /// ASE al que pertenece el detalle.
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Total D104:D108 del ASE (Σ visibles leaf Q2, composición V0.4). Es el valor que Excel
    /// computa en CONSOLIDADO D104:D108 para este ASE.
    /// </summary>
    public decimal TotalD104 { get; set; }

    /// <summary>
    /// Detalle DetRetri-D (entero) = <see cref="DetRetriRounder.Round"/>(<see cref="TotalD104"/>).
    /// Es el valor que se escribe en <c>DetRetri2026072!D9:D13</c> (V0.4).
    /// </summary>
    public decimal Detalle => DetRetriRounder.Round(TotalD104);
}