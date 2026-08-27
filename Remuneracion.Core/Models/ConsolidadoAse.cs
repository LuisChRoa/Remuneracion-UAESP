using Remuneracion.Core.Models;

namespace Remuneracion.Core.Models;

/// <summary>
/// Consolidado de remuneración de un ASE, agrupando los valores provenientes de los
/// reportes R1, R2, R4 y los ajustes (Fase 1 y posteriores).
/// </summary>
public class ConsolidadoAse
{
    /// <summary>
    /// ASE al que corresponde el consolidado.
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Total oportuno (D[n]): proveniente del reporte R1.
    /// </summary>
    public decimal TotOpt { get; set; }

    /// <summary>
    /// Total oportuno del reporte R2 (D[n]).
    /// </summary>
    public decimal R2TotalOportuno { get; set; }

    /// <summary>
    /// Valor extemporáneo (D[n]): proveniente del reporte R1.
    /// </summary>
    public decimal Extemp { get; set; }

    /// <summary>
    /// Reversión del reporte R4 (D[n]).
    /// </summary>
    public decimal ReversionR4 { get; set; }

    /// <summary>
    /// Ajustes por saldos a favor y retribución negativa (D[n]). En Fase 1 es 0.
    /// </summary>
    public decimal AjustesSfT { get; set; }

    /// <summary>
    /// Total del ASE, suma de todos los conceptos.
    /// </summary>
    public decimal TotalAse => TotOpt + R2TotalOportuno + Extemp + ReversionR4 + AjustesSfT;
}
