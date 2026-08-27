using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Define el cálculo de la remuneración quincenal a partir de los datos de los reportes.
/// </summary>
public interface ICalculoRemuneracion
{
    /// <summary>
    /// Calcula el consolidado de un ASE individual.
    /// </summary>
    /// <param name="ase">ASE a liquidar.</param>
    /// <param name="r1">Datos del reporte R1.</param>
    /// <param name="r2">Datos del reporte R2.</param>
    /// <param name="r4">Datos del reporte R4.</param>
    /// <returns>Consolidado del ASE.</returns>
    ConsolidadoAse Calcular(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4);

    /// <summary>
    /// Calcula el consolidado global de todos los ASE para un periodo.
    /// </summary>
    /// <param name="periodo">Periodo quincenal.</param>
    /// <param name="datos">Tuplas con el ASE y sus reportes asociados.</param>
    /// <returns>Resultado global de la remuneración.</returns>
    ResultadoRemuneracion CalcularConsolidado(
        Periodo periodo,
        List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4)> datos);
}
