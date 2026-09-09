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

    /// <summary>
    /// HU-11 (2.5, D1): calcula el consolidado de un ASE individual con ajustes AJUSTES-SF-T
    /// (Q2). El path Q1 (<see cref="Calcular(Ase, RecaudoComponenteR1, SaldosFavorR2, ReversionR4)"/>)
    /// queda intacto; esta sobrecarga solo difiere en <paramref name="ajustesSfT"/>.
    /// </summary>
    /// <param name="ase">ASE a liquidar.</param>
    /// <param name="r1">Datos del reporte R1.</param>
    /// <param name="r2">Datos del reporte R2.</param>
    /// <param name="r4">Datos del reporte R4.</param>
    /// <param name="ajustesSfT">Ajustes por saldos a favor y retribución negativa (Q2).</param>
    /// <returns>Consolidado del ASE con <see cref="ConsolidadoAse.AjustesSfT"/> poblado.</returns>
    ConsolidadoAse Calcular(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4, decimal ajustesSfT);

    /// <summary>
    /// HU-11 (2.5, D1): calcula el consolidado global de un período Q2 con ajustes por ASE.
    /// El overload viejo (sin ajustes) queda intacto como red de seguridad: sigue lanzando
    /// <c>CalculoInvalidoException</c> para Q2 sin modelar.
    /// </summary>
    /// <param name="periodo">Periodo quincenal (Q1 o Q2).</param>
    /// <param name="datos">Tuplas con ASE, reportes y ajustes AJUSTES-SF-T por ASE.</param>
    /// <returns>Resultado global de la remuneración con <see cref="ConsolidadoAse.AjustesSfT"/>.</returns>
    ResultadoRemuneracion CalcularConsolidado(
        Periodo periodo,
        List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4, decimal ajustesSfT)> datos);
}
