namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 53 (H1.1, D1): fila-empresa ya TIPADA por sección — el índice (opaco para el resolvedor) de
/// la fila en la fuente y su <see cref="ContratoSeccionPeriodFreeR1.SeccionPorCierre"/>. El tipado lo
/// hace el reader con el predicado vigente (Plan 46 §2.2); el resolvedor NO reimplementa el scan de
/// fronteras (D3.4).
/// </summary>
/// <param name="Fila">Identificador de la fila (índice crudo en la fuente).</param>
/// <param name="Seccion">Sección de la fila (frontera que la cierra).</param>
public readonly record struct FilaSeccionadaR1(int Fila, SeccionFronteraArriba Seccion);
