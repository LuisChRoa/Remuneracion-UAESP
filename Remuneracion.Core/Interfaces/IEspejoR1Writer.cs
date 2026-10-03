using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Plan 21 (T3, R-E-2/R-E-3/R-E-4): capacidad de escritura espejo acotada a los bloques de la
/// hoja <c>Reporte Componentes R1</c>. Dimensiona el bloque destino de cada ASE a la forma de la
/// fuente del período actual (inserta/borra filas preservando estilos y fórmulas), reancla toda
/// fórmula y nombre definido que referencie el bloque mutado (referencias A1 puras; ±delta por
/// filas) y escribe los valores por ENCABEZADO de columna (nunca por posición).
///
/// La implementación NO recalcula la aritmética de negocio: solo mueve filas, reancla referencias
/// y copia los valores de la fuente observada. El guard anti-fórmula sigue vigente: si una celda
/// de valor es fórmula tras la mutación, lanza <c>CalculoInvalidoException</c> (<c>ERR-PLANTILLA</c>),
/// nunca sobrescritura silenciosa.
/// </summary>
public interface IEspejoR1Writer
{
    /// <summary>
    /// Aplica el espejo estructural R1 sobre una COPIA de la plantilla: por cada ASE (procesado de
    /// abajo hacia arriba, ASE 5 → ASE 1, para que cada mutación vea las direcciones de los bloques
    /// inferiores ya estabilizadas) redimensiona el bloque y escribe la secuencia fuente.
    /// </summary>
    /// <param name="rutaPlantillaOrigen">Plantilla origen (no se muta).</param>
    /// <param name="rutaSalida">Ruta del archivo generado.</param>
    /// <param name="bloques">Bloques espejo por ASE (secuencia fuente del período actual).</param>
    void EscribirEspejoR1(string rutaPlantillaOrigen, string rutaSalida, IReadOnlyList<BloqueEspejoAseInputs> bloques);
}
