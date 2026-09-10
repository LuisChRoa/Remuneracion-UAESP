using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Contrato para generar un workbook real copiando la plantilla de origen y escribiendo solamente
/// las celdas leaf editables que alimentan las fórmulas del libro.
/// </summary>
public interface IWorkbookLeafWriter
{
    /// <summary>
    /// Genera un workbook de salida a partir de la plantilla fuente y los inputs leaf validados.
    /// </summary>
    /// <param name="rutaPlantillaOrigen">Ruta de la plantilla de origen.</param>
    /// <param name="rutaSalida">Ruta del archivo generado.</param>
    /// <param name="resultado">Resultado agregado del cálculo para la validación de coherencia.</param>
    /// <param name="leafInputs">Inputs leaf que alimentan las fórmulas del workbook.</param>
    void GenerarWorkbook(string rutaPlantillaOrigen, string rutaSalida, ResultadoRemuneracion resultado, WorkbookLeafInputs leafInputs);

    /// <summary>
    /// Genera un workbook multi-ASE (modo período) en UNA sola llamada de escritura: una copia,
    /// una validación pre/post con el mapa ampliado y la escritura de cada bloque por ASE.
    /// </summary>
    /// <param name="rutaPlantillaOrigen">Ruta de la plantilla de origen.</param>
    /// <param name="rutaSalida">Ruta del archivo generado (carpeta + <see cref="Periodo.NombreArchivo"/>).</param>
    /// <param name="resultado">Resultado con los consolidados por ASE.</param>
    /// <param name="leafInputs">Inputs leaf, uno por ASE (5 en modo período).</param>
    void GenerarWorkbook(string rutaPlantillaOrigen, string rutaSalida, ResultadoRemuneracion resultado, IReadOnlyList<WorkbookLeafInputs> leafInputs);
}
