using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Define la validación de un resultado de remuneración.
/// </summary>
public interface IValidador
{
    /// <summary>
    /// Valida el resultado de la remuneración y retorna la lista de advertencias o errores.
    /// </summary>
    /// <param name="resultado">Resultado a validar.</param>
    /// <returns>Lista de mensajes de validación (vacía si no hay problemas).</returns>
    List<string> Validar(ResultadoRemuneracion resultado);

    /// <summary>
    /// Valida el resultado utilizando además los inputs leaf que alimentan las fórmulas del workbook.
    /// </summary>
    /// <param name="resultado">Resultado a validar.</param>
    /// <param name="leaf">Inputs leaf asociados a la salida que se quiere certificar.</param>
    /// <returns>Lista de mensajes de validación (vacía si no hay problemas).</returns>
    List<string> Validar(ResultadoRemuneracion resultado, WorkbookLeafInputs leaf);
}
