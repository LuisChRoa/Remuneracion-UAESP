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
}
