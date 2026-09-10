using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Caso de uso de remuneración de un período completo (los 5 ASE) en un solo proceso.
/// </summary>
public interface IProcesadorPeriodo
{
    /// <summary>
    /// Ejecuta la liquidación multi-ASE: resuelve las 5 carpetas, lee y calcula por ASE,
    /// valida la coherencia y escribe UNA sola vez el workbook de salida.
    /// </summary>
    /// <param name="solicitud">Datos del período a procesar.</param>
    /// <param name="progreso">Reporte de progreso por ASE y por paso (opcional).</param>
    /// <returns>Resultado con los 5 consolidados, los 5 leafs y la ruta de salida.</returns>
    ResultadoProcesoPeriodo Ejecutar(SolicitudProcesoPeriodo solicitud, IProgress<string>? progreso = null);
}
