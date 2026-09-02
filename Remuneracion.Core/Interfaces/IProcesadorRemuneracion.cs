using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Orquesta la ejecución de la remuneración para un único ASE.
/// </summary>
public interface IProcesadorRemuneracion
{
    ResultadoProcesoAse Ejecutar(SolicitudProcesoAse solicitud, IProgress<string>? progreso = null);
}
