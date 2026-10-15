using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Plan 42 (T2, R-REP-1, D-E): contrato del REPORTE DURABLE de divergencias declaradas. Cuando una
/// corrida cierra con una divergencia declarada reconocida, la app deja DOS huellas obligatorias:
/// este reporte (artefacto legible por el administrativo, junto a la salida) y la traza en Serilog.
/// Sin reporte NO hay ÉXITO (D-E: declarada ≠ tolerada en silencio).
///
/// La implementación vive en Infrastructure (I/O); el orquestador de Core solo la invoca con los
/// datos ya clasificados por el validador (una sola decisión de clasificación, R-DOBLE-PUERTA).
/// </summary>
public interface IReporteDivergenciasDeclaradas
{
    /// <summary>
    /// Emite el reporte durable de las divergencias declaradas reconocidas en la corrida y devuelve
    /// la ruta del archivo generado (junto a la salida). El texto es de negocio (sin jerga de
    /// código: prohibido <c>ERR-*</c>, <c>&lt;f&gt;</c>, nombres de clase, etc.).
    /// </summary>
    /// <param name="rutaSalida">Ruta del archivo de salida generado (el reporte se emite junto a él).</param>
    /// <param name="periodo">Período de la corrida.</param>
    /// <param name="runId">RunId de la corrida (correlación con el log).</param>
    /// <param name="clasificacion">Clasificación de las divergencias (declaradas + no declaradas).</param>
    /// <param name="usuario">
    /// Plan 43 (O-3): usuario del sistema que aceptó la divergencia declarada. Aditivo; <c>null</c>/vacío
    /// = se resuelve <see cref="Environment.UserName"/> (o "No disponible" si el entorno no lo expone).
    /// </param>
    /// <param name="canal">
    /// Plan 43 (O-3): canal por el que se aceptó (UI o CLI). Aditivo; <c>null</c>/vacío = "No especificado".
    /// </param>
    /// <returns>Ruta absoluta del reporte emitido.</returns>
    string Emitir(
        string rutaSalida,
        Periodo periodo,
        Guid runId,
        ClasificacionDivergenciasDeclaradas clasificacion,
        string? usuario = null,
        string? canal = null);
}
