namespace Remuneracion.Core.Models;

/// <summary>
/// Solicitud para ejecutar la liquidación de un período completo (los 5 ASE) en un solo proceso.
/// </summary>
public sealed class SolicitudProcesoPeriodo
{
    /// <summary>
    /// Periodo quincenal a liquidar.
    /// </summary>
    public Periodo Periodo { get; set; } = new();

    /// <summary>
    /// Carpeta del período (ej. "REMUNERACION 2026071") que contiene las 5 carpetas de ASE.
    /// </summary>
    public string CarpetaPeriodo { get; set; } = string.Empty;

    /// <summary>
    /// Ruta de la plantilla de origen.
    /// </summary>
    public string RutaPlantilla { get; set; } = string.Empty;

    /// <summary>
    /// Ruta de salida (carpeta + <see cref="Periodo.NombreArchivo"/>; igual que HU-06).
    /// </summary>
    public string RutaSalida { get; set; } = string.Empty;
}