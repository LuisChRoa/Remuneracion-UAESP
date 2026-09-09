namespace Remuneracion.Core.Models;

/// <summary>
/// Solicitud para ejecutar la liquidación de un único ASE.
/// </summary>
public sealed class SolicitudProcesoAse
{
    public Ase Ase { get; set; } = new();
    public Periodo Periodo { get; set; } = new();
    public string RutaR1 { get; set; } = string.Empty;
    public string RutaR2 { get; set; } = string.Empty;
    public string RutaR4 { get; set; } = string.Empty;
    public string RutaPlantilla { get; set; } = string.Empty;
    public string RutaSalida { get; set; } = string.Empty;

    /// <summary>
    /// HU-15 (W-2.1, D4): RunId inyectado por el frontend (UI/CLI) para correlacionar TODOS
    /// los eventos de la ejecución (UI → procesador → writer) en el log. <c>null</c> = el
    /// procesador genera uno (compat HU-14: <c>Guid.NewGuid()</c>).
    /// </summary>
    public Guid? RunId { get; set; }
}
