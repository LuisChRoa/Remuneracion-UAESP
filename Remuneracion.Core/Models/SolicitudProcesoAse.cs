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
}
