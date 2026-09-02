namespace Remuneracion.Core.Models;

/// <summary>
/// Resultado del caso de uso de cálculo para un ASE.
/// </summary>
public sealed class ResultadoProcesoAse
{
    public ResultadoRemuneracion Resultado { get; set; } = new();
    public WorkbookLeafInputs Leaf { get; set; } = new();
    public string RutaSalida { get; set; } = string.Empty;
}
