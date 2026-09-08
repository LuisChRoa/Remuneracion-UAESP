namespace Remuneracion.Core.Models;

/// <summary>
/// Resultado del caso de uso de remuneración de un período completo (5 ASE).
/// </summary>
public sealed class ResultadoProcesoPeriodo
{
    /// <summary>
    /// Resultado del cálculo con los 5 consolidados.
    /// </summary>
    public ResultadoRemuneracion Resultado { get; set; } = new();

    /// <summary>
    /// Inputs leaf, uno por ASE (5).
    /// </summary>
    public IReadOnlyList<WorkbookLeafInputs> Leafs { get; set; } = [];

    /// <summary>
    /// Ruta del archivo de salida generado.
    /// </summary>
    public string RutaSalida { get; set; } = string.Empty;
}