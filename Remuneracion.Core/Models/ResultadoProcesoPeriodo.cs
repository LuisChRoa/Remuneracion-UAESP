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

    /// <summary>
    /// HU-13 (2.7): veredictos de validaciones cruzadas por ASE (bloque VALIDACIONES del log/UI).
    /// Cada entrada es una línea ya formateada (cierra / diverge documentada). Lista vacía =
    /// sin gates 2.7 (reader ausente o snapshot ausente → HU-12 puro).
    /// </summary>
    public IReadOnlyList<string> Validaciones { get; set; } = [];
}
