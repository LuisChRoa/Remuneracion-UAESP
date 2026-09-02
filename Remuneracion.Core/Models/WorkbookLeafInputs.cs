namespace Remuneracion.Core.Models;

/// <summary>
/// Modelo raíz para los inputs leaf del workbook que alimentan las fórmulas reales del libro.
/// Este contrato representa los datos editables reales en R1, R2 y R4 sin depender del agregado final.
/// </summary>
public sealed class WorkbookLeafInputs
{
    /// <summary>
    /// ASE asociado al workbook que se está materializando.
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Periodo asociado al workbook de salida.
    /// </summary>
    public Periodo Periodo { get; set; } = new();

    /// <summary>
    /// Inputs leaf de la hoja <c>Reporte Componentes R1</c>.
    /// </summary>
    public WorkbookLeafInputsR1 R1 { get; set; } = new();

    /// <summary>
    /// Inputs leaf de la hoja <c>Rem. Anticipos R2</c>.
    /// </summary>
    public WorkbookLeafInputsR2 R2 { get; set; } = new();

    /// <summary>
    /// Inputs leaf de la hoja <c>Reversion Pagos R4</c>.
    /// </summary>
    public WorkbookLeafInputsR4 R4 { get; set; } = new();
}
