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

    /// <summary>
    /// HU-08 (2.2): conciliación por empresa de facturación para este ASE
    /// (operandos R1/R2/R4 + visibles esperados). Lista vacía = comportamiento HU-07 puro.
    /// </summary>
    public IReadOnlyList<ConciliacionEmpresaInputs> Conciliacion { get; set; } = [];

    /// <summary>
    /// HU-08 (2.2): hojas <c>Recaudo *</c> en valores (5 empresas). Lista vacía = HU-07 puro.
    /// </summary>
    public IReadOnlyList<RecaudoEmpresaInputs> Recaudos { get; set; } = [];

    /// <summary>
    /// HU-09 (2.3): inputs de la hoja <c>REPORTE RECAUDO x BANCO</c> para este ASE (bloque en
    /// valores desde el Resumen del <c>ReportePagosxBanco_*</c> + C59). <c>null</c> =
    /// comportamiento HU-08 puro (compatibilidad hacia atrás por construcción, plan §0.3 G4).
    /// </summary>
    public ReporteBancoInputs? ReporteBanco { get; set; }

    /// <summary>
    /// HU-10 (2.4): inputs de la hoja <c>BCE SC POR FACT.</c> para este ASE (Contribución +
    /// Subsidio desde el "Total General" del <c>R4-BalanceSubsidioyContribuciones_*</c>).
    /// <c>null</c> = comportamiento HU-09 puro (compatibilidad hacia atrás por construcción,
    /// plan §0.4 G6 / §8).
    /// </summary>
    public BalanceScInputs? BalanceSc { get; set; }
}
