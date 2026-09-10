namespace Remuneracion.Core.Models;

/// <summary>
/// Inputs 2.2 (HU-08) de la conciliación por empresa para UN ASE: operandos editables en
/// R1/R2/R4 + visibles esperados post-Excel por hoja. Los operandos se escriben en la misma
/// pasada HU-07; los visibles alimentan el gate Σ empresas = visible de bloque (§2.5).
///
/// Las claves de <see cref="CeldasR1"/>/<see cref="CeldasR2"/>/<see cref="CeldasR4"/> son
/// referencias del template (p. ej. "F33", "E18", "D8"), congeladas por T0 en
/// <c>WorkbookLeafCellMapPorEmpresa</c>.
/// </summary>
public sealed class ConciliacionEmpresaInputs
{
    /// <summary>
    /// Empresa de facturación (catálogo 2.2).
    /// </summary>
    public EmpresaFacturacion Empresa { get; set; } = new();

    /// <summary>
    /// ASE al que pertenecen estos inputs.
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Operandos editables por empresa en <c>Reporte Componentes R1</c> (ref del template → valor).
    /// </summary>
    public IReadOnlyDictionary<string, decimal> CeldasR1 { get; set; } = new Dictionary<string, decimal>();

    /// <summary>
    /// Operandos editables por empresa en <c>Rem. Anticipos R2</c>.
    /// </summary>
    public IReadOnlyDictionary<string, decimal> CeldasR2 { get; set; } = new Dictionary<string, decimal>();

    /// <summary>
    /// Operandos editables por empresa en <c>Reversion Pagos R4</c>.
    /// </summary>
    public IReadOnlyDictionary<string, decimal> CeldasR4 { get; set; } = new Dictionary<string, decimal>();

    /// <summary>
    /// Visible esperado post-Excel del detalle por empresa en R1 (Σ operandos según cadena T0).
    /// </summary>
    public decimal VisibleR1 { get; set; }

    /// <summary>
    /// Visible esperado post-Excel del detalle por empresa en R2.
    /// </summary>
    public decimal VisibleR2 { get; set; }

    /// <summary>
    /// Visible esperado post-Excel del detalle por empresa en R4.
    /// </summary>
    public decimal VisibleR4 { get; set; }
}
