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

    /// <summary>
    /// Plan 50 (A1, §2.3): huella auditable de las FUENTES DE RECAUDO AUSENTES en este ASE para esta
    /// empresa, con formato <c>{Reporte}:{label}</c> (p. ej. <c>R1:NUEVO ESQUEMA</c>,
    /// <c>R4:OCCIDENTE</c>). Vacía = comportamiento anterior (todas las fuentes del mapa vinieron).
    ///
    /// La ausencia es un caso NORMAL del negocio (§0.2 del plan): el reader materializa el 0
    /// explícito en <see cref="CeldasR1"/>/<see cref="CeldasR2"/>/<see cref="CeldasR4"/> y registra
    /// aquí la huella; los invariantes de cierre del reporte siguen siendo fail-fast.
    /// </summary>
    public IReadOnlyList<string> FuentesAusentes { get; set; } = [];

    /// <summary>
    /// Plan 50 (A3, §2.5 — control mínimo alternativo): evidencia READ-ONLY de etiquetas de fuente
    /// observadas en los reportes de este ASE que NO resuelven a ninguna empresa del mapa (ni por la
    /// sinonimia citada de <see cref="SinonimosEmpresaR1"/>), con formato <c>{Reporte}:{label}</c>
    /// (p. ej. <c>R2:EAB</c>).
    ///
    /// Es evidencia del REPORTE: el reader la replica en las 5 entradas del ASE (no hay contrato por
    /// ASE en <c>LeerConciliacionEmpresas</c>); el validador la deduplica. El juicio (fail-fast sólo
    /// si además hay una fuente ausente en el mismo reporte) vive en <c>ValidadorBasico</c>.
    /// </summary>
    public IReadOnlyList<string> EtiquetasNoReconocidas { get; set; } = [];
}
