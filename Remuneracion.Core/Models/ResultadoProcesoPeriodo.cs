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

    /// <summary>
    /// Plan 42 (T1, R-CLS-3): clasificación de las divergencias de <c>VALIDACION_TOTAL</c> en
    /// declaradas vs no declaradas, con referencia a la entrada del catálogo y montos/Δ. La UI y el
    /// CLI la PRESENTAN sin re-evaluar (R-DOBLE-PUERTA). <c>null</c>/vacía = comportamiento anterior
    /// (sin divergencias o sin oráculo de validaciones). Aditivo (compat HU-13).
    /// </summary>
    public ClasificacionDivergenciasDeclaradas? DivergenciasDeclaradas { get; set; }

    /// <summary>
    /// Plan 42 (T2, R-REP-1): ruta del reporte durable de divergencias declaradas emitido junto a la
    /// salida. Vacío = no hubo cierre con declarada reconocida (no se emite reporte).
    /// </summary>
    public string ReporteDivergenciasDeclaradas { get; set; } = string.Empty;

    /// <summary>
    /// Plan 50 (A1, §2.3): huella durable de las FUENTES DE RECAUDO AUSENTES por (ASE, reporte, label),
    /// con formato <c>ASE {n} · {Reporte}:{label}</c> (p. ej. <c>ASE 2 · R1:NUEVO ESQUEMA</c>).
    /// Vacía = comportamiento anterior (todas las fuentes del período vinieron). La ausencia es un caso
    /// NORMAL: su celda destino quedó en 0 (no es divergencia ni error).
    /// </summary>
    public IReadOnlyList<string> FuentesAusentes { get; set; } = [];

    /// <summary>
    /// Plan 50 (A3, §2.5): evidencia de etiquetas de fuente observadas que no resuelven a ninguna
    /// empresa del mapa, con formato <c>ASE {n} · {Reporte}:{label}</c>. Sólo se convierte en
    /// fail-fast si coexiste con una <see cref="FuentesAusentes"/> del MISMO reporte (posible fuente
    /// presente no reconocida). Vacía = sin hallazgos.
    /// </summary>
    public IReadOnlyList<string> EtiquetasNoReconocidas { get; set; } = [];
}
