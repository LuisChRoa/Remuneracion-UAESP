namespace Remuneracion.Core.Models;

/// <summary>
/// HU-09 (2.3): una empresa de facturación dentro del bloque ASE del reporte por banco.
/// Los cuatro conceptos del bloque (filas 3–6 del bloque en la hoja
/// <c>REPORTE RECAUDO x BANCO</c>) provienen del "Resumen Recaudo Aplicado Por Servicio"
/// del <c>ReportePagosxBanco_*.xlsx</c> (búsqueda dinámica desde el final, T0-0.7).
/// </summary>
public sealed class ReporteBancoEmpresaInputs
{
    /// <summary>
    /// Nombre de la empresa de facturación tal como aparece en el Resumen de la fuente
    /// ("ENEL", "NUEVO ESQUEMA", "ENERBIT", "OCCIDENTE"). Nunca se inventa: si la columna
    /// esperada no existe en la fuente, el reader falla nombrando ASE + empresa-columna.
    /// </summary>
    public string Empresa { get; set; } = string.Empty;

    /// <summary>
    /// Concepto 1: "1-APLICADOS A FACTURACION".
    /// </summary>
    public decimal AplicadosFacturacion { get; set; }

    /// <summary>
    /// Concepto 2: "2-SALDOS A FAVOR GENERADOS".
    /// </summary>
    public decimal SaldosFavorGenerados { get; set; }

    /// <summary>
    /// Concepto 3: "3-APLICADOS A FINANCIACIONES NUEVAS" (match por prefijo normalizado
    /// tolera la variante sin "NUEVAS", T0-0.4).
    /// </summary>
    public decimal FinanciacionesNuevas { get; set; }

    /// <summary>
    /// Concepto 7: "7-APLICADOS A RECIBOS SERVICIOS ESPECIALES". En Q1 ninguna fuente trae
    /// la fila (fila ausente = 0 demostrable; el Total de la fuente lo confirma).
    /// </summary>
    public decimal RecibosServEspeciales { get; set; }

    /// <summary>
    /// Total de la empresa en la fila "Total" del Resumen de la fuente. Es la contraparte del
    /// gate D5(i) "bloque template == resumen fuente" ±0.5 (§2.5.2): el validador compara
    /// <see cref="Total"/> contra este valor sin abrir ningún xlsx.
    /// </summary>
    public decimal TotalFuente { get; set; }

    /// <summary>
    /// Total del bloque por empresa = Σ de los 4 conceptos (análogo a
    /// <see cref="ConsolidadoAse.TotalAse"/>). Debe coincidir con <see cref="TotalFuente"/> ±0.5.
    /// </summary>
    public decimal Total => AplicadosFacturacion + SaldosFavorGenerados + FinanciacionesNuevas + RecibosServEspeciales;
}

/// <summary>
/// HU-09 (2.3): bloque de 10 filas de un ASE en la hoja <c>REPORTE RECAUDO x BANCO</c>
/// (etiqueta ASE + Resumen + header + 4 conceptos + Total; filas 9–58 del template).
/// </summary>
public sealed class ReporteBancoAseInputs
{
    /// <summary>
    /// ASE al que pertenece el bloque.
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Empresas del bloque con sus 4 conceptos. Solo las empresas <b>presentes</b> en el Resumen
    /// de la fuente: la fuente ausente no se emite (el writer itera presentes y deja el 0 de la
    /// plantilla). Ver <see cref="FuentesAusentesBanco"/>.
    /// </summary>
    public IReadOnlyList<ReporteBancoEmpresaInputs> Empresas { get; set; } = [];

    /// <summary>
    /// Plan 51 (A2, §2.2): huella durable de las fuentes de recaudo del catálogo que la plantilla
    /// declara para este ASE (<c>WorkbookLeafCellMapReporteBanco.EditablesPorAse[ase].Keys</c>) y
    /// que <b>no vinieron</b> en el Resumen del reporte por banco. Formato <c>BANCO:{label}</c>
    /// (p. ej. <c>BANCO:NUEVO ESQUEMA</c>). Vacía = todas las fuentes esperadas vinieron
    /// (comportamiento anterior). La ausencia es un caso NORMAL: sus celdas destino quedan en el 0
    /// de la plantilla (no se emite fila, no hay <c>TotalFuente = 0</c> que confundir con un total
    /// real) y no es divergencia ni error.
    /// </summary>
    public IReadOnlyList<string> FuentesAusentesBanco { get; set; } = [];

    /// <summary>
    /// Plan 51 (A3, §2.3): evidencia de celdas de header del Resumen que no resuelven a ninguna
    /// fuente del vocabulario del negocio (catálogo + rótulo vigente por ASE + sinonimia citada) y
    /// que no son <c>Total</c>. Formato <c>BANCO:{label}</c>. Solo se convierte en fail-fast si
    /// coexiste con una <see cref="FuentesAusentesBanco"/> del MISMO ASE (posible fuente presente
    /// no reconocida: cambio de etiqueta). Vacía = sin hallazgos.
    /// </summary>
    public IReadOnlyList<string> EtiquetasBancoNoReconocidas { get; set; } = [];
}

/// <summary>
/// HU-09 (2.3): inputs de la hoja <c>REPORTE RECAUDO x BANCO</c> para un ASE (dentro de
/// <see cref="WorkbookLeafInputs.ReporteBanco"/>). En el flujo período cada leaf carga SU
/// bloque (<see cref="Ases"/> con un solo elemento); el validador multi-ASE agrega la Σ por
/// empresa a través de los 5 leafs para el gate D5(ii).
/// </summary>
public sealed class ReporteBancoInputs
{
    /// <summary>
    /// Bloques ASE del reporte. En el flujo por-ASE contiene el bloque del leaf actual.
    /// </summary>
    public IReadOnlyList<ReporteBancoAseInputs> Ases { get; set; } = [];

    /// <summary>
    /// Quincena para la celda C59 (1 o 2). Proviene de <see cref="Periodo.NumeroQuincena"/>
    /// (dominio, nunca de la fuente; plan §2.2 D7 / Requirement 4).
    /// </summary>
    public int Quincena { get; set; }

    /// <summary>
    /// Consolidado filas 1–7. T0-0.1 determinó que esas filas son FÓRMULAS en el template
    /// (D2(b)): este campo queda <c>null</c> y las filas 1–7 entran al mapa de fórmulas
    /// protegidas; solo se verifica Σ (gate D5(ii) contra el caché golden en Capa A).
    /// </summary>
    public IReadOnlyDictionary<string, decimal>? Consolidado { get; set; }
}
