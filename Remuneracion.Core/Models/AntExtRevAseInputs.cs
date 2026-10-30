namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 64 (H1.1, §2.2 / D-03/D-04): inputs de la hoja <c>ANT EXT-REV</c> de UN ASE, leídos del
/// insumo <c>RecaudosReversados_*.xlsx</c> (hoja <c>Sheet1</c>, bloque 2 = fuente, bloque 3 =
/// control — veredicto D-10 sustentado por bytes en <c>Docs/T0-Plan64-Ant-Ext-Rev.md</c> §6).
///
/// <para><b>Modelo PURO.</b> No abre archivos ni conoce períodos: el lector
/// (<c>ExcelDataReaderWorkbookLeafInputReader.LeerAntExtRev</c>) lo puebla y el escritor values-only
/// lo consume dentro de la pasada atómica. La opcionalidad vive fuera: <c>null</c> en
/// <see cref="WorkbookLeafInputs.AntExtRev"/> = comportamiento anterior intacto.</para>
///
/// <para><b>Regla de presencia (D-04).</b> Una fuente AUSENTE del insumo no se inventa: se omite de
/// <see cref="Fuentes"/> y queda como traza en <see cref="FuentesAusentes"/> (sus celdas conservan el
/// cero del template). Un CONCEPTO ausente dentro de una fuente PRESENTE viaja como <c>0</c>
/// explícito en el registro de esa fuente.</para>
/// </summary>
public sealed class AntExtRevAseInputs
{
    /// <summary>
    /// ASE al que pertenece la lectura.
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Fuentes PRESENTES en el insumo, en el orden en que sus grupos cierran en el archivo.
    /// </summary>
    public IReadOnlyList<AntExtRevFuenteInputs> Fuentes { get; set; } = [];

    /// <summary>
    /// Fuentes del catálogo (<see cref="CatalogoFuentesAntExtRev"/>) que el insumo NO trae, con la
    /// traza del Plan 51: <c>ANTEXTREV:{fuente}</c>. Lista vacía = todas las fuentes del catálogo
    /// vinieron.
    /// </summary>
    public IReadOnlyList<string> FuentesAusentes { get; set; } = [];

    /// <summary>
    /// Etiquetas de fuente OBSERVADAS en el insumo que el catálogo no tabula (evidencia A3 del
    /// Plan 50/51: <c>ANTEXTREV:{etiqueta}</c>). Su sola presencia no bloquea; su coexistencia con una
    /// ausencia del catálogo sí es fail-fast (gemelo A3, §2.4 paso 6).
    /// </summary>
    public IReadOnlyList<string> EtiquetasNoReconocidas { get; set; } = [];

    /// <summary>
    /// SONDA DE ETIQUETAS del insumo (Plan 64 H5.2, R-04): una entrada por cada etiqueta de fuente
    /// observada en el bloque 2 que <see cref="CatalogoFuentesAntExtRev"/> no tabula, en la forma
    /// administrativa <c>ASE n · RecaudosReversados · Sheet1 · {etiqueta}</c>.
    ///
    /// <para>Es la evidencia DURABLE del hallazgo: su sola presencia no bloquea la corrida. Si la etiqueta
    /// desconocida COEXISTE con una fuente del catálogo ausente, el lector ya hizo fail-fast
    /// <c>ERR-FORMATO-FUENTE</c> (gemelo A3 de los Planes 50/51) y no hay lectura que reportar.</para>
    /// </summary>
    public IReadOnlyList<string> HallazgosEtiquetasNoMapeadas { get; set; } = [];

    /// <summary>
    /// Σ de los totales de fuente del bloque 2 (cierre del bloque leído). Es la magnitud contra la que
    /// se contrasta el control del bloque 3.
    /// </summary>
    public decimal TotalBloque2 => Fuentes.Sum(fuente => fuente.TotalFuente);

    /// <summary>
    /// Total del bloque 3 (<c>Resumen Recaudo Aplicado Por Servicio</c>), leído SOLO como control
    /// (§2.4 paso 7): <c>null</c> cuando el bloque no declara su columna <c>Total</c> o su fila de
    /// cierre. Nunca es fuente ni destino de escritura.
    /// </summary>
    public decimal? TotalControlBloque3 { get; set; }

    /// <summary>
    /// Hallazgo auditable del control del bloque 3 (ASE + Δ observado). <c>null</c> = el control cierra
    /// contra <see cref="TotalBloque2"/> dentro de ±0,5. El bloque 3 no discrimina por fuente, así que
    /// una divergencia aquí se REGISTRA: no bloquea la entrega (§2.4 paso 7).
    /// </summary>
    public string? HallazgoControlBloque3 { get; set; }
}

/// <summary>
/// Plan 64 (§2.2 / §1.2): los cuatro conceptos de una fuente PRESENTE, en el orden congelado por T0
/// (<c>AplicadoR3</c>, <c>AnticiposR3</c>, <c>AplicadoR4</c>, <c>AnticiposR4</c>), más el total de la
/// fuente que cierra el grupo en el insumo.
///
/// <para>El signo viaja <b>tal cual</b> viene del origen (los R4 ya son negativos; la plantilla no
/// re-signa). <see cref="FuenteEtiqueta"/> es la etiqueta OBSERVADA normalizada (mayúsculas, sin
/// diacríticos); el mapeo a la etiqueta de bloque del template lo resuelve
/// <see cref="CatalogoFuentesAntExtRev"/>.</para>
/// </summary>
/// <param name="FuenteEtiqueta">Etiqueta de fuente observada en el bloque 2 (columna A del total).</param>
/// <param name="AplicadoR3">Concepto <c>1-APLICADOS A FACTURACION</c> (fila APLICADO, REPORTE 3).</param>
/// <param name="AnticiposR3">Concepto <c>2-SALDOS A FAVOR GENERADOS</c> (fila ANTICIPOS, REPORTE 3).</param>
/// <param name="AplicadoR4">Concepto <c>9-REVERSION PAGOS A COMPONENTES</c> (fila APLICADO, REPORTE 4).</param>
/// <param name="AnticiposR4">Concepto <c>9-REVERSION SALDOS A FAVOR</c> (fila ANTICIPOS, REPORTE 4).</param>
/// <param name="TotalFuente">Total de la fuente que cierra el grupo en el insumo (columna de total).</param>
public sealed record AntExtRevFuenteInputs(
    string FuenteEtiqueta,
    decimal AplicadoR3,
    decimal AnticiposR3,
    decimal AplicadoR4,
    decimal AnticiposR4,
    decimal TotalFuente)
{
    /// <summary>
    /// Σ de los cuatro conceptos. Invariante T0 §4.4 / §1.4: ≡ <see cref="TotalFuente"/> ±0,5 (el
    /// lector hace fail-fast si no cierra; aquí queda como cierre auditable del modelo).
    /// </summary>
    public decimal SumaConceptos => AplicadoR3 + AnticiposR3 + AplicadoR4 + AnticiposR4;
}
