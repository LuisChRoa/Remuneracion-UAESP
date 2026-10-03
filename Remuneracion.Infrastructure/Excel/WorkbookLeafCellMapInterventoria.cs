namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// HU-16: mapa hermano congelado por T0 (Plan 16 §4 Fase 0) para la hoja <c>INTERVENTORIA</c>
/// (D2b) y las L-Especiales menores del <c>Reporte Componentes R1</c> (D3a). Por doctrina
/// HU-07 D1 / HU-12 D1: explícito por (<see cref="Remuneracion.Core.Models.Ase.Id"/>, período),
/// prohibidos offsets aritméticos y filas fijas; los mapas HU-07..HU-12 quedan INTACTOS.
///
/// ── Veredicto T0-INTERVENTORIA (T0-0.2/0.3/0.4, ambos canónicos + Q2-raíz idénticos) ──
/// El bloque real (dump OpenXML <c>&lt;c&gt;</c> raw) es:
///   fila 25  = headers (VALOR OFICIAL MES | ASE | Segunda quincena | Primera quincena)
///   filas 26..30 = un ASE por fila (K = valor oficial mes, L = Id ASE, M = 2ª quincena,
///                  N = 1ª quincena) — TODOS VALORES <c>&lt;v&gt;</c>, sin fórmulas ni refs externas
///   fila 31  = K31/M31/N31 = <c>SUM(K26:K30)</c>/<c>SUM(M26:M30)</c>/<c>SUM(N26:N30)</c> (fórmulas)
///   fila 32  = K32 = <c>SUM(M31:N31)</c> (fórmula; M32/N32 ausentes)
/// Valores Q1=Q2=Q2-raíz: [378371975, 533160511, 309577071, 240782166, 257980891];
/// M+N=K ±1 por mitades; ΣK=1719872614; ΣM=859936309; ΣN=859936305.
/// Búsqueda exhaustiva normalizada <c>*nterventoria*</c> en <c>Docs/Insumos/</c> (ambos períodos):
/// SIN FUENTE (V8). → Desenlace D2(b): insumo externo anual declarado; hoja PROTEGIDA intacta +
/// assert estructural (bloque presente con el carácter T0) + log "insumo externo — hoja intacta";
/// NUNCA se escribe ni se inventa. No hay finder/reader/modelo de interventoría (D4 no aplica).
///
/// ── Veredicto T0-L-ESPECIALES (T0-0.5/0.6) ──
/// La columna L del template R1 espeja la columna <c>SERVICIO ESPECIALES</c> de la fuente
/// <c>Recaudoporcomponente_*</c> (misma secuencia de etiquetas/bloques; verificado 5/5 ASE en
/// ambos períodos contra el golden ±0.5). Celdas L fuera del set V4 (HU-07 T0) y del mapa HU-08:
///   - numéricas no-cero que cierran contra la fuente → D3a: mapa editable por Id (abajo).
///   - cero en ambos canónicos → D3b: NO se escriben; el test Capa A A8 fija el 0 (stale-guard).
/// El reader resuelve cada rol por ETIQUETA (col A/B/D/E) con ocurrencia 0-based congelada,
/// distinguiendo "leído 0" de "slot ausente" (fail-fast nombra ASE + hoja + celda).
/// </summary>
public static class WorkbookLeafCellMapInterventoria
{
    public const string HojaInterventoria = Remuneracion.Core.Constants.InterventoriaDeclarada.Hoja;
    public const string HojaR1 = WorkbookLeafCellMap.HojaR1;

    /// <summary>
    /// Fila del header del bloque INTERVENTORIA (T0-0.2: fila 25; K25:L25:M25:N25).
    /// </summary>
    public const int FilaHeader = Remuneracion.Core.Constants.InterventoriaDeclarada.FilaHeader;

    /// <summary>
    /// Primera fila de datos ASE (T0-0.2: filas 26..30 = ASE 1..5).
    /// </summary>
    public const int FilaPrimerAse = Remuneracion.Core.Constants.InterventoriaDeclarada.FilaPrimerAse;

    /// <summary>
    /// Fila de los totales por quincena (T0-0.2: K31/M31/N31 = SUM(K26:K30)…).
    /// </summary>
    public const int FilaTotales = 31;

    /// <summary>
    /// Fila del gran total (T0-0.2: K32 = SUM(M31:N31)).
    /// </summary>
    public const int FilaGranTotal = 32;

    /// <summary>
    /// Valores oficiales de mes por ASE (K26..K30) congelados por T0 — ambos canónicos idénticos.
    /// NO son fuente (no hay archivo de interventoría); son el insumo externo DECLARADO que la
    /// plantilla ya contiene y esta HU preserva (D2b). Alias de <see cref="Remuneracion.Core.Constants.InterventoriaDeclarada"/>.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, decimal> ValorOficialMesPorAse =
        Remuneracion.Core.Constants.InterventoriaDeclarada.ValorOficialMesPorAse;

    /// <summary>
    /// Segunda quincena (M26..M30) por ASE — misma tabla anual (T0-0.2).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, decimal> SegundaQuincenaPorAse =
        Remuneracion.Core.Constants.InterventoriaDeclarada.SegundaQuincenaPorAse;

    /// <summary>
    /// Primera quincena (N26..N30) por ASE — misma tabla anual (T0-0.2).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, decimal> PrimeraQuincenaPorAse =
        Remuneracion.Core.Constants.InterventoriaDeclarada.PrimeraQuincenaPorAse;

    // Historial (Plan 21 T5): este archivo conservaba aquí el mapa L-menores por rol+ocurrencia
    // (enum RolLMenor + LMenoresPorAse/LMenoresPorAseQ2 + ObtenerLMenores) congelado por el
    // T0-0.5 de HU-16. Ese path quedó SUPERSEDED por el espejo estructural R1
    // (ExcelDataReaderWorkbookLeafInputReader.LeerEspejoR1 + OpenXmlEspejoR1Mutador): la columna L
    // del template (encabezado 'SERVICIO ESPECIALES') la escribe el espejo por encabezado en toda
    // la secuencia fuente, y la absorción se probó 15/15 (5 ASE × Q1/Q2/agosto) en
    // EspejoR1AbsorcionTests. Retirado por Plan 21 T5; el resto del archivo (INTERVENTORIA D2b)
    // sigue vigente.

    /// <summary>
    /// Fórmulas protegidas del bloque INTERVENTORIA (D2b) — comunes a ambos períodos:
    /// totales por quincena (K31/M31/N31 = SUM) y gran total (K32 = SUM(M31:N31)). El mapa Q2
    /// (ProtegidasAdicionalesQ2) ya cubre K31/M31/N31; aquí se unifican para validar también Q1
    /// y el gran total K32 en ambos períodos (Requirement 5: INTERVENTORIA siempre protegida).
    /// </summary>
    public static readonly (string Hoja, string Celda, string[] Fragmentos)[] FormulasProtegidas =
    [
        (HojaInterventoria, "K31", ["SUM", "K26", "K30"]),
        (HojaInterventoria, "M31", ["SUM", "M26", "M30"]),
        (HojaInterventoria, "N31", ["SUM", "N26", "N30"]),
        (HojaInterventoria, "K32", ["SUM", "M31", "N31"])
    ];

    /// <summary>
    /// Celda (columna) del bloque por rol: K=valor oficial mes, L=ASE, M=segunda, N=primera.
    /// </summary>
    public static string CeldaBloque(string columna, int fila) => $"{columna}{fila}";

    /// <summary>
    /// Celdas del bloque que DEBEN ser VALORES (no fórmula) por el carácter T0 (D2b): las filas
    /// 26..30 en K/L/M/N. El assert estructural falla si alguna es fórmula (plantilla incompatible).
    /// </summary>
    public static IEnumerable<string> CeldasValoresBloque()
    {
        foreach (var fila in Enumerable.Range(FilaPrimerAse, 5))
        {
            foreach (var col in new[] { "K", "L", "M", "N" })
            {
                yield return $"{col}{fila}";
            }
        }
    }
}
