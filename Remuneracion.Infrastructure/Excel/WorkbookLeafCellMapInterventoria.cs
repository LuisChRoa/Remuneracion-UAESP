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

    /// <summary>
    /// Rol de la fila FUENTE (Recaudoporcomponente) que alimenta una L-menor del template.
    /// Resuelto por etiqueta (col A/B/D/E) con ocurrencia 0-based dentro de la zona Componente
    /// (filas anteriores a la fila A='Componente' B='Total'); TotalFinal se resuelve DESPUÉS de
    /// la fila Componente (fila A='Total' B vacío). Ocurrencias congeladas por T0-0.5.
    /// </summary>
    public enum RolLMenor
    {
        /// <summary>Fila E='Vlr Servicio' (enésima ocurrencia).</summary>
        VlrServicio,

        /// <summary>Fila E='Vlr Intereses' (enésima ocurrencia).</summary>
        VlrIntereses,

        /// <summary>Fila D='E' ∧ E='Total' (bloque ENEL).</summary>
        TotalD_E,

        /// <summary>Fila D='H' ∧ E='Total' (bloque H).</summary>
        TotalD_H,

        /// <summary>Fila D='O' ∧ E='Total' (bloque O).</summary>
        TotalD_O,

        /// <summary>Fila D='T' ∧ E='Total' (bloque OCCIDENTE).</summary>
        TotalD_T,

        /// <summary>Fila D='Total' ∧ E vacío (fila display del Total).</summary>
        TotalDisplay,

        /// <summary>Fila A='Componente' B='Total' (cierre del bloque Componente).</summary>
        CompTotal,

        /// <summary>Fila A='Total' B vacío (fila Total final, después de la zona Componente).</summary>
        TotalFinal
    }

    /// <summary>
    /// L-menores Q1 por ASE (mapa congelado T0-0.5): celda del template → (rol fuente, ocurrencia).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Celda, RolLMenor Rol, int Ocurrencia)[]> LMenoresPorAse =
        new Dictionary<int, (string, RolLMenor, int)[]>
        {
            [1] =
            [
                ("L11", RolLMenor.VlrServicio, 1), ("L12", RolLMenor.VlrIntereses, 1),
                ("L13", RolLMenor.TotalD_E, 0), ("L21", RolLMenor.VlrServicio, 4),
                ("L22", RolLMenor.VlrIntereses, 4), ("L23", RolLMenor.TotalD_T, 1),
                ("L26", RolLMenor.CompTotal, 0), ("L43", RolLMenor.TotalFinal, 0)
            ],
            [2] =
            [
                ("L92", RolLMenor.VlrServicio, 3), ("L93", RolLMenor.VlrIntereses, 0),
                ("L94", RolLMenor.TotalD_E, 1), ("L106", RolLMenor.VlrServicio, 7),
                ("L107", RolLMenor.VlrIntereses, 4), ("L108", RolLMenor.TotalD_O, 1),
                ("L109", RolLMenor.VlrServicio, 8), ("L111", RolLMenor.TotalD_T, 0),
                ("L114", RolLMenor.CompTotal, 0), ("L132", RolLMenor.TotalFinal, 0)
            ],
            [3] =
            [
                ("L224", RolLMenor.VlrServicio, 3), ("L225", RolLMenor.VlrIntereses, 1),
                ("L226", RolLMenor.TotalD_E, 1), ("L231", RolLMenor.VlrServicio, 5),
                ("L233", RolLMenor.TotalD_O, 1), ("L234", RolLMenor.VlrServicio, 6),
                ("L235", RolLMenor.VlrIntereses, 4), ("L236", RolLMenor.TotalD_T, 1),
                ("L239", RolLMenor.CompTotal, 0), ("L256", RolLMenor.TotalFinal, 0)
            ],
            [4] =
            [
                ("L370", RolLMenor.VlrServicio, 5), ("L371", RolLMenor.VlrIntereses, 2),
                ("L372", RolLMenor.TotalD_E, 2), ("L384", RolLMenor.VlrServicio, 9),
                ("L385", RolLMenor.VlrIntereses, 6), ("L386", RolLMenor.TotalD_O, 2),
                ("L387", RolLMenor.VlrServicio, 10), ("L389", RolLMenor.TotalD_T, 1),
                ("L392", RolLMenor.CompTotal, 0), ("L419", RolLMenor.TotalFinal, 0)
            ],
            [5] =
            [
                ("L480", RolLMenor.VlrServicio, 2), ("L481", RolLMenor.VlrIntereses, 0),
                ("L482", RolLMenor.TotalD_E, 1), ("L485", RolLMenor.VlrIntereses, 1),
                ("L486", RolLMenor.TotalD_E, 2), ("L494", RolLMenor.VlrServicio, 6),
                ("L495", RolLMenor.VlrIntereses, 4), ("L496", RolLMenor.TotalD_T, 0),
                ("L499", RolLMenor.CompTotal, 0), ("L515", RolLMenor.TotalFinal, 0)
            ]
        };

    /// <summary>
    /// L-menores Q2 por ASE (mapa congelado T0-0.5 sobre el canónico "Plantilla 8 agos…" y el
    /// caché golden <c>Remuneracion 202607-2 Total.xlsx</c>; fuente R1-Q2 del ASE).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Celda, RolLMenor Rol, int Ocurrencia)[]> LMenoresPorAseQ2 =
        new Dictionary<int, (string, RolLMenor, int)[]>
        {
            [1] =
            [
                ("L18", RolLMenor.VlrServicio, 3), ("L19", RolLMenor.VlrIntereses, 0),
                ("L20", RolLMenor.TotalD_E, 2), ("L21", RolLMenor.TotalDisplay, 3),
                ("L25", RolLMenor.VlrServicio, 5), ("L27", RolLMenor.TotalD_O, 1),
                ("L28", RolLMenor.VlrServicio, 6), ("L29", RolLMenor.VlrIntereses, 3),
                ("L30", RolLMenor.TotalD_T, 0), ("L31", RolLMenor.TotalDisplay, 4),
                ("L33", RolLMenor.CompTotal, 0), ("L50", RolLMenor.TotalFinal, 0)
            ],
            [2] =
            [
                ("L108", RolLMenor.VlrServicio, 5), ("L109", RolLMenor.VlrIntereses, 3),
                ("L110", RolLMenor.TotalD_E, 2), ("L111", RolLMenor.TotalDisplay, 4),
                ("L122", RolLMenor.VlrServicio, 9), ("L123", RolLMenor.VlrIntereses, 7),
                ("L124", RolLMenor.TotalD_H, 0), ("L125", RolLMenor.VlrServicio, 10),
                ("L127", RolLMenor.TotalD_O, 2), ("L128", RolLMenor.VlrServicio, 11),
                ("L129", RolLMenor.VlrIntereses, 9), ("L130", RolLMenor.TotalD_T, 1),
                ("L131", RolLMenor.TotalDisplay, 6), ("L133", RolLMenor.CompTotal, 0),
                ("L162", RolLMenor.TotalFinal, 0)
            ],
            [3] =
            [
                ("L251", RolLMenor.VlrServicio, 2), ("L252", RolLMenor.VlrIntereses, 1),
                ("L253", RolLMenor.TotalD_E, 1), ("L254", RolLMenor.TotalDisplay, 2),
                ("L255", RolLMenor.VlrServicio, 3), ("L257", RolLMenor.TotalD_H, 0),
                ("L258", RolLMenor.VlrServicio, 4), ("L260", RolLMenor.TotalD_O, 1),
                ("L261", RolLMenor.VlrServicio, 5), ("L262", RolLMenor.VlrIntereses, 4),
                ("L263", RolLMenor.TotalD_T, 0), ("L264", RolLMenor.TotalDisplay, 3),
                ("L266", RolLMenor.CompTotal, 0), ("L283", RolLMenor.TotalFinal, 0)
            ],
            [4] =
            [
                ("L398", RolLMenor.VlrServicio, 5), ("L399", RolLMenor.VlrIntereses, 2),
                ("L400", RolLMenor.TotalD_E, 2), ("L401", RolLMenor.TotalDisplay, 5),
                ("L415", RolLMenor.VlrServicio, 10), ("L417", RolLMenor.TotalD_O, 2),
                ("L418", RolLMenor.VlrServicio, 11), ("L419", RolLMenor.VlrIntereses, 8),
                ("L420", RolLMenor.TotalD_T, 0), ("L421", RolLMenor.TotalDisplay, 7),
                ("L423", RolLMenor.CompTotal, 0), ("L450", RolLMenor.TotalFinal, 0)
            ],
            [5] =
            [
                ("L513", RolLMenor.VlrServicio, 3), ("L514", RolLMenor.VlrIntereses, 2),
                ("L515", RolLMenor.TotalD_E, 0), ("L516", RolLMenor.TotalDisplay, 1),
                ("L524", RolLMenor.VlrServicio, 6), ("L526", RolLMenor.TotalD_O, 1),
                ("L527", RolLMenor.VlrServicio, 7), ("L528", RolLMenor.VlrIntereses, 6),
                ("L529", RolLMenor.TotalD_T, 1), ("L530", RolLMenor.TotalDisplay, 3),
                ("L532", RolLMenor.CompTotal, 0), ("L554", RolLMenor.TotalFinal, 0)
            ]
        };

    /// <summary>
    /// Obtiene el mapa L-menores del período (Q1/Q2); lanza si el ASE no está soportado.
    /// </summary>
    public static (string Celda, RolLMenor Rol, int Ocurrencia)[] ObtenerLMenores(int aseId, int numeroQuincena)
    {
        var mapa = numeroQuincena == 2 ? LMenoresPorAseQ2 : LMenoresPorAse;
        return mapa.TryGetValue(aseId, out var celdas)
            ? celdas
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay cell-map de L-menores para el ASE {aseId}.");
    }

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
