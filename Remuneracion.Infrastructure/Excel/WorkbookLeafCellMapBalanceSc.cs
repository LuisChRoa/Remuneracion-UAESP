namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// HU-10 (2.4): mapa de celdas de la hoja <c>BCE SC POR FACT.</c> congelado por T0-0.7
/// (evidencia en <c>plans/10 - HU-10 BCE balance subsidios Fase 2.md</c> §0.2/§4 Fase 0 y en el
/// dump OpenXML del golden <c>Docs/Insumos/Remuneracion 202607-1 Total.xlsx</c> sheet9 + lectura
/// de los 5 <c>R4-BalanceSubsidioyContribuciones_*</c> Q1).
///
/// Veredictos T0 (probados, no asumidos):
/// - T0-0.1 valor-vs-fórmula: C3:C7 = valores (ids ASE 1..5, ya presentes → solo se verifican);
///   D3:D7 / E3:E7 = VALORES editables (contribución/subsidio); F3:F7 = fórmulas <c>Dn+En</c>;
///   H3:H7 = fórmulas <c>DetRetri2026071!J9..J13</c>; I3:I7 = fórmulas <c>Fn-Hn</c>; filas 9/10/11/12/13
///   y bloque 18–24 = fórmulas → TODAS protegidas.
/// - T0-0.2 veredicto D/E = <b>CORREGIDO (Plan 29 T1 / T0-V4)</b>: el header de la plantilla
///   MANDA — <c>D2=SUBSIDIO</c>, <c>E2=CONTRIBUCION</c>. Por tanto template-D ← columna E-fuente
///   (Subsidio, negativo) y template-E ← columna F-fuente (Contribución, positivo).
///   El veredicto original del Plan 10 ("template-D = CONTRIBUCION") quedó refutado por el header
///   + el archivo manual (10/10 celdas en ambos períodos, T0 V4); como <c>F = D + E</c> es
///   conmutativa, el golden ±0.5 vigente no podía detectar el swap (solo permuta D/E visibles).
///   El reader NUNCA estuvo invertido (<see cref="ColumnaSubsidioFuente"/> = 4 = E-fuente).
/// - T0-0.6 veredicto H = D2(b): H es FÓRMULA → protegida, jamás se escribe; solo gate H≈F (Capa A).
///
/// Columnas fuente (V7 + T0-0.3): B=productor, C=usuarios-res, D=usuarios-no-res, E=Subsidio,
/// F=Contribución, G=Valor. La etiqueta "Total General" está en la ÚLTIMA fila del Sheet1 en los
/// 5 ASE (R118/R110/R38/R42/R28) → búsqueda dinámica desde el final, nunca fila fija (D1).
/// </summary>
public static class WorkbookLeafCellMapBalanceSc
{
    public const string HojaBce = "BCE SC POR FACT.";

    /// <summary>
    /// Etiqueta canónica "Total General" normalizada (minúsculas, sin espacios) para el match
    /// exacto desde el final del Sheet1 (T0-0.4). Match EXACTO, no por prefijo: las filas
    /// "Total" de cada localidad NO deben matchear.
    /// </summary>
    public const string EtiquetaTotalGeneral = "totalgeneral";

    /// <summary>
    /// Índices de columna (0-based) de la fila "Total General" en la fuente (T0-0.3): las 5
    /// fuentes Q1 tienen el mismo layout B..G (B=productor vacío en la fila total).
    /// </summary>
    public const int ColumnaSubsidioFuente = 4;     // E-fuente (negativo)
    public const int ColumnaContribucionFuente = 5; // F-fuente (positivo)
    public const int ColumnaTotalFuente = 6;        // G-fuente (Valor = E + F)

    /// <summary>
    /// Celdas editables por ASE (T0-0.7): fila = 2 + <see cref="Ase.Id"/> (C3=1..C7=5 en el
    /// template). D = SUBSIDIO (← E-fuente), E = CONTRIBUCION (← F-fuente) según el header de la
    /// plantilla y el veredicto corregido T0-V4 del Plan 29. La tupla conserva los nombres de
    /// dominio (<c>Contribucion</c>/<c>Subsidio</c>): <c>Subsidio</c> resuelve a la celda D y
    /// <c>Contribucion</c> a la celda E. La columna C (id ASE) ya trae el valor correcto en el
    /// template → solo se verifica. La columna H (SISTEMA) es fórmula (D2(b)) → fuera de editables.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Contribucion, string Subsidio)> EditablesPorAse =
        new Dictionary<int, (string, string)>
        {
            [1] = ("E3", "D3"),
            [2] = ("E4", "D4"),
            [3] = ("E5", "D5"),
            [4] = ("E6", "D6"),
            [5] = ("E7", "D7")
        };

    // ── Plan 59 (H5): la foto congelada de fórmulas protegidas salió del camino productivo ──
    //
    // Aquí vivía `Protegidas`: 78 celdas literales de la geometría v1 (F3:F7, H3:H7, I3:I7, filas
    // 9/10/11/12/13, bloque 18-24, CONSOLIDADO J/K/M y las referencias cruzadas de DetRetri y
    // DetValiRetri) cuyos fragmentos escribían el nombre de hoja de UN PERÍODO CONCRETO.
    // El camino productivo ya no la consume: `VerificadorBalanceScDerivado` verifica las mismas
    // protecciones por IDENTIDAD + FORMA sobre lo que la hoja declara, y resuelve la hoja del período
    // leyéndola del workbook (`HojasPeriodo`).
    //
    // La foto queda como material histórico de fixtures (`BalanceScGoldenV1`) porque es la evidencia
    // del T0 con la que se catalogaron los roles; no es fuente de coordenadas para producción.
    /// <summary>
    /// Obtiene las celdas editables (D=Subsidio, E=Contribucion) de un ASE; lanza si no está soportado.
    /// </summary>
    public static (string Contribucion, string Subsidio) ObtenerEditables(int aseId) =>
        EditablesPorAse.TryGetValue(aseId, out var mapa)
            ? mapa
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay cell-map de balance SC para el ASE {aseId}.");
}
