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
/// - T0-0.2 veredicto D/E = HIPÓTESIS LÍDER PROBADA en los 5 ASE (fuente vs golden ±0.5):
///   template-D (CONTRIBUCION, positivo) ← columna F-fuente; template-E (SUBSIDIO, negativo) ←
///   columna E-fuente. El texto del doc base (D=subsidio/E=contribución) queda descartado por prueba.
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
    /// template). D = CONTRIBUCION (← F-fuente), E = SUBSIDIO (← E-fuente) según veredicto
    /// T0-0.2. La columna C (id ASE) ya trae el valor correcto en el template → solo se verifica.
    /// La columna H (SISTEMA) es fórmula (D2(b)) → fuera de editables.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Contribucion, string Subsidio)> EditablesPorAse =
        new Dictionary<int, (string, string)>
        {
            [1] = ("D3", "E3"),
            [2] = ("D4", "E4"),
            [3] = ("D5", "E5"),
            [4] = ("D6", "E6"),
            [5] = ("D7", "E7")
        };

    /// <summary>
    /// Fórmulas protegidas 2.4 (jamás se escriben; D6): F3:F7 (D+E), H3:H7 (DetRetri), I3:I7
    /// (diferencia, shared-follower en I5:I7), filas 9/10/11/12/13, bloque 18–24 de validación
    /// (2.7), CONSOLIDADO J9:J13 + K9:K13 + M9:M13 (downstream automático, V5) y las fórmulas de
    /// DetRetri/DetValiRetri que referencian <c>'BCE SC POR FACT.'!F3..F7</c>. Fragmentos con la
    /// hoja normalizada ("BCE SC POR FACT." → "BCESCPORFACT." tras quitar comillas y espacios).
    /// </summary>
    public static readonly (string Hoja, string Celda, string[] Fragmentos)[] Protegidas =
    [
        // Filas ASE 3..7: F = D+E (protector del vínculo con el mapa editable).
        (HojaBce, "F3", ["D3", "E3"]),
        (HojaBce, "F4", ["D4", "E4"]),
        (HojaBce, "F5", ["D5", "E5"]),
        (HojaBce, "F6", ["D6", "E6"]),
        (HojaBce, "F7", ["D7", "E7"]),
        // H = SISTEMA (fórmula → veredicto D2(b)).
        (HojaBce, "H3", ["DetRetri2026071", "J9"]),
        (HojaBce, "H4", ["DetRetri2026071", "J10"]),
        (HojaBce, "H5", ["DetRetri2026071", "J11"]),
        (HojaBce, "H6", ["DetRetri2026071", "J12"]),
        (HojaBce, "H7", ["DetRetri2026071", "J13"]),
        // I = diferencia F−H (I3/I4 explícitas; I5..I7 shared-followers → solo presencia).
        (HojaBce, "I3", ["F3", "H3"]),
        (HojaBce, "I4", ["F4", "H4"]),
        (HojaBce, "I5", []),
        (HojaBce, "I6", []),
        (HojaBce, "I7", []),
        // Fila 9: total F/H + diferencia.
        (HojaBce, "F9", ["SUM", "F3", "F7"]),
        (HojaBce, "H9", ["SUM", "H3", "H7"]),
        (HojaBce, "I9", ["F9", "H9"]),
        // Fila 10: verificación H9 == DetRetri J14.
        (HojaBce, "H10", ["H9", "DetRetri2026071", "J14"]),
        // Fila 11: sumas D/E/F.
        (HojaBce, "D11", ["SUM", "D3", "D7"]),
        (HojaBce, "E11", ["SUM", "E3", "E7"]),
        (HojaBce, "F11", ["SUM", "D3", "E7"]),
        // Filas 12/13: verificación ROUND.
        (HojaBce, "F12", ["ROUND", "F11", "F9"]),
        (HojaBce, "F13", ["ROUND", "F11", "F9"]),
        // Bloque 18–24: validación vs CONSOLIDADO J104:J108 (2.7, solo se protege).
        (HojaBce, "E19", ["CONSOLIDADO", "J104"]),
        (HojaBce, "E20", ["CONSOLIDADO", "J105"]),
        (HojaBce, "E21", ["CONSOLIDADO", "J106"]),
        (HojaBce, "E22", ["CONSOLIDADO", "J107"]),
        (HojaBce, "E23", ["CONSOLIDADO", "J108"]),
        (HojaBce, "E24", ["SUM", "E19", "E23"]),
        (HojaBce, "F19", ["E19", "F3"]),
        (HojaBce, "F20", ["E20", "F4"]),
        (HojaBce, "F21", ["E21", "F5"]),
        (HojaBce, "F22", ["E22", "F6"]),
        (HojaBce, "F23", ["E23", "F7"]),
        (HojaBce, "F24", ["E24", "F9"]),
        (HojaBce, "G19", ["E19", "H3"]),
        (HojaBce, "G20", ["E20", "H4"]),
        (HojaBce, "G21", ["E21", "H5"]),
        (HojaBce, "G22", ["E22", "H6"]),
        (HojaBce, "G23", ["E23", "H7"]),
        (HojaBce, "G24", ["E24", "H9"]),
        (HojaBce, "H19", ["E19", "H3"]),
        (HojaBce, "H20", ["E20", "H4"]),
        (HojaBce, "H21", ["E21", "H5"]),
        (HojaBce, "H22", ["E22", "H6"]),
        (HojaBce, "H23", ["E23", "H7"]),
        (HojaBce, "H24", ["E24", "H9"]),
        // CONSOLIDADO downstream (V5): J9:J13 = BCE F3:F7; K = D−E−F−G−H−I−J; M = K−L.
        ("CONSOLIDADO_TOTAL RECAUDO", "J9", ["BCESCPORFACT.!F3"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "J10", ["BCESCPORFACT.!F4"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "J11", ["BCESCPORFACT.!F5"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "J12", ["BCESCPORFACT.!F6"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "J13", ["BCESCPORFACT.!F7"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "K9", ["J9"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "K10", ["J10"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "K11", ["J11"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "K12", ["J12"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "K13", ["J13"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "M9", ["K9", "L9"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "M10", ["K10", "L10"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "M11", ["K11", "L11"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "M12", ["K12", "L12"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "M13", ["K13", "L13"]),
        // DetRetri2026071 (sheet36): fórmulas grandes que referencian BCE F3..F7 (2.6, protegidas).
        ("DetRetri2026071", "J32", ["BCESCPORFACT.!F3"]),
        ("DetRetri2026071", "J33", ["BCESCPORFACT.!F4"]),
        ("DetRetri2026071", "J34", ["BCESCPORFACT.!F5"]),
        ("DetRetri2026071", "J35", ["BCESCPORFACT.!F6"]),
        ("DetRetri2026071", "J36", ["BCESCPORFACT.!F7"]),
        ("DetRetri2026071", "K32", ["BCESCPORFACT.!F3"]),
        ("DetRetri2026071", "K33", ["BCESCPORFACT.!F4"]),
        ("DetRetri2026071", "K34", ["BCESCPORFACT.!F5"]),
        ("DetRetri2026071", "K35", ["BCESCPORFACT.!F6"]),
        ("DetRetri2026071", "K36", ["BCESCPORFACT.!F7"]),
        ("DetRetri2026071", "M32", ["BCESCPORFACT.!F3"]),
        ("DetRetri2026071", "M33", ["BCESCPORFACT.!F4"]),
        ("DetRetri2026071", "M34", ["BCESCPORFACT.!F5"]),
        ("DetRetri2026071", "M35", ["BCESCPORFACT.!F6"]),
        ("DetRetri2026071", "M36", ["BCESCPORFACT.!F7"]),
        // DetValiRetri2026071 (sheet37): fórmulas grandes que referencian BCE F3..F7 (2.7, protegidas).
        ("DetValiRetri2026071", "M24", ["BCESCPORFACT.!F3"]),
        ("DetValiRetri2026071", "M25", ["BCESCPORFACT.!F4"]),
        ("DetValiRetri2026071", "M26", ["BCESCPORFACT.!F5"]),
        ("DetValiRetri2026071", "M27", ["BCESCPORFACT.!F6"]),
        ("DetValiRetri2026071", "M28", ["BCESCPORFACT.!F7"]),
        ("DetValiRetri2026071", "M29", ["BCESCPORFACT.!F3"]),
        ("DetValiRetri2026071", "O24", ["BCESCPORFACT.!F3"]),
        ("DetValiRetri2026071", "O25", ["BCESCPORFACT.!F4"]),
        ("DetValiRetri2026071", "O26", ["BCESCPORFACT.!F5"]),
        ("DetValiRetri2026071", "O27", ["BCESCPORFACT.!F6"]),
        ("DetValiRetri2026071", "O28", ["BCESCPORFACT.!F7"]),
        ("DetValiRetri2026071", "O29", ["BCESCPORFACT.!F3"])
    ];

    /// <summary>
    /// Obtiene las celdas editables (D=Contribucion, E=Subsidio) de un ASE; lanza si no está soportado.
    /// </summary>
    public static (string Contribucion, string Subsidio) ObtenerEditables(int aseId) =>
        EditablesPorAse.TryGetValue(aseId, out var mapa)
            ? mapa
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay cell-map de balance SC para el ASE {aseId}.");
}
