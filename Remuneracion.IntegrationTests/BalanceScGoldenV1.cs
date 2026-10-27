namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (H5, §0.1/§0.2): FOTO HISTÓRICA de las protecciones de la hoja <c>BCE SC POR FACT.</c>
/// (78 celdas de la geometría v1, con los nombres de hoja del período que la v1 declaraba).
///
/// La foto vivió en producción hasta H5 (<c>WorkbookLeafCellMapBalanceSc.Protegidas</c>) y se usaba
/// como gate de "esta celda sigue siendo fórmula". Salió del camino productivo porque codificaba
/// coordenadas y un nombre de hoja de período: el gate real es ahora
/// <c>VerificadorBalanceScDerivado</c> (identidad + forma, hoja del período leída del workbook).
///
/// Se conserva SOLO como evidencia histórica del T0 (los fixtures de catalogación la citan para
/// demostrar que el mapa catalogado sigue describiendo el golden). No es fuente de coordenadas para
/// producción y ningún camino de escritura la consume.
/// </summary>
internal static class BalanceScGoldenV1
{
    /// <summary>Hoja del mapa histórico (vocabulario de la plantilla).</summary>
    private const string HojaBce = "BCE SC POR FACT.";

    /// <summary>Celdas protegidas según la foto v1: (hoja, celda, fragmentos esperados).</summary>
    internal static readonly (string Hoja, string Celda, string[] Fragmentos)[] Protegidas =
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
        // DetRetri{periodo} de la v1: fórmulas grandes que referencian BCE F3..F7.
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
        // DetValiRetri{periodo} de la v1: fórmulas grandes que referencian BCE F3..F7.
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
}
