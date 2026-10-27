using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// HU-13 (2.7): mapa-oráculo de validaciones cruzadas, hermano de los mapas HU-07..HU-12.
/// Congelado por T0 (Plan 13 §4 Fase 0) con evidencia en AMBOS canónicos (Q1 golden y Q2
/// valores golden + canónico "8 agos").
///
/// Veredictos T0 (probados, no asumidos):
/// - T0-0.2: <c>VALIDACION_*</c> ×5: filas 3..7 = ASE 1..5 (columna B numérica), fila 8 = "X",
///   fila 9 = Σ. Columna O = H−N (Recaudo * vs REMUNERACION_*), columna P = INT(O)=0 (booleano).
///   Caché Q1/Q2 golden: O = 0 y P = TRUE en toda fila-ASE. Refs Recaudo por columna D (Q1)
///   vs F (Q2) — el mapa-oráculo NO depende de la columna Recaudo (solo lee O/P).
/// - T0-0.3: <c>DetValiRetri2026071/2026072</c> (sheets 36/37): col D rows 16..20 = D9..D13 vs
///   CONSOLIDADO U104..U108 (diferencias float ≤ ±0.5 en ambos goldens); D24..D28 booleanos de
///   composición TRUE; D21 (fila Total) DIVERGE en ambos canónicos (Q1 −0.62, Q2 −1.71) y D9:D14
///   son valores fuente ≠ ROUND (hallazgo HU-12) → EXCLUIDOS del gate, documentados (D6).
/// - T0-0.4: <c>VALIDACION_TOTAL</c>: filas 3..7 por ASE + fila 9 Σ con la misma semántica O/P
///   (O9 = 0, P9 = TRUE en ambos goldens); sub-bloques booleanos (C15/D25/O25/D34/F34) TRUE.
/// - T0-0.5: <c>Valida -Remunera / -Anticipos / -Control Recaudo</c> = vistas lado-a-lado
///   (CONSOLIDADO vs REMUNERACION_*) SIN celda de diferencia/booleana estable → protegidas
///   (fórmulas presentes donde aplique), sin gate numérico (recorte honesto). Valida - Control
///   Recaudo mezcla VALORES (F10) con booleanos (F7) → protegida-valor: no exigir fórmula.
/// - T0-0.6: <c>GERENTES_*</c> ×5 = fórmulas puras (espejo REMUNERACION_*); cierra el pendiente
///   Plan 08 T0-0.5. Sin gate numérico.
///
/// Fe de erratas Plan 12 (G7/A8): las hojas DetRetri/DetValiRetri son las sheets 36/37; todo
/// artefacto nuevo usa los NOMBRES con sufijo 2026071/2026072 (nunca números 93/94).
/// </summary>
public static class WorkbookLeafCellMapValidaciones
{
    public const string HojaValidacionTotal = "VALIDACION_TOTAL";
    public const string HojaValidaRemunera = "Valida -Remunera";
    public const string HojaValidaAnticipos = "Valida - Anticipos";
    public const string HojaValidaControlRecaudo = "Valida - Control Recaudo";

    /// <summary>
    /// Plan 29 (T5, Unidad P — R-P-3/T0f): celda <c>'Valida - Control Recaudo'!F10</c>, literal de
    /// control que <c>VALIDACION_TOTAL!C14</c> consume en la cadena <c>C15 = C14 = C9</c>. Su origen
    /// workbook-interno (por TEXTO de fórmula, D-E) es <c>VALIDACION_TOTAL!C9 = SUM(C3:C8)</c>, donde
    /// <c>C3:C8</c> = Σ de las hojas <c>Recaudo *</c> <c>F21:F26</c> (bloque TOTAL). El writer la
    /// puebla como literal; sin ella, C9 (que recibe los recaudos del app) ≠ 0 = F10 ⇒ C15 False.
    /// </summary>
    public const string CeldaControlRecaudoTotal = "F10";

    /// <summary>
    /// Plan 48 (H1.2 — T0 §1.5/§4.2): filas del bloque de control de recaudo de
    /// <c>Valida - Control Recaudo</c>. En la v2 la plantilla deja <c>D7:D10 = 0</c> y convierte
    /// <c>E7 = SUM(D7:D9)</c>, <c>E10 = SUM(D10)</c> y <c>F10 = E7+E10</c> en fórmulas. La App debe
    /// poblar <c>D7:D10</c> desde la fila 27 (TOTAL) de cada <c>Recaudo *</c> — la fila se localiza
    /// POR ETIQUETA de la columna <c>C</c>, nunca por índice. Tabla de DATOS citada por bytes
    /// (manual v2: D7=68.144.618.192, D8=28.645.323, D9=40.117.400, D10=4.284.569.052).
    /// El label <c>ENEL / RECIP.</c> mapea a ENEL (id 2) y <c>EAAB / RECIP.</c> a Reciprocidad (id 1):
    /// son los rótulos heredados del template, no una regla por período/empresa nueva.
    /// </summary>
    public static readonly IReadOnlyList<(string Etiqueta, int EmpresaId)> ControlRecaudoPorEtiqueta =
    [
        ("ENEL / RECIP.", 2),
        ("ENERBIT", 3),
        ("EAAB / RECIP.", 1),
        ("DIRECTA OCCIDENTE", 4),
    ];

    /// <summary>
    /// Plan 48 (H1b — S-1 del review): celdas cuyo <c>&lt;v&gt;</c> sella el writer tras escribir
    /// <c>D7:D10</c> con el valor EVALUADO por read-back (jamás inventado). Punto ÚNICO de
    /// congelamiento: el writer consume esta lista, no un arreglo inline. En la v2 las tres son
    /// fórmulas (<c>E7=SUM(D7:D9)</c>, <c>E10=SUM(D10)</c>, <c>F10=E7+E10</c>); en la v1 son literales
    /// y el sellado es no-op.
    /// </summary>
    public static readonly IReadOnlyList<string> CeldasSelladoControlRecaudo = ["E7", "E10", "F10"];

    /// <summary>
    /// HU-14 (W-1): sub-bloques booleanos de <c>VALIDACION_TOTAL</c> gateados TRUE exacto por
    /// ASE. Amparo T0-0.4 HU-13 (TRUE en ambos goldens; re-verificado por los tests W-1).
    /// </summary>
    public static readonly IReadOnlyList<string> SubBloquesValidacionTotal = ["C15", "D25", "O25", "D34", "F34"];

    /// <summary>
    /// Hoja DetRetri del período (D-A, Plan 30): se compone del dominio vía
    /// <see cref="NombresHojaPeriodo.DetRetri(Periodo)"/>, nunca se detecta enumerando hojas.
    /// </summary>
    public static string HojaDetRetri(Periodo periodo) => NombresHojaPeriodo.DetRetri(periodo);

    /// <summary>
    /// Hoja DetValiRetri del período (D-A, Plan 30): se compone del dominio vía
    /// <see cref="NombresHojaPeriodo.DetValiRetri(Periodo)"/>.
    /// </summary>
    public static string HojaDetValiRetri(Periodo periodo) => NombresHojaPeriodo.DetValiRetri(periodo);

    /// <summary>
    /// Fila de la hoja VALIDACION_* / VALIDACION_TOTAL que corresponde a un ASE
    /// (T0: filas 3..7 = ASE 1..5).
    /// </summary>
    public static int FilaValidacionPorAse(int aseId) => aseId + 2;

    /// <summary>
    /// Fila D de DetValiRetri con la diferencia del ASE (T0: D16..D20 = 15+aseId).
    /// </summary>
    public static int FilaDiferenciaDetValiRetri(int aseId) => 15 + aseId;

    /// <summary>
    /// Fila D de DetValiRetri con la verificación booleana del ASE (T0: D24..D28 = 23+aseId).
    /// </summary>
    public static int FilaVerificacionDetValiRetri(int aseId) => 23 + aseId;

    /// <summary>
    /// Mapa protegido extendido 2.7 (G2/D4): TODAS las hojas de validación entran al mapa
    /// protegido por período (el writer falla si alguna deja de ser fórmula donde T0 exige
    /// fórmula). No duplica el mapa HU-08/HU-12 existente (VALIDACION_* O3 / GERENTES D9 /
    /// DetValiRetri-Q2 ya cubiertos); suma lo que faltaba: P3 y filas 4..7 de VALIDACION_*,
    /// VALIDACION_TOTAL O/P, detalle col-D DetRetri/DetValiRetri Q1 y representantes Valida -*.
    /// </summary>
    public static IReadOnlyList<(string Hoja, string Celda, string[] Fragmentos)> ProtegidasValidacionesParaPeriodo(Periodo periodo)
    {
        var lista = new List<(string, string, string[])>();

        // VALIDACION_* ×5: columnas O/P en las filas 3..7 (ASE 1..5). La fila 3 suele traer el
        // texto maestro; las filas 4..7 followers con shared-index → presencia de <f> basta.
        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var fila = FilaValidacionPorAse(aseId);
                lista.Add((empresa.HojaValidacion, $"O{fila}", fila == 3 ? ["H3", "N3"] : []));
                lista.Add((empresa.HojaValidacion, $"P{fila}", fila == 3 ? ["INT(O3)=0"] : []));
            }
        }

        // VALIDACION_TOTAL: filas 3..9 (ASE 1..5 + X + Σ) columnas O/P.
        for (var fila = 3; fila <= 9; fila++)
        {
            lista.Add((HojaValidacionTotal, $"O{fila}", fila == 9 ? ["SUM", "O3", "O8"] : []));
            lista.Add((HojaValidacionTotal, $"P{fila}", []));
        }

        // HU-14 (W-1): sub-bloques booleanos de VALIDACION_TOTAL (C15/D25/O25/D34/F34) entran
        // al mapa protegido con presencia de <f> (assert W2); el gate TRUE exacto vive en el
        // validador (amparo T0-0.4 HU-13).
        foreach (var celda in SubBloquesValidacionTotal)
        {
            lista.Add((HojaValidacionTotal, celda, []));
        }

        // Plan 59 (§11 — FRENTE 2): aquí vivía el mapa Q1 del detalle del período con CELDAS Y
        // OPERANDOS congelados (DetRetri D23..D28 'CONSOLIDADO D104+i', D32..D36; DetValiRetri
        // D16..D21 'CONSOLIDADO U104+i', D24..D29). Se retiró: la verificación por ROL + FORMA
        // (`DetallePeriodoDerivado.Validar`) corre ahora en AMBAS quincenas desde el tronco común del
        // writer, cubre esas mismas columnas y no caduca con la próxima geometría (§2-D2.3-bis:
        // prohibido "actualizar la lista").

        // Valida -*: representantes de fórmula (vistas lado-a-lado; T0-0.5 sin gate numérico).
        // Valida - Control Recaudo mezcla VALORES → no se exige fórmula (protegida-valor).
        lista.Add((HojaValidaRemunera, "D9", ["CONSOLIDADO", "D104"]));
        lista.Add((HojaValidaRemunera, "D26", ["D104"]));
        lista.Add((HojaValidaAnticipos, "D6", ["CONSOLIDADO", "D125"]));
        lista.Add((HojaValidaAnticipos, "D21", ["D125"]));

        // GERENTES_*: fórmulas puras (T0-0.6) — D9 ya protegida por HU-08; se suma el total
        // SUM por hoja (D14 es SUM(D9:D13) en ambos canónicos).
        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            lista.Add((empresa.HojaGerentes, "D14", ["SUM", "D9", "D13"]));
        }

        return lista;
    }
}
