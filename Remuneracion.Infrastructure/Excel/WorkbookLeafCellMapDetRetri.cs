using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 29 (T4, Unidad D — desglose DetRetri/DetValiRetri trazable): mapa de destino y de
/// procedencia de las columnas C..O de <c>DetRetri2026072</c> y D..O de
/// <c>DetValiRetri2026072</c>.
///
/// Gate T0d (<c>plans/29-T0-Evidencia-B.md</c> §1.4): cada columna del MANUAL es igual, ±0.5, a
/// la suma de la fila <c>104+(ASE-1)</c> de <c>CONSOLIDADO_TOTAL RECAUDO</c> (cols D..O para
/// DetRetri; U..AE para DetValiRetri). Esa fila 104 es, a su vez, la suma de 5 filas-sección
/// (9/28/47/66/85) cuyas celdas son FÓRMULAS workbook-internas
/// (<c>='Reporte Componentes R1'!G53</c>, <c>='Rem. Anticipos R2'!F43</c>,
/// <c>='Reversion Pagos R4'!E73</c>, <c>='AJUSTES - SF-T'!E47</c>, …).
///
/// Regla del plan (§2.2, fila «D: solo trazable»): rellenar esas celdas desde el motor C#
/// duplicaría la cadena de fórmulas fuera del workbook (doble fuente de verdad) → PROHIBIDO.
/// Por eso aquí solo se declaran como TRAZABLES las columnas cuyo valor es una celda
/// workbook-interna de UNA sola arista (celda literal que la propia app ya materializa):
///
///   - DetRetri <c>J</c> = <c>BCE SC POR FACT.!F3</c> = <c>D3+E3</c>, con D3/E3 escritos por la
///     app (Unidad B) → J = ROUND(Subsidio + Contribución).
///   - DetRetri <c>L</c> = <c>INTERVENTORIA!F15</c> = costo de interventoría de la 2ª quincena
///     (tabla anual DECLARADA, celda LITERAL en la plantilla).
///   - DetValiRetri <c>I</c> = <c>Z104</c> = <c>L9*-1</c> = −interventoría.
///   - DetValiRetri <c>O</c> = bandera de validación (0, sin desglose).
///
/// El resto de columnas queda EXCLUIDO con motivo explícito (<see cref="ColumnasExcluidas"/>):
/// su origen es una fórmula de la cadena R1/R2/R4/AJUSTES que el plan prohíbe replicar. Destacar
/// el SALE de T0d: <c>DetValiRetri!J</c> (AJUSTE A LA DECENA) no tiene origen 10/10 (manual-externo).
///
/// Guard anti-fórmula: la escritura usa <see cref="OpenXmlPlantillaWriter"/> → si una celda-destino
/// fuera fórmula, lanza <c>ERR-PLANTILLA</c>; nunca sobrescritura.
/// </summary>
public static class WorkbookLeafCellMapDetRetri
{
    /// <summary>Hoja DetRetri del período (D-A, Plan 30): se compone del dominio.</summary>
    public static string HojaDetRetri(Periodo periodo) => WorkbookLeafCellMapQ2.HojaDetRetri(periodo);

    /// <summary>Hoja DetValiRetri del período (D-A, Plan 30): se compone del dominio.</summary>
    public static string HojaDetValiRetri(Periodo periodo) => WorkbookLeafCellMapQ2.HojaDetValiRetri(periodo);

    /// <summary>Primer ASE soportado.</summary>
    public const int PrimerAse = 1;

    /// <summary>Último ASE soportado.</summary>
    public const int UltimoAse = 5;

    /// <summary>
    /// Fila destino del ASE en <c>DetRetri*/DetValiRetri*</c> (T0-0.4): D9..D13 = ASE 1..5
    /// (<c>fila = 8 + Ase.Id</c>). La fila 14 (Total) está fuera de este rango.
    /// </summary>
    public static int ObtenerFila(int aseId) =>
        aseId is >= PrimerAse and <= UltimoAse
            ? 8 + aseId
            : throw new ArgumentOutOfRangeException(
                nameof(aseId),
                aseId,
                $"No hay fila DetRetri/DetValiRetri declarada para el ASE {aseId} (rango {PrimerAse}..{UltimoAse}).");

    /// <summary>Procedencia permitida de una columna trazable de <c>DetRetri2026072</c>.</summary>
    public enum FuenteDetRetri
    {
        /// <summary><c>BCE SC POR FACT.!F3</c> = Subsidio(D3) + Contribución(E3), escritos por la app.</summary>
        BalanceBceF3,

        /// <summary><c>INTERVENTORIA!F15</c> = costo de interventoría 2ª quincena (literal declarado).</summary>
        InterventoriaF15
    }

    /// <summary>Procedencia permitida de una columna trazable de <c>DetValiRetri2026072</c>.</summary>
    public enum FuenteDetValiRetri
    {
        /// <summary><c>Z104 = L9*-1</c> = −(costo de interventoría).</summary>
        InterventoriaF15Negada,

        /// <summary>Bandera de validación sin desglose (0).</summary>
        Cero
    }

    /// <summary>
    /// Columnas trazables de <c>DetRetri2026072</c> (columna destino → procedencia). Solo las
    /// columnas cuyo valor es una celda workbook-interna de una arista. NO incluye C (código de
    /// ASE, no es columna trazada por T0d) ni D (ya la escribe el path HU-12 V0.4).
    /// </summary>
    public static readonly IReadOnlyList<(string Columna, FuenteDetRetri Fuente)> ColumnasDetRetri =
    [
        ("J", FuenteDetRetri.BalanceBceF3),
        ("L", FuenteDetRetri.InterventoriaF15)
    ];

    /// <summary>
    /// Columnas trazables de <c>DetValiRetri2026072</c> (columna destino → procedencia). Solo las
    /// de una arista. NO incluye D (ya la escribe el path HU-12 V0.4).
    /// </summary>
    public static readonly IReadOnlyList<(string Columna, FuenteDetValiRetri Fuente)> ColumnasDetValiRetri =
    [
        ("I", FuenteDetValiRetri.InterventoriaF15Negada),
        ("O", FuenteDetValiRetri.Cero)
    ];

    /// <summary>Motivo común de las columnas cuya celda-origen es una fórmula de la cadena.</summary>
    public const string MotivoCadenaFormula =
        "origen = celda-fórmula workbook-interna (CONSOLIDADO_TOTAL RECAUDO fila 104 = Σ R1/R2/R4/AJUSTES); " +
        "reproducirla en C# duplicaría la cadena de fórmulas (plan 29 §2.2, alternativa 'motor C#' DESCARTADA)";

    /// <summary>Motivo del SALE explícito de T0d.</summary>
    public const string MotivoSaleAjusteDecena =
        "SALE (T0d): 'AJUSTE A LA DECENA' manual-externo, sin origen workbook-interno 10/10 (±0.5); no se inventa";

    /// <summary>
    /// Columnas NO escritas, cada una con su motivo (R-D-1: «columnas excluidas listadas con
    /// motivo»). El SALE de T0d va primero y con motivo propio. Las hojas se resuelven por período
    /// (D-A, Plan 30).
    /// </summary>
    public static IReadOnlyList<(string Hoja, string Columna, string Motivo)> ColumnasExcluidas(Periodo periodo)
    {
        var hojaDetRetri = HojaDetRetri(periodo);
        var hojaDetValiRetri = HojaDetValiRetri(periodo);
        return
        [
            (hojaDetValiRetri, "J", MotivoSaleAjusteDecena),
            (hojaDetRetri, "E", MotivoCadenaFormula),
            (hojaDetRetri, "F", MotivoCadenaFormula),
            (hojaDetRetri, "G", MotivoCadenaFormula),
            (hojaDetRetri, "H", MotivoCadenaFormula),
            (hojaDetRetri, "I", MotivoCadenaFormula),
            (hojaDetRetri, "K", MotivoCadenaFormula),
            (hojaDetRetri, "M", MotivoCadenaFormula),
            (hojaDetRetri, "N", MotivoCadenaFormula),
            (hojaDetRetri, "O", MotivoCadenaFormula),
            (hojaDetValiRetri, "D", MotivoCadenaFormula),
            (hojaDetValiRetri, "E", MotivoCadenaFormula),
            (hojaDetValiRetri, "F", MotivoCadenaFormula),
            (hojaDetValiRetri, "G", MotivoCadenaFormula),
            (hojaDetValiRetri, "H", MotivoCadenaFormula),
            (hojaDetValiRetri, "K", MotivoCadenaFormula),
            (hojaDetValiRetri, "L", MotivoCadenaFormula),
            (hojaDetValiRetri, "M", MotivoCadenaFormula),
            (hojaDetValiRetri, "N", MotivoCadenaFormula)
        ];
    }
}
