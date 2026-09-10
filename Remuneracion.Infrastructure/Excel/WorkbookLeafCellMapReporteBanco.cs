namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// HU-09 (2.3): mapa de celdas de la hoja <c>REPORTE RECAUDO x BANCO</c> congelado por T0
/// (evidencia en <c>plans/09 - HU-09 Reporte recaudo por banco Fase 2.md</c> §4 Fase 0 y en el
/// dump OpenXML del golden <c>Docs/Insumos/Remuneracion 202607-1 Total.xlsx</c>).
///
/// Veredictos T0-0.1/0.6 (D2): las filas 1–7 (consolidado), los Total de bloque (fila 7 del
/// bloque), las columnas I/J de verificación, las filas 62–79 y el TOTAL (fila 81) son
/// FÓRMULAS → van al mapa <see cref="Protegidas"/> y jamás se escriben. C59 es VALOR → se
/// escribe con <see cref="Core.Models.ReporteBancoInputs.Quincena"/>. Solo las celdas de
/// concepto de los bloques (filas 3–6 de cada bloque) son valores escribibles.
///
/// El mapa de LECTURA de la fuente (columna del Resumen por ASE) es heterogéneo:
/// ASE1 ENEL=1/OCC=2; ASE2 ENEL=1/NUEVO ESQUEMA=2/OCC=3; ASE3 ENEL=1/OCC=2;
/// ASE4 ENEL=1/NUEVO ESQUEMA=2/OCC=3; ASE5 ENEL=1/ENERBIT=2/OCC=3 (V10 + T0-0.3).
/// Prohibidos offsets y filas fijas para el Resumen (D1: el conteo de filas diarias varía).
/// </summary>
public static class WorkbookLeafCellMapReporteBanco
{
    public const string HojaBanco = "REPORTE RECAUDO x BANCO";

    /// <summary>
    /// Prefijo normalizado de la etiqueta "Resumen Recaudo Aplicado Por Servicio" (búsqueda
    /// desde el final). Match por prefijo, nunca string exacto ni fila fija (V4/V6).
    /// </summary>
    public const string PrefijoEtiquetaResumen = "resumenrecaudoaplicadoporservicio";

    /// <summary>
    /// Prefijos normalizados de los conceptos (T0-0.4): "3-APLICADOS A FINANCIACIONES NUEVAS"
    /// matchea por prefijo sin depender de ±" NUEVAS". Los prefijos ya vienen normalizados
    /// (minúsculas, sin espacios, sin guiones entre palabras) como la salida de
    /// <c>NormalizarEtiqueta</c> del reader.
    /// </summary>
    public static readonly (int Indice, string Prefijo)[] Conceptos =
    [
        (0, "1-aplicadosafacturacion"),
        (1, "2-saldosafavor"),
        (2, "3-aplicadosafinanciaciones"),
        (3, "7-aplicados")
    ];

    /// <summary>
    /// Índice de columna de cada empresa en la fila de headers del Resumen de la fuente
    /// (1-based sobre las columnas de datos; la columna 0 es la etiqueta de concepto).
    /// Congelado por T0-0.3.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, IReadOnlyDictionary<string, int>> ColumnasFuentePorAse =
        new Dictionary<int, IReadOnlyDictionary<string, int>>
        {
            [1] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { ["ENEL"] = 1, ["OCCIDENTE"] = 2 },
            [2] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { ["ENEL"] = 1, ["NUEVO ESQUEMA"] = 2, ["OCCIDENTE"] = 3 },
            [3] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { ["ENEL"] = 1, ["OCCIDENTE"] = 2 },
            [4] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { ["ENEL"] = 1, ["NUEVO ESQUEMA"] = 2, ["OCCIDENTE"] = 3 },
            [5] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { ["ENEL"] = 1, ["ENERBIT"] = 2, ["OCCIDENTE"] = 3 }
        };

    /// <summary>
    /// Celdas de concepto editables en el template por ASE × empresa (4 celdas por empresa:
    /// filas de concepto 1/2/3/7 del bloque). El índice del arreglo coincide con
    /// <see cref="Conceptos"/> (0..3). Congelado por T0-0.7.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, IReadOnlyDictionary<string, string[]>> EditablesPorAse =
        new Dictionary<int, IReadOnlyDictionary<string, string[]>>
        {
            [1] = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                ["ENEL"] = ["C13", "C14", "C15", "C16"],
                ["OCCIDENTE"] = ["H13", "H14", "H15", "H16"]
            },
            [2] = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                ["ENEL"] = ["C23", "C24", "C25", "C26"],
                ["NUEVO ESQUEMA"] = ["F23", "F24", "F25", "F26"],
                ["OCCIDENTE"] = ["H23", "H24", "H25", "H26"]
            },
            [3] = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                ["ENEL"] = ["C32", "C33", "C34", "C35"],
                ["OCCIDENTE"] = ["H32", "H33", "H34", "H35"]
            },
            [4] = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                ["ENEL"] = ["C41", "C42", "C43", "C44"],
                ["NUEVO ESQUEMA"] = ["F41", "F42", "F43", "F44"],
                ["OCCIDENTE"] = ["H41", "H42", "H43", "H44"]
            },
            [5] = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
            {
                ["ENEL"] = ["C51", "C52", "C53", "C54"],
                ["ENERBIT"] = ["G51", "G52", "G53", "G54"],
                ["OCCIDENTE"] = ["H51", "H52", "H53", "H54"]
            }
        };

    /// <summary>
    /// Fórmulas protegidas 2.3 (jamás se escriben; D6). Incluye el consolidado filas 1–7
    /// (D2(b): T0 determinó que son fórmulas), los Total de bloque, la verificación I/J del
    /// bloque, las filas 62–79 (validación vs <c>Recaudo *</c> — lógica 2.7, solo se protege)
    /// y el TOTAL RECAUDO (fila 81). C59 queda FUERA (es valor).
    /// </summary>
    public static readonly (string Hoja, string Celda, string[] Fragmentos)[] Protegidas =
    [
        // Consolidado filas 1–7 (maestros + verificaciones).
        (HojaBanco, "C2", ["C13", "C23", "C32", "C41", "C51"]),
        (HojaBanco, "I2", ["SUM"]),
        (HojaBanco, "C3", ["C14", "C24", "C33", "C42", "C52"]),
        (HojaBanco, "I3", ["SUM"]),
        (HojaBanco, "C4", ["C15", "C25", "C34", "C43", "C53"]),
        (HojaBanco, "I4", ["SUM"]),
        (HojaBanco, "C5", ["C16", "C26", "C35", "C44", "C54"]),
        (HojaBanco, "I5", ["SUM"]),
        (HojaBanco, "C6", ["SUM"]),
        (HojaBanco, "I6", ["SUM"]),
        (HojaBanco, "C7", ["C6"]),
        (HojaBanco, "I7", ["I6"]),
        // Total de bloque (fila 7 de cada bloque) + verificación I del bloque.
        (HojaBanco, "C17", ["SUM", "C13"]),
        (HojaBanco, "H17", ["SUM", "H13"]),
        (HojaBanco, "I17", ["SUM"]),
        (HojaBanco, "J17", ["SUM", "I17"]),
        (HojaBanco, "C27", ["SUM", "C23"]),
        (HojaBanco, "F27", ["SUM", "F23"]),
        (HojaBanco, "H27", ["SUM", "H23"]),
        (HojaBanco, "I27", ["SUM"]),
        (HojaBanco, "J27", ["SUM", "I27"]),
        (HojaBanco, "C36", ["SUM", "C32"]),
        (HojaBanco, "H36", ["SUM", "H32"]),
        (HojaBanco, "I36", ["SUM"]),
        (HojaBanco, "J36", ["SUM", "I36"]),
        (HojaBanco, "C45", ["SUM", "C41"]),
        (HojaBanco, "F45", ["SUM", "F41"]),
        (HojaBanco, "H45", ["SUM", "H41"]),
        (HojaBanco, "I45", ["SUM"]),
        (HojaBanco, "J45", ["SUM", "I45"]),
        (HojaBanco, "C55", ["SUM", "C51"]),
        (HojaBanco, "G55", ["SUM", "G51"]),
        (HojaBanco, "H55", ["SUM", "H51"]),
        (HojaBanco, "I55", ["SUM"]),
        (HojaBanco, "J55", ["SUM", "I55"]),
        // Validación 59–80 (2.7: IF($C$59=1,...) contra hojas Recaudo *; solo se protege).
        // El fragmento "$C$59" (con $) es el literal exacto de la fórmula; "C59" no matchea "C$59".
        (HojaBanco, "C62", ["$C$59", "Recaudo ENEL"]),
        (HojaBanco, "D62", ["C17", "C62"]),
        (HojaBanco, "E62", ["C17", "C62"]),
        (HojaBanco, "C75", ["$C$59", "Recaudo Directa Occidente"]),
        (HojaBanco, "E75", ["H17", "C75"]),
        // TOTAL RECAUDO (fila 81).
        (HojaBanco, "C81", ["SUM", "C62"]),
        (HojaBanco, "D81", ["C81", "I6"])
    ];

    /// <summary>
    /// Obtiene las celdas de concepto editables de un ASE; lanza si el ASE no está soportado.
    /// </summary>
    public static IReadOnlyDictionary<string, string[]> ObtenerEditables(int aseId) =>
        EditablesPorAse.TryGetValue(aseId, out var mapa)
            ? mapa
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay cell-map de reporte banco para el ASE {aseId}.");
}
