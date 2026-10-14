namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Mapa de celdas editables leaf y fórmulas protegidas del workbook real.
/// El writer solo escribe las celdas editables; las fórmulas de visibles y consolidado se preservan.
/// </summary>
public static class WorkbookLeafCellMap
{
    public const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";
    public const string HojaR1 = "Reporte Componentes R1";
    public const string HojaR2 = "Rem. Anticipos R2";
    public const string HojaR4 = "Reversion Pagos R4";

    public static readonly (string Hoja, string Celda, string Nombre)[] EditableLeafCells =
    [
        (HojaR1, "F25", "R1.F25"),
        (HojaR1, "F41", "R1.F41"),
        (HojaR1, "L25", "R1.L25"),
        (HojaR1, "F30", "R1.F30"),
        (HojaR1, "F10", "R1.F10"),
        (HojaR1, "L10", "R1.L10"),
        (HojaR2, "E15", "R2.E15"),
        (HojaR2, "E26", "R2.E26"),
        (HojaR2, "K15", "R2.K15"),
        (HojaR4, "D9", "R4.D9"),
        (HojaR4, "P9", "R4.P9")
    ];

    public static readonly (string Hoja, string Celda, string[] Fragmentos)[] ProtectedFormulas =
    [
        (HojaR1, "F75", ["F47", "F70", "F16", "L16", "L47"]),
        (HojaR1, "F77", ["F59", "F32", "L32"]),
        (HojaR2, "E47", ["E21", "E32", "K21"]),
        (HojaR4, "D73", ["D15", "P15"]),
        (HojaConsolidado, "D9", ["F75", "Reporte Componentes R1"]),
        (HojaConsolidado, "D28", ["E47", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D47", ["F77", "Reporte Componentes R1"]),
        (HojaConsolidado, "D66", ["D73", "Reversion Pagos R4"]),
        (HojaConsolidado, "D104", ["D9", "D28", "D47", "D66", "D85"]),
        (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
    ];
}
