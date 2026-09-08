namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Mapa de celdas editables leaf y fórmulas visibles protegidas POR BLOQUE ASE.
/// Congelado por T0 en <c>plans/07 - HU-07 T0 Evidencia.md</c> (§0.1/§0.3) con dump OpenXML
/// celda-por-celda del template golden <c>Docs/Insumos/Remuneracion 202607-1 Total.xlsx</c>.
///
/// R1 es HETEROGÉNEO por bloque (ASE1: 3 términos; ASE2..5: 5 términos; F178/F521 valor 0 sin
/// fórmula) — por eso el mapa es explícito por <see cref="Ase.Id"/> y está PROHIBIDA cualquier
/// aritmética de offsets (plan G2/D1).
/// </summary>
public static class WorkbookLeafCellMapPorAse
{
    public const string HojaConsolidado = WorkbookLeafCellMap.HojaConsolidado;
    public const string HojaR1 = WorkbookLeafCellMap.HojaR1;
    public const string HojaR2 = WorkbookLeafCellMap.HojaR2;
    public const string HojaR4 = WorkbookLeafCellMap.HojaR4;

    /// <summary>
    /// Celdas leaf editables por ASE (visibles por bloque), congeladas por T0.
    /// El writer escribe SOLO estas celdas para cada bloque.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Hoja, string Celda, string Nombre)[]> EditableLeafCellsPorAse =
        new Dictionary<int, (string Hoja, string Celda, string Nombre)[]>
        {
            [1] =
            [
                (HojaR1, "F25", "R1.F25"), (HojaR1, "F41", "R1.F41"), (HojaR1, "L25", "R1.L25"),
                (HojaR1, "F30", "R1.F30"), (HojaR1, "F10", "R1.F10"), (HojaR1, "L10", "R1.L10"),
                (HojaR2, "E15", "R2.E15"), (HojaR2, "E26", "R2.E26"), (HojaR2, "K15", "R2.K15"),
                (HojaR4, "D9", "R4.D9"), (HojaR4, "P9", "R4.P9")
            ],
            [2] =
            [
                (HojaR1, "F113", "R1.F113"), (HojaR1, "F130", "R1.F130"), (HojaR1, "L113", "R1.L113"),
                (HojaR1, "F90", "R1.F90"), (HojaR1, "L90", "R1.L90"),
                (HojaR2, "E85", "R2.E85"), (HojaR2, "E103", "R2.E103"), (HojaR2, "K85", "R2.K85"),
                (HojaR4, "D98", "R4.D98"), (HojaR4, "P98", "R4.P98")
            ],
            [3] =
            [
                (HojaR1, "F238", "R1.F238"), (HojaR1, "F254", "R1.F254"), (HojaR1, "F217", "R1.F217"),
                (HojaR1, "L217", "R1.L217"), (HojaR1, "L238", "R1.L238"),
                (HojaR1, "F243", "R1.F243"), (HojaR1, "F223", "R1.F223"), (HojaR1, "L223", "R1.L223"),
                (HojaR2, "E167", "R2.E167"), (HojaR2, "E178", "R2.E178"), (HojaR2, "K167", "R2.K167"),
                (HojaR4, "D193", "R4.D193"), (HojaR4, "P193", "R4.P193")
            ],
            [4] =
            [
                (HojaR1, "F391", "R1.F391"), (HojaR1, "F417", "R1.F417"), (HojaR1, "F357", "R1.F357"),
                (HojaR1, "L357", "R1.L357"), (HojaR1, "L391", "R1.L391"),
                (HojaR1, "F401", "R1.F401"), (HojaR1, "F369", "R1.F369"), (HojaR1, "L369", "R1.L369"),
                (HojaR2, "E290", "R2.E290"), (HojaR2, "E308", "R2.E308"), (HojaR2, "K290", "R2.K290"),
                (HojaR4, "D236", "R4.D236"), (HojaR4, "P236", "R4.P236")
            ],
            [5] =
            [
                (HojaR1, "F513", "R1.F513"), (HojaR1, "F498", "R1.F498"), (HojaR1, "F478", "R1.F478"),
                (HojaR1, "L478", "R1.L478"), (HojaR1, "L498", "R1.L498"),
                (HojaR2, "E385", "R2.E385"), (HojaR2, "E374", "R2.E374"), (HojaR2, "K374", "R2.K374"),
                (HojaR4, "D344", "R4.D344"), (HojaR4, "P344", "R4.P344")
            ]
        };

    /// <summary>
    /// Fórmulas visibles protegidas por bloque (R1/R2/R4) + filas CONSOLIDADO por ASE.
    /// ASE2/ASE5 NO incluyen F178/F521: son VALOR 0 estático sin fórmula (T0-0.2) y no se escriben.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Hoja, string Celda, string[] Fragmentos)[]> ProtectedFormulasPorAse =
        new Dictionary<int, (string Hoja, string Celda, string[] Fragmentos)[]>
        {
            [1] =
            [
                (HojaR1, "F46", ["F25", "F41", "L25"]),
                (HojaR1, "F48", ["F30", "F10", "L10"]),
                (HojaR2, "E41", ["E15", "E26", "K15"]),
                (HojaR4, "D67", ["D9", "P9"]),
                (HojaConsolidado, "D9", ["F46", "Reporte Componentes R1"]),
                (HojaConsolidado, "D28", ["E41", "Rem. Anticipos R2"]),
                (HojaConsolidado, "D47", ["F48", "Reporte Componentes R1"]),
                (HojaConsolidado, "D66", ["D67", "Reversion Pagos R4"]),
                (HojaConsolidado, "D104", ["D9", "D28", "D47", "D66", "D85"]),
                (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
            ],
            [2] =
            [
                (HojaR1, "F176", ["F113", "F130", "L113", "F90", "L90"]),
                (HojaR2, "E135", ["E85", "E103", "K85"]),
                (HojaR4, "D161", ["D98", "P98"]),
                (HojaConsolidado, "D10", ["F176", "Reporte Componentes R1"]),
                (HojaConsolidado, "D29", ["E135", "Rem. Anticipos R2"]),
                (HojaConsolidado, "D48", ["F178", "Reporte Componentes R1"]),
                (HojaConsolidado, "D67", ["D161", "Reversion Pagos R4"]),
                (HojaConsolidado, "D105", ["D10", "D29", "D48", "D67", "D86"]),
                (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
            ],
            [3] =
            [
                (HojaR1, "F316", ["F238", "F254", "F217", "L217", "L238"]),
                (HojaR1, "F318", ["F243", "F223", "L223"]),
                (HojaR2, "E247", ["E167", "E178", "K167"]),
                (HojaR4, "D198", ["D193", "P193"]),
                (HojaConsolidado, "D11", ["F316", "Reporte Componentes R1"]),
                (HojaConsolidado, "D30", ["E247", "Rem. Anticipos R2"]),
                (HojaConsolidado, "D49", ["F318", "Reporte Componentes R1"]),
                (HojaConsolidado, "D68", ["D198", "Reversion Pagos R4"]),
                (HojaConsolidado, "D106", []), // shared follower del maestro D104 (texto vacío; solo presencia de fórmula)
                (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
            ],
            [4] =
            [
                (HojaR1, "F437", ["F391", "F417", "F357", "L357", "L391"]),
                (HojaR1, "F439", ["F401", "F369", "L369"]),
                (HojaR2, "E343", ["E290", "E308", "K290"]),
                (HojaR4, "D312", ["D236", "P236"]),
                (HojaConsolidado, "D12", ["F437", "Reporte Componentes R1"]),
                (HojaConsolidado, "D31", ["E343", "Rem. Anticipos R2"]),
                (HojaConsolidado, "D50", ["F439", "Reporte Componentes R1"]),
                (HojaConsolidado, "D69", ["D312", "Reversion Pagos R4"]),
                (HojaConsolidado, "D107", ["D12", "D31", "D50", "D69", "D88"]),
                (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
            ],
            [5] =
            [
                (HojaR1, "F519", ["F513", "F498", "F478", "L478", "L498"]),
                (HojaR2, "E413", ["E385", "E374", "K374"]),
                (HojaR4, "D347", ["D344", "P344"]),
                (HojaConsolidado, "D13", ["F519", "Reporte Componentes R1"]),
                (HojaConsolidado, "D32", ["E413", "Rem. Anticipos R2"]),
                (HojaConsolidado, "D51", ["F521", "Reporte Componentes R1"]),
                (HojaConsolidado, "D70", ["D347", "Reversion Pagos R4"]),
                (HojaConsolidado, "D108", ["D13", "D32", "D51", "D70", "D89"]),
                (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
            ]
        };

    /// <summary>
    /// Obtiene las celdas leaf editables para un ASE; lanza si el ASE no está soportado.
    /// </summary>
    public static (string Hoja, string Celda, string Nombre)[] ObtenerEditables(int aseId) =>
        EditableLeafCellsPorAse.TryGetValue(aseId, out var celdas)
            ? celdas
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay cell-map para el ASE {aseId}.");

    /// <summary>
    /// Obtiene las fórmulas protegidas de un bloque ASE (incluye sus filas CONSOLIDADO).
    /// </summary>
    public static (string Hoja, string Celda, string[] Fragmentos)[] ObtenerProtegidos(int aseId) =>
        ProtectedFormulasPorAse.TryGetValue(aseId, out var protegidos)
            ? protegidos
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay fórmulas protegidas para el ASE {aseId}.");
}