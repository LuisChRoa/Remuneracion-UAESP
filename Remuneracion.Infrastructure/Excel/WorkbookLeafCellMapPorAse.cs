namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Mapa de celdas editables leaf y fórmulas visibles protegidas POR BLOQUE ASE.
/// Congelado por T0 en <c>plans/07 - HU-07 T0 Evidencia.md</c> (§0.1/§0.3) con dump OpenXML
/// celda-por-celda del template golden <c>Docs/Insumos/Remuneracion 202607-1 Total.xlsx</c>.
///
/// R1 es HETEROGÉNEO por bloque (ASE1: 3 términos; ASE2..5: 5 términos; F178/F521 valor 0 sin
/// fórmula) — por eso el mapa es explícito por <see cref="Ase.Id"/> y está PROHIBIDA cualquier
/// aritmética de offsets (plan G2/D1).
///
/// <para>
/// <b>LEGACY (geometría v1) — Plan 54, H1.3/D3:</b> este mapa NO es el productivo desde el Plan 54.
/// La geometría vigente (plantilla v2) es <see cref="WorkbookLeafCellMapQ2.R1Q2EditablesPorAse"/>, que
/// rige AMBAS quincenas. Aquí quedan las direcciones congeladas de la v1 (F25/F41/L25/F30/F10/L10 para
/// R1) para el API histórico de leaf y los fixtures de caracterización. Prohibido re-apuntar producción
/// a este mapa o añadirle lectores nuevos.
/// </para>
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
                (HojaR2, "E21", "R2.E21"), (HojaR2, "E32", "R2.E32"), (HojaR2, "K21", "R2.K21"),
                (HojaR4, "D15", "R4.D15"), (HojaR4, "P15", "R4.P15")
            ],
            [2] =
            [
                (HojaR1, "F113", "R1.F113"), (HojaR1, "F130", "R1.F130"), (HojaR1, "L113", "R1.L113"),
                (HojaR1, "F90", "R1.F90"), (HojaR1, "L90", "R1.L90"),
                (HojaR2, "E101", "R2.E101"), (HojaR2, "E123", "R2.E123"), (HojaR2, "K101", "R2.K101"),
                (HojaR4, "D111", "R4.D111"), (HojaR4, "P111", "R4.P111")
            ],
            [3] =
            [
                (HojaR1, "F238", "R1.F238"), (HojaR1, "F254", "R1.F254"), (HojaR1, "F217", "R1.F217"),
                (HojaR1, "L217", "R1.L217"), (HojaR1, "L238", "R1.L238"),
                (HojaR1, "F243", "R1.F243"), (HojaR1, "F223", "R1.F223"), (HojaR1, "L223", "R1.L223"),
                (HojaR2, "E193", "R2.E193"), (HojaR2, "E204", "R2.E204"), (HojaR2, "K193", "R2.K193"),
                (HojaR4, "D206", "R4.D206"), (HojaR4, "P206", "R4.P206")
            ],
            [4] =
            [
                (HojaR1, "F391", "R1.F391"), (HojaR1, "F417", "R1.F417"), (HojaR1, "F357", "R1.F357"),
                (HojaR1, "L357", "R1.L357"), (HojaR1, "L391", "R1.L391"),
                (HojaR1, "F401", "R1.F401"), (HojaR1, "F369", "R1.F369"), (HojaR1, "L369", "R1.L369"),
                (HojaR2, "E327", "R2.E327"), (HojaR2, "E349", "R2.E349"), (HojaR2, "K327", "R2.K327"),
                (HojaR4, "D249", "R4.D249"), (HojaR4, "P249", "R4.P249")
            ],
            [5] =
            [
                (HojaR1, "F513", "R1.F513"), (HojaR1, "F498", "R1.F498"), (HojaR1, "F478", "R1.F478"),
                (HojaR1, "L478", "R1.L478"), (HojaR1, "L498", "R1.L498"),
                (HojaR2, "E426", "R2.E426"), (HojaR2, "E440", "R2.E440"), (HojaR2, "K426", "R2.K426"),
                (HojaR4, "D358", "R4.D358"), (HojaR4, "P358", "R4.P358")
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
            ],
            [2] =
            [
                (HojaR1, "F254", ["F171", "F208", "L171", "F120", "L120"]),
                (HojaR2, "E155", ["E101", "E123", "K101"]),
                (HojaR4, "D174", ["D111", "P111"]),
                (HojaConsolidado, "D10", ["F254", "Reporte Componentes R1"]),
                (HojaConsolidado, "D29", ["E155", "Rem. Anticipos R2"]),
                (HojaConsolidado, "D48", ["F256", "Reporte Componentes R1"]),
                (HojaConsolidado, "D67", ["D174", "Reversion Pagos R4"]),
                (HojaConsolidado, "D105", ["D10", "D29", "D48", "D67", "D86"]),
                (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
            ],
            [3] =
            [
                (HojaR1, "F415", ["F330", "F353", "F299", "L299", "L330"]),
                (HojaR1, "F417", ["F342", "F315", "L315"]),
                (HojaR2, "E273", ["E193", "E204", "K193"]),
                (HojaR4, "D211", ["D206", "P206"]),
                (HojaConsolidado, "D11", ["F415", "Reporte Componentes R1"]),
                (HojaConsolidado, "D30", ["E273", "Rem. Anticipos R2"]),
                (HojaConsolidado, "D49", ["F417", "Reporte Componentes R1"]),
                (HojaConsolidado, "D68", ["D211", "Reversion Pagos R4"]),
                (HojaConsolidado, "D106", []), // shared follower del maestro D104 (texto vacío; solo presencia de fórmula)
                (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
            ],
            [4] =
            [
                (HojaR1, "F568", ["F511", "F548", "F460", "L460", "L511"]),
                (HojaR1, "F570", ["F530", "F486", "L486"]),
                (HojaR2, "E384", ["E327", "E349", "K327"]),
                (HojaR4, "D326", ["D249", "P249"]),
                (HojaConsolidado, "D12", ["F568", "Reporte Componentes R1"]),
                (HojaConsolidado, "D31", ["E384", "Rem. Anticipos R2"]),
                (HojaConsolidado, "D50", ["F570", "Reporte Componentes R1"]),
                (HojaConsolidado, "D69", ["D326", "Reversion Pagos R4"]),
                (HojaConsolidado, "D107", ["D12", "D31", "D50", "D69", "D88"]),
                (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
            ],
            [5] =
            [
                (HojaR1, "F680", ["F674", "F648", "F613", "L613", "L648"]),
                (HojaR2, "E468", ["E440", "E426", "K426"]),
                (HojaR4, "D361", ["D358", "P358"]),
                (HojaConsolidado, "D13", ["F680", "Reporte Componentes R1"]),
                (HojaConsolidado, "D32", ["E468", "Rem. Anticipos R2"]),
                (HojaConsolidado, "D51", ["F682", "Reporte Componentes R1"]),
                (HojaConsolidado, "D70", ["D361", "Reversion Pagos R4"]),
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
