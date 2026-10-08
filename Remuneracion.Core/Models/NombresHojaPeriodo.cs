namespace Remuneracion.Core.Models;

/// <summary>
/// Naming de las hojas cuyo nombre depende del período (D-A, Plan 30). Fuente ÚNICA de la regla:
/// el nombre se COMPONE del dominio (<see cref="Periodo.CodigoCompleto"/>), nunca se detecta
/// enumerando las hojas del workbook (detección por contenido frágil ante hojas heredadas de otro
/// período). Puro: sin I/O ni dependencias.
/// </summary>
public static class NombresHojaPeriodo
{
    /// <summary>Hoja DetRetri del período: <c>DetRetri{AAAAMMQ}</c> (p. ej. 2026071 → DetRetri2026071).</summary>
    public static string DetRetri(string codigoCompleto) => $"DetRetri{codigoCompleto}";

    /// <summary>Hoja DetValiRetri del período: <c>DetValiRetri{AAAAMMQ}</c>.</summary>
    public static string DetValiRetri(string codigoCompleto) => $"DetValiRetri{codigoCompleto}";

    /// <summary>Hoja Informe AFaseo del período: <c>Informe AFaseo Recaudo {AAAAMM}-{Q}</c>.</summary>
    public static string InformeAFaseo(string codigoAaaamm, int quincena) =>
        $"Informe AFaseo Recaudo {codigoAaaamm}-{quincena}";

    /// <summary>Hoja DetRetri del período a partir del dominio.</summary>
    public static string DetRetri(Periodo periodo) => DetRetri(periodo.CodigoCompleto);

    /// <summary>Hoja DetValiRetri del período a partir del dominio.</summary>
    public static string DetValiRetri(Periodo periodo) => DetValiRetri(periodo.CodigoCompleto);

    /// <summary>Hoja Informe AFaseo del período a partir del dominio.</summary>
    public static string InformeAFaseo(Periodo periodo) =>
        InformeAFaseo(periodo.CodigoAAAAMM, periodo.NumeroQuincena);
}
