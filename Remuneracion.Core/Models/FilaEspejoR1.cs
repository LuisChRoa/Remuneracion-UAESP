namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 21 (T1, R-E-1): fila tipada del espejo estructural de la hoja <c>Reporte Componentes R1</c>.
///
/// Modelo PURO (sin I/O) que representa una fila de la fuente <c>Recaudoporcomponente</c>
/// (hoja <c>Sheet1</c>) como la tupla de etiquetas de firma (columnas A–E, en su orden físico)
/// más los valores numéricos por ENCABEZADO de columna (F..T), nunca por índice fijo.
///
/// La firma y la secuencia las fija la evidencia congelada T0e
/// (<c>plans/21 - T0 Evidencia.md</c> §6): la secuencia observada de firmas A–E ES la
/// especificación del período (no se exigen cardinalidades ni ocurrencias — decisión D-B/R-E-1).
/// <c>Mes</c> y <c>AFaseo</c> NO son invariantes; <see cref="EsComponenteTotal"/>,
/// <see cref="EsSubsContTotal"/> y <see cref="EsTotalFinal"/> SÍ lo son (corte de zona y cierre
/// de bloque).
/// </summary>
public sealed class FilaEspejoR1
{
    /// <summary>Etiqueta de la columna A (firma).</summary>
    public string A { get; init; } = string.Empty;

    /// <summary>Etiqueta de la columna B (firma).</summary>
    public string B { get; init; } = string.Empty;

    /// <summary>Etiqueta de la columna C (firma).</summary>
    public string C { get; init; } = string.Empty;

    /// <summary>Etiqueta de la columna D (firma).</summary>
    public string D { get; init; } = string.Empty;

    /// <summary>Etiqueta de la columna E (firma).</summary>
    public string E { get; init; } = string.Empty;

    /// <summary>
    /// Valores de la fila por ENCABEZADO de columna (p. ej. <c>"Total"</c>,
    /// <c>"SERVICIO ESPECIALES"</c>). <c>null</c> = celda vacía en la fuente (distinguir
    /// "ausente" de "cero"; R-E-5).
    /// </summary>
    public IReadOnlyDictionary<string, decimal?> ValoresPorColumna { get; init; } =
        new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Firma de la fila como tupla <c>A|B|C|D|E</c> (mismo criterio de la evidencia T0a/T0e).
    /// </summary>
    public string Firma => $"{A}|{B}|{C}|{D}|{E}";

    /// <summary>
    /// Invariante dura T0e: <c>A='Componente' B='Total'</c> (corte de la zona Componente).
    /// </summary>
    public bool EsComponenteTotal =>
        string.Equals(A, "Componente", StringComparison.OrdinalIgnoreCase)
        && string.Equals(B, "Total", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Invariante dura T0e: <c>A='Subs/Cont' B='Total'</c> (cierre Subs/Cont).
    /// </summary>
    public bool EsSubsContTotal =>
        string.Equals(A, "Subs/Cont", StringComparison.OrdinalIgnoreCase)
        && string.Equals(B, "Total", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Invariante dura T0e: <c>A='Total' B</c> vacío (fila Total final, cierre de bloque).
    /// </summary>
    public bool EsTotalFinal =>
        string.Equals(A, "Total", StringComparison.OrdinalIgnoreCase)
        && string.IsNullOrWhiteSpace(B);

    // ── Plan 25 (T1, R-F-1/2/3): predicados de firma de los roles R1-Q2 ─────────────────────
    // Fuente ÚNICA de cada firma (R-DOBLE-FUENTE): el lector filtra la secuencia por estos
    // predicados (misma normalización de etiquetas A–E que LeerEspejoR1), nunca por índices.

    /// <summary>
    /// Plan 25 (T1, R-F-1): firma del rol <c>Mes/Total</c> (<c>B='Mes' ∧ C='Total'</c>),
    /// por orden de aparición (Mes0/Mes1/Mes2). Rol OBLIGATORIO (T0e §6.2): su ausencia es
    /// fail-fast, nunca 0 silencioso.
    /// </summary>
    public bool EsMesTotal =>
        string.Equals(B, "Mes", StringComparison.OrdinalIgnoreCase)
        && string.Equals(C, "Total", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Plan 25 (T1, R-F-2): firma del rol <c>Aplicacion nuevos x reversion/Total</c>
    /// (<c>B</c> contiene <c>Aplicacion nuevos x reversion</c> ∧ <c>C='Total'</c>), por orden de
    /// aparición (Aplic0/Aplic1/…). Rol OPCIONAL (T0e §6.1): ausente = 0 explícito.
    /// </summary>
    public bool EsAplicacionTotal =>
        B.Contains("Aplicacion nuevos x reversion", StringComparison.OrdinalIgnoreCase)
        && string.Equals(C, "Total", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Plan 25 (T1, T0e §6.1): firma de la fila <c>Subsidio(-)/Contribucion(+)</c>
    /// (<c>E</c> contiene el label). Predicado de cierre/documentación: tras D-A
    /// (<c>F37/F270 → Aplic1</c>) no tiene consumidor en Q2.
    /// </summary>
    public bool EsSubsidio =>
        E.Contains("Subsidio(-)/Contribucion(+)", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Plan 25 (T1): firma de la fila <c>Vlr Servicio</c> (<c>E='Vlr Servicio'</c>).
    /// </summary>
    public bool EsVlrServicio =>
        string.Equals(E, "Vlr Servicio", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Plan 25 (T1): firma de la fila <c>Vlr Intereses</c> (<c>E='Vlr Intereses'</c>).
    /// </summary>
    public bool EsVlrIntereses =>
        string.Equals(E, "Vlr Intereses", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Plan 25 (T1): firma del código de empresa de la columna D (uno de E/H/O/T).
    /// </summary>
    public bool EsCodigoD =>
        string.Equals(D, "E", StringComparison.OrdinalIgnoreCase)
        || string.Equals(D, "H", StringComparison.OrdinalIgnoreCase)
        || string.Equals(D, "O", StringComparison.OrdinalIgnoreCase)
        || string.Equals(D, "T", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Valor numérico de la fila para el encabezado indicado; <c>null</c> si la columna no
    /// existe o la celda está vacía.
    /// </summary>
    public decimal? Valor(string encabezado) =>
        ValoresPorColumna.TryGetValue(encabezado, out var valor) ? valor : null;
}
