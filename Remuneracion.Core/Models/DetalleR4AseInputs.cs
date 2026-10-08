namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 29 (T2, Unidad R — SOLO LECTURA): matriz de detalle de la hoja <c>Reversion Pagos R4</c>
/// extraída de la fuente <c>ReversiónPorComponente_*</c>, resuelta por LABEL de fila (A–C) y por
/// ENCABEZADO de componente (col D..), NUNCA por corrimiento fijo.
///
/// Motivo (veredicto T0b, <c>plans/29-T0-Evidencia-B.md</c> §2.3): la fuente trae el mismo
/// esqueleto de labels que el template (r5..r17 → destino r3..r15 en julio), pero la malla de
/// filas deriva en agosto (PROMO 13→6, CIUDAD LIMPIA 13→7, BOGOTA LIMPIA 19→18) y las columnas
/// destino P..R (<c>SERVICIO ESPECIALES</c>, <c>Ajuste Recargo Aseo</c>, <c>Otros</c>) no tienen
/// contraparte en la fuente → 0. La resolución por label/encabezado es la única robusta.
///
/// Modelo PURO de dominio (sin I/O ni celdas de plantilla): expone la observación fiel de la
/// fuente. El mapeo a celdas-destino es responsabilidad de la ESCRITURA (T3).
///
/// Convive con <see cref="WorkbookLeafInputsR4"/> (agregado D9/P9): no lo reemplaza.
/// </summary>
public sealed class DetalleR4AseInputs
{
    /// <summary>ASE asociado a la fuente leída.</summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Encabezados de componente detectados en la fila de headers de la fuente, en orden físico
    /// (p. ej. <c>"Total"</c>, <c>"Componente TDF"</c>, …, <c>"Ajuste Decena"</c>). Es la clave de
    /// <see cref="DetalleR4Fila.ValoresPorComponente"/>.
    /// </summary>
    public IReadOnlyList<string> Componentes { get; set; } = [];

    /// <summary>
    /// Filas de detalle en orden físico, desde la primera <c>Vlr Servicio</c> hasta la fila de
    /// cierre <c>Total</c> (A='Total') inclusive (la secuencia observada ES la especificación del
    /// período; no se normaliza ni se reordena).
    /// </summary>
    public IReadOnlyList<DetalleR4Fila> Filas { get; set; } = [];
}

/// <summary>
/// Fila tipada del detalle R4: firma de labels (A–C) + valores por ENCABEZADO de componente.
/// </summary>
public sealed class DetalleR4Fila
{
    /// <summary>Etiqueta de la columna A (empresa del bloque <c>ENEL</c>/<c>OCCIDENTE</c>/<c>NUEVO ESQUEMA</c>, o cierre <c>Total</c>).</summary>
    public string A { get; init; } = string.Empty;

    /// <summary>Etiqueta de la columna B (p. ej. <c>Componente</c>/<c>Subs/Cont</c>/<c>Total</c>).</summary>
    public string B { get; init; } = string.Empty;

    /// <summary>Etiqueta de la columna C (concepto: <c>Vlr Servicio</c>, <c>Vlr Intereses</c>, <c>Total</c>, <c>Subsidio(-)/Contribucion(+)</c>).</summary>
    public string C { get; init; } = string.Empty;

    /// <summary>Firma de la fila como tupla <c>A|B|C</c> (mismo criterio de la evidencia T0b).</summary>
    public string Firma => $"{A}|{B}|{C}";

    /// <summary>
    /// Valores de la fila por ENCABEZADO de componente (<c>"Total"</c>, <c>"Componente TDF"</c>, …,
    /// <c>"Ajuste Decena"</c>). <c>null</c> = celda vacía en la fuente (distingue "ausente" de "cero").
    /// Las columnas destino sin contraparte en la fuente (<c>SERVICIO ESPECIALES</c>, <c>Ajuste
    /// Recargo Aseo</c>, <c>Otros</c>) NO se agregan aquí: quedan en 0 en la plantilla y T3 no las
    /// escribe (T0b §2.3).
    /// </summary>
    public IReadOnlyDictionary<string, decimal?> ValoresPorComponente { get; init; } =
        new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);

    /// <summary>Valor de la fila para el encabezado de componente indicado; <c>null</c> si no existe/está vacío.</summary>
    public decimal? Valor(string componente) =>
        ValoresPorComponente.TryGetValue(componente, out var valor) ? valor : null;
}
