namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 29 (T2, Unidad R — SOLO LECTURA): matriz de detalle de la hoja <c>Rem. Anticipos R2</c>
/// extraída de la fuente <c>RerpoteDetalleSaldosaFavor_*</c>, resuelta por LABEL de fila (A–D) y
/// por ENCABEZADO de componente (col E..), NUNCA por corrimiento fijo.
///
/// Motivo (veredicto T0b, <c>plans/29-T0-Evidencia-B.md</c> §2.1): la malla DERIVA entre períodos
/// —julio trae E..P sin <c>Especiales</c>; agosto trae E..Q con <c>Especiales</c> en K— y las
/// filas LIME/BOGOTA-agosto agregan un sub-bloque de componente. Un mapa posicional se rompe; la
/// resolución por label/encabezado es la única robusta (réplica del patrón de firmas del Plan 25).
///
/// Modelo PURO de dominio (sin I/O ni celdas de plantilla): expone la observación fiel de la
/// fuente. El mapeo a las celdas-destino de la plantilla es responsabilidad de la ESCRITURA (T3),
/// que sí conoce el template; este modelo entrega la matriz que T3 mapea por label/encabezado.
///
/// Este tipo NO reemplaza a <see cref="WorkbookLeafInputsR2"/> (agregados E15/E26/K15): convive
/// con él. Aquel alimenta la cadena de fórmulas; este, el desglose-detalle por componente.
/// </summary>
public sealed class DetalleR2AseInputs
{
    /// <summary>ASE asociado a la fuente leída.</summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Encabezados de componente detectados en la fila de headers de la fuente, en orden físico
    /// (p. ej. <c>"Total"</c>, <c>"Componente TDF"</c>, …). Es la clave de <see cref="DetalleR2Fila.ValoresPorComponente"/>.
    /// </summary>
    public IReadOnlyList<string> Componentes { get; set; } = [];

    /// <summary>
    /// Verdad si la fuente trae la columna <c>Especiales</c> (agosto, posición K). En su ausencia
    /// (julio) cada fila expone <c>Especiales = 0</c> de forma EXPLÍCITA (invariante: nunca se
    /// inventa otro valor; el <c>K</c>-destino queda en 0). Réplica del patrón
    /// <c>ConstruirFilaEspejo</c> (R-E-6) de la columna <c>SERVICIO ESPECIALES</c>.
    /// </summary>
    public bool TieneColumnaEspeciales { get; set; }

    /// <summary>
    /// Filas de detalle en orden físico, desde la primera <c>Vlr Servicio</c> hasta la fila de
    /// cierre <c>Total</c> (A='Total') inclusive. Incluye las filas de bloque-empresa y de código
    /// de componente tal como aparecen en la fuente (la secuencia observada ES la especificación
    /// del período; no se normaliza ni se reordena).
    /// </summary>
    public IReadOnlyList<DetalleR2Fila> Filas { get; set; } = [];
}

/// <summary>
/// Fila tipada del detalle R2: firma de labels (A–D) + valores por ENCABEZADO de componente.
/// </summary>
public sealed class DetalleR2Fila
{
    /// <summary>Etiqueta de la columna A (p. ej. sección <c>Componente</c>/<c>Subs/Cont</c>/<c>Total</c> o una empresa del bloque ENEL).</summary>
    public string A { get; init; } = string.Empty;

    /// <summary>Etiqueta de la columna B (p. ej. empresa <c>ENEL</c>/<c>OCCIDENTE</c> o <c>Total</c>).</summary>
    public string B { get; init; } = string.Empty;

    /// <summary>Etiqueta de la columna C (p. ej. código de sub-bloque <c>E</c>/<c>H</c>/<c>O</c>/<c>T</c>/<c>1</c>/<c>2</c>/<c>5</c>/<c>9</c>/<c>A</c> o <c>Total</c>).</summary>
    public string C { get; init; } = string.Empty;

    /// <summary>Etiqueta de la columna D (concepto: <c>Vlr Servicio</c>, <c>Vlr Intereses</c>, <c>Total</c>, <c>Subsidio(-)/Contribucion(+)</c>).</summary>
    public string D { get; init; } = string.Empty;

    /// <summary>Firma de la fila como tupla <c>A|B|C|D</c> (mismo criterio de la evidencia T0b).</summary>
    public string Firma => $"{A}|{B}|{C}|{D}";

    /// <summary>
    /// Valores de la fila por ENCABEZADO de componente (<c>"Total"</c>, <c>"Componente TDF"</c>, …,
    /// <c>"Especiales"</c>). <c>null</c> = celda vacía en la fuente (distingue "ausente" de "cero";
    /// R-E-5). Cuando la fuente no trae <c>Especiales</c> se agrega <c>0</c> explícito (nunca <c>null</c>).
    /// </summary>
    public IReadOnlyDictionary<string, decimal?> ValoresPorComponente { get; init; } =
        new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);

    /// <summary>Valor de la fila para el encabezado de componente indicado; <c>null</c> si no existe/está vacío.</summary>
    public decimal? Valor(string componente) =>
        ValoresPorComponente.TryGetValue(componente, out var valor) ? valor : null;
}
