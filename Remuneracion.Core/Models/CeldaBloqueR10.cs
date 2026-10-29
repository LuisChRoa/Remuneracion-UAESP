namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 62 (H1, D-H1.4): una celda del bloque del R10 que se copia a la hoja homónima de la
/// plantilla (<c>DetRetri{AAAAMMQ}</c> / <c>DetValiRetri{AAAAMMQ}</c>).
///
/// <para>
/// <b>Contrato de contenido (exactamente una forma por celda):</b>
/// </para>
/// <list type="bullet">
/// <item><see cref="Texto"/>: el TEXTO exacto de la celda (rótulos, sello de proceso, fechas y
/// horas del encabezado, pie del bloque).</item>
/// <item><see cref="Numerico"/>: el valor exacto de la celda numérica y de los montos del R10
/// (que la fuente trae como texto monetario es-CO: <c>"$17.450.250.543"</c>,
/// <c>"-$5.093.259.797"</c>). <b>Sin redondear</b>: el redondeo a entero es una regla de
/// ESCRITURA del CONSOLIDADO, no de lectura; redondear aquí perdería el dato de la fuente antes
/// de que la regla de escritura lo necesite.</item>
/// <item><b>Celda vacía</b>: <see cref="Texto"/> y <see cref="Numerico"/> nulos
/// (<see cref="Vacia"/>). La celda NO se omite del bloque: el rango <c>A3:O19</c> /
/// <c>A3:O14</c> se representa completo para que el pegado y su read-back recorran las mismas
/// coordenadas (una celda ausente y una vacía no son lo mismo).</item>
/// </list>
///
/// <para>
/// Puro: sin I/O ni dependencias de Excel. Las coordenadas son relativas al contrato del bloque
/// (<see cref="Fila"/> = fila de la hoja, <see cref="Columna"/> = letra de la columna).
/// </para>
/// </summary>
public sealed class CeldaBloqueR10
{
    /// <summary>Inicializa la celda del bloque con sus coordenadas y su contenido.</summary>
    /// <param name="fila">Fila de la hoja (1-based).</param>
    /// <param name="columna">Letra de la columna (p. ej. <c>"D"</c>).</param>
    /// <param name="texto">Texto exacto de la celda, o <c>null</c> si la celda no es texto o está vacía.</param>
    /// <param name="numerico">Valor exacto de la celda numérica o del monto, o <c>null</c> si no aplica.</param>
    public CeldaBloqueR10(int fila, string columna, string? texto, decimal? numerico)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(fila);
        ArgumentException.ThrowIfNullOrWhiteSpace(columna);

        Fila = fila;
        Columna = columna;
        Texto = texto;
        Numerico = numerico;
    }

    /// <summary>Fila de la hoja a la que pertenece la celda.</summary>
    public int Fila { get; }

    /// <summary>Letra de la columna (sin anclaje absoluto).</summary>
    public string Columna { get; }

    /// <summary>Texto exacto de la celda cuando su contenido es texto; <c>null</c> en otro caso.</summary>
    public string? Texto { get; }

    /// <summary>Valor exacto cuando la celda es numérica o un monto del R10; <c>null</c> en otro caso.</summary>
    public decimal? Numerico { get; }

    /// <summary>Referencia de la celda (<c>Columna</c> + <c>Fila</c>, p. ej. <c>"D9"</c>).</summary>
    public string Referencia => $"{Columna}{Fila}";

    /// <summary><c>true</c> cuando la celda está dentro del bloque pero no trae contenido.</summary>
    public bool Vacia => Texto is null && Numerico is null;
}
