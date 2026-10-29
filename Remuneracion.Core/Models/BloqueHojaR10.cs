namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 62 (H1, D-H1.4): bloque completo del insumo <c>R10_Remuneracion_{AAAAMMQ}.xlsx</c> que se
/// copia a la hoja homónima de la plantilla de procesamiento (N1).
///
/// <para>
/// Dos bloques por período, con la geometría del contrato T0 (§T0-01/04): <c>DetRetri</c> filas
/// 3–19 y <c>DetValiRetri</c> filas 3–14, columnas A–O. El bloque viaja <b>completo</b>, incluidos
/// los encabezados corporativos de las filas 3–8 (razón social, NIT, código de remuneración del
/// período, fecha y hora de proceso, rango del período y encabezado de la tabla): la aclaración
/// ratificada del Plan 62 (§11, R-Aclareo) establece que el pegado ES el transportador de esos
/// encabezados, no se filtran ni se reconstruyen.
/// </para>
///
/// <para>
/// Puro: sin I/O. La hoja se resuelve con <see cref="NombresHojaPeriodo"/> en el lector de
/// Infrastructure (nunca por enumeración del workbook) y el período es el del dominio: el modelo no
/// tiene ramas por período, quincena ni ASE.
/// </para>
/// </summary>
public sealed class BloqueHojaR10
{
    /// <summary>Inicializa el bloque con su hoja de origen, su período y sus celdas.</summary>
    /// <param name="hoja">Nombre de la hoja del R10 (compuesto con <see cref="NombresHojaPeriodo"/>).</param>
    /// <param name="periodo">Período del dominio al que pertenece el bloque.</param>
    /// <param name="celdas">Celdas del rango del contrato, en orden de fila y columna.</param>
    public BloqueHojaR10(string hoja, Periodo periodo, IReadOnlyList<CeldaBloqueR10> celdas)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(hoja);
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(celdas);

        Hoja = hoja;
        Periodo = periodo;
        Celdas = celdas;
    }

    /// <summary>Nombre de la hoja de la que salió el bloque.</summary>
    public string Hoja { get; }

    /// <summary>Período del dominio (<c>CodigoCompleto</c> = <c>AAAAMMQ</c>).</summary>
    public Periodo Periodo { get; }

    /// <summary>Celdas del bloque, en orden de fila ascendente y columna ascendente.</summary>
    public IReadOnlyList<CeldaBloqueR10> Celdas { get; }

    /// <summary>Primera fila cubierta por el bloque (0 si el bloque no trae celdas).</summary>
    public int FilaDesde => Celdas.Count == 0 ? 0 : Celdas.Min(celda => celda.Fila);

    /// <summary>Última fila cubierta por el bloque (0 si el bloque no trae celdas).</summary>
    public int FilaHasta => Celdas.Count == 0 ? 0 : Celdas.Max(celda => celda.Fila);

    /// <summary>
    /// Devuelve la celda de la coordenada indicada, o <c>null</c> si la coordenada cae fuera del
    /// bloque. Una celda dentro del bloque siempre existe (vacía o con contenido): el <c>null</c>
    /// solo significa "fuera del rango del contrato".
    /// </summary>
    /// <param name="fila">Fila de la hoja.</param>
    /// <param name="columna">Letra de la columna (p. ej. <c>"D"</c>).</param>
    public CeldaBloqueR10? Celda(int fila, string columna) =>
        Celdas.FirstOrDefault(celda =>
            celda.Fila == fila && celda.Columna.Equals(columna, StringComparison.OrdinalIgnoreCase));
}
