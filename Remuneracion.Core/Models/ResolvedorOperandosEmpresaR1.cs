namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 53 (H1.1, D1/D4/D5): RESOLUTOR PURO del operando-empresa de R1 por <b>sección declarada</b>.
///
/// Reemplaza la resolución por ocurrencia/índice sobre listas ordenadas por <c>|F|</c>
/// (<c>ExcelDataReaderWorkbookLeafInputReader.cs:1585-1637</c>: <c>Main=[0]</c>,
/// <c>Medio=[1]</c>, <c>Menor=[count-1]</c>), prohibida por el Plan 52 §13 y el Plan 41 E1-1.
///
/// Contrato (recibe SOLO filas-empresa ya tipadas por el reader; sin I/O, sin período, sin empresa —
/// D-H):
/// <list type="bullet">
/// <item>Las <b>secciones del bloque</b> son las <see cref="FilaSeccionadaR1.Seccion"/> distintas en
/// orden de aparición (la fuente define la forma; T0-52 §7 probó 0 intercalados).</item>
/// <item><see cref="SeccionEsperadaR1.Primera"/> → la primera sección <b>solo si hay ≥ 2</b> (con una
/// sola no existe "el otro" bloque); ausente → la celda queda en 0.</item>
/// <item><see cref="SeccionEsperadaR1.UltimaOportuna"/> / <see cref="SeccionEsperadaR1.UnicaRecaudo"/>
/// → la última sección del bloque (misma resolución; ver el enum para el valor contractual de
/// <c>UnicaRecaudo</c>).</item>
/// <item><see cref="SeccionEsperadaR1.Intermedia"/> → la segunda sección <b>solo si hay ≥ 3</b>
/// (D4); con menos → 0.</item>
/// <item><b>Empate</b> (2 filas-empresa en la MISMA sección) → señal de STOP
/// (<see cref="ResultadoResolucionR1.EsAmbigua"/>): no hay regla de desempate aprobada, el reader
/// lanza nombrando ASE + empresa + sección + filas.</item>
/// <item>Sección declarada <b>ausente</b> → <see cref="ResultadoResolucionR1.EsAusente"/>: el reader
/// materializa <c>0m</c> (la fuente no trae esa sección; es normal, reglas 6 y 7 del
/// project-context).</item>
/// </list>
/// Cero ramas por período/empresa: la forma la define la fuente.
/// </summary>
public static class ResolvedorOperandosEmpresaR1
{
    /// <summary>
    /// Resuelve la fila de la sección declarada dentro del bloque de la empresa.
    /// </summary>
    /// <param name="filasBloque">
    /// Filas-empresa del bloque (Totales o Subs, según el operando) ya tipadas, en orden de aparición.
    /// </param>
    /// <param name="esperada">Sección declarada por la celda del contrato.</param>
    public static ResultadoResolucionR1 Resolver(
        IReadOnlyList<FilaSeccionadaR1> filasBloque,
        SeccionEsperadaR1 esperada)
    {
        ArgumentNullException.ThrowIfNull(filasBloque);

        var secciones = Secciones(filasBloque);

        // Gate de empate (Plan 53 §6/R2): dos filas-empresa en la misma sección = ambigüedad sin regla
        // aprobada. Se evalúa ANTES de elegir rol (la ambigüedad no depende del operando).
        foreach (var seccion in secciones)
        {
            var enSeccion = filasBloque.Where(f => f.Seccion == seccion).Select(f => f.Fila).ToList();
            if (enSeccion.Count > 1)
            {
                return ResultadoResolucionR1.Ambiguedad(seccion, enSeccion);
            }
        }

        SeccionFronteraArriba? objetivo = esperada switch
        {
            SeccionEsperadaR1.Primera => secciones.Count >= 2 ? secciones[0] : null,
            SeccionEsperadaR1.UltimaOportuna or SeccionEsperadaR1.UnicaRecaudo =>
                secciones.Count >= 1 ? secciones[^1] : null,
            SeccionEsperadaR1.Intermedia => secciones.Count >= 3 ? secciones[1] : null,
            _ => null,
        };

        if (objetivo is null)
        {
            return ResultadoResolucionR1.Ausente;
        }

        var fila = filasBloque.First(f => f.Seccion == objetivo.Value);
        return ResultadoResolucionR1.Resuelto(fila.Fila, objetivo.Value);
    }

    private static List<SeccionFronteraArriba> Secciones(IReadOnlyList<FilaSeccionadaR1> filas)
    {
        var secciones = new List<SeccionFronteraArriba>();
        foreach (var fila in filas)
        {
            if (!secciones.Contains(fila.Seccion))
            {
                secciones.Add(fila.Seccion);
            }
        }

        return secciones;
    }
}

/// <summary>
/// Resultado de <see cref="ResolvedorOperandosEmpresaR1.Resolver"/>: fila resuelta, sección declarada
/// ausente (0 explícito) o empate de sección (señal de STOP para el reader).
/// </summary>
public sealed class ResultadoResolucionR1
{
    private ResultadoResolucionR1(int? fila, SeccionFronteraArriba? seccion, IReadOnlyList<int> filasAmbiguas)
    {
        Fila = fila;
        Seccion = seccion;
        FilasAmbiguas = filasAmbiguas;
    }

    /// <summary>La sección declarada no existe en la fuente: la celda queda en 0 explícito.</summary>
    public static ResultadoResolucionR1 Ausente { get; } = new(null, null, []);

    /// <summary>Fila de la sección declarada.</summary>
    public int? Fila { get; }

    /// <summary>Sección de la fila resuelta (o la sección con empate).</summary>
    public SeccionFronteraArriba? Seccion { get; }

    /// <summary>Filas-empresa que comparten la sección con empate (vacío si no hay empate).</summary>
    public IReadOnlyList<int> FilasAmbiguas { get; }

    /// <summary>Verdad si la sección declarada no existe en la fuente (0 explícito).</summary>
    public bool EsAusente => Fila is null && FilasAmbiguas.Count == 0;

    /// <summary>Verdad si hay 2+ filas-empresa en la misma sección (STOP).</summary>
    public bool EsAmbigua => FilasAmbiguas.Count > 0;

    /// <summary>Construye un resultado resuelto.</summary>
    public static ResultadoResolucionR1 Resuelto(int fila, SeccionFronteraArriba seccion) =>
        new(fila, seccion, []);

    /// <summary>Construye una señal de empate de sección.</summary>
    public static ResultadoResolucionR1 Ambiguedad(SeccionFronteraArriba seccion, IReadOnlyList<int> filas) =>
        new(null, seccion, filas);
}
