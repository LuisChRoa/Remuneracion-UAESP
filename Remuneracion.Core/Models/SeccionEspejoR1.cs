namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 41 (T1, R-SEC-1/R-SEC-3 — D-A/D-H): tipo de SECCIÓN del bloque de <c>Reporte Componentes R1</c>
/// según la FRONTERA INMEDIATAMENTE ARRIBA de cada fila.
///
/// El bloque es una secuencia de secciones separadas por las filas-frontera <c>Mes/Total</c> y
/// <c>Aplicacion nuevos x reversion/Total</c>. La sección de una fila la determina la frontera que
/// la precede (arriba): <c>Mes</c> → Oportuno, <c>Aplicacion</c> → Aplicación, sin frontera arriba →
/// 1.ª sección. El <see cref="SeccionFronteraArriba.Ordinal"/> distingue las ocurrencias del mismo
/// tipo (Aplic1/Aplic2, Oportuno1/Oportuno2) — necesario porque la fuente y el oficial pueden
/// ordenar las secciones de forma distinta.
///
/// Es un predicado PURO (sin I/O, sin OpenXML, sin períodos, sin nombres de empresa — D-H): lo
/// consume el escritor values-only por slot (<c>EscritorValoresSlotR1</c>) para colocar cada
/// fila-dato de empresa en el slot cuya sección coincide. Convive con
/// <see cref="R1FirmaInterior.EsAplicacionPorFrontera"/> (frontera por DEBAJO, Plan 35-T2, otro
/// consumidor): semánticas distintas que NO se mezclan (este plan no toca ese predicado).
/// </summary>
public enum TipoSeccionEspejoR1
{
    /// <summary>1.ª sección del bloque: sin frontera <c>Mes</c>/<c>Aplicacion</c> por arriba.</summary>
    Primera,

    /// <summary>Sección Oportuno: la frontera inmediatamente arriba es <c>Mes/Total</c>.</summary>
    Oportuno,

    /// <summary>Sección Aplicación: la frontera inmediatamente arriba es <c>Aplicacion…/Total</c>.</summary>
    Aplicacion,
}

/// <summary>
/// Sección del bloque R1: <see cref="Tipo"/> + <see cref="Ordinal"/> de la frontera de ese tipo
/// que la inicia (0 para la 1.ª sección). Dos filas pertenecen a la misma sección si y solo si su
/// <see cref="SeccionFronteraArriba"/> es igual.
/// </summary>
public readonly record struct SeccionFronteraArriba(TipoSeccionEspejoR1 Tipo, int Ordinal);

/// <summary>
/// Plan 41 (T1, R-SEC-1 — D-A): resolución de la sección de una fila del bloque R1 por su frontera
/// inmediatamente ARRIBA (nunca por debajo, nunca por posición física).
/// </summary>
public static class SeccionEspejoR1
{
    /// <summary>Sección de la 1.ª sección del bloque (sin frontera arriba).</summary>
    public static readonly SeccionFronteraArriba Primera = new(TipoSeccionEspejoR1.Primera, 0);

    /// <summary>
    /// Sección de la fila <paramref name="indice"/> = la frontera inmediatamente ARRIBA
    /// (<see cref="FilaEspejoR1.EsMesTotal"/> → Oportuno, <see cref="FilaEspejoR1.EsAplicacionTotal"/>
    /// → Aplicación), con el ordinal de esa ocurrencia. Sin frontera arriba → 1.ª sección
    /// (<see cref="TipoSeccionEspejoR1.Primera"/>, ordinal 0).
    /// </summary>
    public static SeccionFronteraArriba PorFronteraArriba(IReadOnlyList<FilaEspejoR1> filas, int indice)
    {
        ArgumentNullException.ThrowIfNull(filas);
        ArgumentOutOfRangeException.ThrowIfNegative(indice);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(indice, filas.Count);

        var mes = 0;
        var aplicacion = 0;
        var tipo = TipoSeccionEspejoR1.Primera;
        var ordinal = 0;

        for (var j = 0; j < indice; j++)
        {
            var fila = filas[j];
            if (fila.EsMesTotal)
            {
                mes++;
                tipo = TipoSeccionEspejoR1.Oportuno;
                ordinal = mes;
            }
            else if (fila.EsAplicacionTotal)
            {
                aplicacion++;
                tipo = TipoSeccionEspejoR1.Aplicacion;
                ordinal = aplicacion;
            }
        }

        return new SeccionFronteraArriba(tipo, ordinal);
    }
}
