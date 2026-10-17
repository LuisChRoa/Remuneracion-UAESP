namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 45 (H1.1) + Plan 46 (H1.1/S2.1/S2.2): contrato de SECCIÓN <b>PERIOD-FREE</b> por ASE para las
/// filas del interior de <c>Reporte Componentes R1</c>.
///
/// Problema probado en bytes (T0 del Plan 44 §2 + T0 del Plan 45 + T0 del Plan 46): la fuente del
/// período define la forma de sus secciones (fronteras <c>Mes/Total</c> y
/// <c>Aplicacion nuevos x reversion/Total</c>) y esa forma NO es fija. Un emparejamiento por
/// <c>(tipo, ordinal)</c> genérico desalinea ASE5, y tipar por la frontera de ARRIBA desalinea la 1.ª
/// sección (Plan 46, DEFECTO 2).
///
/// Regla derivada (geometría pura, sin I/O, sin período, sin empresa — D-H): cada sección (fuente y
/// oficial) se tipa por la frontera que la <b>CIERRA</b> (<see cref="SeccionPorCierre"/>,
/// <c>Mes</c>→Oportuno, <c>Aplicacion</c>→Aplicación, ordinal = ocurrencia de esa frontera) y la
/// k-ésima sección de la FUENTE se alinea con la próxima sección del OFICIAL <b>del mismo tipo</b> aún
/// no consumida (<see cref="Alinear"/>, emparejamiento monotónico/greedy por tipo). Formas vigentes:
/// <list type="bullet">
/// <item><c>[O#1, A#1, O#2, A#2, O#3]</c> (julio2 ASE1-4, agosto mayoría, Q1-ASE3/4) → identidad;</item>
/// <item><c>[A#1, O#1, A#2, O#2]</c> (julio2 ASE5, Q1-ASE1) →
/// <c>[A#1(615-628), O#2(630-647), A#2(649-659), O#3(661-673)]</c>: la 1.ª oficial queda sin
/// contraparte (cero) y el bloque <c>Componente/Subsidio</c> fuente cae en el <c>A#2</c>;</item>
/// <item><c>[O#1, O#2, O#3]</c> (agosto-ASE3 / Q1-ASE2/5, sin <c>Aplicacion</c>) →
/// <c>[O#1, O#2, O#3]</c>: se saltan las secciones <c>Aplic</c> del oficial.</item>
/// </list>
/// Además, las filas sin <c>D</c> propio se colocan por su <b>CANAL EFECTIVO</b>
/// (<see cref="CanalEfectivo"/>, Plan 46 §2.1). El contrato recibe SOLO la forma (tipos de sección
/// observados); nunca un string de período/empresa.
/// </summary>
public static class ContratoSeccionPeriodFreeR1
{
    /// <summary>
    /// Verdad si la fila PARTICIPA de la alineación por sección (filas de datos/componentes). Las
    /// filas-frontera (<c>Mes/Total</c>, <c>Aplicacion…/Total</c>) y AFaseo NO participan: se localizan
    /// por firma (su rótulo ya las identifica) y conservan el comportamiento histórico.
    /// </summary>
    public static bool ParticipaDeSeccionConformada(FilaEspejoR1 fila)
    {
        ArgumentNullException.ThrowIfNull(fila);
        return !fila.EsMesTotal
            && !fila.EsAplicacionTotal
            && !EsFirmaAFaseo(fila.A, fila.B, fila.D, fila.E);
    }

    /// <summary>
    /// Plan 46 (§2.2, DEFECTO 2): SECCIÓN de la fila por la frontera que la <b>CIERRA</b> (la
    /// inmediatamente DEBAJO: su fila resumen <c>Mes</c>/<c>Aplicacion…</c> cuyo importe es su
    /// subtotal), no por la que la abre. Geometría pura (cero período/empresa, D-H):
    /// <list type="bullet">
    /// <item>fila de datos → la sección cerrada por la próxima frontera hacia abajo;</item>
    /// <item>fila-frontera → la sección que ELLA cierra (<c>Mes</c>→Oportuno, <c>Aplicacion</c>→Aplicación,
    /// ordinal = ocurrencia de esa frontera dentro del bloque);</item>
    /// <item>AFaseo → <c>null</c> (se localiza por firma);</item>
    /// <item>cola sin frontera debajo (Subs/Cont, Total) → la sección de la frontera inmediatamente
    /// ARRIBA (<see cref="SeccionEspejoR1.PorFronteraArriba"/>, que conserva su API intacta).</item>
    /// </list>
    /// Es la resolución ÚNICA del interior por sección en el contrato PERIOD-FREE-FREE: absorbe las 3
    /// formas probadas (<c>[M,A,M,A,M]</c>, <c>[A,M,A,M]</c>, <c>[M,M,M]</c>) sin ramas por
    /// período/empresa. NO sustituye ni altera <see cref="SeccionEspejoR1.PorFronteraArriba"/>
    /// (consumidor Plan 35-T2, <see cref="R1FirmaInterior.EsAplicacionPorFrontera"/>, intacto).
    /// </summary>
    public static SeccionFronteraArriba? SeccionPorCierre(IReadOnlyList<FilaEspejoR1> filas, int indice)
    {
        ArgumentNullException.ThrowIfNull(filas);
        ArgumentOutOfRangeException.ThrowIfNegative(indice);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(indice, filas.Count);

        var fila = filas[indice];
        if (EsFirmaAFaseo(fila.A, fila.B, fila.D, fila.E))
        {
            return null;
        }

        if (fila.EsMesTotal)
        {
            return new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, CuentaFronteraHasta(filas, indice, mes: true));
        }

        if (fila.EsAplicacionTotal)
        {
            return new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, CuentaFronteraHasta(filas, indice, mes: false));
        }

        for (var j = indice + 1; j < filas.Count; j++)
        {
            if (filas[j].EsMesTotal || filas[j].EsAplicacionTotal)
            {
                return SeccionPorCierre(filas, j);
            }
        }

        // Cola sin frontera de cierre: se queda con la sección de la frontera inmediatamente arriba.
        return SeccionEspejoR1.PorFronteraArriba(filas, indice);
    }

    private static int CuentaFronteraHasta(IReadOnlyList<FilaEspejoR1> filas, int indice, bool mes)
    {
        var cuenta = 0;
        for (var j = 0; j <= indice; j++)
        {
            if (mes ? filas[j].EsMesTotal : filas[j].EsAplicacionTotal)
            {
                cuenta++;
            }
        }

        return cuenta;
    }

    /// <summary>
    /// Plan 46 (§2.1, DEFECTO 1): CANAL EFECTIVO de la fila — el <c>D</c> de la fila resumen
    /// <c>…/Total</c> que CIERRA su grupo de canal (la próxima hacia abajo con <c>D</c> no vacío y
    /// distinto del literal <c>"Total"</c>). Solo aplica a las filas SIN <c>D</c> propio (las que traen
    /// <c>D</c> conservan el suyo: totales, <c>C=… D=Total</c>); las filas que no son de canal (traen
    /// <c>A</c> o <c>B</c>: AFaseo, <c>Componente</c>, fronteras, <c>Subs/Cont</c>, <c>Total</c>) no
    /// tienen canal efectivo (<c>null</c>). Predicado PURO, sin I/O, sin período, sin empresa (D-H).
    /// El MISMO cálculo se aplica a las filas de la fuente y a los slots del oficial.
    /// </summary>
    public static string? CanalEfectivo(IReadOnlyList<FilaEspejoR1> filas, int indice)
    {
        ArgumentNullException.ThrowIfNull(filas);
        ArgumentOutOfRangeException.ThrowIfNegative(indice);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(indice, filas.Count);

        var fila = filas[indice];
        if (!string.IsNullOrWhiteSpace(fila.D))
        {
            // Con D propio (totales de canal, C=… D=Total): conserva su D, no se recalcula.
            return null;
        }

        if (!string.IsNullOrWhiteSpace(fila.A) || !string.IsNullOrWhiteSpace(fila.B))
        {
            // No es fila de canal (AFaseo/Componente/frontera/Subs/Cont/Total): se localiza por firma.
            return null;
        }

        for (var j = indice + 1; j < filas.Count; j++)
        {
            var d = filas[j].D;
            if (!string.IsNullOrWhiteSpace(d)
                && !string.Equals(d.Trim(), "Total", StringComparison.OrdinalIgnoreCase))
            {
                return d.Trim();
            }
        }

        return null;
    }

    /// <summary>
    /// Alinea las secciones de la FUENTE (en orden) con las del OFICIAL (en orden) por tipo,
    /// monotónico greedy: cada sección fuente consume la próxima sección oficial del mismo
    /// <see cref="TipoSeccionEspejoR1"/> aún no consumida. Una sección fuente sin contraparte del
    /// mismo tipo (contrato agotado) queda SIN entrada → sus filas caerán a cero documentado o
    /// <c>ERR-SIN-SLOT</c> (nunca remap por posición).
    /// </summary>
    public static IReadOnlyDictionary<SeccionFronteraArriba, SeccionFronteraArriba> Alinear(
        IReadOnlyList<SeccionFronteraArriba> seccionesFuente,
        IReadOnlyList<SeccionFronteraArriba> seccionesOficial)
    {
        ArgumentNullException.ThrowIfNull(seccionesFuente);
        ArgumentNullException.ThrowIfNull(seccionesOficial);

        var mapa = new Dictionary<SeccionFronteraArriba, SeccionFronteraArriba>();
        var cursor = 0;
        foreach (var fuente in seccionesFuente)
        {
            var encontrado = -1;
            for (var j = cursor; j < seccionesOficial.Count; j++)
            {
                if (seccionesOficial[j].Tipo == fuente.Tipo)
                {
                    encontrado = j;
                    break;
                }
            }

            if (encontrado < 0)
            {
                continue;
            }

            mapa[fuente] = seccionesOficial[encontrado];
            cursor = encontrado + 1;
        }

        return mapa;
    }

    private static bool EsFirmaAFaseo(string a, string b, string d, string e) =>
        (Igual(a, "AFaseo") && Igual(b, "Total"))
        || (Igual(d, "AFASEO") && Igual(e, "Total"));

    private static bool Igual(string? x, string? y) =>
        string.Equals(x?.Trim(), y?.Trim(), StringComparison.OrdinalIgnoreCase);
}
