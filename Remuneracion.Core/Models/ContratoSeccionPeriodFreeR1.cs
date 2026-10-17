namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 45 (H1.1, S2.1 — §2.1/§2.2 del plan; Enmienda E-44.1): contrato de SECCIÓN <b>PERIOD-FREE</b>
/// por ASE para las filas del interior de <c>Reporte Componentes R1</c>.
///
/// Problema probado en bytes (T0 del Plan 44 §2 + T0 del Plan 45): la fuente del período define la
/// forma de sus secciones (fronteras <c>Mes/Total</c> y <c>Aplicacion nuevos x reversion/Total</c>) y
/// esa forma NO es fija: ASE5-julio-2 arranca con <c>Aplicacion</c> (5 secciones) mientras el oficial
/// declara siempre 6 (<c>1.ª · Oportuno#1 · Aplic#1 · Oportuno#2 · Aplic#2 · Oportuno#3</c>). Un
/// emparejamiento por <c>(tipo, ordinal)</c> genérico desalinea ASE5.
///
/// Regla derivada (geometría pura, sin I/O, sin período, sin empresa — D-H): la k-ésima sección de la
/// FUENTE se alinea con la próxima sección del OFICIAL <b>del mismo tipo</b> aún no consumida
/// (emparejamiento monotónico/greedy por tipo). Reproduce el mapeo del plan §2.2:
/// <list type="bullet">
/// <item><c>[1.ª, Aplic#1, Oportuno#1, Aplic#2, Oportuno#2]</c> (ASE5-julio2) →
/// <c>[1.ª, Aplic#1, Oportuno#2, Aplic#2, Oportuno#3]</c>: el bloque <c>Componente/Subsidio</c> fuente
/// cae en el <c>Oportuno#2</c> oficial (649-659), no en el <c>Oportuno#1</c>;</item>
/// <item><c>[1.ª, Oportuno#1, Oportuno#2, Oportuno#3]</c> (agosto-ASE3 / Q1 sin <c>Aplicacion</c>) →
/// <c>[1.ª, Oportuno#1, Oportuno#2, Oportuno#3]</c>: se saltan las secciones <c>Aplic</c> del oficial;</item>
/// <item><c>[1.ª, Oportuno#1, Aplic#1, Oportuno#2, Aplic#2, Oportuno#3]</c> (forma completa) → identidad.</item>
/// </list>
/// El contrato recibe SOLO la forma (tipos de sección observados); nunca un string de período/empresa.
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
