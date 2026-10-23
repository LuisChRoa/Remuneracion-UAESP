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

    // ═══════════════════════════════════════════════════════════════════════════════════════════
    // [P56-JERARQUIA-AB-INICIO] camino de resolución de la identidad de cuadro (Plan 56 §2-D2/D3).
    // ═══════════════════════════════════════════════════════════════════════════════════════════
    //
    // Identidad del cuadro = (sección A, subsección B, ocurrencia de B dentro de A), LEÍDA de las
    // etiquetas de cierre de la propia hoja (nunca declarada aquí):
    //   · cierre de SUBsección: `B=<sub>` con `C='Total'` (A vacía);
    //   · cierre de SECCIÓN:    `A=<sec>` con `B='Total'`;
    //   · gran total del bloque: `A='Total'` con `B` vacía.
    // Dentro del cuadro la llave es fuente C → canal D (efectivo) → concepto E, y los valores son
    // componentes por ENCABEZADO (los resuelve el writer, no este predicado).
    //
    // Reglas (§2-D3): ausencia en la fuente ⇒ cero del template (no hay fila que colocar, no es
    // error); combinación presente y ausente en TODO el bloque ⇒ fallo accionable (sin escribir);
    // fila sin destino en SU cuadro pero declarada en el bloque ⇒ cero documentado si el cuadro
    // cierra al centavo por sus filas-hoja, fallo si no cierra.
    //
    // Geometría pura (D-H): cero identificadores de bloque/ASE, cero nombres de empresa (solo el
    // catálogo abierto <see cref="SinonimosEmpresaR1"/>), cero ramas por período, cero direcciones
    // de celda congeladas. No hay fallback posicional: el emparejamiento es por identidad y firma.
    // ═══════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>Tolerancia del cierre aritmético del cuadro (misma del resto del dominio).</summary>
    private const decimal ToleranciaCuadroR1 = 0.5m;

    /// <summary>
    /// Identidad del cuadro: <paramref name="Seccion"/> (columna A), <paramref name="Subseccion"/>
    /// (columna B) y <see cref="Ocurrencia"/> de esa subsección DENTRO de su sección (1-based).
    /// </summary>
    public readonly record struct CuadroR1(string Seccion, string Subseccion, int Ocurrencia);

    /// <summary>
    /// Subsección medida por su etiqueta de cierre (<c>B=&lt;sub&gt;</c> con <c>C='Total'</c>).
    /// <see cref="Inicio"/>..<see cref="Fin"/> es inclusivo y <see cref="Fin"/> ES la fila de cierre
    /// (las filas de datos del cuadro son <c>Inicio..Fin-1</c>).
    /// </summary>
    public sealed record SubseccionR1(string Etiqueta, int Ocurrencia, int Inicio, int Fin);

    /// <summary>
    /// Sección medida por su etiqueta de cierre (<c>A=&lt;sec&gt;</c> con <c>B='Total'</c>).
    /// <see cref="Fin"/> es la fila de cierre de la sección; sus subsecciones viven antes.
    /// </summary>
    public sealed record SeccionR1(
        string Etiqueta,
        int Ocurrencia,
        int Inicio,
        int Fin,
        IReadOnlyList<SubseccionR1> Subsecciones)
    {
        /// <summary>Subsección con esa etiqueta y ocurrencia dentro de esta sección, o <c>null</c>.</summary>
        public SubseccionR1? BuscarSubseccion(string etiqueta, int ocurrencia) =>
            Subsecciones.FirstOrDefault(s => Igual(s.Etiqueta, etiqueta) && s.Ocurrencia == ocurrencia);
    }

    /// <summary>
    /// Jerarquía leída de una secuencia de filas (bloque completo de una hoja). <see cref="FilaTotal"/>
    /// es el índice del gran total del bloque (<c>A='Total'</c>, <c>B</c> vacía) o <c>null</c>.
    /// </summary>
    public sealed record JerarquiaR1(IReadOnlyList<SeccionR1> Secciones, int? FilaTotal)
    {
        /// <summary>Verdad si la secuencia declara al menos una sección por su etiqueta de cierre.</summary>
        public bool DeclaraSecciones => Secciones.Count > 0;

        /// <summary>Sección con esa etiqueta y ocurrencia dentro del bloque, o <c>null</c>.</summary>
        public SeccionR1? BuscarSeccion(string etiqueta, int ocurrencia) =>
            Secciones.FirstOrDefault(s => Igual(s.Etiqueta, etiqueta) && s.Ocurrencia == ocurrencia);
    }

    /// <summary>
    /// Fila fuente sin destino en su cuadro que quedó en el CERO del template porque el cuadro
    /// cierra al centavo sin ella (registro durable: sección, subsección, ocurrencia, firma y valor).
    /// </summary>
    public sealed record CeroDocumentadoR1(
        CuadroR1 Cuadro,
        string Firma,
        string Fuente,
        string Canal,
        string Concepto,
        decimal? Total);

    /// <summary>
    /// Punto de fallo accionable (§2-D3): la fila fuente que no tiene destino declarado en el bloque
    /// o cuyo cuadro no cierra sin ella. Nombra sección + subsección + fuente + canal + concepto para
    /// que el writer componga el mensaje en vocabulario de la propietaria.
    /// </summary>
    public sealed record FalloJerarquiaR1(
        int IndiceFilaFuente,
        string? Seccion,
        string? Subseccion,
        int Ocurrencia,
        string Fuente,
        string Canal,
        string Concepto,
        string Motivo);

    /// <summary>
    /// Plan de colocación por identidad A/B. <see cref="Colocaciones"/> empareja índices de la fuente
    /// con índices de la plantilla (ambos del mismo bloque). <see cref="AplicaJerarquia"/> es falso
    /// cuando la fuente NO declara secciones (secuencia sintética sin cierres A/B): en ese caso el
    /// writer conserva el camino histórico por firma, que no decide posición alguna.
    /// </summary>
    public sealed record PlanJerarquiaR1(
        bool AplicaJerarquia,
        IReadOnlyList<(int Fuente, int Slot)> Colocaciones,
        IReadOnlyList<CeroDocumentadoR1> Ceros,
        FalloJerarquiaR1? Fallo)
    {
        /// <summary>La fuente no declara jerarquía: el writer usa el camino histórico por firma.</summary>
        public static PlanJerarquiaR1 SinJerarquia { get; } = new(false, [], [], null);
    }

    /// <summary>
    /// Particiona un bloque por sus ETIQUETAS de cierre (columnas A y B): secciones y subsecciones con
    /// su ocurrencia, rango de filas y fila de cierre. Cero I/O, cero empresa/período (D-H).
    /// Las filas que no caen en ninguna sección (p. ej. una cola sin cierre de subsección) quedan
    /// fuera de la jerarquía y el plan las reporta como fallo: nunca se descartan en silencio.
    /// </summary>
    public static JerarquiaR1 ParticionarJerarquia(IReadOnlyList<FilaEspejoR1> filas)
    {
        ArgumentNullException.ThrowIfNull(filas);

        var secciones = new List<SeccionR1>();
        var ocurrencias = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        int? filaTotal = null;
        var inicio = 0;
        for (var i = 0; i < filas.Count; i++)
        {
            var fila = filas[i];
            if (EsTotalBloqueR1(fila))
            {
                filaTotal ??= i;
                continue;
            }

            if (!EsCierreSeccionR1(fila))
            {
                continue;
            }

            var etiqueta = fila.A.Trim();
            ocurrencias.TryGetValue(etiqueta, out var cuenta);
            ocurrencias[etiqueta] = ++cuenta;
            secciones.Add(new SeccionR1(
                etiqueta,
                cuenta,
                inicio,
                i,
                ParticionarSubseccionesR1(filas, inicio, i - 1)));
            inicio = i + 1;
        }

        return new JerarquiaR1(secciones, filaTotal);
    }

    /// <summary>
    /// Firma de la fila DENTRO de su cuadro: <c>C | D efectivo | E</c>, con el canal efectivo buscado
    /// solo hacia abajo y SOLO dentro del cuadro (<paramref name="limiteExclusivo"/> = fila de cierre,
    /// no incluida). Cero I/O (D-H).
    /// </summary>
    public static string FirmaCuadroR1(IReadOnlyList<FilaEspejoR1> filas, int indice, int limiteExclusivo)
    {
        ArgumentNullException.ThrowIfNull(filas);
        ArgumentOutOfRangeException.ThrowIfNegative(indice);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(indice, filas.Count);

        var fila = filas[indice];
        return $"{fila.C.Trim()}|{CanalEfectivoEnCuadroR1(filas, indice, limiteExclusivo)}|{fila.E.Trim()}";
    }

    /// <summary>
    /// Planifica el bloque completo por identidad <c>(A, B, ocurrencia)</c> + firma
    /// <c>C/D(efectivo)/E</c>, aplicando las reglas de ausencia y de destino inexistente de §2-D3.
    /// NO escribe nada y NO elige posiciones: cada fila va a la fila que la plantilla declara para su
    /// identidad y su firma (k-ésima ↔ k-ésima cuando la firma se repite dentro del cuadro).
    /// </summary>
    public static PlanJerarquiaR1 PlanificarJerarquiaR1(
        IReadOnlyList<FilaEspejoR1> fuente,
        IReadOnlyList<FilaEspejoR1> plantilla)
    {
        ArgumentNullException.ThrowIfNull(fuente);
        ArgumentNullException.ThrowIfNull(plantilla);

        var jerarquiaFuente = ParticionarJerarquia(fuente);
        if (!jerarquiaFuente.DeclaraSecciones)
        {
            return PlanJerarquiaR1.SinJerarquia;
        }

        var jerarquiaPlantilla = ParticionarJerarquia(plantilla);
        var colocaciones = new List<(int Fuente, int Slot)>();
        var pendientes = new List<(int Indice, CuadroR1 Cuadro, SubseccionR1 Sub, string Firma)>();
        var cubiertas = new bool[fuente.Count];

        foreach (var seccionFuente in jerarquiaFuente.Secciones)
        {
            var seccionPlantilla = jerarquiaPlantilla.BuscarSeccion(seccionFuente.Etiqueta, seccionFuente.Ocurrencia);
            if (seccionPlantilla is null)
            {
                return PlanJerarquiaR1.SinJerarquia with
                {
                    AplicaJerarquia = true,
                    Fallo = FalloDeSeccionR1(fuente, seccionFuente),
                };
            }

            // Cierre de SECCIÓN: destino declarado por la etiqueta A (no participa de la firma C/D/E).
            colocaciones.Add((seccionFuente.Fin, seccionPlantilla.Fin));
            cubiertas[seccionFuente.Fin] = true;

            foreach (var subFuente in seccionFuente.Subsecciones)
            {
                var cuadro = new CuadroR1(seccionFuente.Etiqueta, subFuente.Etiqueta, subFuente.Ocurrencia);
                var subPlantilla = seccionPlantilla.BuscarSubseccion(subFuente.Etiqueta, subFuente.Ocurrencia);
                if (subPlantilla is null)
                {
                    return PlanJerarquiaR1.SinJerarquia with
                    {
                        AplicaJerarquia = true,
                        Fallo = FalloDeCuadroR1(fuente, subFuente, cuadro),
                    };
                }

                // Cierre de SUBsección: destino declarado por la identidad del cuadro.
                colocaciones.Add((subFuente.Fin, subPlantilla.Fin));
                MarcarRangoR1(cubiertas, subFuente.Inicio, subFuente.Fin);

                var destinos = IndiceFirmasR1(plantilla, subPlantilla);
                var usadas = new List<(string Canal, FilaEspejoR1 Fila)>();
                for (var i = subFuente.Inicio; i < subFuente.Fin; i++)
                {
                    var fila = fuente[i];
                    var canal = CanalEfectivoEnCuadroR1(fuente, i, subFuente.Fin);
                    var ocurrencia = usadas.Count(u => MismaFirmaR1(u.Fila, u.Canal, fila, canal));
                    usadas.Add((canal, fila));

                    var destino = destinos
                        .Where(g => MismaFirmaR1(fila, canal, g.Representante, g.Canal))
                        .FirstOrDefault(g => ocurrencia < g.Filas.Count);
                    if (destino is null)
                    {
                        pendientes.Add((i, cuadro, subFuente, $"{fila.C.Trim()}|{canal}|{fila.E.Trim()}"));
                        continue;
                    }

                    colocaciones.Add((i, destino.Filas[ocurrencia]));
                }
            }
        }

        if (jerarquiaFuente.FilaTotal is { } totalFuente)
        {
            if (jerarquiaPlantilla.FilaTotal is not { } totalPlantilla)
            {
                return PlanJerarquiaR1.SinJerarquia with
                {
                    AplicaJerarquia = true,
                    Fallo = FalloTotalR1(fuente, totalFuente),
                };
            }

            colocaciones.Add((totalFuente, totalPlantilla));
            cubiertas[totalFuente] = true;
        }

        // Fila fuera de toda sección/subsección/total: no se descarta en silencio (R-SILENCIO).
        for (var i = 0; i < fuente.Count; i++)
        {
            if (!cubiertas[i])
            {
                return PlanJerarquiaR1.SinJerarquia with
                {
                    AplicaJerarquia = true,
                    Fallo = FalloDeFilaR1(fuente, i, seccion: null, subseccion: null, ocurrencia: 0, "fila fuera de la jerarquía declarada (sin cierre A/B que la contenga)"),
                };
            }
        }

        var ceros = new List<CeroDocumentadoR1>();
        var firmasDelBloque = FirmasDelBloqueR1(plantilla, jerarquiaPlantilla);
        foreach (var grupo in pendientes.GroupBy(p => (p.Cuadro, p.Sub.Inicio, p.Sub.Fin)))
        {
            var sub = grupo.First().Sub;
            var cuadro = grupo.First().Cuadro;

            // §2-D3 (destino inexistente): la combinación no está declarada en NINGÚN cuadro del
            // bloque ⇒ error accionable + alto, sin escribir. No es una fila "fuera de su cuadro".
            foreach (var pendiente in grupo)
            {
                var fila = fuente[pendiente.Indice];
                var canal = CanalEfectivoEnCuadroR1(fuente, pendiente.Indice, sub.Fin);
                if (firmasDelBloque.Any(f => MismaFirmaR1(fila, canal, f.Fila, f.Canal)))
                {
                    continue;
                }

                return PlanJerarquiaR1.SinJerarquia with
                {
                    AplicaJerarquia = true,
                    Fallo = FalloDeFilaR1(
                        fuente,
                        pendiente.Indice,
                        cuadro.Seccion,
                        cuadro.Subseccion,
                        cuadro.Ocurrencia,
                        "la combinación (fuente/canal/concepto) no está declarada en ningún cuadro del bloque de la plantilla",
                        canal),
                };
            }

            if (!CuadroCierraSinLasFilasR1(fuente, sub))
            {
                var primera = grupo.First();
                return PlanJerarquiaR1.SinJerarquia with
                {
                    AplicaJerarquia = true,
                    Fallo = FalloDeFilaR1(
                        fuente,
                        primera.Indice,
                        cuadro.Seccion,
                        cuadro.Subseccion,
                        cuadro.Ocurrencia,
                        "la fila no tiene destino en su cuadro y el subtotal del cuadro NO cierra sin ella",
                        CanalEfectivoEnCuadroR1(fuente, primera.Indice, sub.Fin)),
                };
            }

            foreach (var pendiente in grupo)
            {
                var fila = fuente[pendiente.Indice];
                var canal = CanalEfectivoEnCuadroR1(fuente, pendiente.Indice, sub.Fin);
                ceros.Add(new CeroDocumentadoR1(
                    cuadro,
                    pendiente.Firma,
                    fila.C.Trim(),
                    canal,
                    fila.E.Trim(),
                    fila.Valor("Total")));
            }
        }

        return new PlanJerarquiaR1(true, colocaciones, ceros, null);
    }

    private static bool EsCierreSeccionR1(FilaEspejoR1 fila) =>
        fila.A.Trim().Length > 0 && Igual(fila.B, "Total");

    private static bool EsCierreSubseccionR1(FilaEspejoR1 fila) =>
        fila.A.Trim().Length == 0 && fila.B.Trim().Length > 0 && Igual(fila.C, "Total");

    private static bool EsTotalBloqueR1(FilaEspejoR1 fila) =>
        Igual(fila.A, "Total") && fila.B.Trim().Length == 0;

    private static List<SubseccionR1> ParticionarSubseccionesR1(IReadOnlyList<FilaEspejoR1> filas, int desde, int hasta)
    {
        var subsecciones = new List<SubseccionR1>();
        var ocurrencias = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var inicio = desde;
        for (var i = desde; i <= hasta; i++)
        {
            if (!EsCierreSubseccionR1(filas[i]))
            {
                continue;
            }

            var etiqueta = filas[i].B.Trim();
            ocurrencias.TryGetValue(etiqueta, out var cuenta);
            ocurrencias[etiqueta] = ++cuenta;
            subsecciones.Add(new SubseccionR1(etiqueta, cuenta, inicio, i));
            inicio = i + 1;
        }

        return subsecciones;
    }

    private static string CanalEfectivoEnCuadroR1(IReadOnlyList<FilaEspejoR1> filas, int indice, int limiteExclusivo)
    {
        var propio = filas[indice].D.Trim();
        if (propio.Length > 0)
        {
            return propio;
        }

        for (var j = indice + 1; j < limiteExclusivo && j < filas.Count; j++)
        {
            var canal = filas[j].D.Trim();
            if (canal.Length > 0 && !Igual(canal, "Total"))
            {
                return canal;
            }
        }

        return string.Empty;
    }

    /// <summary>
    /// Verdad si la fuente y la fila de la plantilla tienen la MISMA firma dentro del cuadro:
    /// <c>C</c> por el catálogo abierto de sinonimia (ambas vacías cuentan como iguales), canal
    /// efectivo y concepto por igualdad literal normalizada. Cero ramas por empresa (D-H).
    /// </summary>
    private static bool MismaFirmaR1(FilaEspejoR1 fuente, string canalFuente, FilaEspejoR1 plantilla, string canalPlantilla)
    {
        if (!Igual(canalFuente, canalPlantilla) || !Igual(fuente.E, plantilla.E))
        {
            return false;
        }

        var cFuente = fuente.C.Trim();
        var cPlantilla = plantilla.C.Trim();
        if (cFuente.Length == 0 || cPlantilla.Length == 0)
        {
            return cFuente.Length == cPlantilla.Length;
        }

        return SinonimosEmpresaR1.SonMismaEmpresa(cFuente, cPlantilla);
    }

    private sealed class GrupoFirmaR1
    {
        public required FilaEspejoR1 Representante { get; init; }

        public required string Canal { get; init; }

        public List<int> Filas { get; } = [];
    }

    /// <summary>
    /// Índice firma → filas declaradas (en orden físico) de los DATOS del cuadro de la plantilla
    /// (<c>Inicio..Fin-1</c>; la fila de cierre no es dato). Las firmas repetidas se emparejan por
    /// ordinal dentro de la firma: k-ésima fuente ↔ k-ésima declarada.
    /// </summary>
    private static List<GrupoFirmaR1> IndiceFirmasR1(IReadOnlyList<FilaEspejoR1> plantilla, SubseccionR1 cuadro)
    {
        var grupos = new List<GrupoFirmaR1>();
        for (var i = cuadro.Inicio; i < cuadro.Fin; i++)
        {
            var fila = plantilla[i];
            var canal = CanalEfectivoEnCuadroR1(plantilla, i, cuadro.Fin);
            var grupo = grupos.FirstOrDefault(g => MismaFirmaR1(fila, canal, g.Representante, g.Canal));
            if (grupo is null)
            {
                grupo = new GrupoFirmaR1 { Representante = fila, Canal = canal };
                grupos.Add(grupo);
            }

            grupo.Filas.Add(i);
        }

        return grupos;
    }

    /// <summary>
    /// Criterio de "no perder dato" de §2-D3, calculado en runtime y sin listas: el cuadro cierra al
    /// centavo por sus FILAS-HOJA (filas sin A/B/C/D: conceptos, no agregados de canal/empresa ni
    /// cierres) contra el subtotal declarado en su fila de cierre. Si cierra, la fila sin destino no
    /// aporta dinero que el cuadro no lleve ya en su subtotal; si no cierra, se perdería dato real.
    /// </summary>
    private static bool CuadroCierraSinLasFilasR1(IReadOnlyList<FilaEspejoR1> fuente, SubseccionR1 cuadro)
    {
        var subtotal = fuente[cuadro.Fin].Valor("Total");
        if (subtotal is null)
        {
            return false;
        }

        var suma = 0m;
        for (var i = cuadro.Inicio; i < cuadro.Fin; i++)
        {
            var fila = fuente[i];
            if (fila.A.Trim().Length > 0
                || fila.B.Trim().Length > 0
                || fila.C.Trim().Length > 0
                || fila.D.Trim().Length > 0)
            {
                continue;
            }

            suma += fila.Valor("Total") ?? 0m;
        }

        return Math.Abs(suma - subtotal.Value) <= ToleranciaCuadroR1;
    }

    /// <summary>
    /// Firmas (fila + canal efectivo) declaradas por TODOS los cuadros del bloque de la plantilla.
    /// Es la evidencia de "destino inexistente en todo el bloque" de §2-D3: si la firma de la fuente
    /// no aparece aquí, la plantilla no declara ese destino en ninguna sección (cero remap).
    /// </summary>
    private static List<(FilaEspejoR1 Fila, string Canal)> FirmasDelBloqueR1(
        IReadOnlyList<FilaEspejoR1> plantilla,
        JerarquiaR1 jerarquia)
    {
        var firmas = new List<(FilaEspejoR1 Fila, string Canal)>();
        foreach (var seccion in jerarquia.Secciones)
        {
            foreach (var cuadro in seccion.Subsecciones)
            {
                for (var i = cuadro.Inicio; i < cuadro.Fin; i++)
                {
                    firmas.Add((plantilla[i], CanalEfectivoEnCuadroR1(plantilla, i, cuadro.Fin)));
                }
            }
        }

        return firmas;
    }

    private static void MarcarRangoR1(bool[] cubiertas, int desde, int hasta)
    {
        for (var i = desde; i <= hasta && i < cubiertas.Length; i++)
        {
            cubiertas[i] = true;
        }
    }

    private static FalloJerarquiaR1 FalloDeSeccionR1(IReadOnlyList<FilaEspejoR1> fuente, SeccionR1 seccion) =>
        FalloDeFilaR1(
            fuente,
            Math.Min(seccion.Fin, fuente.Count - 1),
            seccion.Etiqueta,
            subseccion: null,
            ocurrencia: 0,
            "la sección de la fuente no está declarada en el bloque de la plantilla");

    private static FalloJerarquiaR1 FalloDeCuadroR1(IReadOnlyList<FilaEspejoR1> fuente, SubseccionR1 sub, CuadroR1 cuadro) =>
        FalloDeFilaR1(
            fuente,
            Math.Min(sub.Fin, fuente.Count - 1),
            cuadro.Seccion,
            cuadro.Subseccion,
            cuadro.Ocurrencia,
            "el cuadro de la fuente no está declarado en el bloque de la plantilla");

    private static FalloJerarquiaR1 FalloTotalR1(IReadOnlyList<FilaEspejoR1> fuente, int indice) =>
        FalloDeFilaR1(
            fuente,
            indice,
            seccion: null,
            subseccion: null,
            ocurrencia: 0,
            "el bloque de la plantilla no declara la fila de gran total");

    private static FalloJerarquiaR1 FalloDeFilaR1(
        IReadOnlyList<FilaEspejoR1> fuente,
        int indice,
        string? seccion,
        string? subseccion,
        int ocurrencia,
        string motivo,
        string? canal = null)
    {
        var fila = fuente[indice];
        var canalEfectivo = canal ?? fila.D.Trim();
        return new FalloJerarquiaR1(
            indice,
            seccion,
            subseccion,
            ocurrencia,
            fila.C.Trim(),
            canalEfectivo,
            fila.E.Trim(),
            motivo);
    }

    // [P56-JERARQUIA-AB-FIN]
}
