using System.Globalization;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Rules;

/// <summary>
/// Plan 60 (§3.2/§4.5): COORDENADA de auditoría de un importe reclasificado. Responde, para
/// cualquier valor del entregable, «este importe salió de esta fuente, de esta sección, de esta
/// subsección, de este canal y de este concepto».
///
/// <para>Las dimensiones se derivan de la PROPIA secuencia del detalle R2: la sección es la
/// <c>A</c> del cierre de sección que encierra la fila, la subsección es la <c>B</c> del primer
/// cierre de empresa por debajo y el canal es la <c>C</c> del cierre del grupo. En R2 la columna
/// <c>C</c> porta el CANAL (no la fuente); la fuente documental del ASE viaja en
/// <see cref="Fuente"/> — ver la nota de vocabulario del plan (§4.5).</para>
/// </summary>
/// <param name="Hoja">Hoja de la plantilla donde se aplica la regla.</param>
/// <param name="Fuente">Ruta del insumo R2 del ASE (<see cref="DetalleR2AseInputs.Fuente"/>).</param>
/// <param name="AseId">ASE del bloque al que pertenece el importe.</param>
/// <param name="Seccion">Sección (<c>A</c> del cierre de sección que encierra la fila).</param>
/// <param name="Subseccion">Empresa/subsección (<c>B</c> del primer cierre de empresa por debajo).</param>
/// <param name="Canal">Canal (<c>C</c> del cierre del grupo; en un agregado, los canales de su tramo).</param>
/// <param name="Concepto">Concepto (<c>D</c> de la fila, o la etiqueta estructural del agregado).</param>
public sealed record CoordenadaConcepto(
    string Hoja,
    string Fuente,
    int AseId,
    string Seccion,
    string Subseccion,
    string Canal,
    string Concepto);

/// <summary>
/// Plan 60 (§3.2): un movimiento de reclasificación con su coordenada completa y el motivo. La
/// lista de movimientos es la TRAZA de auditoría que el writer escribe al log con el RunId.
/// </summary>
/// <param name="IndiceFila">Posición de la fila dentro de <see cref="DetalleR2AseInputs.Filas"/>.</param>
/// <param name="Firma">Firma <c>A|B|C|D</c> de la fila (dato de auditoría, no coordenada de celda).</param>
/// <param name="Coordenada">Coordenada completa del importe movido.</param>
/// <param name="MovidoQ">Importe sacado de <c>Deb/Cred</c> (negativo = el importe era negativo).</param>
/// <param name="Motivo">Motivo estable: <see cref="ReclasificadorDebCredR2.MotivoCierreCanal"/> o
/// <see cref="ReclasificadorDebCredR2.MotivoAgregado"/>, con el detalle que lo justifica.</param>
public sealed record MovimientoReclasificacion(
    int IndiceFila,
    string Firma,
    CoordenadaConcepto Coordenada,
    decimal MovidoQ,
    string Motivo);

/// <summary>
/// Plan 60 (§3.2): resultado puro de la regla. <see cref="Clasificado"/> es una instancia NUEVA
/// (la de entrada no se muta) y <see cref="Movimientos"/> es la traza de auditoría.
/// </summary>
/// <param name="Clasificado">Modelo con los valores ya reclasificados.</param>
/// <param name="Movimientos">Traza de auditoría, ordenada por posición física de la fila.</param>
public sealed record ResultadoReclasificacion(
    DetalleR2AseInputs Clasificado,
    IReadOnlyList<MovimientoReclasificacion> Movimientos);

/// <summary>
/// Plan 60 (H1, D1/D2/D6 — §3.2/§3.3/§4): RECLASIFICACIÓN <c>Deb/Cred</c> → <c>Deb/Cred-otros</c>
/// por CANAL de recaudo en la pestaña <c>Rem. Anticipos R2</c>.
///
/// <para><b>Regla de negocio (texto de la administradora, §0.1):</b> cuando el cierre de un canal de
/// recaudo trae en <c>Deb/Cred</c> un importe significativo, el BLOQUE COMPLETO de ese canal
/// (sus conceptos y su cierre) se traslada a <c>Deb/Cred-otros</c> y los agregados estructurales
/// que son literales se re-propagan de forma consistente. Lo que es fórmula no se toca: se
/// recalcula solo.</para>
///
/// <para><b>Función PURA:</b> sin I/O, sin período, sin ASE concreto, sin plantilla. Opera sobre la
/// secuencia tipada <see cref="DetalleR2Fila"/> (A/B/C/D + valores por encabezado) — no conoce
/// filas, celdas ni columnas de la hoja (ver §3.4: la regla NO depende de Infrastructure porque
/// resuelve QUÉ VALOR porta cada fila, no DÓNDE va). Devuelve una instancia NUEVA: mutar la entrada
/// queda prohibido (el writer puede re-ejecutarse y los gates comparan el modelo del reader).</para>
///
/// <para><b>Algoritmo (§3.3, identidad = INSTANCIA y nunca etiqueta):</b></para>
/// <list type="number">
/// <item>Detectar los CIERRES DE CANAL de la secuencia (<c>D = Total</c> con <c>C</c> no vacío).</item>
/// <item>Grupo = tramo contiguo que termina en ese cierre y empieza tras el cierre estructural
/// anterior (de canal, empresa, sección o gran total). La etiqueta <c>O</c> puede repetirse en dos
/// cuadros: agrupar por etiqueta colapsaría ambos y movería el equivocado.</item>
/// <item>Si <c>|Q_cierre| &gt; umbral</c> se valida la PERTENENCIA completa (§4.5) y, solo entonces,
/// cada fila del grupo pasa a <c>Q' = 0</c>, <c>R' = R + Q</c>.</item>
/// <item>Para cada <b>cierre estructural</b> (empresa/sección/gran total) con <c>Q ≠ 0</c>: se busca
/// el tramo contiguo MÁS LARGO de cierres de canal anteriores cuya suma sea su <c>Q</c> (tolerancia
/// 0,01) y se reconcilia (<c>Q' = Q − Σ</c>, <c>R' = R + Σ</c>). Si no hay tramo se falla con
/// <c>ERR-RECLASIFICACION</c>. Las filas de DETALLE (<c>Vlr Servicio</c>, <c>Vlr Intereses</c>,
/// <c>Subsidio(-)/Contribucion(+)</c>) **nunca** se reconcilian: su <c>Q/R</c> es el importe de su
/// propio concepto, aunque coincida por casualidad con la suma de un tramo. <b>Este paso solo corre
/// si algún grupo disparó:</b> sin disparo la regla es no-op (no hay nada que reconciliar y el
/// período sale intacto).</item>
/// </list>
///
/// <para><b>Invariantes:</b> conservación del par por fila (<c>Q' + R' = Q + R</c>, §4.3 — el
/// CONSOLIDADO y <c>VALIDACION_TOTAL</c> siguen cerrando) e idempotencia (§4.4:
/// <c>Aplicar(Aplicar(x)) = Aplicar(x)</c>).</para>
/// </summary>
public static class ReclasificadorDebCredR2
{
    /// <summary>
    /// Plan 60 (D1 — §2): umbral de significancia. Procedencia: regla de la administradora
    /// ratificada por el Ingeniero el 2026-10-27 («algún concepto <c>Total</c> de algún canal de
    /// recaudo con valor mayor a 500.000»). Disparo con VALOR ABSOLUTO (<c>|Q| &gt; umbral</c>,
    /// opción recomendada de D1): un canal con −6M es igual de significativo. Un único punto; el
    /// literal NO se duplica.
    /// </summary>
    public const decimal UmbralReclasificacionDebCred = 500_000m;

    /// <summary>Hoja de la plantilla a la que la regla está acotada (D3 — §1.3).</summary>
    public const string HojaRemAnticiposR2 = "Rem. Anticipos R2";

    /// <summary>Encabezado de componente que porta <c>Deb/Cred</c> (dato de la fuente, no coordenada).</summary>
    public const string ComponenteDebCred = "Deb/Cred";

    /// <summary>Encabezado de componente que porta <c>Deb/Cred-otros</c> (destino de la reclasificación).</summary>
    public const string ComponenteDebCredOtros = "Deb/Cred-otros";

    /// <summary>Motivo estable de un movimiento de concepto (prefijo de <see cref="MovimientoReclasificacion.Motivo"/>).</summary>
    public const string MotivoCierreCanal = "cierre-canal-umbral";

    /// <summary>Motivo estable de un agregado reconciliado (prefijo de <see cref="MovimientoReclasificacion.Motivo"/>).</summary>
    public const string MotivoAgregado = "agregado-tramo";

    /// <summary>Marca literal de cierre de la gramática de la hoja (única palabra cerrada).</summary>
    private const string MarcaTotal = "Total";

    /// <summary>Reconciliación de tramo al centavo (§4.6): la comparación con el oráculo usa 0,5 aparte.</summary>
    private const decimal ToleranciaTramo = 0.01m;

    /// <summary>Rol estructural de una fila del detalle R2 (gramática de cierres de la hoja).</summary>
    private enum RolFila
    {
        /// <summary>Fila de detalle (concepto: no cierra nada).</summary>
        Dato,

        /// <summary>Cierre de canal: <c>D = Total</c> con <c>C</c> no vacío.</summary>
        CierreCanal,

        /// <summary>Cierre de empresa (subsección): <c>B</c> no vacía con <c>C = Total</c>.</summary>
        CierreEmpresa,

        /// <summary>Cierre de sección: <c>A</c> no vacía con <c>B = Total</c>.</summary>
        CierreSeccion,

        /// <summary>Gran total del bloque: <c>A = Total</c> con <c>B</c> vacía.</summary>
        GranTotal,
    }

    /// <summary>
    /// Aplica la regla al detalle R2 del ASE y devuelve el modelo clasificado + la traza.
    /// </summary>
    /// <param name="detalle">Observación fiel de la fuente (no se muta).</param>
    /// <param name="umbral">Umbral de significancia (el call site productivo usa el default).</param>
    /// <param name="hoja">Hoja destino, solo para la coordenada de auditoría.</param>
    /// <exception cref="CalculoInvalidoException">
    /// <c>ERR-RECLASIFICACION</c> cuando un cierre que dispara no tiene coordenada completa o cuando
    /// un cierre estructural no reconcilia con ningún tramo (nunca se mueve «lo que se pueda»).
    /// </exception>
    public static ResultadoReclasificacion Aplicar(
        DetalleR2AseInputs detalle,
        decimal umbral = UmbralReclasificacionDebCred,
        string hoja = HojaRemAnticiposR2)
    {
        ArgumentNullException.ThrowIfNull(detalle);

        var filas = detalle.Filas;
        var total = filas.Count;
        var roles = new RolFila[total];
        var debCred = new decimal[total];
        var debCredOtros = new decimal[total];
        for (var i = 0; i < total; i++)
        {
            roles[i] = Clasificar(filas[i]);
            debCred[i] = Valor(filas[i], ComponenteDebCred);
            debCredOtros[i] = Valor(filas[i], ComponenteDebCredOtros);
        }

        var nuevoQ = (decimal[])debCred.Clone();
        var nuevoR = (decimal[])debCredOtros.Clone();
        var movidos = new bool[total];
        var movimientos = new List<MovimientoReclasificacion>();

        // (1)-(3) Grupos de canal que disparan: se valida la pertenencia ANTES de mover.
        for (var cierre = 0; cierre < total; cierre++)
        {
            if (roles[cierre] != RolFila.CierreCanal)
            {
                continue;
            }

            if (Math.Abs(debCred[cierre]) <= umbral)
            {
                continue;
            }

            var inicio = InicioDelGrupo(roles, cierre);
            var canal = filas[cierre].C.Trim();
            var seccion = ValidarPertenencia(detalle, filas, roles, inicio, cierre, hoja);
            var subseccion = EmpresaQueEncierra(filas, roles, cierre) ?? string.Empty;

            for (var i = inicio; i <= cierre; i++)
            {
                movidos[i] = true;
                var antes = debCred[i];
                nuevoQ[i] = 0m;
                nuevoR[i] = debCredOtros[i] + antes;
                if (antes == 0m)
                {
                    continue;
                }

                movimientos.Add(new MovimientoReclasificacion(
                    i,
                    filas[i].Firma,
                    new CoordenadaConcepto(
                        hoja,
                        detalle.Fuente,
                        detalle.Ase.Id,
                        seccion,
                        subseccion,
                        canal,
                        ConceptoDe(filas[i])),
                    antes,
                    $"{MotivoCierreCanal}: canal '{canal}' |Q| = {Importe(debCred[cierre])} > umbral {Importe(umbral)}"));
            }
        }

        // (4) Propagación a los AGREGADOS estructurales que son literales.
        //
        // SOLO tiene sentido si algún grupo disparó: la reconciliación existe para re-propagar lo que
        // se movió. Sin disparo no hay nada que reconciliar y el período debe salir INTACTO (D5), de
        // modo que la regla no convierte una validación post-movimiento en un falso fallo de un
        // período sano: un cierre estructural que no sea la suma exacta de un tramo solo es un
        // problema cuando hay un movimiento que justifique mirarlo.
        var cierresCanal = Enumerable.Range(0, total).Where(i => roles[i] == RolFila.CierreCanal).ToList();
        if (!movidos.Any(movido => movido))
        {
            return NoOp(detalle, nuevoQ, nuevoR, movimientos);
        }

        for (var i = 0; i < total; i++)
        {
            // Solo los CIERRES estructurales (empresa, sección, gran total) se re-propagan. Una fila
            // de DETALLE (Vlr Servicio, Vlr Intereses, Subsidio(-)/Contribucion(+)) NUNCA se
            // reconcilia: su Q/R es el importe de SU concepto, no un agregado, y aunque su importe
            // coincida por casualidad con la suma de un tramo de cierres anteriores, moverlo sería
            // corromper el desglose (y rompería la aritmética del grupo al que pertenece).
            if (roles[i] is RolFila.Dato or RolFila.CierreCanal || movidos[i] || debCred[i] == 0m)
            {
                continue;
            }

            var tramo = BuscarTramo(detalle, filas, roles, cierresCanal, i, debCred, debCred[i], hoja, umbral);
            if (tramo is null)
            {
                throw Error(
                    detalle,
                    filas,
                    i,
                    hoja,
                    SeccionQueEncierra(filas, roles, i) ?? string.Empty,
                    EmpresaQueEncierra(filas, roles, i) ?? string.Empty,
                    string.Empty,
                    debCred[i],
                    $"es un cierre estructural ({roles[i]}) con Q = {Importe(debCred[i])} y su importe no es la suma de NINGÚN tramo contiguo de cierres de canal anteriores (tolerancia {ToleranciaTramo}). " +
                    "Nunca se remapea ni se deja el agregado inconsistente: revise la fuente o la plantilla.");
            }

            var suma = tramo.Value.Cierres.Where(c => movidos[c]).Sum(c => debCred[c]);
            if (suma == 0m)
            {
                continue; // el tramo no contiene ningún canal movido: la reconciliación es vacía.
            }

            var canales = CanalesDelTramo(filas, tramo.Value.Cierres);
            nuevoQ[i] = debCred[i] - suma;
            nuevoR[i] = debCredOtros[i] + suma;
            movimientos.Add(new MovimientoReclasificacion(
                i,
                filas[i].Firma,
                new CoordenadaConcepto(
                    hoja,
                    detalle.Fuente,
                    detalle.Ase.Id,
                    SeccionQueEncierra(filas, roles, i) ?? string.Empty,
                    roles[i] == RolFila.CierreEmpresa ? filas[i].B.Trim() : string.Empty,
                    string.Join(",", canales),
                    ConceptoDe(filas[i])),
                suma,
                $"{MotivoAgregado}: tramo [{string.Join(",", canales)}] suma {Importe(tramo.Value.Suma)}"));
        }

        return new ResultadoReclasificacion(
            Clonar(detalle, nuevoQ, nuevoR),
            movimientos.OrderBy(m => m.IndiceFila).ToList());
    }

    /// <summary>
    /// Plan 60 (H1, D4/D5 — corrección de borde): camino NO-OP. Ningún cierre de canal superó el
    /// umbral ⇒ no hay nada que mover ni que reconciliar: se devuelve el modelo CLONADO sin cambios
    /// (la entrada sigue sin mutarse) y la traza vacía. La regla no valida tramos de agregados porque
    /// no existe movimiento que los justifique: un período sin disparo sale intacto, nunca en
    /// <c>ERR-RECLASIFICACION</c>.
    /// </summary>
    private static ResultadoReclasificacion NoOp(
        DetalleR2AseInputs detalle,
        IReadOnlyList<decimal> debCred,
        IReadOnlyList<decimal> debCredOtros,
        IReadOnlyList<MovimientoReclasificacion> movimientos) =>
        new(Clonar(detalle, debCred, debCredOtros), movimientos);

    // ── Pertenencia (D6 / §4.5) ─────────────────────────────────────────────────────────────────

    /// <summary>
    /// Plan 60 (D6/§4.5): valida la coordenada completa del grupo ANTES de mover un solo valor y
    /// devuelve la sección derivada. Reglas duras, todas fail-fast:
    /// (1) el grupo es contiguo y termina en su cierre de canal; (2) ninguna fila del tramo trae un
    /// canal distinto ni una etiqueta de sección/empresa que la expulse; (3) hay sección y
    /// subsección (empresa) derivables; (4) cada fila tiene concepto; (5) la fuente está declarada y
    /// el ASE del bloque identificado.
    /// </summary>
    private static string ValidarPertenencia(
        DetalleR2AseInputs detalle,
        IReadOnlyList<DetalleR2Fila> filas,
        IReadOnlyList<RolFila> roles,
        int inicio,
        int cierre,
        string hoja)
    {
        var canal = filas[cierre].C.Trim();

        // (5) Fuente declarada + ASE del bloque identificado.
        if (string.IsNullOrWhiteSpace(detalle.Fuente))
        {
            throw Error(
                detalle,
                filas,
                cierre,
                hoja,
                string.Empty,
                string.Empty,
                canal,
                Valor(filas[cierre], ComponenteDebCred),
                "la fuente R2 del ASE no está declarada en el modelo (DetalleR2AseInputs.Fuente): sin fuente no hay coordenada de auditoría y no se mueve ningún valor.");
        }

        if (detalle.Ase.Id <= 0)
        {
            throw Error(
                detalle,
                filas,
                cierre,
                hoja,
                string.Empty,
                string.Empty,
                canal,
                Valor(filas[cierre], ComponenteDebCred),
                "el ASE del bloque no está identificado (Ase.Id): sin ASE no hay coordenada de auditoría y no se mueve ningún valor.");
        }

        for (var i = inicio; i <= cierre; i++)
        {
            var fila = filas[i];

            // (2) Pertenencia al canal: la etiqueta de sección/empresa expulsa la fila del grupo.
            if (fila.A.Trim().Length > 0 || fila.B.Trim().Length > 0)
            {
                throw Error(
                    detalle,
                    filas,
                    i,
                    hoja,
                    fila.A.Trim(),
                    fila.B.Trim(),
                    canal,
                    ConceptoDe(fila),
                    Valor(fila, ComponenteDebCred),
                    $"la fila trae etiqueta de sección/empresa (A='{fila.A.Trim()}' B='{fila.B.Trim()}') y no pertenece al grupo del canal '{canal}': el tramo no es contiguo. No se mueve nada del grupo.");
            }

            var etiquetaCanal = fila.C.Trim();
            if (etiquetaCanal.Length > 0 && !string.Equals(etiquetaCanal, canal, StringComparison.OrdinalIgnoreCase))
            {
                throw Error(
                    detalle,
                    filas,
                    i,
                    hoja,
                    fila.A.Trim(),
                    fila.B.Trim(),
                    canal,
                    ConceptoDe(fila),
                    Valor(fila, ComponenteDebCred),
                    $"la fila trae la etiqueta de canal '{etiquetaCanal}' y el grupo cierra en el canal '{canal}': el tramo mezcla canales. No se mueve nada del grupo.");
            }

            // (4) Concepto obligatorio.
            if (fila.D.Trim().Length == 0)
            {
                throw Error(
                    detalle,
                    filas,
                    i,
                    hoja,
                    fila.A.Trim(),
                    fila.B.Trim(),
                    canal,
                    string.Empty,
                    Valor(fila, ComponenteDebCred),
                    $"la fila del grupo del canal '{canal}' no declara concepto (columna D vacía): sin concepto no hay coordenada completa. No se mueve nada del grupo.");
            }
        }

        // (3) Sección y subsección (empresa) derivables de los cierres de la propia secuencia.
        var seccion = SeccionQueEncierra(filas, roles, cierre);
        if (seccion is null)
        {
            throw Error(
                detalle,
                filas,
                cierre,
                hoja,
                string.Empty,
                string.Empty,
                canal,
                ConceptoDe(filas[cierre]),
                Valor(filas[cierre], ComponenteDebCred),
                $"el grupo del canal '{canal}' no tiene sección derivable: ningún cierre de sección (A con B = '{MarcaTotal}') lo encierra. No se mueve nada del grupo.");
        }

        if (EmpresaQueEncierra(filas, roles, cierre) is null)
        {
            throw Error(
                detalle,
                filas,
                cierre,
                hoja,
                seccion,
                string.Empty,
                canal,
                ConceptoDe(filas[cierre]),
                Valor(filas[cierre], ComponenteDebCred),
                $"el grupo del canal '{canal}' no tiene subsección (empresa) derivable: ningún cierre de empresa (B con C = '{MarcaTotal}') lo enciende por debajo. No se mueve nada del grupo.");
        }

        return seccion;
    }

    // ── Tramo de cierres (§4.1/§4.2/§4.6) ───────────────────────────────────────────────────────

    /// <summary>
    /// Tramo válido de una fila agregada: la ventana CONTIGUA de cierres de canal anteriores cuya
    /// suma es su <c>Q</c> (tolerancia de un centavo). Se busca de mayor a menor longitud: el más
    /// largo gana. Dos ventanas empatadas en la longitud ganadora son una ambigüedad de la fuente y
    /// se fallan (nunca se elige una por orden).
    /// </summary>
    private static (IReadOnlyList<int> Cierres, decimal Suma)? BuscarTramo(
        DetalleR2AseInputs detalle,
        IReadOnlyList<DetalleR2Fila> filas,
        IReadOnlyList<RolFila> roles,
        IReadOnlyList<int> cierresCanal,
        int indice,
        IReadOnlyList<decimal> debCred,
        decimal objetivo,
        string hoja,
        decimal umbral)
    {
        var anteriores = cierresCanal.Where(c => c < indice).ToList();
        for (var largo = anteriores.Count; largo >= 1; largo--)
        {
            List<int>? encontrado = null;
            var sumaEncontrada = 0m;
            var coincidencias = 0;
            for (var inicio = 0; inicio + largo <= anteriores.Count; inicio++)
            {
                var ventana = anteriores.GetRange(inicio, largo);
                var suma = ventana.Sum(c => debCred[c]);
                if (Math.Abs(suma - objetivo) > ToleranciaTramo)
                {
                    continue;
                }

                coincidencias++;
                encontrado = ventana;
                sumaEncontrada = suma;
            }

            if (coincidencias > 1)
            {
                throw Error(
                    detalle,
                    filas,
                    indice,
                    hoja,
                    SeccionQueEncierra(filas, roles, indice) ?? string.Empty,
                    EmpresaQueEncierra(filas, roles, indice) ?? string.Empty,
                    string.Join(",", encontrado!.Select(c => filas[c].C.Trim())),
                    objetivo,
                    $"hay {coincidencias} tramos contiguos de {largo} cierres de canal cuya suma es su Q = {Importe(objetivo)}: la fuente es ambigua y no se elige uno por orden. Umbral vigente {Importe(umbral)}.");
            }

            if (encontrado is not null)
            {
                return (encontrado, sumaEncontrada);
            }
        }

        return null;
    }

    // ── Derivaciones de la secuencia ────────────────────────────────────────────────────────────

    /// <summary>Rol estructural que declara la fila, según la gramática de cierres de la hoja.</summary>
    private static RolFila Clasificar(DetalleR2Fila fila)
    {
        var a = fila.A.Trim();
        var b = fila.B.Trim();
        var c = fila.C.Trim();
        var d = fila.D.Trim();

        if (c.Length > 0 && EsTotal(d))
        {
            return RolFila.CierreCanal;
        }

        if (EsTotal(a) && b.Length == 0)
        {
            return RolFila.GranTotal;
        }

        if (a.Length == 0 && b.Length > 0 && EsTotal(c))
        {
            return RolFila.CierreEmpresa;
        }

        if (a.Length > 0 && EsTotal(b) && c.Length == 0)
        {
            return RolFila.CierreSeccion;
        }

        return RolFila.Dato;
    }

    /// <summary>
    /// Inicio del grupo que termina en el cierre indicado: la fila siguiente al último cierre
    /// ESTRUCTURAL anterior (de canal, empresa, sección o gran total). El cierre de empresa que
    /// precede a un grupo no forma parte de él: cierra el bloque anterior, no este.
    /// </summary>
    private static int InicioDelGrupo(IReadOnlyList<RolFila> roles, int cierre)
    {
        for (var j = cierre - 1; j >= 0; j--)
        {
            if (roles[j] != RolFila.Dato)
            {
                return j + 1;
            }
        }

        return 0;
    }

    /// <summary>Sección que encierra la fila: la <c>A</c> del primer cierre de sección en o por debajo de ella.</summary>
    private static string? SeccionQueEncierra(IReadOnlyList<DetalleR2Fila> filas, IReadOnlyList<RolFila> roles, int indice)
    {
        for (var j = indice; j < filas.Count; j++)
        {
            if (roles[j] == RolFila.CierreSeccion)
            {
                return filas[j].A.Trim();
            }
        }

        return null;
    }

    /// <summary>Subsección (empresa) que encierra la fila: la <c>B</c> del primer cierre de empresa por debajo.</summary>
    private static string? EmpresaQueEncierra(IReadOnlyList<DetalleR2Fila> filas, IReadOnlyList<RolFila> roles, int indice)
    {
        for (var j = indice + 1; j < filas.Count; j++)
        {
            if (roles[j] == RolFila.CierreEmpresa)
            {
                return filas[j].B.Trim();
            }
        }

        return null;
    }

    /// <summary>
    /// Canales del tramo, sin repetir y en orden de aparición: el gran total reconcilia los DOS
    /// cuadros de la hoja (los mismos canales dos veces), y la coordenada debe nombrar los canales
    /// reales del tramo, no la lista física con duplicados.
    /// </summary>
    private static IReadOnlyList<string> CanalesDelTramo(IReadOnlyList<DetalleR2Fila> filas, IReadOnlyList<int> cierres)
    {
        var vistos = new List<string>(cierres.Count);
        foreach (var cierre in cierres)
        {
            var canal = filas[cierre].C.Trim();
            if (!vistos.Contains(canal, StringComparer.OrdinalIgnoreCase))
            {
                vistos.Add(canal);
            }
        }

        return vistos;
    }

    /// <summary>Concepto de la fila: la columna <c>D</c> o, en un agregado, su etiqueta estructural.</summary>
    private static string ConceptoDe(DetalleR2Fila fila)
    {
        foreach (var etiqueta in new[] { fila.D, fila.C, fila.A, fila.B })
        {
            var texto = etiqueta.Trim();
            if (texto.Length > 0)
            {
                return texto;
            }
        }

        return string.Empty;
    }

    private static bool EsTotal(string texto) =>
        string.Equals(texto.Trim(), MarcaTotal, StringComparison.OrdinalIgnoreCase);

    private static decimal Valor(DetalleR2Fila fila, string componente) =>
        fila.Valor(componente) ?? 0m;

    /// <summary>
    /// Modelo NUEVO con los valores reclasificados: se copian todos los encabezados de la fuente y
    /// se fijan los dos puertos de la regla. La instancia y las filas de ENTRADA no se tocan.
    /// </summary>
    private static DetalleR2AseInputs Clonar(DetalleR2AseInputs detalle, IReadOnlyList<decimal> debCred, IReadOnlyList<decimal> debCredOtros)
    {
        var filas = new List<DetalleR2Fila>(detalle.Filas.Count);
        for (var i = 0; i < detalle.Filas.Count; i++)
        {
            var origen = detalle.Filas[i];
            var valores = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);
            foreach (var (componente, valor) in origen.ValoresPorComponente)
            {
                valores[componente] = valor;
            }

            valores[ComponenteDebCred] = debCred[i];
            valores[ComponenteDebCredOtros] = debCredOtros[i];

            filas.Add(new DetalleR2Fila
            {
                A = origen.A,
                B = origen.B,
                C = origen.C,
                D = origen.D,
                ValoresPorComponente = valores,
            });
        }

        return new DetalleR2AseInputs
        {
            Ase = detalle.Ase,
            Componentes = detalle.Componentes,
            TieneColumnaEspeciales = detalle.TieneColumnaEspeciales,
            Fuente = detalle.Fuente,
            Filas = filas,
        };
    }

    // ── Mensajes (G-TRAZA: concepto + ASE + hoja + fuente + fila + importe) ─────────────────────

    private static CalculoInvalidoException Error(
        DetalleR2AseInputs detalle,
        IReadOnlyList<DetalleR2Fila> filas,
        int indice,
        string hoja,
        string seccion,
        string subseccion,
        string canal,
        decimal importe,
        string motivo) =>
        Error(detalle, filas, indice, hoja, seccion, subseccion, canal, ConceptoDe(filas[indice]), importe, motivo);

    private static CalculoInvalidoException Error(
        DetalleR2AseInputs detalle,
        IReadOnlyList<DetalleR2Fila> filas,
        int indice,
        string hoja,
        string seccion,
        string subseccion,
        string canal,
        string concepto,
        decimal importe,
        string motivo)
    {
        var fila = filas[indice];
        return new CalculoInvalidoException(
            CodigoError.Reclasificacion,
            $"ERR-RECLASIFICACION: hoja '{hoja}' · ASE {detalle.Ase.Id} · fuente '{detalle.Fuente}' · " +
            $"coordenada (sección '{seccion}' · subsección '{subseccion}' · canal '{canal}' · concepto '{concepto}') · " +
            $"fila #{indice + 1} de la secuencia del detalle (firma '{fila.Firma}') · importe {Importe(importe)}: {motivo}");
    }

    /// <summary>Importe para los mensajes: sin separador de miles y con el separador decimal de la cultura.</summary>
    private static string Importe(decimal valor) => valor.ToString("0.##", CultureInfo.CurrentCulture);
}
