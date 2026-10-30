using System.Globalization;
using System.Text;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Serilog;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 64 (H1.3, §2.4 — D-04/D-06/D-10): lectura del insumo <c>RecaudosReversados_*.xlsx</c> para la
/// hoja <c>ANT EXT-REV</c>.
///
/// <para><b>Qué se lee y qué no.</b> El <b>bloque 2</b> (<c>Pagos Pendientes por Aplicar</c>) es la
/// FUENTE: una fila de conceptos por fuente y una fila de total rotulada con la etiqueta de esa fuente
/// que CIERRA el grupo. El <b>bloque 3</b> (<c>Resumen Recaudo Aplicado Por Servicio</c>) es SOLO
/// CONTROL: no discrimina por fuente (<c>RECAUDO OCCIDENTE</c> / <c>RECAUDO OTROS</c> / <c>Total</c>, y 8
/// de los 30 archivos ni traen la columna OCCIDENTE), así que usarlo como fuente obligaría a inventar el
/// reparto (veredicto D-10, T0 §6).</para>
///
/// <para><b>Cómo se ubica cada cosa (cero coordenadas de la plantilla).</b>
/// <list type="bullet">
/// <item>Los dos marcadores se localizan por TEXTO normalizado (§2.4 paso 2); faltantes o invertidos son
/// fail-fast.</item>
/// <item>El encabezado del bloque 2 se reconoce por el par de rótulos que declara (<c>ASEO</c> seguido de
/// <c>Total</c>), no por número de fila.</item>
/// <item>La columna de CONCEPTO se DERIVA: es la única columna cuyas celdas traen rótulos del catálogo
/// <see cref="CatalogoConceptosAntExtRev"/>.</item>
/// <item>La columna del rótulo <c>Total</c> que cierra cada grupo se DERIVA igual (es la única que
/// declara <c>Total</c> con una etiqueta a su izquierda).</item>
/// <item><b>TERN A DE VALORES (enmienda M4, BLOQUEANTE).</b> El insumo NO tiene encabezados semánticos
/// para las tres columnas del monto (el encabezado dice <c>Cantidad</c>/<c>Valor</c> y pertenece a la
/// tabla de pendientes). La única propiedad estable medida en 30/30 es que <b>las tres columnas son
/// copias del mismo monto</b>. Por eso la terna se identifica validando esa INVARIANTE sobre todas las
/// filas de concepto, y el valor del concepto se lee de la TERCERA columna de la terna (la de cierre).
/// Prohibido fijarla por letra o por índice: si la terna se desplaza, se sigue leyendo el mismo monto.</item>
/// </list></para>
///
/// <para><b>Fail-fast (§2.4, D-06) con <c>ERR-FORMATO-FUENTE</c>:</b> marcador ausente/invertido o
/// repetido; encabezado ausente; cero conceptos reconocidos; fila no reconocida; celda de valor no
/// numérica; las tres columnas del monto divergen &gt; 0,5; Σ de los cuatro conceptos ≠ total de la
/// fuente &gt; 0,5; etiqueta de fuente duplicada; cero fuentes del catálogo; etiqueta no tabulada que
/// COEXISTE con una ausencia del catálogo (gemelo A3, Plan 50/51). Cada mensaje nombra ASE + archivo +
/// marcador/etiqueta/concepto + observado vs esperado.</para>
/// </summary>
public sealed partial class ExcelDataReaderWorkbookLeafInputReader
{
    /// <summary>
    /// Hoja del insumo. Se abre por NOMBRE (Plan 27 D-A): un insumo con la hoja renombrada es un
    /// fail-fast que nombra archivo + hoja, nunca una lectura a ciegas de la primera hoja.
    /// </summary>
    private const string HojaInsumoRecaudosReversados = "Sheet1";

    /// <summary>Marcador del bloque 2 (fuente) — texto medido en 30/30 (T0 §4).</summary>
    private const string MarcadorBloque2 = "Pagos Pendientes por Aplicar";

    /// <summary>Marcador del bloque 3 (control) — texto medido en 30/30 (T0 §4).</summary>
    private const string MarcadorBloque3 = "Resumen Recaudo Aplicado Por Servicio";

    /// <summary>Rótulo del total (de canal y de fuente) — DATO del formato del insumo (T0 §4.2).</summary>
    private const string RotuloTotal = "TOTAL";

    /// <summary>Rótulo de la cabecera real del bloque 2 — DATO del formato del insumo (T0 §4.1).</summary>
    private const string RotuloAseo = "ASEO";

    /// <summary>Nombre administrativo del insumo para las trazas y los mensajes.</summary>
    private const string NombreInsumo = "RecaudosReversados";

    /// <summary>Tolerancia de comparación del plan (S0): ±0,5 en toda comparación de dinero.</summary>
    private const decimal ToleranciaAntExtRev = 0.5m;

    /// <summary>Prefijo de traza del Plan 51 aplicado a este insumo (§2.4 paso 5).</summary>
    private const string PrefijoTraza = "ANTEXTREV:";

    /// <inheritdoc />
    public AntExtRevAseInputs LeerAntExtRev(Ase ase, Periodo periodo, string rutaRecaudosReversados)
    {
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(rutaRecaudosReversados);

        // El PERÍODO no parametriza la lectura: el insumo aplica en AMBAS quincenas y su gramática no cambia
        // (D-01/D-07). Recibe el dominio solo para no romper la simetría del contrato: cero ramas por quincena.
        var archivo = Path.GetFileName(rutaRecaudosReversados);
        var filas = ExcelWorksheetNavigator.LeerFilas(rutaRecaudosReversados, HojaInsumoRecaudosReversados);

        // §2.4 paso 2: marcadores por texto normalizado; faltantes, repetidos o invertidos = fail-fast.
        var filaMarcador2 = LocalizarMarcadorInsumo(filas, MarcadorBloque2, ase, archivo);
        var filaMarcador3 = LocalizarMarcadorInsumo(filas, MarcadorBloque3, ase, archivo);
        if (filaMarcador3 < filaMarcador2)
        {
            throw FormatoFuenteInsumo(
                ase,
                archivo,
                $"los marcadores están invertidos: '{MarcadorBloque3}' (fila {filaMarcador3 + 1}) aparece ANTES de " +
                $"'{MarcadorBloque2}' (fila {filaMarcador2 + 1})");
        }

        var filaEncabezado = LocalizarEncabezadoBloque2(filas, filaMarcador2, filaMarcador3, ase, archivo);
        var (columnaConcepto, columnaCanal, columnaFuente) =
            LocalizarColumnasBloque2(filas, filaEncabezado + 1, filaMarcador3, ase, archivo);
        var terna = LocalizarTernaDeValores(filas, filaEncabezado + 1, filaMarcador3, columnaConcepto, ase, archivo);

        // §2.4 pasos 3–4: recorrido del bloque 2 agrupando por fuente; el grupo cierra en su fila de total.
        var fuentes = new List<AntExtRevFuenteInputs>();
        var etiquetasNoTabuladas = new List<string>();
        var grupo = new List<FilaGrupoAntExtRev>();

        for (var i = filaEncabezado + 1; i < filaMarcador3; i++)
        {
            var a = TextoNormalizadoInsumo(filas, i, columnaFuente);
            var b = TextoNormalizadoInsumo(filas, i, columnaCanal);
            var c = TextoNormalizadoInsumo(filas, i, columnaConcepto);
            var tercerValor = NumeroDeInsumo(CeldaInsumo(filas, i, terna.Tercera));

            var vacia = a.Length == 0 && b.Length == 0 && c.Length == 0
                && NumeroDeInsumo(CeldaInsumo(filas, i, terna.Primera)) is null
                && NumeroDeInsumo(CeldaInsumo(filas, i, terna.Segunda)) is null
                && tercerValor is null;
            if (vacia)
            {
                // Filas en blanco y la banda del encabezado (donde el "66" es TEXTO y no un monto).
                continue;
            }

            var concepto = CatalogoConceptosAntExtRev.Buscar(c);
            if (concepto is not null)
            {
                if (a.Length > 0 || b.Length > 0)
                {
                    throw FormatoFuenteInsumo(
                        ase,
                        archivo,
                        $"fila {i + 1}: trae el concepto '{concepto.Etiqueta}' pero además rotula fuente/canal " +
                        $"(A='{a}', B='{b}'); la fila de concepto no puede cerrar un grupo");
                }

                grupo.Add(LeerFilaDeValores(
                    filas, i, terna, $"concepto '{concepto.Etiqueta}'", concepto, ase, archivo));
                continue;
            }

            if (b.Length > 0 && c == RotuloTotal)
            {
                // Total de CANAL: informativo (no es destino), pero su aritmética también debe cerrar (M4).
                grupo.Add(LeerFilaDeValores(filas, i, terna, $"total de canal '{b}'", null, ase, archivo));
                continue;
            }

            if (a.Length > 0 && b == RotuloTotal)
            {
                // Total de FUENTE: cierra el grupo y aporta la etiqueta con la que el insumo la rotula.
                var cierre = LeerFilaDeValores(filas, i, terna, $"total de fuente '{a}'", null, ase, archivo);
                CerrarGrupoAntExtRev(ase, archivo, a, grupo, cierre, fuentes, etiquetasNoTabuladas);
                grupo.Clear();
                continue;
            }

            if (a == RotuloTotal && b.Length == 0)
            {
                // Gran total del bloque 2: informativo (el cierre por fuente ya se validó grupo a grupo).
                continue;
            }

            throw FormatoFuenteInsumo(
                ase,
                archivo,
                $"fila {i + 1} no reconocida en el bloque 2 (A='{a}', B='{b}', C='{c}'): no es concepto del " +
                "catálogo, ni total de canal (B + 'Total'), ni total de fuente (A + 'Total'), ni gran total");
        }

        if (grupo.Count > 0)
        {
            throw FormatoFuenteInsumo(
                ase,
                archivo,
                $"el grupo que empieza en la fila {grupo[0].Fila + 1} no cierra con su fila de total de fuente " +
                $"antes del marcador '{MarcadorBloque3}' (fila {filaMarcador3 + 1}); la agrupación por fuente es " +
                "incompleta y no se puede atribuir el dinero");
        }

        // §2.4 paso 5: fuentes del catálogo ausentes ⇒ omisión + traza (nunca un cero indistinguible).
        var ausentes = new List<string>();
        foreach (var filaCatalogo in CatalogoFuentesAntExtRev.Filas)
        {
            if (fuentes.Any(fuente => EsFuenteDeLaFila(fuente, filaCatalogo)))
            {
                continue;
            }

            ausentes.Add(PrefijoTraza + filaCatalogo.EtiquetasFuente[0]);
        }

        // §2.4 paso 6: etiqueta no tabulada que COEXISTE con una ausencia = estructura rota (gemelo A3).
        if (etiquetasNoTabuladas.Count > 0 && ausentes.Count > 0)
        {
            throw FormatoFuenteInsumo(
                ase,
                archivo,
                $"el bloque 2 trae la(s) etiqueta(s) no tabulada(s) [{string.Join(", ", etiquetasNoTabuladas)}] " +
                $"mientras falta(n) del catálogo [{string.Join(", ", ausentes)}]: no se " +
                "puede decidir si la etiqueta desconocida sustituye a una fuente ausente o es dinero de otra " +
                "fuente (sin cita no se mapea, y nunca se asume 0)");
        }

        if (fuentes.Count == 0)
        {
            throw FormatoFuenteInsumo(
                ase,
                archivo,
                "no se reconoció NINGUNA fuente del catálogo en el bloque 2 (estructura rota: el insumo debe " +
                "traer al menos una fuente con su total)");
        }

        // §2.4 paso 7: bloque 3 leído SOLO como control; su divergencia se registra, no bloquea.
        var totalBloque2 = fuentes.Sum(fuente => fuente.TotalFuente);
        var (totalControl, hallazgo) = LeerControlBloque3(filas, filaMarcador3, columnaFuente, totalBloque2, ase);

        RegistrarObservabilidadAntExtRev(ase, fuentes, ausentes, etiquetasNoTabuladas);

        return new AntExtRevAseInputs
        {
            Ase = ase,
            Fuentes = fuentes,
            FuentesAusentes = ausentes,
            EtiquetasNoReconocidas = etiquetasNoTabuladas.Select(etiqueta => PrefijoTraza + etiqueta).ToArray(),
            HallazgosEtiquetasNoMapeadas = HallazgosDeEtiquetas(ase, etiquetasNoTabuladas),
            TotalControlBloque3 = totalControl,
            HallazgoControlBloque3 = hallazgo
        };
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Agrupación por fuente y cierre aritmético (§2.4 pasos 3–4)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Fila del bloque 2 ya leída con la terna de valores resuelta. La igualdad de la terna se valida al
    /// CERRAR el grupo, porque solo entonces se conoce la etiqueta de la fuente con la que el mensaje
    /// debe nombrar el defecto (D-06: ASE + fuente + concepto + D/E/F observados).
    /// </summary>
    private sealed record FilaGrupoAntExtRev(
        int Fila,
        string Identidad,
        CatalogoConceptosAntExtRev.Concepto? Destino,
        decimal Primera,
        decimal Segunda,
        decimal Tercera);

    /// <summary>
    /// Lee la terna de una fila de datos del bloque 2 y exige que sus tres celdas sean numéricas (una
    /// celda no numérica donde el insumo declara dinero es estructura inesperada: <c>ERR-FORMATO-FUENTE</c>).
    /// </summary>
    private static FilaGrupoAntExtRev LeerFilaDeValores(
        List<object?[]> filas,
        int indice,
        (int Primera, int Segunda, int Tercera) terna,
        string identidad,
        CatalogoConceptosAntExtRev.Concepto? destino,
        Ase ase,
        string archivo)
    {
        var primera = NumeroDeInsumo(CeldaInsumo(filas, indice, terna.Primera));
        var segunda = NumeroDeInsumo(CeldaInsumo(filas, indice, terna.Segunda));
        var tercera = NumeroDeInsumo(CeldaInsumo(filas, indice, terna.Tercera));
        if (primera is null || segunda is null || tercera is null)
        {
            throw FormatoFuenteInsumo(
                ase,
                archivo,
                $"fila {indice + 1} ({identidad}): la terna de valores tiene una celda no numérica " +
                $"(observado: {DescribirCelda(CeldaInsumo(filas, indice, terna.Primera))} / " +
                $"{DescribirCelda(CeldaInsumo(filas, indice, terna.Segunda))} / " +
                $"{DescribirCelda(CeldaInsumo(filas, indice, terna.Tercera))})");
        }

        return new FilaGrupoAntExtRev(indice, identidad, destino, primera.Value, segunda.Value, tercera.Value);
    }

    /// <summary>
    /// Cierra el grupo con la fila de total de la fuente: exige la igualdad de la terna en cada fila
    /// (M4) y <c>Σ conceptos ≡ total de fuente ±0,5</c>, mapea la etiqueta al catálogo y acumula la
    /// fuente presente. Una etiqueta no tabulada no se mapea: queda como evidencia (y es fail-fast si
    /// además hay una fuente del catálogo ausente).
    /// </summary>
    private static void CerrarGrupoAntExtRev(
        Ase ase,
        string archivo,
        string etiquetaFuente,
        List<FilaGrupoAntExtRev> grupo,
        FilaGrupoAntExtRev cierre,
        List<AntExtRevFuenteInputs> fuentes,
        List<string> etiquetasNoTabuladas)
    {
        if (grupo.Count == 0)
        {
            throw FormatoFuenteInsumo(
                ase,
                archivo,
                $"fila {cierre.Fila + 1}: la fuente '{etiquetaFuente}' declara su total sin ninguna fila de " +
                "concepto antes; el grupo no aporta dinero y la estructura está rota");
        }

        ExigirTernaDeValores(ase, archivo, etiquetaFuente, cierre, cierre.Identidad);

        var acumulador = new decimal[CatalogoConceptosAntExtRev.Conceptos.Count];
        foreach (var fila in grupo)
        {
            ExigirTernaDeValores(ase, archivo, etiquetaFuente, fila, fila.Identidad);
            if (fila.Destino is not null)
            {
                acumulador[fila.Destino.Indice] += fila.Tercera;
            }
        }

        var suma = acumulador.Sum();
        if (Math.Abs(suma - cierre.Tercera) > ToleranciaAntExtRev)
        {
            throw FormatoFuenteInsumo(
                ase,
                archivo,
                $"la fuente '{etiquetaFuente}' no cierra: Σ de sus cuatro conceptos = {suma} y el total de la " +
                $"fuente = {cierre.Tercera} (Δ {suma - cierre.Tercera}, fila {cierre.Fila + 1}). El dinero del " +
                "grupo no se puede atribuir sin cuadrar");
        }

        var filaCatalogo = CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente(etiquetaFuente);
        if (filaCatalogo is null)
        {
            if (!etiquetasNoTabuladas.Contains(etiquetaFuente, StringComparer.Ordinal))
            {
                etiquetasNoTabuladas.Add(etiquetaFuente);
            }

            return;
        }

        foreach (var presente in fuentes)
        {
            var filaPresente = CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente(presente.FuenteEtiqueta);
            if (filaPresente is not null && CatalogoFuentesAntExtRev.EsMismaFuente(filaPresente, filaCatalogo))
            {
                throw FormatoFuenteInsumo(
                    ase,
                    archivo,
                    $"la fuente '{etiquetaFuente}' aparece MÁS DE UNA VEZ en el bloque 2 (ya venía como " +
                    $"'{presente.FuenteEtiqueta}', misma fuente '{filaCatalogo.EtiquetaBloque}'); el destino sería " +
                    "ambiguo y no se elige por orden");
            }
        }

        fuentes.Add(new AntExtRevFuenteInputs(
            etiquetaFuente,
            acumulador[0],
            acumulador[1],
            acumulador[2],
            acumulador[3],
            cierre.Tercera));
    }

    /// <summary>
    /// Exige la INVARIANTE de la terna (enmienda M4): las tres columnas del insumo son copias del mismo
    /// monto. Si difieren más de ±0,5 el insumo está roto y el mensaje nombra ASE + fuente + concepto +
    /// los tres valores observados. No hay "columna que mande": hay igualdad.
    /// </summary>
    private static void ExigirTernaDeValores(
        Ase ase,
        string archivo,
        string etiquetaFuente,
        FilaGrupoAntExtRev fila,
        string concepto)
    {
        if (Math.Abs(fila.Primera - fila.Segunda) <= ToleranciaAntExtRev
            && Math.Abs(fila.Segunda - fila.Tercera) <= ToleranciaAntExtRev)
        {
            return;
        }

        throw FormatoFuenteInsumo(
            ase,
            archivo,
            $"fuente '{etiquetaFuente}', {concepto} (fila {fila.Fila + 1}): las tres columnas del monto deben ser " +
            $"copias del mismo valor y divergen — observado {fila.Primera} / {fila.Segunda} / {fila.Tercera} " +
            $"(Δ {fila.Primera - fila.Segunda} y {fila.Segunda - fila.Tercera}); el insumo está roto y no se " +
            "elige una columna por descarte");
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Localización de anclas del insumo (por CONTENIDO, nunca por fila/letra fija)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>Índice (0-based) de la única fila que trae el marcador; 0 o &gt;1 ocurrencias = fail-fast.</summary>
    private static int LocalizarMarcadorInsumo(List<object?[]> filas, string marcador, Ase ase, string archivo)
    {
        var esperado = NormalizarTextoInsumo(marcador);
        var encontradas = new List<int>();
        for (var i = 0; i < filas.Count; i++)
        {
            var fila = filas[i];
            for (var j = 0; j < fila.Length; j++)
            {
                if (string.Equals(
                        NormalizarTextoInsumo(ExcelWorksheetNavigator.CeldaTexto(fila[j])),
                        esperado,
                        StringComparison.Ordinal))
                {
                    encontradas.Add(i);
                    break;
                }
            }
        }

        if (encontradas.Count == 1)
        {
            return encontradas[0];
        }

        var detalle = encontradas.Count == 0
            ? $"no aparece el marcador '{marcador}' (el bloque que delimita no existe en la hoja)"
            : $"el marcador '{marcador}' aparece {encontradas.Count} veces (filas " +
              $"{string.Join(", ", encontradas.Select(i => i + 1))}); se esperaba exactamente una";
        throw FormatoFuenteInsumo(ase, archivo, detalle);
    }

    /// <summary>
    /// Índice (0-based) de la cabecera real del bloque 2: la única fila del rango que declara el par de
    /// rótulos <c>ASEO</c> seguido de <c>Total</c> (§4.1 de T0). Se busca por contenido, no por fila.
    /// </summary>
    private static int LocalizarEncabezadoBloque2(
        List<object?[]> filas,
        int filaMarcador2,
        int filaMarcador3,
        Ase ase,
        string archivo)
    {
        var encontradas = new List<int>();
        for (var i = filaMarcador2 + 1; i < filaMarcador3; i++)
        {
            var fila = filas[i];
            for (var j = 0; j + 1 < fila.Length; j++)
            {
                if (string.Equals(NormalizarTextoInsumo(ExcelWorksheetNavigator.CeldaTexto(fila[j])), RotuloAseo, StringComparison.Ordinal)
                    && string.Equals(NormalizarTextoInsumo(ExcelWorksheetNavigator.CeldaTexto(fila[j + 1])), RotuloTotal, StringComparison.Ordinal))
                {
                    encontradas.Add(i);
                    break;
                }
            }
        }

        if (encontradas.Count == 1)
        {
            return encontradas[0];
        }

        throw FormatoFuenteInsumo(
            ase,
            archivo,
            encontradas.Count == 0
                ? $"no se encontró la cabecera del bloque 2 (par de rótulos '{RotuloAseo}' seguido de 'Total' " +
                  $"entre las filas {filaMarcador2 + 1} y {filaMarcador3})"
                : $"la cabecera del bloque 2 aparece {encontradas.Count} veces (filas " +
                  $"{string.Join(", ", encontradas.Select(i => i + 1))}): la malla de la fuente es ambigua");
    }

    /// <summary>
    /// Deriva las columnas del bloque 2 a partir de su CONTENIDO: la columna de concepto es la única que
    /// trae rótulos del catálogo; la columna del rótulo <c>Total</c> que cierra los grupos de fuente es la
    /// única —distinta de la de concepto, donde el insumo rotula también el total de canal— que declara
    /// <c>Total</c> con una etiqueta a su izquierda (esa etiqueta es la fuente); la de canal es su vecina
    /// derecha y la de fuente su vecina izquierda.
    /// </summary>
    private static (int Concepto, int Canal, int Fuente) LocalizarColumnasBloque2(
        List<object?[]> filas,
        int inicio,
        int fin,
        Ase ase,
        string archivo)
    {
        var ancho = 0;
        var porConcepto = new Dictionary<int, int>();
        for (var i = inicio; i < fin; i++)
        {
            var fila = filas[i];
            ancho = Math.Max(ancho, fila.Length);
            for (var j = 1; j <= fila.Length; j++)
            {
                var texto = NormalizarTextoInsumo(ExcelWorksheetNavigator.CeldaTexto(fila[j - 1]));
                if (texto.Length > 0 && CatalogoConceptosAntExtRev.Buscar(texto) is not null)
                {
                    porConcepto[j] = porConcepto.GetValueOrDefault(j) + 1;
                }
            }
        }

        if (porConcepto.Count != 1)
        {
            throw FormatoFuenteInsumo(
                ase,
                archivo,
                porConcepto.Count == 0
                    ? "no se reconoció NINGÚN concepto del catálogo en el bloque 2 (se buscaron sus rótulos en " +
                      $"las {ancho} columnas de las filas {inicio + 1}..{fin}): la fuente no trae el bloque de conceptos"
                    : $"los conceptos del catálogo aparecen en {porConcepto.Count} columnas distintas (" +
                      $"{string.Join(", ", porConcepto.Keys.Order())}): la columna de concepto es ambigua");
        }

        var columnaConcepto = porConcepto.Keys.Single();

        var porCierre = new Dictionary<int, int>();
        for (var i = inicio; i < fin; i++)
        {
            var fila = filas[i];
            for (var j = 1; j <= fila.Length; j++)
            {
                if (j == columnaConcepto)
                {
                    continue;
                }

                var texto = NormalizarTextoInsumo(ExcelWorksheetNavigator.CeldaTexto(fila[j - 1]));
                if (texto != RotuloTotal || j <= 1)
                {
                    continue;
                }

                // Un total de fuente es "etiqueta de fuente + 'Total' + su monto a la derecha". La banda del
                // encabezado trae "'66' + 'Total'" sin monto: exigir dinero a la derecha la descarta sin
                // nombrar ninguna celda.
                if (NormalizarTextoInsumo(ExcelWorksheetNavigator.CeldaTexto(fila[j - 2])).Length > 0
                    && HayNumeroALaDerecha(fila, j))
                {
                    porCierre[j] = porCierre.GetValueOrDefault(j) + 1;
                }
            }
        }

        if (porCierre.Count != 1)
        {
            throw FormatoFuenteInsumo(
                ase,
                archivo,
                porCierre.Count == 0
                    ? "no se encontró la columna del rótulo 'Total' que cierra los grupos de fuente en el bloque 2"
                    : $"el rótulo 'Total' con una etiqueta a su izquierda aparece en {porCierre.Count} columnas " +
                      $"({string.Join(", ", porCierre.Keys.Order())}): la columna de cierre es ambigua");
        }

        var columnaCanal = porCierre.Keys.Single();
        if (columnaCanal >= columnaConcepto || columnaCanal <= 1)
        {
            throw FormatoFuenteInsumo(
                ase,
                archivo,
                $"la columna de cierre ('Total') queda en {columnaCanal} y la de concepto en {columnaConcepto}: " +
                "la gramática medida del bloque 2 (fuente — 'Total' — concepto) no se cumple");
        }

        return (columnaConcepto, columnaCanal, columnaCanal - 1);
    }

    /// <summary>
    /// Deriva la TERNA de columnas del monto (enmienda M4): entre la columna de concepto y el borde de la
    /// malla, la única terna de columnas consecutivas que, en TODAS las filas de concepto, son numéricas y
    /// copias del mismo valor ±0,5. El valor del concepto se lee de la TERCERA (la de cierre). Si no hay
    /// terna válida, el insumo está roto y el mensaje muestra los tres valores observados donde el formato
    /// medido los declara.
    /// </summary>
    private static (int Primera, int Segunda, int Tercera) LocalizarTernaDeValores(
        List<object?[]> filas,
        int inicio,
        int fin,
        int columnaConcepto,
        Ase ase,
        string archivo)
    {
        var filasConConcepto = new List<int>();
        var ancho = 0;
        for (var i = inicio; i < fin; i++)
        {
            var fila = filas[i];
            ancho = Math.Max(ancho, fila.Length);
            var texto = NormalizarTextoInsumo(ExcelWorksheetNavigator.CeldaTexto(
                columnaConcepto <= fila.Length ? fila[columnaConcepto - 1] : null));
            if (CatalogoConceptosAntExtRev.Buscar(texto) is not null)
            {
                filasConConcepto.Add(i);
            }
        }

        var candidatas = new List<(int Primera, int Segunda, int Tercera)>();
        for (var primera = columnaConcepto + 1; primera + 2 <= ancho; primera++)
        {
            var cumple = true;
            foreach (var i in filasConConcepto)
            {
                if (NumeroDeInsumo(CeldaInsumo(filas, i, primera)) is not decimal primero
                    || NumeroDeInsumo(CeldaInsumo(filas, i, primera + 1)) is not decimal segundo
                    || NumeroDeInsumo(CeldaInsumo(filas, i, primera + 2)) is not decimal tercero
                    || Math.Abs(primero - segundo) > ToleranciaAntExtRev
                    || Math.Abs(segundo - tercero) > ToleranciaAntExtRev)
                {
                    cumple = false;
                    break;
                }
            }

            if (cumple)
            {
                candidatas.Add((primera, primera + 1, primera + 2));
            }
        }

        if (candidatas.Count == 1)
        {
            return candidatas[0];
        }

        var observado = "sin filas de concepto en el bloque 2";
        foreach (var i in filasConConcepto)
        {
            var concepto = TextoNormalizadoInsumo(filas, i, columnaConcepto);
            var primera = CeldaInsumo(filas, i, columnaConcepto + 1);
            var segunda = CeldaInsumo(filas, i, columnaConcepto + 2);
            var tercera = CeldaInsumo(filas, i, columnaConcepto + 3);
            if (NumeroDeInsumo(primera) is not decimal valorPrimera
                || NumeroDeInsumo(segunda) is not decimal valorSegunda
                || NumeroDeInsumo(tercera) is not decimal valorTercera
                || Math.Abs(valorPrimera - valorSegunda) > ToleranciaAntExtRev
                || Math.Abs(valorSegunda - valorTercera) > ToleranciaAntExtRev)
            {
                observado =
                    $"concepto '{concepto}' (fila {i + 1}): {DescribirCelda(primera)} / " +
                    $"{DescribirCelda(segunda)} / {DescribirCelda(tercera)}";
                break;
            }
        }

        throw FormatoFuenteInsumo(
            ase,
            archivo,
            candidatas.Count == 0
                ? "no se encontró la terna de columnas del monto: no hay tres columnas consecutivas a la derecha " +
                  $"del concepto que sean copias del mismo valor en todas las filas de concepto (observado en " +
                  $"{observado}). La igualdad de las tres columnas es la única propiedad estable del formato y no " +
                  "se sustituye por una columna fija"
                : $"la terna de columnas del monto es ambigua: {candidatas.Count} ternas consecutivas cumplen la " +
                  "igualdad de sus tres columnas y elegir una por posición inventaría el destino del dinero");
    }

    /// <summary>
    /// Lee el bloque 3 SOLO como control (§2.4 paso 7): columnas por encabezado (variables por archivo),
    /// fila de cierre por su rótulo <c>Total</c> y comparación informativa contra el bloque 2. Devuelve el
    /// total de control y el hallazgo auditable (Δ &gt; 0,5), sin bloquear la entrega.
    /// </summary>
    private static (decimal? Total, string? Hallazgo) LeerControlBloque3(
        List<object?[]> filas,
        int filaMarcador3,
        int columnaFuente,
        decimal totalBloque2,
        Ase ase)
    {
        var filaEncabezado = filaMarcador3 + 1;
        if (filaEncabezado >= filas.Count)
        {
            return (null, $"ASE {ase.Id} · {NombreInsumo} · bloque 3: el marcador '{MarcadorBloque3}' no tiene " +
                          "fila de encabezados debajo");
        }

        var fila = filas[filaEncabezado];
        var columnas = new List<(int Columna, string Etiqueta)>();
        for (var j = 1; j <= fila.Length; j++)
        {
            var texto = NormalizarTextoInsumo(ExcelWorksheetNavigator.CeldaTexto(fila[j - 1]));
            if (texto.Length > 0)
            {
                columnas.Add((j, texto));
            }
        }

        var totales = columnas.Where(columna => columna.Etiqueta == RotuloTotal).ToArray();
        if (totales.Length != 1)
        {
            return (null, $"ASE {ase.Id} · {NombreInsumo} · bloque 3: declara {totales.Length} columnas 'Total' " +
                          "(se esperaba una)");
        }

        var filaCierre = -1;
        for (var i = filaEncabezado + 1; i < filas.Count; i++)
        {
            if (TextoNormalizadoInsumo(filas, i, columnaFuente) == RotuloTotal)
            {
                filaCierre = i;
            }
        }

        if (filaCierre < 0)
        {
            return (null, $"ASE {ase.Id} · {NombreInsumo} · bloque 3: no trae fila de cierre 'Total'; el control " +
                          "no se puede leer");
        }

        var valor = NumeroDeInsumo(CeldaInsumo(filas, filaCierre, totales[0].Columna));
        if (valor is null)
        {
            return (null, $"ASE {ase.Id} · {NombreInsumo} · bloque 3: la celda de total (fila {filaCierre + 1}, " +
                          $"columna {totales[0].Columna}) no es numérica");
        }

        var delta = valor.Value - totalBloque2;
        return Math.Abs(delta) <= ToleranciaAntExtRev
            ? (valor, null)
            : (valor, $"ASE {ase.Id} · {NombreInsumo} · bloque 3 vs bloque 2: el total de control es {valor} y la " +
                      $"Σ de las fuentes del bloque 2 es {totalBloque2} (Δ {delta}); se registra como hallazgo " +
                      "auditable (el bloque 3 no discrimina por fuente y no puede juzgar al bloque 2)");
    }

    /// <summary>
    /// Plan 64 (HU-14, §2.9 / H5.2): UN <c>Warning</c> por (ASE, fuente ausente) con el formato del Plan 51,
    /// UN <c>Warning</c> por etiqueta sin mapear con la forma de la sonda (<c>ASE n · RecaudosReversados ·
    /// Sheet1 · {etiqueta}</c>) y UN hito <c>Information</c> por ASE con el balance de la lectura. Lenguaje
    /// administrativo (Plan 26 D-E).
    /// </summary>
    private static void RegistrarObservabilidadAntExtRev(
        Ase ase,
        IReadOnlyList<AntExtRevFuenteInputs> fuentes,
        IReadOnlyList<string> ausentes,
        IReadOnlyList<string> etiquetasNoTabuladas)
    {
        foreach (var ausente in ausentes)
        {
            var fuente = SinPrefijoTraza(ausente);
            Log.ForContext("AseId", ase.Id)
                .ForContext("Reporte", "ANT EXT-REV")
                .ForContext("Fuente", fuente)
                .Warning(
                    "ASE {AseId} · ANT EXT-REV: la fuente {Fuente} no viene en el recaudo reversado del período; sus celdas quedan en el 0 del template (ausencia registrada, no es error).",
                    ase.Id,
                    fuente);
        }

        // Sonda de etiquetas (H5.2): una etiqueta que el catálogo no tabula se registra con su forma
        // administrativa. Sin ausencia en el mismo archivo es evidencia; con ausencia, el lector ya falló.
        foreach (var hallazgo in HallazgosDeEtiquetas(ase, etiquetasNoTabuladas))
        {
            Log.ForContext("AseId", ase.Id)
                .ForContext("Reporte", "ANT EXT-REV")
                .Warning(
                    "{Hallazgo}: la etiqueta no está tabulada en el catálogo de fuentes de ANT EXT-REV; se registra como hallazgo auditable (sin cita no se mapea y nunca se asume 0).",
                    hallazgo);
        }

        Log.ForContext("AseId", ase.Id).Information(
            "ASE {AseId} · ANT EXT-REV leído: {Presentes} fuente(s) presente(s), {Ausentes} ausente(s), {SinMapear} etiqueta(s) sin mapear.",
            ase.Id,
            fuentes.Count,
            ausentes.Count,
            etiquetasNoTabuladas.Count);
    }

    /// <summary>
    /// Hallazgos de la sonda de etiquetas con la forma administrativa
    /// <c>ASE n · RecaudosReversados · Sheet1 · {etiqueta}</c> (H5.2). Se derivan de la evidencia del
    /// lector: una sola función para el modelo y para el log, sin dos formatos que puedan divergir.
    /// </summary>
    private static IReadOnlyList<string> HallazgosDeEtiquetas(Ase ase, IEnumerable<string> etiquetas) =>
        etiquetas
            .Select(etiqueta => $"ASE {ase.Id} · {NombreInsumo} · {HojaInsumoRecaudosReversados} · {etiqueta}")
            .ToArray();

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Utilidades de la lectura (puras)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>Verdad si la fuente presente corresponde a la fila del catálogo (misma etiqueta de bloque).</summary>
    private static bool EsFuenteDeLaFila(AntExtRevFuenteInputs fuente, CatalogoFuentesAntExtRev.FilaFuente fila)
    {
        var deLaFuente = CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente(fuente.FuenteEtiqueta);
        return deLaFuente is not null && CatalogoFuentesAntExtRev.EsMismaFuente(deLaFuente, fila);
    }

    private static string SinPrefijoTraza(string traza) =>
        traza.StartsWith(PrefijoTraza, StringComparison.Ordinal) ? traza[PrefijoTraza.Length..] : traza;

    private static object? CeldaInsumo(List<object?[]> filas, int indice, int columna) =>
        indice >= 0 && indice < filas.Count && columna >= 1 && columna <= filas[indice].Length
            ? filas[indice][columna - 1]
            : null;

    private static string TextoNormalizadoInsumo(List<object?[]> filas, int indice, int columna) =>
        NormalizarTextoInsumo(ExcelWorksheetNavigator.CeldaTexto(CeldaInsumo(filas, indice, columna)));

    /// <summary>Verdad si la fila trae al menos un monto numérico a la derecha de <paramref name="columna"/>.</summary>
    private static bool HayNumeroALaDerecha(object?[] fila, int columna)
    {
        for (var j = columna + 1; j <= fila.Length; j++)
        {
            if (NumeroDeInsumo(fila[j - 1]) is not null)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Número de una celda del insumo. SOLO tipos numéricos: un texto (p. ej. el "66" de la banda del
    /// encabezado) NO es un monto, y una celda de dinero con texto es estructura inesperada que el
    /// llamador convierte en fail-fast.
    /// </summary>
    private static decimal? NumeroDeInsumo(object? valor) => valor switch
    {
        decimal m => m,
        double d => (decimal)d,
        float f => (decimal)f,
        long l => l,
        int i => i,
        short s => s,
        byte b => b,
        _ => null
    };

    private static string DescribirCelda(object? valor) => valor switch
    {
        null => "<vacía>",
        decimal m => m.ToString(CultureInfo.InvariantCulture),
        double d => d.ToString(CultureInfo.InvariantCulture),
        _ => $"'{ExcelWorksheetNavigator.CeldaTexto(valor)}'"
    };

    /// <summary>
    /// Normalización del §2.4: mayúsculas invariantes, sin diacríticos, <c>Trim</c> y espacios colapsados.
    /// Es la regla con la que se comparan marcadores, rótulos y etiquetas del insumo entre sí.
    /// </summary>
    private static string NormalizarTextoInsumo(string? texto)
    {
        if (string.IsNullOrWhiteSpace(texto))
        {
            return string.Empty;
        }

        var descompuesto = texto.Normalize(NormalizationForm.FormD);
        var constructor = new StringBuilder(descompuesto.Length);
        foreach (var caracter in descompuesto)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(caracter) != UnicodeCategory.NonSpacingMark)
            {
                constructor.Append(caracter);
            }
        }

        return string.Join(
            ' ',
            constructor.ToString().ToUpperInvariant().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    private static CalculoInvalidoException FormatoFuenteInsumo(Ase ase, string archivo, string detalle) =>
        new(
            CodigoError.FormatoFuente,
            $"ERR-FORMATO-FUENTE: ASE {ase.Id} · {NombreInsumo} · hoja {HojaInsumoRecaudosReversados} " +
            $"({archivo}): {detalle}. El insumo no se corrige ni se reinterpreta: se corrige el archivo y se " +
            "vuelve a correr (nunca un valor inventado, nunca un 0 silencioso).");
}
