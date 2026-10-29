using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 39 (T1, R-B-1/R-B-2/R-B-4 — D-B/D-C/D-E): <b>Phase 0</b> — preparación AUTOMÁTICA de la
/// base de un período desde la plantilla oficial, sobre una COPIA de trabajo (el oficial NUNCA se
/// abre en modo escritura). Código MOVIDO (no reescrito) del modo <c>--plantilla-unica</c> de
/// <c>Herramientas/PreparadorBasePeriodo</c> + paso nuevo 4d (selección de columna L9:L13 por
/// quincena, R-INTERV-Q).
///
/// Qué hace (BCL puro: <see cref="ZipArchive"/> + <see cref="Regex"/>; sin OpenXML/ExcelDataReader,
/// sin Excel/COM):
///   (a) renombra las 3 hojas de período vía el token de período (workbook.xml + app.xml;
///       rIds/sheetIds intactos);
///   (b) reescribe las 11 <c>&lt;f&gt;</c> del token de período (base detectada por forma
///       <c>DetRetri{7dígitos}</c>);
///   (c) sella <c>CONSOLIDADO_TOTAL RECAUDO!N3 = periodo.CodigoCompleto</c> (valor plano; si N3 trae
///       <c>&lt;f&gt;</c> → FAIL, nunca sobrescritura);
///   (d) NUEVO: selecciona <c>L9:L13</c> = <c>INTERVENTORIA!N26:N30</c> (Q1) / <c>M26:M30</c> (Q2)
///       por DOMINIO (<see cref="Periodo.NumeroQuincena"/>, jamás por contenido);
///   (e) autoverificación embebida (0 token viejo, N hojas, rIds intactos, N3 sellado, variante L
///       == esperada) + escritura atómica (tmp + verificación + move);
///   (f) Plan 62 (T4.1, D-H4): la familia E1-4 se extiende a <c>CONSOLIDADO!E142:E146</c>
///       (Σ <c>'Recaudo *'</c> por ASE en filas 21:25, columna de la quincena del run).
///
/// Invariante (D-D): esta fase SÍ puede tocar <c>&lt;f&gt;</c> (allow-list T0: 11 tokens de período
/// + 5 L9:L13). El PROCESAMIENTO de negocio (slot writer values-only y todo lo downstream) NUNCA.
///
/// D-H: cero literales de período en el cuerpo; el destino llega como <see cref="Periodo"/> y la base
/// se DETECTA por forma. Cero ramas por empresa (empresas no tabuladas = identidad).
/// </summary>
public static class PreparadorBasePipeline
{
    private const string PartWorkbook = "xl/workbook.xml";
    private const string PartApp = "docProps/app.xml";
    private const string PartShared = "xl/sharedStrings.xml";

    /// <summary>Hoja de metadatos base (mismo nombre que <c>WorkbookLeafCellMap.HojaConsolidado</c>).</summary>
    private const string HojaConsolidado = WorkbookLeafCellMap.HojaConsolidado;

    /// <summary>Nombre de hoja DetRetri del período base en la plantilla única (detección por forma).</summary>
    private static readonly Regex HojaDetRetriBase = new("^DetRetri(?<periodo>\\d{7})$", RegexOptions.Compiled);

    private static readonly Regex ReferenciaHojaDet = new(
        "(?:DetRetri(?<p1>\\d{7}))|(?:DetValiRetri(?<p2>\\d{7}))",
        RegexOptions.Compiled);

    // Prefijo XML opcional de un elemento: el oficial (Excel) no usa prefijo, pero cualquier
    // workbook re-serializado por el SDK OpenXML lo escribe como `x:` (p. ej. `<x:sheet>`). La
    // cirugía mecánica debe operar sobre AMBOS formatos sin cambiar los bytes del oficial.
    private const string P = "(?:[A-Za-z_][A-Za-z0-9_]*:)?";

    private static readonly Regex CeldaRegex = new(
        $"<(?<p>{P})?c r=\"(?<ref>[A-Z]+\\d+)\"(?<attrs>[^>]*?)(?:/>|>(?<inner>.*?)</{P}c>)",
        RegexOptions.Singleline | RegexOptions.Compiled);

    private static readonly Regex FormulaRegex = new(
        $"<{P}f(?<attrs>[^>]*?)>(?<texto>[^<]*)</{P}f>",
        RegexOptions.Singleline | RegexOptions.Compiled);

    private static readonly Regex ValorRegex = new($"<{P}v>(?<v>[^<]*)</{P}v>", RegexOptions.Compiled);
    private static readonly Regex SiRegex = new($"<{P}si>(?<cuerpo>.*?)</{P}si>", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex TextoRegex = new($"<{P}t[^>]*>(?<t>[^<]*)</{P}t>", RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex HojaRegex = new("^xl/worksheets/sheet\\d+\\.xml$", RegexOptions.Compiled);

    private static readonly UTF8Encoding Utf8SinBom = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: false);

    // ── Plan 39 (T0/R-INTERV-Q): L9:L13 -> INTERVENTORIA!{columna}{26..30} ────────────────────────
    // Geometría congelada por T0 (celda L -> fila INTERVENTORIA del ASE). La COLUMNA la decide la
    // quincena del dominio (tabla ColumnaLPorQuincena), nunca un literal suelto.

    private static readonly (string Celda, int FilaInterventoria)[] CeldasL9L13 =
    [
        ("L9", 26),
        ("L10", 27),
        ("L11", 28),
        ("L12", 29),
        ("L13", 30),
    ];

    /// <summary>
    /// R-INTERV-Q (Plan 39 §2.2): columna de <c>INTERVENTORIA</c> que gobierna <c>L9:L13</c> por
    /// quincena. Q1 (Primera) → N; Q2 (Segunda) → M.
    /// </summary>
    private static readonly IReadOnlyDictionary<int, string> ColumnaLPorQuincena = new Dictionary<int, string>
    {
        [1] = "N",
        [2] = "M",
    };

    // ── Plan 40 (T3, R-F-1 + E1-4): variantes por quincena de la familia F-NEW-1 + bloqueante Q1 ──
    //
    // Tres familias de la plantilla fijan una COLUMNA de quincena y no varían con el run:
    //   (a) E1-4: VALIDACION_TOTAL!C3:C8 referencia 'Recaudo *'!F21:F26, pero el recaudo se escribe
    //       en la columna D en Q1 (el reader usa D27 en Q1 / F27 en Q2) ⇒ C9=0 ≠ C14 ⇒ C15 FALSE.
    //   (b) F-NEW-1: DetRetri{per}!L32:L36, DetRetri{per}!O32:O36 y DetValiRetri{per}!I24:I29
    //       hardcodean INTERVENTORIA!M26:M31; R-INTERV-Q conmuta L9:L13 → N en Q1 ⇒ evaluarían
    //       FALSE (diferencia material ±1 por ASE por el split de redondeo).
    // La REGLA es la MISMA que R-INTERV-Q: la columna la decide el DOMINIO (NumeroQuincena), nunca
    // el contenido del workbook (D-H). La geometría (celda → fila) es del T0 (Plan 39 §6 / Plan 40 T3).

    /// <summary>
    /// E1-4: celdas de <c>VALIDACION_TOTAL</c> que suman el bloque <c>'Recaudo *'</c> y su fila
    /// (geometría T0). La COLUMNA la decide <see cref="ColumnaRecaudoPorQuincena"/>.
    /// </summary>
    private static readonly (string Celda, int FilaRecaudo)[] CeldasRecaudoValidacion =
    [
        ("C3", 21),
        ("C4", 22),
        ("C5", 23),
        ("C6", 24),
        ("C7", 25),
        ("C8", 26),
    ];

    /// <summary>
    /// E1-4: columna de valor del bloque <c>'Recaudo *'</c> por quincena (dominio). Q1 → D; Q2 → F.
    /// Es el mismo eje que el reader de recaudo (<c>D27</c> en Q1 / <c>F27</c> en Q2).
    /// </summary>
    private static readonly IReadOnlyDictionary<int, string> ColumnaRecaudoPorQuincena = new Dictionary<int, string>
    {
        [1] = "D",
        [2] = "F",
    };

    // ── Plan 62 (T4.1, D-H4 / D-E1-4EXT): la familia E1-4 se extiende al bloque 142–147 ──────────
    //
    // El bloque "VALIDACION DEL RECAUDO VERSUS REMUNERACION" del CONSOLIDADO suma el recaudo de los
    // 5 canales en una COLUMNA FIJA (F, filas 21:25) — la de la 2.ª quincena —, así que en Q1 el
    // bloque queda en 0 mientras el recaudo se escribe en D (desalineación T0-12 del Plan 62). La
    // REGLA es la MISMA de E1-4: la columna la decide el DOMINIO (NumeroQuincena) vía
    // <see cref="ColumnaRecaudoDe"/>, nunca el contenido del workbook (D-H). La geometría
    // (E142→fila 21 … E146→fila 25) es del T0-10/T0-11 (Plan 62); la fila 147 (cierre) NO se toca.

    /// <summary>
    /// Plan 62 (T4.1, D-H4): celdas de <c>CONSOLIDADO_TOTAL RECAUDO!E142:E146</c> que suman el
    /// bloque <c>'Recaudo *'</c> y su fila en cada hoja de canal (geometría T0-11: E142→21 …
    /// E146→25). La COLUMNA la decide <see cref="ColumnaRecaudoPorQuincena"/> (misma fuente que E1-4).
    /// </summary>
    private static readonly (string Celda, int FilaRecaudo)[] CeldasRecaudoValidacionConsolidado142 =
    [
        ("E142", 21),
        ("E143", 22),
        ("E144", 23),
        ("E145", 24),
        ("E146", 25),
    ];

    /// <summary>F-NEW-1: <c>DetRetri{per}!L32:L36</c> = <c>INTERVENTORIA!M{26..30}+L{23..27}=L{9..13}</c>.</summary>
    private static readonly (string Celda, int FilaInterventoria)[] CeldasAutoChequeoDetRetriL =
    [
        ("L32", 26), ("L33", 27), ("L34", 28), ("L35", 29), ("L36", 30),
    ];

    /// <summary>F-NEW-1: <c>DetRetri{per}!O32:O36</c> = reconciliación que resta <c>INTERVENTORIA!M{26..30}</c>.</summary>
    private static readonly (string Celda, int FilaInterventoria)[] CeldasAutoChequeoDetRetriO =
    [
        ("O32", 26), ("O33", 27), ("O34", 28), ("O35", 29), ("O36", 30),
    ];

    /// <summary>F-NEW-1: <c>DetValiRetri{per}!I24:I29</c> = <c>-INTERVENTORIA!M{26..31}+I{16..21}=I{9..14}</c>.</summary>
    private static readonly (string Celda, int FilaInterventoria)[] CeldasAutoChequeoDetValiRetriI =
    [
        ("I24", 26), ("I25", 27), ("I26", 28), ("I27", 29), ("I28", 30), ("I29", 31),
    ];

    /// <summary>
    /// Plan 40 (T3, E1-4): columna del bloque <c>'Recaudo *'</c> de la quincena del run. Fuente
    /// ÚNICA del dato (la consumen Phase 0 y el sellado del writer). Quincena no soportada = FAIL.
    /// </summary>
    public static string ColumnaRecaudoDe(Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        if (!ColumnaRecaudoPorQuincena.TryGetValue(periodo.NumeroQuincena, out var columna))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"E1-4: quincena no soportada ({periodo.NumeroQuincena}); solo 1 (Primera) o 2 (Segunda).");
        }

        return columna;
    }

    /// <summary>
    /// Ejecuta Phase 0: lee <paramref name="rutaEntrada"/> (read-only) y escribe la base preparada
    /// del período en <paramref name="rutaSalida"/> de forma atómica (tmp + verificación + move).
    /// Endurecido (R-INPLACE): la entrada NUNCA se abre en modo escritura ni se usa como salida
    /// (<c>rutaEntrada != rutaSalida</c>); el oficial se abre SOLO lectura y la copia de trabajo es
    /// la salida.
    /// </summary>
    /// <param name="rutaEntrada">Plantilla de entrada (oficial o copia). Se abre SOLO lectura.</param>
    /// <param name="periodo">Período destino (dominio: hojas + tokens + N3 + columna L).</param>
    /// <param name="rutaSalida">Ruta de la base preparada (distinta de <paramref name="rutaEntrada"/>).</param>
    /// <param name="periodoBaseEsperado">Período base esperado (opcional; si se indica y no coincide con el detectado → FAIL).</param>
    /// <param name="log">Sink de log opcional (una línea por evento).</param>
    public static void Ejecutar(
        string rutaEntrada,
        Periodo periodo,
        string rutaSalida,
        string? periodoBaseEsperado = null,
        Action<string>? log = null)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        log ??= _ => { };

        Exigir(!string.IsNullOrWhiteSpace(rutaEntrada), "La ruta de la plantilla de entrada es requerida.");
        Exigir(!string.IsNullOrWhiteSpace(rutaSalida), "La ruta de salida es requerida.");
        Exigir(File.Exists(rutaEntrada), $"No existe la plantilla de entrada: {rutaEntrada}");

        // R-INPLACE endurecido: Phase 0 jamás escribe in-place sobre su entrada (el oficial nunca se
        // abre en modo escritura; la copia de trabajo es la salida).
        Exigir(!string.Equals(Path.GetFullPath(rutaEntrada), Path.GetFullPath(rutaSalida), StringComparison.OrdinalIgnoreCase),
            "Phase 0 no puede escribir in-place sobre su entrada: la plantilla de origen nunca se modifica (oficial != trabajo).");

        var periodoDestino = periodo.CodigoCompleto;
        Exigir(periodoDestino.Length == 7, $"El período destino debe ser AAAAMMQ; se recibió '{periodoDestino}'.");

        log($"Plantilla entrada : {rutaEntrada}");
        log($"Período destino   : {periodoDestino} (Q{periodo.NumeroQuincena})");
        log($"Salida            : {rutaSalida}");

        var orden = new List<string>();
        var partes = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        var resumen = new List<string>();
        string basePeriodo;
        string workbookOriginal;
        string hojaConsolidadoParte;
        string hojaValidacionTotalParte;
        string hojaDetRetriParte;
        string hojaDetValiRetriParte;
        string hojaInterventoriaParte;
        List<(string Hoja, string Parte)> hojasRecaudoValidacion = [];
        var formulasPeriodoBase = 0;

        // Se lee el ZIP completo y se libera ANTES de escribir: habilita el trabajo in-place sobre
        // la copia de trabajo (el move atómico borra/mueve el destino).
        using (var origenZip = ZipFile.OpenRead(rutaEntrada))
        {
            var entradasOrigen = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (var entrada in origenZip.Entries)
            {
                Exigir(!entradasOrigen.ContainsKey(entrada.FullName), $"Entrada ZIP duplicada: {entrada.FullName}");
                entradasOrigen[entrada.FullName] = entrada;
                orden.Add(entrada.FullName);
            }

            workbookOriginal = LeerTexto(entradasOrigen[PartWorkbook]);
            var baseDetectada = DetectarPeriodoBase(workbookOriginal);

            if (baseDetectada is null)
            {
                // La plantilla no trae hoja base `DetRetri{AAAAMMQ}`: NO es una base de período, así
                // que no hay token de período que reescribir. Phase 0 no transforma (se comporta como
                // antes de este plan: copia) y las validaciones del writer siguen siendo la red
                // (fail-fast honesto que nombra la hoja del período). El OFICIAL siempre trae base.
                log("[base] sin hoja 'DetRetri{AAAAMMQ}': Phase 0 no transforma la entrada (copia)");
                foreach (var entrada in origenZip.Entries)
                {
                    partes[entrada.FullName] = LeerBytes(entrada);
                }

                EscribirBytesAtomico(SerializarZip(orden, partes), rutaSalida);
                log($"OK: entrada copiada sin transformar en {rutaSalida}");
                return;
            }

            basePeriodo = baseDetectada;
            if (!string.IsNullOrWhiteSpace(periodoBaseEsperado))
            {
                Exigir(string.Equals(periodoBaseEsperado, basePeriodo, StringComparison.Ordinal),
                    $"El período base indicado ('{periodoBaseEsperado}') no coincide con el interno detectado ('{basePeriodo}').");
            }

            log($"[base detectada] período interno de la plantilla: {basePeriodo}");

            var tokenBase = new Regex($"(?<![0-9]){basePeriodo}(?![0-9])", RegexOptions.Compiled);
            var informeBase = $"{basePeriodo[..6]}-{basePeriodo[6]}";
            var informeDestino = $"{periodoDestino[..6]}-{periodoDestino[6]}";
            hojaConsolidadoParte = ResolverParteDeHoja(origenZip, workbookOriginal, HojaConsolidado);

            // Plan 40 (T3): partes de las 3 familias por-quincena (geometría del T0). Se resuelven
            // por NOMBRE (nunca enumerando hojas); el oficial trae las hojas con la base detectada.
            hojaValidacionTotalParte = ResolverParteDeHoja(origenZip, workbookOriginal, WorkbookLeafCellMapValidaciones.HojaValidacionTotal);
            hojaDetRetriParte = ResolverParteDeHoja(origenZip, workbookOriginal, NombresHojaPeriodo.DetRetri(basePeriodo));
            hojaDetValiRetriParte = ResolverParteDeHoja(origenZip, workbookOriginal, NombresHojaPeriodo.DetValiRetri(basePeriodo));
            hojaInterventoriaParte = ResolverParteDeHoja(origenZip, workbookOriginal, WorkbookLeafCellMapAjustesSfT.HojaInterventoria);
            hojasRecaudoValidacion = new List<(string Hoja, string Parte)>
            {
                (WorkbookLeafCellMapValidaciones.HojaValidacionTotal, hojaValidacionTotalParte),
            };
            foreach (var empresa in EmpresaFacturacion.Catalogo.OrderBy(e => e.Id))
            {
                hojasRecaudoValidacion.Add((empresa.HojaValidacion, ResolverParteDeHoja(origenZip, workbookOriginal, empresa.HojaValidacion)));
            }

            // sharedStrings: token de período reemplazado (sin append: este modo no agrega textos).
            partes[PartShared] = Utf8SinBom.GetBytes(tokenBase.Replace(LeerTexto(entradasOrigen[PartShared]), periodoDestino));

            foreach (var entrada in origenZip.Entries)
            {
                var nombre = entrada.FullName;
                if (string.Equals(nombre, PartShared, StringComparison.OrdinalIgnoreCase))
                {
                    continue; // ya emitido arriba.
                }

                var bytes = LeerBytes(entrada);
                if (!EsParteTextoTransformable(nombre))
                {
                    partes[nombre] = bytes;
                    continue;
                }

                var texto = Encoding.UTF8.GetString(bytes);
                var nToken = tokenBase.Matches(texto).Count;
                var nFormulaPeriodo = ContarFormulasConPeriodo(texto, basePeriodo);
                formulasPeriodoBase += nFormulaPeriodo;

                texto = tokenBase.Replace(texto, periodoDestino);
                texto = texto.Replace(informeBase, informeDestino, StringComparison.Ordinal);

                if (string.Equals(nombre, hojaConsolidadoParte, StringComparison.OrdinalIgnoreCase))
                {
                    // Sello del código de período de la base (NUNCA sobre fórmula: si el destino trae
                    // <f>, SetValorNumerico lo rechaza y la autoverificación lo delata).
                    texto = SetValorNumerico(texto, "N3", periodoDestino, out var aplicado, out var motivo);
                    resumen.Add(aplicado
                        ? $"  {nombre}: CONSOLIDADO!N3 = {periodoDestino} (sello de período de la base)"
                        : $"  {nombre}: CONSOLIDADO!N3 OMITIDO ({motivo})");

                    // Plan 39 (T1, paso 4d / R-INTERV-Q): variante L9:L13 por quincena (dominio).
                    texto = AplicarVarianteL(texto, periodo, resumen);

                    // Plan 62 (T4.1, D-H4 / D-E1-4EXT): el bloque 142–147 del CONSOLIDADO suma el
                    // recaudo de los 5 canales en la columna de la quincena del run (Q1 → D,
                    // Q2 → F). Sin esto el bloque lee F21:F25 en Q1 (donde el recaudo aterriza en D)
                    // y queda en 0 (desalineación T0-12).
                    texto = AplicarVarianteRecaudoValidacionConsolidado142(texto, periodo, resumen);
                }

                // Plan 40 (T3, E1-4): las celdas que suman el bloque 'Recaudo *'
                // (VALIDACION_TOTAL!C3:C8 + VALIDACION_*!C3:C8) → columna de la quincena
                // (D en Q1 / F en Q2). Sin esto, C9/D24=0 en Q1 ⇒ C15/D25 FALSE.
                foreach (var (hojaRecaudo, parteRecaudo) in hojasRecaudoValidacion)
                {
                    if (string.Equals(nombre, parteRecaudo, StringComparison.OrdinalIgnoreCase))
                    {
                        texto = AplicarVarianteRecaudoValidacion(texto, hojaRecaudo, periodo, resumen);
                    }
                }

                // Plan 40 (T3, F-NEW-1): auto-chequeos de DetRetri/DetValiRetri que hardcodean
                // INTERVENTORIA!M26:M31 → columna de la quincena (N en Q1 / M en Q2).
                if (string.Equals(nombre, hojaDetRetriParte, StringComparison.OrdinalIgnoreCase))
                {
                    var hojaDetRetri = NombresHojaPeriodo.DetRetri(periodoDestino);
                    texto = AplicarVarianteAutoChequeos(texto, hojaDetRetri, periodo, CeldasAutoChequeoDetRetriL, resumen);
                    texto = AplicarVarianteAutoChequeos(texto, hojaDetRetri, periodo, CeldasAutoChequeoDetRetriO, resumen);
                }

                if (string.Equals(nombre, hojaDetValiRetriParte, StringComparison.OrdinalIgnoreCase))
                {
                    texto = AplicarVarianteAutoChequeos(texto, NombresHojaPeriodo.DetValiRetri(periodoDestino), periodo, CeldasAutoChequeoDetValiRetriI, resumen);
                }

                // Plan 41 (T3): el auto-chequeo S26:S31 compara Q (sigue la selección L por quincena)
                // contra R (= tabla J fija = Q2). Phase 0 reescribe el maestro R26 a la columna de la
                // quincena del run para que el chequeo sea honesto en Q1 y Q2.
                if (string.Equals(nombre, hojaInterventoriaParte, StringComparison.OrdinalIgnoreCase))
                {
                    texto = AplicarVarianteTablaJ(texto, periodo, resumen);
                }

                partes[nombre] = Utf8SinBom.GetBytes(texto);

                if (nToken > 0)
                {
                    resumen.Add($"  {nombre}: token {basePeriodo}->{periodoDestino} x{nToken}; <f> del período x{nFormulaPeriodo}");
                }
            }
        }

        log("[cambios]");
        foreach (var linea in resumen)
        {
            log(linea);
        }

        log("[autoverificación embebida]");
        foreach (var linea in VerificarBase(partes, workbookOriginal, basePeriodo, periodoDestino, hojaConsolidadoParte, hojasRecaudoValidacion, hojaDetRetriParte, hojaDetValiRetriParte, hojaInterventoriaParte, formulasPeriodoBase, periodo))
        {
            log("  " + linea);
        }

        // Plan 58 (H1.1/H1.2): la base preparada se valida POR BYTES contra la plantilla ANTES de
        // escribirse. Ninguna celda con <f> puede desaparecer ni perder sus atributos shared
        // (t/ref/si) y no puede aparecer ninguna fórmula nueva. Se serializa en memoria: un defecto
        // aborta sin producir archivo (nada de "arreglar después").
        var paquete = SerializarZip(orden, partes);
        ValidadorIntegridadPaquete.ValidarPreparacion(File.ReadAllBytes(rutaEntrada), paquete);

        EscribirBytesAtomico(paquete, rutaSalida);
        log($"OK: base canónica del período {periodoDestino} escrita en {rutaSalida}");
    }

    // ── Plan 39 (T1, paso 4d): selección de la columna L9:L13 por quincena (R-INTERV-Q) ──────────

    /// <summary>
    /// R-INTERV-Q: fija <c>CONSOLIDADO_TOTAL RECAUDO!L9:L13</c> a la variante de la quincena del
    /// run (texto-<c>&lt;f&gt;</c> + caché <c>&lt;v&gt;</c>). Si la celda ya trae la variante
    /// esperada, NO se toca (paridad byte-preserving). Si trae la variante contraria conocida
    /// (M/N del mismo ASE), se reescribe. Cualquier tercera forma = FAIL (<c>ERR-PLANTILLA</c>).
    /// </summary>
    private static string AplicarVarianteL(string xml, Periodo periodo, List<string> resumen)
    {
        if (!ColumnaLPorQuincena.TryGetValue(periodo.NumeroQuincena, out var columna))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"R-INTERV-Q: quincena no soportada ({periodo.NumeroQuincena}); solo 1 (Primera) o 2 (Segunda).");
        }

        // La mitad declarada la decide el DOMINIO en un solo lugar (primera 'N' / segunda 'M').
        var valores = InterventoriaDeclarada.PorQuincena(periodo.NumeroQuincena);

        foreach (var (celda, filaInterventoria) in CeldasL9L13)
        {
            var aseId = filaInterventoria - InterventoriaDeclarada.FilaPrimerAse + 1;
            var formulaEsperada = $"INTERVENTORIA!{columna}{filaInterventoria}";
            var valorEsperado = valores[aseId].ToString(CultureInfo.InvariantCulture);

            xml = AsegurarVarianteL(xml, celda, filaInterventoria, formulaEsperada, valorEsperado, periodo, out var cambiado);
            resumen.Add(cambiado
                ? $"  {HojaConsolidado}!{celda}: L-variante = {formulaEsperada} (Q{periodo.NumeroQuincena}, R-INTERV-Q)"
                : $"  {HojaConsolidado}!{celda}: L-variante ya correcta ({formulaEsperada})");
        }

        return xml;
    }

    // ── Plan 58 (H1.1): cirugía de <f> que NO puede romper el grupo shared ──────────────────────

    /// <summary>
    /// Plan 58 (H1.1) — sustituye SOLO el nodo de texto de la <c>&lt;f&gt;</c> de una celda.
    ///
    /// Regla estructural: los atributos del elemento (<c>t</c>, <c>ref</c>, <c>si</c>) NO se
    /// reconstruyen; se reutiliza el texto crudo que capturó <see cref="FormulaRegex"/>, así que
    /// viajan intactos por construcción y no por acuerdo. Lo único que cambia es el texto.
    ///
    /// Por qué importa (causa raíz del P0 del Plan 58): <c>INTERVENTORIA!R26</c> es el maestro de un
    /// grupo <c>shared ref="R26:R31" si="0"</c>. Reescribir el elemento completo
    /// (<c>&lt;f&gt;N26&lt;/f&gt;</c>) borraba <c>t/ref/si</c> y dejaba <c>R27:R31</c> como
    /// seguidoras de un grupo sin maestro; Excel detecta el grupo huérfano y repara el libro
    /// ELIMINANDO las fórmulas.
    /// </summary>
    /// <param name="inner">Contenido crudo de la celda (lo que está entre <c>&lt;c ...&gt;</c> y <c>&lt;/c&gt;</c>).</param>
    /// <param name="prefijo">Prefijo XML del elemento (<c>""</c> o <c>"x:"</c>).</param>
    /// <param name="texto">Nuevo texto de la fórmula.</param>
    /// <param name="hoja">Hoja (solo para el mensaje de error).</param>
    /// <param name="celda">Celda (solo para el mensaje de error).</param>
    private static string SustituirTextoFormula(string inner, string prefijo, string texto, string hoja, string celda)
    {
        var match = FormulaRegex.Match(inner);
        Exigir(match.Success,
            $"Plan 58: la celda {hoja}!{celda} no trae una <f> simple cuando la cirugía de la base la requiere.");

        var attrsFormula = match.Groups["attrs"].Value;
        var sustituta = $"<{prefijo}f{attrsFormula}>{texto}</{prefijo}f>";
        var resultado = string.Concat(
            inner.AsSpan(0, match.Index),
            sustituta,
            inner.AsSpan(match.Index + match.Length));

        // Invariante verificada (no solo por construcción): la <f> resultante conserva EXACTAMENTE
        // los mismos atributos que la original. Si un cambio futuro rompe esto, el fail-fast aparece
        // acá y no en el Excel del Ingeniero.
        var despues = FormulaRegex.Match(resultado);
        Exigir(despues.Success
                && string.Equals(despues.Groups["attrs"].Value, attrsFormula, StringComparison.Ordinal),
            $"Plan 58: la cirugía de fórmula de {hoja}!{celda} alteraría los atributos de <f> " +
            $"('{attrsFormula}' -> '{despues.Groups["attrs"].Value}'), lo que rompería el grupo shared. " +
            "Se aborta sin escribir (un grupo huérfano hace que Excel borre las fórmulas).");

        return resultado;
    }

    private static string AsegurarVarianteL(
        string xml,
        string celda,
        int filaInterventoria,
        string formulaEsperada,
        string valorEsperado,
        Periodo periodo,
        out bool cambiado)
    {
        var match = BuscarCelda(xml, celda)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"R-INTERV-Q: falta la celda {HojaConsolidado}!{celda} en la base del período {periodo.CodigoCompleto}.");

        var (prefijo, _, inner, autocierre, tieneFormula) = AnalizarCelda(match);
        if (!tieneFormula || autocierre || inner is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"R-INTERV-Q: {HojaConsolidado}!{celda} no es una <f> simple en la base del período {periodo.CodigoCompleto} (forma inesperada).");
        }

        var formulaActual = FormulaRegex.Match(inner).Groups["texto"].Value;
        if (!EsVarianteLConocida(formulaActual, filaInterventoria))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"R-INTERV-Q: {HojaConsolidado}!{celda} trae la forma inesperada '{formulaActual}' " +
                $"(se esperaba {formulaEsperada}); fail-fast, nunca se sobrescribe a ciegas.");
        }

        var valorActual = ValorRegex.Match(inner).Groups["v"].Value;
        if (string.Equals(formulaActual, formulaEsperada, StringComparison.Ordinal)
            && string.Equals(valorActual, valorEsperado, StringComparison.Ordinal))
        {
            cambiado = false;
            return xml; // no-op byte-preserving (paridad Q2 sobre el oficial).
        }

        var nuevoInner = SustituirTextoFormula(inner, prefijo, formulaEsperada, HojaConsolidado, celda);
        nuevoInner = ValorRegex.IsMatch(nuevoInner)
            ? ValorRegex.Replace(nuevoInner, _ => $"<{prefijo}v>{valorEsperado}</{prefijo}v>")
            : nuevoInner + $"<{prefijo}v>{valorEsperado}</{prefijo}v>";
        var nuevaCelda = $"<{prefijo}c r=\"{celda}\"{match.Groups["attrs"].Value}>{nuevoInner}</{prefijo}c>";
        cambiado = true;
        return xml[..match.Index] + nuevaCelda + xml[(match.Index + match.Length)..];
    }

    private static bool EsVarianteLConocida(string formula, int filaInterventoria) =>
        string.Equals(formula, $"INTERVENTORIA!M{filaInterventoria}", StringComparison.Ordinal)
        || string.Equals(formula, $"INTERVENTORIA!N{filaInterventoria}", StringComparison.Ordinal);

    /// <summary>
    /// R-INTERV-Q: <c>&lt;f&gt;</c> esperadas de <c>CONSOLIDADO_TOTAL RECAUDO!L9:L13</c> para la
    /// quincena del run. Fuente ÚNICA de la regla (la consumen la autoverificación de Phase 0 y el
    /// preflight <see cref="ValidadorBasePeriodo"/>).
    /// </summary>
    public static IReadOnlyList<(string Celda, string Formula)> FormulasLEsperadas(Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        if (!ColumnaLPorQuincena.TryGetValue(periodo.NumeroQuincena, out var columna))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"R-INTERV-Q: quincena no soportada ({periodo.NumeroQuincena}); solo 1 (Primera) o 2 (Segunda).");
        }

        return CeldasL9L13
            .Select(c => (c.Celda, $"INTERVENTORIA!{columna}{c.FilaInterventoria}"))
            .ToList();
    }

    // ── Plan 40 (T3, E1-4 + F-NEW-1): variantes de quincena de la familia por-columna ─────────────

    /// <summary>
    /// Plan 40 (T3, E1-4): reescribe <c>VALIDACION_TOTAL!C3:C8</c> a la columna del bloque
    /// <c>'Recaudo *'</c> de la quincena del run (Q1 → D, Q2 → F). La forma de cada término se
    /// aserta (Σ de referencias al catálogo en la fila esperada); cualquier tercera forma = FAIL.
    /// </summary>
    private static string AplicarVarianteRecaudoValidacion(string xml, string hoja, Periodo periodo, List<string> resumen)
    {
        var columna = ColumnaRecaudoDe(periodo);
        foreach (var (celda, fila) in CeldasRecaudoValidacion)
        {
            xml = AsegurarVarianteRecaudo(xml, hoja, celda, fila, columna, periodo, out var cambiado);
            resumen.Add(cambiado
                ? $"  {hoja}!{celda}: recaudo → columna {columna} (Q{periodo.NumeroQuincena}, E1-4)"
                : $"  {hoja}!{celda}: recaudo ya en columna {columna} (Q{periodo.NumeroQuincena})");
        }

        return xml;
    }

    /// <summary>
    /// Plan 62 (T4.1, D-H4 / D-E1-4EXT): reescribe <c>CONSOLIDADO_TOTAL RECAUDO!E142:E146</c> a la
    /// columna del bloque <c>'Recaudo *'</c> de la quincena del run (Q1 → D, Q2 → F).
    ///
    /// <para><b>Reuso (decisión, cero duplicación):</b> la construcción de la fórmula es EXACTAMENTE
    /// la de la familia E1-4 — <see cref="AsegurarVarianteRecaudo"/> + <see cref="ReconstruirFormulaRecaudo"/>
    /// —, que son genéricas (hoja + celda + fila + columna) y ya validan cada término contra el
    /// catálogo abierto de <c>'Recaudo *'</c> y su fila. No se agrega un constructor de fórmula
    /// nuevo: solo una tabla de celdas (geometría T0) y este bucle. La etiqueta de error sigue
    /// siendo <c>E1-4</c> porque la extensión ES la familia (D-E1-4EXT), no una familia distinta.</para>
    ///
    /// <para>Alcance cerrado: SOLO el texto de <c>E142:E146</c>. <c>D/F..M</c>, la fila 147 y las
    /// fórmulas compartidas <c>L145</c> (<c>si=68</c>) / <c>M142</c> (<c>si=67</c>) quedan
    /// byte-idénticas (R5-GATE): <see cref="SustituirTextoFormula"/> reutiliza los atributos crudos
    /// de la <c>&lt;f&gt;</c> y <c>ValidadorIntegridadPaquete.ValidarPreparacion</c> aborta si
    /// <c>t</c>/<c>ref</c>/<c>si</c> cambian.</para>
    /// </summary>
    private static string AplicarVarianteRecaudoValidacionConsolidado142(string xml, Periodo periodo, List<string> resumen)
    {
        var columna = ColumnaRecaudoDe(periodo);
        foreach (var (celda, fila) in CeldasRecaudoValidacionConsolidado142)
        {
            xml = AsegurarVarianteRecaudo(xml, HojaConsolidado, celda, fila, columna, periodo, out var cambiado);
            resumen.Add(cambiado
                ? $"  {HojaConsolidado}!{celda}: recaudo → columna {columna} (Q{periodo.NumeroQuincena}, E1-4/142-147)"
                : $"  {HojaConsolidado}!{celda}: recaudo ya en columna {columna} (Q{periodo.NumeroQuincena})");
        }

        return xml;
    }

    private static string AsegurarVarianteRecaudo(
        string xml,
        string hoja,
        string celda,
        int fila,
        string columna,
        Periodo periodo,
        out bool cambiado)
    {
        var match = BuscarCelda(xml, celda)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"E1-4: falta la celda {hoja}!{celda} en la base del período {periodo.CodigoCompleto}.");

        var (prefijo, _, inner, autocierre, tieneFormula) = AnalizarCelda(match);
        if (!tieneFormula || autocierre || inner is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"E1-4: {hoja}!{celda} no es una <f> simple en la base del período {periodo.CodigoCompleto} (forma inesperada).");
        }

        var formulaActual = FormulaRegex.Match(inner).Groups["texto"].Value;
        var esperada = ReconstruirFormulaRecaudo(formulaActual, hoja, celda, fila, columna, periodo);
        if (string.Equals(formulaActual, esperada, StringComparison.Ordinal))
        {
            cambiado = false;
            return xml;
        }

        var nuevoInner = SustituirTextoFormula(inner, prefijo, esperada, hoja, celda);
        var nuevaCelda = $"<{prefijo}c r=\"{celda}\"{match.Groups["attrs"].Value}>{nuevoInner}</{prefijo}c>";
        cambiado = true;
        return xml[..match.Index] + nuevaCelda + xml[(match.Index + match.Length)..];
    }

    /// <summary>
    /// Valida que <paramref name="formula"/> sea (suma de) referencias al bloque <c>'Recaudo *'</c>
    /// (hojas del catálogo abierto) en la fila esperada, con una columna de quincena CONOCIDA
    /// (<c>D</c>/<c>F</c>), y devuelve la misma forma con la columna objetivo.
    /// </summary>
    private static string ReconstruirFormulaRecaudo(string formula, string hoja, string celda, int fila, string columna, Periodo periodo)
    {
        var hojas = EmpresaFacturacion.Catalogo.Select(e => e.HojaRecaudo).ToHashSet(StringComparer.Ordinal);
        var terminos = SepararTerminosSuma(formula);
        var partes = new List<string>(terminos.Count);
        foreach (var termino in terminos)
        {
            var m = Regex.Match(termino, "^'(?<hoja>[^']+)'!(?<col>[A-Z]+)(?<fila>\\d+)$");
            if (!m.Success
                || !hojas.Contains(m.Groups["hoja"].Value)
                || !int.TryParse(m.Groups["fila"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var filaTermino)
                || filaTermino != fila
                || !EsColumnaRecaudoConocida(m.Groups["col"].Value))
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"E1-4: {hoja}!{celda} trae la forma inesperada '{formula}' " +
                    $"(se esperaba Σ 'Recaudo *'!{{col}}{fila} con col ∈ {{D,F}}; Q{periodo.NumeroQuincena}); fail-fast, nunca se sobrescribe a ciegas.");
            }

            partes.Add($"'{m.Groups["hoja"].Value}'!{columna}{fila}");
        }

        return string.Join("+", partes);
    }

    /// <summary>
    /// Parte una suma por <c>+</c> respetando los nombres de hoja entrecomillados (que pueden
    /// contener <c>+</c>, p. ej. <c>'Recaudo EAAB + Ciud Limp'</c>). Público: lo comparte el
    /// preflight <see cref="ValidadorBasePeriodo"/>.
    /// </summary>
    public static IReadOnlyList<string> SepararTerminosSuma(string formula)
    {
        var terminos = new List<string>();
        var enComillas = false;
        var inicio = 0;
        for (var i = 0; i < formula.Length; i++)
        {
            var ch = formula[i];
            if (ch == '\'')
            {
                enComillas = !enComillas;
            }
            else if (ch == '+' && !enComillas)
            {
                terminos.Add(formula[inicio..i]);
                inicio = i + 1;
            }
        }

        terminos.Add(formula[inicio..]);
        return terminos;
    }

    private static bool EsColumnaRecaudoConocida(string columna) =>
        ColumnaRecaudoPorQuincena.Values.Contains(columna, StringComparer.Ordinal);

    /// <summary>
    /// Plan 40 (T3, F-NEW-1): reescribe los auto-chequeos que hardcodean <c>INTERVENTORIA!M{26..31}</c>
    /// a la columna de la quincena del run (Q1 → N, Q2 → M — misma tabla que R-INTERV-Q).
    /// </summary>
    private static string AplicarVarianteAutoChequeos(
        string xml,
        string hoja,
        Periodo periodo,
        (string Celda, int FilaInterventoria)[] celdas,
        List<string> resumen)
    {
        if (!ColumnaLPorQuincena.TryGetValue(periodo.NumeroQuincena, out var columna))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"F-NEW-1: quincena no soportada ({periodo.NumeroQuincena}); solo 1 (Primera) o 2 (Segunda).");
        }

        foreach (var (celda, fila) in celdas)
        {
            xml = AsegurarVarianteAutoChequeo(xml, hoja, celda, fila, columna, periodo, out var cambiado);
            resumen.Add(cambiado
                ? $"  {hoja}!{celda}: INTERVENTORIA!{columna}{fila} (Q{periodo.NumeroQuincena}, F-NEW-1)"
                : $"  {hoja}!{celda}: ya INTERVENTORIA!{columna}{fila} (Q{periodo.NumeroQuincena})");
        }

        return xml;
    }

    private static string AsegurarVarianteAutoChequeo(
        string xml,
        string hoja,
        string celda,
        int fila,
        string columna,
        Periodo periodo,
        out bool cambiado)
    {
        var match = BuscarCelda(xml, celda)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"F-NEW-1: falta la celda {hoja}!{celda} en la base del período {periodo.CodigoCompleto}.");

        var (prefijo, _, inner, autocierre, tieneFormula) = AnalizarCelda(match);
        if (!tieneFormula || autocierre || inner is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"F-NEW-1: {hoja}!{celda} no es una <f> simple en la base del período {periodo.CodigoCompleto} (forma inesperada).");
        }

        var formulaActual = FormulaRegex.Match(inner).Groups["texto"].Value;
        var objetivo = $"INTERVENTORIA!{columna}{fila}";
        if (formulaActual.Contains(objetivo, StringComparison.Ordinal))
        {
            cambiado = false;
            return xml;
        }

        var contraria = $"INTERVENTORIA!{ColumnaContraria(columna)}{fila}";
        if (!formulaActual.Contains(contraria, StringComparison.Ordinal))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"F-NEW-1: {hoja}!{celda} trae la forma inesperada '{formulaActual}' " +
                $"(se esperaba '{contraria}' o '{objetivo}'); fail-fast, nunca se sobrescribe a ciegas.");
        }

        var nueva = formulaActual.Replace(contraria, objetivo, StringComparison.Ordinal);
        var nuevoInner = SustituirTextoFormula(inner, prefijo, nueva, hoja, celda);
        var nuevaCelda = $"<{prefijo}c r=\"{celda}\"{match.Groups["attrs"].Value}>{nuevoInner}</{prefijo}c>";
        cambiado = true;
        return xml[..match.Index] + nuevaCelda + xml[(match.Index + match.Length)..];
    }

    private static string ColumnaContraria(string columna) =>
        ColumnaLPorQuincena.Values.FirstOrDefault(c => !string.Equals(c, columna, StringComparison.Ordinal)) ?? columna;

    // ── Plan 41 (T3, R-S31-1 — D-G): auto-chequeo S26:S31 = (Q26:R26) de la tabla J ─────────────
    //
    // El oficial trae INTERVENTORIA!R26 = <f t="shared" ref="R26:R31">J15</f> (la tabla anual J =
    // los valores de la 2.ª quincena). Q26:Q31 = 'CONSOLIDADO_TOTAL RECAUDO'!L104:L109, cuya cadena
    // arranca en L9:L13 = INTERVENTORIA!M26:M30 (Q2) / N26:N30 (Q1) por R-INTERV-Q. En Q1, Q26 = N26
    // (189185987) ≠ R26 = J15 (189185988) ⇒ S26 FALSE real. La REGLA honesta (mini-T0 T3): la tabla
    // J es el valor Q2; el valor Q1 vive en N26:N30 — Phase 0 reescribe la <f> del maestro por
    // quincena (Q2 → J15 identidad; Q1 → N26) para que el auto-chequeo compare contra la columna
    // de la quincena del run. La columna la decide el DOMINIO (NumeroQuincena), nunca el contenido.

    /// <summary>
    /// Plan 41 (T3): celda maestra del auto-chequeo <c>S26:S31</c> (shared ref <c>R26:R31</c>) y su
    /// texto-<c>&lt;f&gt;</c> por quincena. Q2 → <c>J15</c> (identidad del oficial); Q1 → <c>N26</c>
    /// (misma forma relativa: R26:R31 = N26:N31).
    /// </summary>
    private static readonly IReadOnlyDictionary<int, string> TablaJMaestroPorQuincena = new Dictionary<int, string>
    {
        [1] = "N26",
        [2] = "J15",
    };

    /// <summary>
    /// Plan 41 (T3): <c>&lt;f&gt;</c> del maestro <c>INTERVENTORIA!R26</c> esperada para la quincena
    /// del run. Fuente ÚNICA (la consumen la autoverificación de Phase 0 y el preflight
    /// <see cref="ValidadorBasePeriodo"/>).
    /// </summary>
    public static string TablaJMaestroDe(Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        if (!TablaJMaestroPorQuincena.TryGetValue(periodo.NumeroQuincena, out var maestro))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"T3 (tabla J): quincena no soportada ({periodo.NumeroQuincena}); solo 1 (Primera) o 2 (Segunda).");
        }

        return maestro;
    }

    /// <summary>Celda maestra del auto-chequeo de la tabla J (contrato de allow-list de Phase 0).</summary>
    public const string CeldaTablaJMaestro = "R26";

    /// <summary>
    /// Plan 41 (T3): reescribe <c>INTERVENTORIA!R26</c> (maestro shared <c>R26:R31</c>) a la variante
    /// de la quincena del run (Q2 → <c>J15</c> identidad; Q1 → <c>N26</c>). Cualquier tercera forma =
    /// FAIL (<c>ERR-PLANTILLA</c>). Si ya trae la variante esperada, NO se toca (paridad
    /// byte-preserving).
    /// </summary>
    private static string AplicarVarianteTablaJ(string xml, Periodo periodo, List<string> resumen)
    {
        var esperado = TablaJMaestroDe(periodo);
        var contraria = periodo.NumeroQuincena == 1 ? "J15" : "N26";
        xml = AsegurarVarianteTablaJ(xml, CeldaTablaJMaestro, esperado, contraria, periodo, out var cambiado);
        resumen.Add(cambiado
            ? $"  INTERVENTORIA!{CeldaTablaJMaestro}: tabla J = {esperado} (Q{periodo.NumeroQuincena}, T3)"
            : $"  INTERVENTORIA!{CeldaTablaJMaestro}: tabla J ya = {esperado} (Q{periodo.NumeroQuincena})");
        return xml;
    }

    private static string AsegurarVarianteTablaJ(
        string xml,
        string celda,
        string esperado,
        string contraria,
        Periodo periodo,
        out bool cambiado)
    {
        var match = BuscarCelda(xml, celda)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"T3 (tabla J): falta la celda INTERVENTORIA!{celda} en la base del período {periodo.CodigoCompleto}.");

        var (prefijo, _, inner, autocierre, tieneFormula) = AnalizarCelda(match);
        if (!tieneFormula || autocierre || inner is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"T3 (tabla J): INTERVENTORIA!{celda} no es una <f> shared en la base del período {periodo.CodigoCompleto} (forma inesperada).");
        }

        var formulaActual = FormulaRegex.Match(inner).Groups["texto"].Value;
        if (string.Equals(formulaActual, esperado, StringComparison.Ordinal))
        {
            cambiado = false;
            return xml; // no-op byte-preserving (paridad Q2 sobre el oficial).
        }

        if (!string.Equals(formulaActual, contraria, StringComparison.Ordinal))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"T3 (tabla J): INTERVENTORIA!{celda} trae la forma inesperada '{formulaActual}' " +
                $"(se esperaba '{contraria}' o '{esperado}'); fail-fast, nunca se sobrescribe a ciegas.");
        }

        var nuevoInner = SustituirTextoFormula(
            inner, prefijo, esperado, WorkbookLeafCellMapAjustesSfT.HojaInterventoria, celda);
        var nuevaCelda = $"<{prefijo}c r=\"{celda}\"{match.Groups["attrs"].Value}>{nuevoInner}</{prefijo}c>";
        cambiado = true;
        return xml[..match.Index] + nuevaCelda + xml[(match.Index + match.Length)..];
    }

    /// <summary>Autoverificación de la tabla J (T3): <c>INTERVENTORIA!R26</c> == la variante del run.</summary>
    private static IEnumerable<string> VerificarTablaJ(byte[] hojaInterventoria, Periodo periodo)
    {
        var esperado = TablaJMaestroDe(periodo);
        var xml = Encoding.UTF8.GetString(hojaInterventoria);
        var match = BuscarCelda(xml, CeldaTablaJMaestro)
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"Autoverificación T3: falta INTERVENTORIA!{CeldaTablaJMaestro}.");
        var (_, _, inner, _, tieneFormula) = AnalizarCelda(match);
        Exigir(tieneFormula && inner is not null, $"Autoverificación T3: INTERVENTORIA!{CeldaTablaJMaestro} no es <f>.");
        var formula = FormulaRegex.Match(inner!).Groups["texto"].Value;
        Exigir(string.Equals(formula, esperado, StringComparison.Ordinal),
            $"Autoverificación T3: INTERVENTORIA!{CeldaTablaJMaestro} = '{formula}' != '{esperado}' (Q{periodo.NumeroQuincena}).");
        yield return $"INTERVENTORIA!{CeldaTablaJMaestro} == {esperado} (Q{periodo.NumeroQuincena}, T3): OK";
    }

    /// <summary>
    /// Plan 40 (T3, F-NEW-1): <c>&lt;f&gt;</c> de <c>INTERVENTORIA</c> esperadas de cada auto-chequeo
    /// por quincena. Fuente ÚNICA de la regla (la consumen la autoverificación de Phase 0 y el
    /// preflight <see cref="ValidadorBasePeriodo"/>).
    /// </summary>
    public static IReadOnlyList<(string Hoja, string Celda, string Fragmento)> FragmentosAutoChequeoEsperados(Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        if (!ColumnaLPorQuincena.TryGetValue(periodo.NumeroQuincena, out var columna))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"F-NEW-1: quincena no soportada ({periodo.NumeroQuincena}); solo 1 (Primera) o 2 (Segunda).");
        }

        var hojaDetRetri = NombresHojaPeriodo.DetRetri(periodo);
        var hojaDetValiRetri = NombresHojaPeriodo.DetValiRetri(periodo);
        var lista = new List<(string, string, string)>();
        foreach (var (celda, fila) in CeldasAutoChequeoDetRetriL)
        {
            lista.Add((hojaDetRetri, celda, $"INTERVENTORIA!{columna}{fila}"));
        }

        foreach (var (celda, fila) in CeldasAutoChequeoDetRetriO)
        {
            lista.Add((hojaDetRetri, celda, $"INTERVENTORIA!{columna}{fila}"));
        }

        foreach (var (celda, fila) in CeldasAutoChequeoDetValiRetriI)
        {
            lista.Add((hojaDetValiRetri, celda, $"INTERVENTORIA!{columna}{fila}"));
        }

        return lista;
    }

    /// <summary>Plan 40 (T3, E1-4): celdas de <c>VALIDACION_TOTAL</c> que suman el bloque <c>'Recaudo *'</c> (celda → fila).</summary>
    public static IReadOnlyList<(string Celda, int FilaRecaudo)> CeldasRecaudoDeValidacionTotal => CeldasRecaudoValidacion;

    /// <summary>
    /// Plan 62 (T4.1, D-H4): celdas de <c>CONSOLIDADO_TOTAL RECAUDO!E142:E146</c> que suman el bloque
    /// <c>'Recaudo *'</c> (celda → fila 21:25). Fuente ÚNICA de la geometría para el preflight
    /// <see cref="ValidadorBasePeriodo"/> (la columna también sale de <see cref="ColumnaRecaudoDe"/>).
    /// </summary>
    public static IReadOnlyList<(string Celda, int FilaRecaudo)> CeldasRecaudoDeConsolidado142 => CeldasRecaudoValidacionConsolidado142;

    // ── Autoverificación (e) ─────────────────────────────────────────────────────────────────────
    /// <summary>
    /// Autoverificación embebida: 0 tokens del período base; las hojas del período destino presentes
    /// en workbook.xml y docProps/app.xml; mismas <c>&lt;f&gt;</c> del período que en la entrada;
    /// rIds/sheetIds intactos; N3 sellado; <b>variante L9:L13 == la de la quincena del run</b>.
    /// </summary>
    private static List<string> VerificarBase(
        IReadOnlyDictionary<string, byte[]> partes,
        string workbookOriginal,
        string basePeriodo,
        string periodoDestino,
        string hojaConsolidadoParte,
        IReadOnlyList<(string Hoja, string Parte)> hojasRecaudoValidacion,
        string hojaDetRetriParte,
        string hojaDetValiRetriParte,
        string hojaInterventoriaParte,
        int formulasPeriodoBase,
        Periodo periodo)
    {
        var lineas = new List<string>();
        var informeBase = $"{basePeriodo[..6]}-{basePeriodo[6]}";

        if (string.Equals(basePeriodo, periodoDestino, StringComparison.Ordinal))
        {
            lineas.Add($"identidad de período ({periodoDestino} == {basePeriodo}): los tokens del período se conservan tal cual");
        }
        else
        {
            var partesConToken = new List<string>();
            foreach (var par in partes)
            {
                if (!HojaRegex.IsMatch(par.Key)
                    && !string.Equals(par.Key, PartWorkbook, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(par.Key, PartApp, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(par.Key, PartShared, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var texto = Encoding.UTF8.GetString(par.Value);
                var tokenSuelto = Regex.IsMatch(texto, $"(?<![0-9]){basePeriodo}(?![0-9])");
                if (tokenSuelto || texto.Contains(informeBase, StringComparison.Ordinal))
                {
                    partesConToken.Add(par.Key);
                }
            }

            Exigir(partesConToken.Count == 0, $"Quedaron apariciones del período base en: {string.Join(", ", partesConToken)}");
            lineas.Add($"0 apariciones de '{basePeriodo}' en worksheets + workbook + app + sharedStrings: OK");
        }

        var formulasDestino = 0;
        foreach (var par in partes)
        {
            if (!HojaRegex.IsMatch(par.Key))
            {
                continue;
            }

            formulasDestino += ContarFormulasConPeriodo(Encoding.UTF8.GetString(par.Value), periodoDestino);
        }

        Exigir(formulasDestino == formulasPeriodoBase,
            $"Las <f> del período cambiaron de {formulasPeriodoBase} a {formulasDestino} (el reemplazo no debe crear ni destruir fórmulas).");
        Exigir(formulasDestino > 0,
            "La base no conserva ninguna <f> del período; el token de período no se reescribió donde debía.");
        lineas.Add($"{formulasDestino} <f> del período {periodoDestino} (== entrada): OK");

        var workbook = Encoding.UTF8.GetString(partes[PartWorkbook]);
        var app = Encoding.UTF8.GetString(partes[PartApp]);
        var nombresDestino = new[]
        {
            NombresHojaPeriodo.DetRetri(periodoDestino),
            NombresHojaPeriodo.DetValiRetri(periodoDestino),
            NombresHojaPeriodo.InformeAFaseo(periodoDestino[..6], int.Parse(periodoDestino[6..], CultureInfo.InvariantCulture)),
        };

        foreach (var nombre in nombresDestino)
        {
            Exigir(workbook.Contains(nombre, StringComparison.Ordinal), $"workbook.xml no contiene '{nombre}'.");
            Exigir(app.Contains(nombre, StringComparison.Ordinal), $"docProps/app.xml no contiene '{nombre}'.");
        }

        lineas.Add($"hojas del período {periodoDestino} presentes en workbook.xml y app.xml: OK");

        var hojasViejas = ExtraerHojas(workbookOriginal);
        var hojasNuevas = ExtraerHojas(workbook);
        Exigir(hojasViejas.Count == hojasNuevas.Count, "Cambió el número de hojas del workbook.");
        for (var i = 0; i < hojasViejas.Count; i++)
        {
            Exigir(string.Equals(hojasViejas[i].SheetId, hojasNuevas[i].SheetId, StringComparison.Ordinal)
                && string.Equals(hojasViejas[i].Rid, hojasNuevas[i].Rid, StringComparison.Ordinal),
                $"Cambió sheetId/rId en la posición {i}: {hojasViejas[i].SheetId}/{hojasViejas[i].Rid} vs {hojasNuevas[i].SheetId}/{hojasNuevas[i].Rid}");
        }

        lineas.Add("rIds/sheetIds intactos: OK");

        var consolidado = ParsearCeldas(
            Encoding.UTF8.GetString(partes[hojaConsolidadoParte]),
            ResolverSharedStrings(Encoding.UTF8.GetString(partes[PartShared])));
        Exigir(consolidado.TryGetValue("N3", out var n3), $"La base producida no tiene {HojaConsolidado}!N3.");
        Exigir(string.Equals(n3!.Texto, periodoDestino, StringComparison.Ordinal),
            $"{HojaConsolidado}!N3 = '{n3.Texto}' != período '{periodoDestino}'.");
        Exigir(!n3.TieneFormula, $"{HojaConsolidado}!N3 es fórmula (el sello debe ser un valor).");
        lineas.Add($"{HojaConsolidado}!N3 == {periodoDestino}: OK");
        Exigir(consolidado.ContainsKey("G7") && consolidado.ContainsKey("K7") && consolidado.ContainsKey("D6"),
            $"La base producida no trae los sellos base {HojaConsolidado}!G7/K7/D6.");
        Exigir(!consolidado["G7"].TieneFormula && !consolidado["K7"].TieneFormula && !consolidado["D6"].TieneFormula,
            $"Los sellos base {HojaConsolidado}!G7/K7/D6 no deben ser fórmula.");
        lineas.Add($"{HojaConsolidado}!G7/K7/D6 presentes y sin fórmula: OK");

        lineas.AddRange(VerificarVarianteL(partes[hojaConsolidadoParte], periodo));
        foreach (var (hojaRecaudo, parteRecaudo) in hojasRecaudoValidacion)
        {
            lineas.AddRange(VerificarVarianteRecaudoValidacion(partes[parteRecaudo], hojaRecaudo, CeldasRecaudoValidacion, periodo));
        }

        // Plan 62 (T4.1): la extensión E1-4 al bloque 142–147 se autoverifica con la MISMA rutina
        // (tabla de celdas + columna de la quincena del run).
        lineas.AddRange(VerificarVarianteRecaudoValidacion(
            partes[hojaConsolidadoParte], HojaConsolidado, CeldasRecaudoValidacionConsolidado142, periodo));

        lineas.AddRange(VerificarAutoChequeos(partes[hojaDetRetriParte], CeldasAutoChequeoDetRetriL, CeldasAutoChequeoDetRetriO, periodo));
        lineas.AddRange(VerificarAutoChequeos(partes[hojaDetValiRetriParte], CeldasAutoChequeoDetValiRetriI, [], periodo));
        lineas.AddRange(VerificarTablaJ(partes[hojaInterventoriaParte], periodo));

        return lineas;
    }

    /// <summary>
    /// Autoverificación de las celdas que suman <c>'Recaudo *'</c> (E1-4 y su extensión 142–147 del
    /// Plan 62) en la columna de la quincena. La tabla de celdas es un parámetro: la rutina es la
    /// MISMA para <c>VALIDACION_*!C3:C8</c> y para <c>CONSOLIDADO!E142:E146</c> (cero duplicación).
    /// </summary>
    private static IEnumerable<string> VerificarVarianteRecaudoValidacion(
        byte[] hoja,
        string nombreHoja,
        IReadOnlyList<(string Celda, int FilaRecaudo)> celdas,
        Periodo periodo)
    {
        var columna = ColumnaRecaudoDe(periodo);
        var xml = Encoding.UTF8.GetString(hoja);
        var hojas = EmpresaFacturacion.Catalogo.Select(e => e.HojaRecaudo).ToHashSet(StringComparer.Ordinal);

        foreach (var (celda, fila) in celdas)
        {
            var match = BuscarCelda(xml, celda)
                ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"Autoverificación E1-4: falta {nombreHoja}!{celda}.");
            var (_, _, inner, _, tieneFormula) = AnalizarCelda(match);
            Exigir(tieneFormula && inner is not null, $"Autoverificación E1-4: {nombreHoja}!{celda} no es <f>.");

            var formula = FormulaRegex.Match(inner!).Groups["texto"].Value;
            foreach (var termino in SepararTerminosSuma(formula))
            {
                var m = Regex.Match(termino, "^'(?<hoja>[^']+)'!(?<col>[A-Z]+)(?<fila>\\d+)$");
                Exigir(m.Success && hojas.Contains(m.Groups["hoja"].Value) && m.Groups["fila"].Value == fila.ToString(CultureInfo.InvariantCulture),
                    $"Autoverificación E1-4: {nombreHoja}!{celda} término '{termino}' fuera de forma.");
                Exigir(string.Equals(m.Groups["col"].Value, columna, StringComparison.Ordinal),
                    $"Autoverificación E1-4: {nombreHoja}!{celda} columna '{m.Groups["col"].Value}' != '{columna}' (Q{periodo.NumeroQuincena}).");
            }

            yield return $"{nombreHoja}!{celda} == Σ 'Recaudo *'!{columna}{fila} (Q{periodo.NumeroQuincena}, E1-4): OK";
        }
    }

    /// <summary>Autoverificación de los auto-chequeos F-NEW-1: fragmento <c>INTERVENTORIA!{col}{fila}</c> por quincena.</summary>
    private static IEnumerable<string> VerificarAutoChequeos(
        byte[] hoja,
        (string Celda, int FilaInterventoria)[] grupo1,
        (string Celda, int FilaInterventoria)[] grupo2,
        Periodo periodo)
    {
        var columna = ColumnaLPorQuincena[periodo.NumeroQuincena];
        var xml = Encoding.UTF8.GetString(hoja);
        foreach (var (celda, fila) in grupo1.Concat(grupo2))
        {
            var match = BuscarCelda(xml, celda)
                ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"Autoverificación F-NEW-1: falta la celda {celda}.");
            var (_, _, inner, _, tieneFormula) = AnalizarCelda(match);
            Exigir(tieneFormula && inner is not null, $"Autoverificación F-NEW-1: {celda} no es <f>.");

            var formula = FormulaRegex.Match(inner!).Groups["texto"].Value;
            var esperado = $"INTERVENTORIA!{columna}{fila}";
            Exigir(formula.Contains(esperado, StringComparison.Ordinal),
                $"Autoverificación F-NEW-1: {celda} = '{formula}' no contiene '{esperado}' (Q{periodo.NumeroQuincena}).");
            yield return $"{celda} ⊇ {esperado} (Q{periodo.NumeroQuincena}, F-NEW-1): OK";
        }
    }

    /// <summary>Autoverificación de la variante L (R-INTERV-Q): L9:L13 == la del run.</summary>
    private static IEnumerable<string> VerificarVarianteL(byte[] hojaConsolidado, Periodo periodo)
    {
        var columna = ColumnaLPorQuincena[periodo.NumeroQuincena];
        // La mitad declarada la decide el DOMINIO en un solo lugar (primera 'N' / segunda 'M').
        var valores = InterventoriaDeclarada.PorQuincena(periodo.NumeroQuincena);
        var xml = Encoding.UTF8.GetString(hojaConsolidado);

        foreach (var (celda, filaInterventoria) in CeldasL9L13)
        {
            var aseId = filaInterventoria - InterventoriaDeclarada.FilaPrimerAse + 1;
            var formulaEsperada = $"INTERVENTORIA!{columna}{filaInterventoria}";
            var valorEsperado = valores[aseId].ToString(CultureInfo.InvariantCulture);

            var match = BuscarCelda(xml, celda)
                ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"Autoverificación L: falta {HojaConsolidado}!{celda}.");
            var (_, _, inner, _, tieneFormula) = AnalizarCelda(match);
            Exigir(tieneFormula && inner is not null, $"Autoverificación L: {HojaConsolidado}!{celda} no es <f>.");

            var formula = FormulaRegex.Match(inner!).Groups["texto"].Value;
            var valor = ValorRegex.Match(inner!).Groups["v"].Value;
            Exigir(string.Equals(formula, formulaEsperada, StringComparison.Ordinal),
                $"Autoverificación L: {HojaConsolidado}!{celda} = '{formula}' != '{formulaEsperada}' (R-INTERV-Q).");
            Exigir(string.Equals(valor, valorEsperado, StringComparison.Ordinal),
                $"Autoverificación L: {HojaConsolidado}!{celda} caché '{valor}' != '{valorEsperado}' (R-INTERV-Q).");

            yield return $"{HojaConsolidado}!{celda} == {formulaEsperada} (Q{periodo.NumeroQuincena}, R-INTERV-Q): OK";
        }
    }

    // ── Detección / resolución ───────────────────────────────────────────────────────────────────

    /// <summary>Período interno de la plantilla: el 7 dígitos de la hoja <c>DetRetri{AAAAMMQ}</c>; <c>null</c> si no la trae.</summary>
    private static string? DetectarPeriodoBase(string workbookXml)
    {
        foreach (Match hoja in Regex.Matches(workbookXml, $"<{P}sheet name=\"(?<n>[^\"]*)\""))
        {
            var match = HojaDetRetriBase.Match(hoja.Groups["n"].Value);
            if (match.Success)
            {
                return match.Groups["periodo"].Value;
            }
        }

        return null;
    }

    /// <summary>Parte del ZIP de una hoja resuelta por NOMBRE (workbook.xml + rels), nunca por número.</summary>
    private static string ResolverParteDeHoja(ZipArchive zip, string workbookXml, string nombreHoja)
    {
        var sheet = Regex.Match(
            workbookXml,
            $"<{P}sheet name=\"{Regex.Escape(nombreHoja)}\"[^>]*r:id=\"(?<rid>[^\"]+)\"");
        if (!sheet.Success)
        {
            sheet = Regex.Match(
                workbookXml,
                $"<{P}sheet name=\"{Regex.Escape(nombreHoja)}\"[^>]*id=\"(?<rid>[^\"]+)\"");
        }

        Exigir(sheet.Success, $"La plantilla no trae la hoja '{nombreHoja}'.");

        var rid = sheet.Groups["rid"].Value;
        var rels = LeerTexto(zip.GetEntry("xl/_rels/workbook.xml.rels")
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "La plantilla no trae xl/_rels/workbook.xml.rels."));
        var rel = Regex.Match(rels, $"<Relationship[^>]*Id=\"{Regex.Escape(rid)}\"[^>]*Target=\"(?<target>[^\"]+)\"");
        if (!rel.Success)
        {
            rel = Regex.Match(rels, $"<Relationship[^>]*Target=\"(?<target>[^\"]+)\"[^>]*Id=\"{Regex.Escape(rid)}\"");
        }

        Exigir(rel.Success, $"No se resolvió la relación '{rid}' de la hoja '{nombreHoja}'.");
        var target = rel.Groups["target"].Value;
        return target.StartsWith('/') ? target[1..] : "xl/" + target;
    }

    /// <summary>Nº de <c>&lt;f&gt;</c> cuyo texto referencia una hoja <c>DetRetri/DetValiRetri</c> del período indicado.</summary>
    private static int ContarFormulasConPeriodo(string xml, string periodo)
    {
        var total = 0;
        foreach (Match match in FormulaRegex.Matches(xml))
        {
            foreach (Match referencia in ReferenciaHojaDet.Matches(match.Groups["texto"].Value))
            {
                var detectado = referencia.Groups["p1"].Success ? referencia.Groups["p1"].Value : referencia.Groups["p2"].Value;
                if (string.Equals(detectado, periodo, StringComparison.Ordinal))
                {
                    total++;
                }
            }
        }

        return total;
    }

    // ── Cirugía de celdas (string-safe, preserva todo el resto del XML) ─────────────────────────

    private static string SetValorNumerico(string xml, string referencia, string valor, out bool aplicado, out string motivo)
    {
        var match = BuscarCelda(xml, referencia);
        if (match is null)
        {
            aplicado = false;
            motivo = "celda inexistente en la base";
            return xml;
        }

        var (prefijo, attrs, inner, autocierre, tieneFormula) = AnalizarCelda(match);
        if (tieneFormula)
        {
            aplicado = false;
            motivo = "destino con fórmula (NUNCA se sobrescribe)";
            return xml;
        }

        attrs = QuitarAtributoTipo(attrs);
        string nueva;
        if (autocierre)
        {
            nueva = $"<{prefijo}c r=\"{referencia}\"{attrs}><{prefijo}v>{valor}</{prefijo}v></{prefijo}c>";
        }
        else
        {
            var cuerpo = inner ?? string.Empty;
            cuerpo = ValorRegex.IsMatch(cuerpo)
                ? ValorRegex.Replace(cuerpo, $"<{prefijo}v>{valor}</{prefijo}v>")
                : cuerpo + $"<{prefijo}v>{valor}</{prefijo}v>";
            nueva = $"<{prefijo}c r=\"{referencia}\"{attrs}>{cuerpo}</{prefijo}c>";
        }

        aplicado = true;
        motivo = "ok";
        return xml[..match.Index] + nueva + xml[(match.Index + match.Length)..];
    }

    private static Match? BuscarCelda(string xml, string referencia)
    {
        foreach (Match match in CeldaRegex.Matches(xml))
        {
            if (string.Equals(match.Groups["ref"].Value, referencia, StringComparison.Ordinal))
            {
                return match;
            }
        }

        return null;
    }

    private static (string Prefijo, string Attrs, string? Inner, bool Autocierre, bool TieneFormula) AnalizarCelda(Match match)
    {
        var prefijo = match.Groups["p"].Value;
        var attrs = match.Groups["attrs"].Value;
        var autocierre = !match.Groups["inner"].Success;
        var inner = match.Groups["inner"].Success ? match.Groups["inner"].Value : null;
        var tieneFormula = inner is not null && Regex.IsMatch(inner, $"<{P}f[ />]");
        return (prefijo, attrs, inner, autocierre, tieneFormula);
    }

    private static string QuitarAtributoTipo(string attrs) =>
        Regex.Replace(attrs, "\\s+t=\"[^\"]*\"", string.Empty, RegexOptions.Compiled);

    // ── Parseo de celdas ─────────────────────────────────────────────────────────────────────────

    private sealed record InfoCelda(bool TieneFormula, bool EsTexto, string? Valor, string? Texto);

    private static Dictionary<string, InfoCelda> ParsearCeldas(string xml, IReadOnlyList<string> sharedStrings)
    {
        var celdas = new Dictionary<string, InfoCelda>(StringComparer.Ordinal);
        foreach (Match match in CeldaRegex.Matches(xml))
        {
            var referencia = match.Groups["ref"].Value;
            var (_, attrs, inner, _, tieneFormula) = AnalizarCelda(match);
            string? valor = null;
            var esTexto = attrs.Contains("t=\"s\"", StringComparison.Ordinal)
                || attrs.Contains("t=\"inlineStr\"", StringComparison.Ordinal)
                || attrs.Contains("t=\"str\"", StringComparison.Ordinal);
            string? texto = null;

            if (inner is not null)
            {
                var valorMatch = ValorRegex.Match(inner);
                if (valorMatch.Success)
                {
                    valor = valorMatch.Groups["v"].Value;
                }

                if (attrs.Contains("t=\"s\"", StringComparison.Ordinal)
                    && int.TryParse(valor, NumberStyles.None, CultureInfo.InvariantCulture, out var indice)
                    && indice >= 0 && indice < sharedStrings.Count)
                {
                    texto = sharedStrings[indice];
                }
                else if (attrs.Contains("t=\"inlineStr\"", StringComparison.Ordinal))
                {
                    texto = string.Concat(TextoRegex.Matches(inner).Select(m => m.Groups["t"].Value));
                }
            }

            texto ??= valor;
            celdas[referencia] = new InfoCelda(tieneFormula, esTexto, valor, texto);
        }

        return celdas;
    }

    private static List<string> ResolverSharedStrings(string xml)
    {
        var lista = new List<string>();
        foreach (Match match in SiRegex.Matches(xml))
        {
            lista.Add(string.Concat(TextoRegex.Matches(match.Groups["cuerpo"].Value).Select(m => m.Groups["t"].Value)));
        }

        return lista;
    }

    // ── Escritura atómica ────────────────────────────────────────────────────────────────────────

    private static byte[] SerializarZip(
        IReadOnlyList<string> orden,
        IReadOnlyDictionary<string, byte[]> partes)
    {
        using var memoria = new MemoryStream();
        using (var salidaZip = new ZipArchive(memoria, ZipArchiveMode.Create, leaveOpen: true))
        {
            // Se recorre el orden original para preservar [Content_Types].xml como primera entrada.
            foreach (var nombre in orden)
            {
                var nueva = salidaZip.CreateEntry(nombre, CompressionLevel.Optimal);
                using var stream = nueva.Open();
                stream.Write(partes[nombre]);
            }
        }

        return memoria.ToArray();
    }

    private static void EscribirBytesAtomico(byte[] contenido, string salidaPath)
    {
        var tmpPath = salidaPath + ".tmp";
        var tmpDir = Path.GetDirectoryName(salidaPath);
        if (!string.IsNullOrEmpty(tmpDir))
        {
            Directory.CreateDirectory(tmpDir);
        }

        if (File.Exists(tmpPath))
        {
            File.Delete(tmpPath);
        }

        File.WriteAllBytes(tmpPath, contenido);

        using (var relectura = ZipFile.OpenRead(tmpPath))
        {
            Exigir(relectura.Entries.Count > 0, "El zip producido no tiene entradas.");
            foreach (var nombre in new[] { PartWorkbook, PartApp, PartShared })
            {
                Exigir(relectura.GetEntry(nombre) is not null, $"El zip producido no contiene {nombre}.");
            }
        }

        if (File.Exists(salidaPath))
        {
            File.Delete(salidaPath);
        }

        File.Move(tmpPath, salidaPath);
    }

    // ── Utilidades ───────────────────────────────────────────────────────────────────────────────

    private sealed record HojaInfo(string Name, string SheetId, string Rid);

    private static List<HojaInfo> ExtraerHojas(string workbookXml)
    {
        var lista = new List<HojaInfo>();
        foreach (Match match in Regex.Matches(workbookXml, $"<{P}sheet name=\"(?<n>[^\"]*)\" sheetId=\"(?<s>\\d+)\"[^>]*r:id=\"(?<r>[^\"]+)\""))
        {
            lista.Add(new HojaInfo(match.Groups["n"].Value, match.Groups["s"].Value, match.Groups["r"].Value));
        }

        return lista;
    }

    private static bool EsParteTextoTransformable(string nombre) =>
        string.Equals(nombre, PartWorkbook, StringComparison.OrdinalIgnoreCase)
        || string.Equals(nombre, PartApp, StringComparison.OrdinalIgnoreCase)
        || string.Equals(nombre, PartShared, StringComparison.OrdinalIgnoreCase)
        || HojaRegex.IsMatch(nombre);

    private static byte[] LeerBytes(ZipArchiveEntry entrada)
    {
        using var stream = entrada.Open();
        using var memoria = new MemoryStream();
        stream.CopyTo(memoria);
        return memoria.ToArray();
    }

    private static string LeerTexto(ZipArchiveEntry entrada) => Encoding.UTF8.GetString(LeerBytes(entrada));

    private static void Exigir(bool condicion, string mensaje)
    {
        if (!condicion)
        {
            throw new CalculoInvalidoException(CodigoError.Plantilla, mensaje);
        }
    }
}
