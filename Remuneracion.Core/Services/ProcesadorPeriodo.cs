using Serilog;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Services;

/// <summary>
/// Orquestador del caso de uso de remuneración de un período completo (los 5 ASE).
/// HU-07: reutiliza el path single-ASE por composición de servicios (lectura, cálculo, leaf),
/// valida multi-ASE y escribe UNA sola vez el workbook de salida (atomicidad).
/// Fail-fast: ante cualquier fallo (carpeta/fuente faltante, validación, escritura) no hay
/// salida certificada.
/// HU-14 (3.1 + 3.2): cada fail-fast porta un <see cref="CodigoError"/> del catálogo (D1);
/// RunId por ejecución vía <c>Serilog.Context.LogContext</c> correlaciona todos los eventos
/// del procesador (D5) y cada paso se registra con nivel y propiedades estructuradas (CA-6).
/// HU-14 (S-3): la creación de <see cref="Ase"/> delega a <see cref="AseFactory.DesdeId"/>.
/// </summary>
public sealed class ProcesadorPeriodo : IProcesadorPeriodo
{
    private readonly IRecaudoReader _recaudoReader;
    private readonly IWorkbookLeafInputReader _leafReader;
    private readonly ICalculoRemuneracion _calculoRemuneracion;
    private readonly IValidador _validador;
    private readonly IWorkbookLeafWriter _workbookLeafWriter;
    private readonly ILocalizadorArchivosAse _localizador;
    private readonly IDetRetriR10Reader _detRetriR10Reader;
    private readonly IValidacionOracleReader? _validacionOracleReader;

    /// <summary>
    /// Plan 42 (T2, D-E): emisor del reporte durable de divergencias declaradas. Opcional para no
    /// romper la composición existente; si una corrida reconoce una declarada y NO hay emisor, el
    /// cierre FALLA (sin reporte no hay ÉXITO — nunca un cierre silencioso).
    /// </summary>
    private readonly IReporteDivergenciasDeclaradas? _reporteDivergencias;

    /// <summary>
    /// Plan 61 (H3.2/H4.1, §2-D-H4): emisor del REPORTE DE NOVEDADES de la corrida. Nullable para no
    /// romper las composiciones de regresión, pero fail-closed: si una corrida produce novedades y no
    /// hay emisor, el procesador FALLA (sin reporte no hay entrega — nunca un cierre silencioso).
    /// </summary>
    private readonly IReporteNovedades? _reporteNovedades;

    /// <summary>
    /// Plan 26 (T2, D-A): preflight de insumos del período. Se compone con el localizador ya
    /// inyectado (mismo contrato que usan los finders runtime) para no ampliar la firma de
    /// composición de UI/CLI.
    /// </summary>
    private readonly ValidadorInsumosPeriodo _validadorInsumos;

    /// <summary>
    /// Composición del caso de uso de período. El oráculo R10 (<see cref="IDetRetriR10Reader"/>)
    /// es una dependencia OBLIGATORIA (HU-20/G3): parte del flujo normal de AMBAS quincenas —
    /// el DetRetri calculado bottom-up se contrasta contra el R10 del período con tolerancia ±0.5
    /// post-redondeo (G3-D1). Solo el oráculo de validaciones cruzadas (<see cref="IValidacionOracleReader"/>)
    /// sigue siendo opcional (compatibilidad de la regresión HU-13 D3).
    /// </summary>
    public ProcesadorPeriodo(
        IRecaudoReader recaudoReader,
        IWorkbookLeafInputReader leafReader,
        ICalculoRemuneracion calculoRemuneracion,
        IValidador validador,
        IWorkbookLeafWriter workbookLeafWriter,
        ILocalizadorArchivosAse localizador,
        IDetRetriR10Reader detRetriR10Reader,
        IValidacionOracleReader? validacionOracleReader = null,
        IReporteDivergenciasDeclaradas? reporteDivergencias = null,
        IReporteNovedades? reporteNovedades = null)
    {
        _recaudoReader = recaudoReader ?? throw new ArgumentNullException(nameof(recaudoReader));
        _leafReader = leafReader ?? throw new ArgumentNullException(nameof(leafReader));
        _calculoRemuneracion = calculoRemuneracion ?? throw new ArgumentNullException(nameof(calculoRemuneracion));
        _validador = validador ?? throw new ArgumentNullException(nameof(validador));
        _workbookLeafWriter = workbookLeafWriter ?? throw new ArgumentNullException(nameof(workbookLeafWriter));
        _localizador = localizador ?? throw new ArgumentNullException(nameof(localizador));
        _detRetriR10Reader = detRetriR10Reader ?? throw new ArgumentNullException(nameof(detRetriR10Reader));
        _validacionOracleReader = validacionOracleReader;
        _reporteDivergencias = reporteDivergencias;
        _reporteNovedades = reporteNovedades;
        _validadorInsumos = new ValidadorInsumosPeriodo(_localizador);
    }

    public ResultadoProcesoPeriodo Ejecutar(SolicitudProcesoPeriodo solicitud, IProgress<string>? progreso = null)
    {
        ArgumentNullException.ThrowIfNull(solicitud);

        // Plan 55 (H1.1/D4): la PLANTILLA es la autoridad de los destinos de la conciliación por
        // empresa (resolución por cruce de etiquetas). Se entrega al reader UNA vez por corrida.
        // Plan 59 (H3): la lectura de la conciliación por empresa corre en ambas quincenas; no hay
        // recorte por quincena que preservar.
        _leafReader.RutaPlantilla = solicitud.RutaPlantilla;

        // HU-14 (3.2, D5): RunId por ejecución correlaciona todos los eventos de ESTE procesador
        // (ambos modos). LogContext es Serilog core; las propiedades Periodo/Quincena/Modo
        // acompañan al RunId para filtrar el log (CA-6).
        // HU-15 (W-2.1, D4): si el frontend (UI/CLI) inyectó un RunId en la solicitud, se respeta
        // (un solo Guid correlaciona UI → procesador → writer); null = genera uno (compat HU-14).
        var runId = solicitud.RunId ?? Guid.NewGuid();
        using var _runIdScope = Serilog.Context.LogContext.PushProperty("RunId", runId);
        using var _periodoScope = Serilog.Context.LogContext.PushProperty("Periodo", solicitud.Periodo.CodigoCompleto);
        using var _quincenaScope = Serilog.Context.LogContext.PushProperty("Quincena", solicitud.Periodo.NumeroQuincena);
        using var _modoScope = Serilog.Context.LogContext.PushProperty("Modo", "5 ASE");

        // Plan 43 (O-3): huella de quién acepta y por qué canal, resueltos UNA vez por corrida (se
        // registran en el reporte durable y en el log al cerrar con declarada reconocida).
        var usuarioAceptacion = ResolverUsuario(solicitud.UsuarioAceptacion);
        var canalAceptacion = string.IsNullOrWhiteSpace(solicitud.Canal) ? HuellaAceptacion.CanalIndefinido : solicitud.Canal;

        // Plan 43 (O-4): mientras la corrida no supere las validaciones, el workbook vive con la
        // marca DETERMINISTA "sin validar"; solo al cerrar bien se promueve al nombre DEFINITIVO.
        // Una corrida que falla después de escribir NUNCA deja el nombre limpio.
        var rutaSinValidar = NombresSalida.RutaSinValidar(solicitud.RutaSalida);

        progreso?.Report("Iniciando proceso del período (5 ASE).");
        Log.Information("Iniciando proceso del período (5 ASE).");
        progreso?.Report($"Periodo: {solicitud.Periodo.CodigoCompleto}; carpeta: {solicitud.CarpetaPeriodo}");
        Log.Information("Periodo: {Periodo}; carpeta: {Carpeta}", solicitud.Periodo.CodigoCompleto, solicitud.CarpetaPeriodo);

        if (string.IsNullOrWhiteSpace(solicitud.CarpetaPeriodo))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.FuenteNoEncontrada, "Debe indicarse la carpeta del período que contiene las 5 carpetas de ASE.");
        }

        if (string.IsNullOrWhiteSpace(solicitud.RutaPlantilla) || string.IsNullOrWhiteSpace(solicitud.RutaSalida))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.Plantilla, "Debe indicarse plantilla y ruta de salida.");
        }

        // Resolver las 5 carpetas de ASE en orden 1..5 (fail-fast si falta alguna).
        var carpetasPeriodo = _localizador.ObtenerCarpetasAse(solicitud.CarpetaPeriodo);
        if (carpetasPeriodo.Length == 0)
        {
            throw new ArchivoFuenteNoEncontradoException(
                CodigoError.FuenteNoEncontrada,
                $"No se encontraron carpetas de ASE en '{solicitud.CarpetaPeriodo}'. Deben existir 5 carpetas con prefijos {string.Join(", ", CarpetasAse.Prefijos)}.");
        }

        // Plan 26 (T2, R-F-1/D-A): preflight fail-fast al inicio. Enumera TODOS los insumos
        // faltantes del período reutilizando los MISMOS finders del runtime (D-C) ANTES de abrir
        // ningún workbook. Con faltantes → UN solo error con la lista numerada completa en
        // lenguaje administrativo (D-B/D-E) y el proceso NO inicia (no lee, no calcula, no escribe).
        progreso?.Report("Verificando insumos del período...");
        Log.Information("Verificando insumos del período...");
        var faltantes = _validadorInsumos.Validar(solicitud.CarpetaPeriodo, solicitud.Periodo);
        if (faltantes.Count > 0)
        {
            Log.Warning(
                "Preflight: faltan {CantidadInsumos} insumo(s) del período; no se inicia el procesamiento.",
                faltantes.Count);
            throw new ArchivoFuenteNoEncontradoException(
                CodigoError.FuenteNoEncontrada,
                FormateadorInsumosFaltantes.Mensaje(solicitud.Periodo, faltantes));
        }

        var datos = new List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4)>();
        var datosConAjustes = new List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4, decimal ajustesSfT)>();
        var leafs = new List<WorkbookLeafInputs>();
        var esQuincena2 = solicitud.Periodo.NumeroQuincena == 2;

        for (var idAse = 1; idAse <= 5; idAse++)
        {
            var ase = AseFactory.DesdeId(idAse);
            var carpetaAse = carpetasPeriodo.FirstOrDefault(c =>
                    Path.GetFileName(c).StartsWith(idAse.ToString(), StringComparison.OrdinalIgnoreCase))
                ?? throw new ArchivoFuenteNoEncontradoException(
                    CodigoError.FuenteNoEncontrada,
                    $"No se encontró la carpeta del ASE {idAse} ({CarpetasAse.Prefijos[idAse - 1]}) en '{solicitud.CarpetaPeriodo}'.");

            progreso?.Report($"ASE {idAse} ({ase.NombreCompleto}): localizando fuentes...");
            Log.Information("ASE {AseId} ({Nombre}): localizando fuentes...", idAse, ase.NombreCompleto);
            var rutaR1 = _localizador.BuscarArchivo(carpetaAse, "Recaudoporcomponente")
                ?? throw new ArchivoFuenteNoEncontradoException(CodigoError.FuenteNoEncontrada, $"No se encontró R1 del ASE {idAse} en {carpetaAse}.");
            var rutaR2 = _localizador.BuscarArchivo(carpetaAse, "RerpoteDetalleSaldosaFavor")
                ?? throw new ArchivoFuenteNoEncontradoException(CodigoError.FuenteNoEncontrada, $"No se encontró R2 del ASE {idAse} en {carpetaAse}.");
            var rutaR4 = _localizador.BuscarArchivo(carpetaAse, "ReversiónPorComponente")
                ?? _localizador.BuscarArchivo(carpetaAse, "ReversionPorComponente")
                ?? throw new ArchivoFuenteNoEncontradoException(CodigoError.FuenteNoEncontrada, $"No se encontró R4 del ASE {idAse} en {carpetaAse}.");

            // Plan 64 (H3.1, §2.7/§2.8): el insumo de ANT EXT-REV se localiza con su FINDER DEDICADO
            // (prefijo normalizado "recaudosreversados", mismo directorio y mismo match
            // case-insensitive de los demás insumos por ASE) y se exige aquí como DEFENSA EN
            // PROFUNDIDAD: el preflight del período ya lo enumera antes de leer (misma fuente de
            // verdad, Plan 26 D-C/D-D), pero el runtime no puede depender de que el preflight haya
            // corrido (Plan 26 R-F-6).
            var rutaRecaudosReversados = _localizador.BuscarRecaudosReversados(carpetaAse)
                ?? throw new ArchivoFuenteNoEncontradoException(
                    CodigoError.FuenteNoEncontrada,
                    $"No se encontró RecaudosReversados (insumo de ANT EXT-REV) del ASE {idAse} en {carpetaAse}.");


            progreso?.Report($"ASE {idAse}: leyendo R1, R2 y R4...");
            Log.Information("ASE {AseId}: leyendo R1, R2 y R4...", idAse);
            var r1 = _recaudoReader.LeerR1(rutaR1);
            var r2 = _recaudoReader.LeerR2(rutaR2);
            var r4 = _recaudoReader.LeerR4(rutaR4);

            progreso?.Report($"ASE {idAse}: leyendo inputs leaf del workbook...");
            Log.Information("ASE {AseId}: leyendo inputs leaf del workbook...", idAse);
            var leaf = _leafReader.LeerLeafInputs(ase, solicitud.Periodo, rutaR1, rutaR2, rutaR4);

            // Plan 64 (H2.1, §2.8): ANT EXT-REV se LEE aquí, antes de cualquier cálculo y antes de
            // escribir. El lector hace fail-fast ERR-FORMATO-FUENTE (ASE + archivo + marcador/etiqueta
            // + observado vs esperado) si el insumo no cierra: nunca se calcula ni se escribe con un
            // insumo roto. Las fuentes AUSENTES del catálogo se omiten con traza (no se inventan).
            progreso?.Report($"ASE {idAse}: leyendo el recaudo reversado (ANT EXT-REV)...");
            Log.Information("ASE {AseId}: leyendo el recaudo reversado (ANT EXT-REV)...", idAse);
            leaf.AntExtRev = _leafReader.LeerAntExtRev(ase, solicitud.Periodo, rutaRecaudosReversados);


            // Plan 21 (T4, R-E-1/R-E-5): bloque espejo estructural del R1 por ASE — la fuente del
            // período actual define la forma (sin cardinalidades ni ocurrencias congeladas). Se lee
            // SIEMPRE antes de la escritura; el writer redimensiona el bloque destino y escribe por
            // encabezado. Los asserts de cierre T0e (Componente/Total, Subs/Cont/Total, Total final)
            // son invariantes duras: si falta una, fail-fast nombrando ASE + reporte + fila.
            progreso?.Report($"ASE {idAse}: leyendo bloque espejo estructural R1...");
            Log.Information("ASE {AseId}: leyendo bloque espejo estructural R1...", idAse);
            // Plan 21 (T5): el espejo R1 es la ÚNICA vía de la columna L-menores; el path legado
            // rol/ocurrencia (LeerLEspecialesMenores + mapa T0-0.5) se retiró al probar la absorción
            // 15/15 (EspejoR1AbsorcionTests). Sin carve-out ni slot ausente: la secuencia observada
            // ES la especificación del período (D-B), sin ocurrencias congeladas que puedan romper.
            var espejoR1 = _leafReader.LeerEspejoR1(ase, rutaR1);
            ValidarInvariantesEspejoR1(ase, espejoR1);
            leaf.EspejoR1 = espejoR1;

            // HU-16 (D2b, T0-0.2/0.3/0.4): INTERVENTORIA = insumo externo anual DECLARADO (sin
            // fuente en Docs/Insumos, V8). La hoja queda protegida intacta + assert estructural
            // (writer); aquí solo se audita el estatuto por ASE con Hoja = "INTERVENTORIA".
            progreso?.Report($"ASE {idAse}: INTERVENTORIA = insumo externo declarado — hoja intacta (bloque anual estático; no se escribe).");
            Log.ForContext("Hoja", "INTERVENTORIA")
                .Information("ASE {AseId}: INTERVENTORIA = insumo externo declarado — hoja intacta (bloque anual estático K26:K30/N26:N30; no se escribe; veredicto D2b T0).", idAse);

            // HU-08 (2.2): conciliación por empresa de facturación de este ASE (fail-fast ASE+empresa).
            // Plan 59 (H3): SIN recorte por quincena. El recorte heredado de HU-20-T0b/G2-D1 se
            // apoyaba en un bloqueo de PLANTILLA que ya no existe: la celda destino del grupo
            // (1,4) `Rem. Anticipos R2!E285` era FÓRMULA en la v2 y en la plantilla vigente es un
            // slot de valor, y el SEGUNDO cuadro de R2 declara el grupo `canal A / empresa EAB` en
            // los 5 ASE. El destino y la protección NO se congelan aquí: los resuelve el reader por
            // cruce de etiquetas contra la plantilla que se está procesando, y el guard
            // anti-fórmula (`EscribirValorNumerico`) sigue abortando si la plantilla futura
            // convierte un destino en fórmula. La lectura corre igual en ambas quincenas.
            progreso?.Report($"ASE {idAse}: leyendo conciliación por empresa (R1/R2/R4)...");
            Log.Information("ASE {AseId}: leyendo conciliación por empresa (R1/R2/R4)...", idAse);
            leaf.Conciliacion = _leafReader.LeerConciliacionEmpresas(ase, solicitud.Periodo, rutaR1, rutaR2, rutaR4);

            // HU-09 (2.3): reporte de recaudo por banco de este ASE (fail-fast ASE+empresa-columna).
            var rutaBanco = _localizador.BuscarReporteBanco(carpetaAse)
                ?? throw new ArchivoFuenteNoEncontradoException(
                    CodigoError.FuenteNoEncontrada,
                    $"No se encontró ReportePagosxBanco del ASE {idAse} en {carpetaAse}.");
            progreso?.Report($"ASE {idAse}: leyendo resumen de recaudo por banco...");
            Log.Information("ASE {AseId}: leyendo resumen de recaudo por banco...", idAse);
            leaf.ReporteBanco = _leafReader.LeerReporteBanco(ase, solicitud.Periodo, rutaBanco);

            // HU-10 (2.4): balance de subsidios y contribuciones de este ASE (fail-fast ASE).
            // Dos prefijos del locator: base R4-BalanceSubsidioyContribuciones_ + variante
            // -Optimizado_ (V8/T0-0.3); Q2 sigue bloqueada aguas arriba (sin cambios).
            var rutaBalance = _localizador.BuscarBalance(carpetaAse)
                ?? throw new ArchivoFuenteNoEncontradoException(
                    CodigoError.FuenteNoEncontrada,
                    $"No se encontró R4-BalanceSubsidioyContribuciones del ASE {idAse} en {carpetaAse}.");
            progreso?.Report($"ASE {idAse}: leyendo balance de subsidios y contribuciones...");
            Log.Information("ASE {AseId}: leyendo balance de subsidios y contribuciones...", idAse);
            leaf.BalanceSc = _leafReader.LeerBalanceSc(ase, solicitud.Periodo, rutaBalance);

            // HU-11 (2.5, path Q2): SALDOS POR NOTA + RETRIBUCION NEGATIVA por ASE (fail-fast
            // nombra ASE + reporte; locator agnóstico a rango y diacríticos, D4). En Q1 el
            // paso 2.5 se OMITE (AjustesSfT = null, comportamiento HU-10 intacto, G3).
            if (esQuincena2)
            {
                var rutaSaldosNotas = _localizador.BuscarSaldosNotas(carpetaAse)
                    ?? throw new ArchivoFuenteNoEncontradoException(
                        CodigoError.FuenteNoEncontrada,
                        $"No se encontró SaldosaFavorAplicadosPorNotas del ASE {idAse} en {carpetaAse}.");
                var rutaRetribucionNegativa = _localizador.BuscarRetribucionNegativa(carpetaAse)
                    ?? throw new ArchivoFuenteNoEncontradoException(
                        CodigoError.FuenteNoEncontrada,
                        $"No se encontró RetribuciónNegativa del ASE {idAse} en {carpetaAse}.");

                progreso?.Report($"ASE {idAse}: leyendo SALDOS POR NOTA y RETRIBUCION NEGATIVA...");
                Log.Information("ASE {AseId}: leyendo SALDOS POR NOTA y RETRIBUCION NEGATIVA...", idAse);
                leaf.AjustesSfT = new AjustesSfTInputs
                {
                    Ase = ase,
                    SaldosNotas = _leafReader.LeerSaldosNotas(ase, rutaSaldosNotas),
                    RetribucionNegativa = _leafReader.LeerRetribucionNegativa(ase, rutaRetribucionNegativa)
                };
            }

            // HU-12 (2.6 ampliada, V0.4) / HU-20 (G3): DetRetri por ASE — composición CONGELADA
            // ROUND(D104:D108,0) vía DetRetriRounder (origen = Σ visibles leaf + AJUSTES-SF-T).
            // Se calcula en AMBAS quincenas: en Q1 AjustesSfT = 0 y el DetRetri es ORÁCULO DE
            // VALIDACIÓN contra el R10 (nunca se escribe — el writer solo escribe DetRetri en Q2);
            // en Q2 suma AJUSTES-SF-T. El ROUND(D104:D108,0) de Q1 cierra contra
            // R10_Remuneracion_{AAAAMMQ} (evidencia T0: D9:D13 = ROUND de D104:D108 en ambas).
            leaf.DetRetriQ2 = new DetRetriQ2Inputs
            {
                Ase = ase,
                TotalD104 = leaf.R1.TotalOportunoEsperadoPorAse
                    + leaf.R2.TotalOportunoEsperado
                    + leaf.R1.ExtemporaneoEsperadoPorAse
                    + leaf.R4.TotalReversionEsperada
                    + (leaf.AjustesSfT?.TotalAjustes ?? 0m)
            };

            progreso?.Report($"ASE {idAse}: DetRetri esperado post-Excel = {leaf.DetRetriQ2.Detalle:0} (origen D104:D108 = {leaf.DetRetriQ2.TotalD104:0.##}).");
            Log.Information("ASE {AseId}: DetRetri esperado post-Excel = {Detalle:0} (origen D104:D108 = {Total:0.##}).", idAse, leaf.DetRetriQ2.Detalle, leaf.DetRetriQ2.TotalD104);

            datos.Add((ase, r1, r2, r4));
            if (esQuincena2)
            {
                datosConAjustes.Add((ase, r1, r2, r4, leaf.AjustesSfT?.TotalAjustes ?? 0m));
            }

            leafs.Add(leaf);
        }

        // HU-08 (2.2) / HU-20: hojas Recaudo * ← {periodo}/Conciliaciones (G1). Se leen UNA vez y
        // se comparten en los 5 leafs; fail-fast nombra la empresa si falta su archivo.
        // HU-20-T0b/G2-D1: el recorte Q2 se LEVANTA para Recaudos (layout RESUMEN MES uniforme
        // verificado en disco). La quincena la gobierna Periodo.NumeroQuincena (G2-D2): en Q2 el
        // reader toma el par de columnas F/G (VALOR 2°Q).
        progreso?.Report("Leyendo hojas Recaudo * desde las conciliaciones por empresa...");
        Log.Information("Leyendo hojas Recaudo * desde las conciliaciones por empresa...");
        var recaudos = _leafReader.LeerRecaudosEmpresa(
            solicitud.Periodo,
            empresa => _localizador.BuscarConciliacion(solicitud.CarpetaPeriodo, empresa.PrefijoConciliacion));
        foreach (var leaf in leafs)
        {
            leaf.Recaudos = recaudos;
        }

        progreso?.Report("Calculando consolidados de los 5 ASE...");
        Log.Information("Calculando consolidados de los 5 ASE...");
        var resultado = esQuincena2
            ? _calculoRemuneracion.CalcularConsolidado(solicitud.Periodo, datosConAjustes)
            : _calculoRemuneracion.CalcularConsolidado(solicitud.Periodo, datos);

        // HU-20 (G3): el R10 del período es ORÁCULO DE VALIDACIÓN del DetRetri calculado (G3-D1).
        // Es un insumo de PERÍODO (no de ASE) y parte del FLUJO NORMAL de AMBAS quincenas: si
        // falta, fail-fast que nombra período + archivo (G3-D2), nunca warning silencioso.
        var rutaR10 = _localizador.BuscarR10(solicitud.CarpetaPeriodo)
            ?? throw new ArchivoFuenteNoEncontradoException(
                CodigoError.FuenteNoEncontrada,
                $"No se encontró R10_Remuneracion_{solicitud.Periodo.CodigoCompleto} en '{solicitud.CarpetaPeriodo}'.");
        progreso?.Report($"Leyendo R10 del período ({Path.GetFileName(rutaR10)}) como oráculo de DetRetri...");
        Log.Information("Leyendo R10 del período ({Archivo}) como oráculo de DetRetri...", Path.GetFileName(rutaR10));
        var r10 = _detRetriR10Reader.LeerDetRetri(solicitud.Periodo, rutaR10);

        // Plan 62 (N1, T2.1/D-ORDEN): el R10 es la FUENTE de las hojas DetRetri/DetValiRetri (N1
        // reemplaza a G3-D1). Los DOS bloques del contrato T0 (A3:O19 y A3:O14) se leen AQUI, antes
        // de escribir: un R10 sin las 2 hojas o con el bloque incompleto falla con
        // ERR-FUENTE-NO-ENCONTRADA / ERR-FORMATO-FUENTE ANTES de tocar la salida (fail-fast sin
        // escritura parcial). El writer los pega en su MISMA pasada atomica, despues del CONSOLIDADO
        // (D104:D108) y de la Phase 0 de la base, y ANTES de los gates de datos (ValidarDetRetri-
        // ContraR10) y del oraculo de validaciones (LeerSnapshots), que asi leen lo pegado.
        progreso?.Report("Leyendo el bloque del R10 que se pega en DetRetri/DetValiRetri...");
        Log.Information("Leyendo el bloque del R10 (DetRetri/DetValiRetri) que se pega en las hojas del periodo...");
        resultado.BloqueDetRetriR10 = _detRetriR10Reader.LeerBloqueDetRetri(solicitud.Periodo, rutaR10);
        resultado.BloqueValiRetriR10 = _detRetriR10Reader.LeerBloqueDetValiRetri(solicitud.Periodo, rutaR10);

        // Plan 28 (Unidad F / D-E): el sello de fechas del período exige que el R10 traiga
        // 'Fecha Desde' (G7) y 'Fecha Hasta' (J7) legibles. Sin fechas → fail-fast que nombra
        // período + archivo + celda; NUNCA se sellan fechas silenciosas de otro período ni se
        // generan salidas con el rango de la plantilla.
        if (r10.FechaDesde == default || r10.FechaHasta == default)
        {
            var celdasFaltantes = new List<string>();
            if (r10.FechaDesde == default)
            {
                celdasFaltantes.Add("G7 ('Fecha Desde')");
            }

            if (r10.FechaHasta == default)
            {
                celdasFaltantes.Add("J7 ('Fecha Hasta')");
            }

            throw new CalculoInvalidoException(
                CodigoError.FormatoFuente,
                $"El R10 del período {solicitud.Periodo.CodigoCompleto} ({Path.GetFileName(rutaR10)}) no trae fecha legible en la celda {string.Join(" y ", celdasFaltantes)} de la hoja DetRetri{solicitud.Periodo.CodigoCompleto}; no se puede sellar el período y no se genera salida.");
        }

        resultado.FechaDesde = r10.FechaDesde;
        resultado.FechaHasta = r10.FechaHasta;

        // Plan 31 (T3, R-S-1/R-S-2/D-F): sello de proceso (Fecha de Proceso / Hora) del encabezado
        // del R10. El mini-T0 congeló que D6/D7 son LITERALES (nunca <f>) y su origen es el propio
        // R10 (coinciden exactamente con el manual del administrativo: julio 04/08/2026 10:15 AM,
        // agosto 02/09/2026 07:42 AM). Sin ese sello legible no se inventa una hora de corrida:
        // fail-fast que nombra período + archivo + celda, sin salida parcial.
        if (r10.FechaProceso is null || r10.HoraProceso is null)
        {
            var celdasProcesoFaltantes = new List<string>();
            if (r10.FechaProceso is null)
            {
                celdasProcesoFaltantes.Add("D6 ('Fecha de Proceso')");
            }

            if (r10.HoraProceso is null)
            {
                celdasProcesoFaltantes.Add("D7 ('Hora')");
            }

            throw new CalculoInvalidoException(
                CodigoError.FormatoFuente,
                $"El R10 del período {solicitud.Periodo.CodigoCompleto} ({Path.GetFileName(rutaR10)}) no trae fecha/hora de proceso legible en la celda {string.Join(" y ", celdasProcesoFaltantes)} de la hoja DetRetri{solicitud.Periodo.CodigoCompleto}; no se puede sellar el proceso y no se genera salida.");
        }

        resultado.FechaProceso = r10.FechaProceso;
        resultado.HoraProceso = r10.HoraProceso;

        progreso?.Report("Generando workbook de salida (una sola escritura)...");
        Log.Information("Generando workbook de salida (una sola escritura)...");
        // Plan 61 (H3.2 / S3.1, §2-D-H3-ii): POLÍTICA DE ENTREGA INVERTIDA. El workbook se escribe
        // ANTES de los gates de DATOS: ante una novedad el usuario SIEMPRE tiene el Excel revisable
        // (marcado "sin validar") para comparar y ajustar mientras la app se corrige. Los gates que
        // siguen aquí abajo ya NO bloquean la escritura: clasifican (ver §2-D-H3). Los gates de
        // PREFLIGHT (insumos faltantes) y los de ESTRUCTURA (plantilla alterada, hoja/destino
        // inexistente, celda-fórmula donde va valor) siguen siendo fallo duro: sin artefacto fiel no
        // hay nada que entregar.
        _workbookLeafWriter.GenerarWorkbook(solicitud.RutaPlantilla, rutaSinValidar, resultado, leafs);

        if (esQuincena2)
        {
            foreach (var leaf in leafs.OrderBy(l => l.Ase.Id))
            {
                var ajustes = leaf.AjustesSfT;
                if (ajustes is not null)
                {
                    progreso?.Report($"ASE {leaf.Ase.Id}: AJUSTES-SF-T esperado post-Excel = {ajustes.TotalAjustes:0.##} (saldos-nota {ajustes.SaldosNotas.TotalSaldosNotas:0.##} + retribución-negativa {ajustes.RetribucionNegativa.TotalRetribucionNegativa:0.##}).");
                    Log.Information("ASE {AseId}: AJUSTES-SF-T esperado post-Excel = {TotalAjustes:0.##} (saldos-nota {SaldosNotas:0.##} + retribución-negativa {RetribucionNegativa:0.##}).",
                        leaf.Ase.Id, ajustes.TotalAjustes, ajustes.SaldosNotas.TotalSaldosNotas, ajustes.RetribucionNegativa.TotalRetribucionNegativa);
                }
            }
        }

        // HU-16 (§2.6, CA-6): resumen INTERVENTORIA + L-Especiales menores por ASE con
        // Hoja = "INTERVENTORIA" (D2b declarado / D3a valores leídos y escritos).
        foreach (var leaf in leafs.OrderBy(l => l.Ase.Id))
        {
            var interventoria = InterventoriaDeclarada.ValorOficialMesPorAse.GetValueOrDefault(leaf.Ase.Id, 0m);
            var seg = InterventoriaDeclarada.SegundaQuincenaPorAse.GetValueOrDefault(leaf.Ase.Id, 0m);
            var pri = InterventoriaDeclarada.PrimeraQuincenaPorAse.GetValueOrDefault(leaf.Ase.Id, 0m);
            progreso?.Report($"ASE {leaf.Ase.Id}: INTERVENTORIA (insumo externo declarado) K={interventoria:0} M(2ª)={seg:0} N(1ª)={pri:0} — hoja intacta, no se escribe.");
            Log.ForContext("Hoja", "INTERVENTORIA")
                // HU-17 (S-4 HU-16): lectura por ASE = DETALLE (Debug), no hito (Information);
                // la doctrina HU-14 D4 exige hitos en Information y valores/lecturas en Debug.
                .Debug("ASE {AseId}: INTERVENTORIA (insumo externo declarado) K={K:0} M(2ª)={M:0} N(1ª)={N:0} — hoja intacta, no se escribe.", leaf.Ase.Id, interventoria, seg, pri);
        }

        // ── Plan 61 (H3.2 / S3.2, §2-D-H3): GATES DE DATOS → NOVEDAD (nunca bloqueo de la entrega) ──
        // Escribir→validar. Lo que sigue ya NO lanza: clasifica. La corrida sigue adelante con el
        // Excel escrito y termina en `EntregadaConNovedades` (marcado + reporte + código ≠ 0).
        var novedades = new List<NovedadEntrega>();

        // Gate multi-ASE (§2.5): Σ visibles-empresa vs visible de bloque (R1/R2/R4), cifras del ASE
        // frente a su consolidado, ajustes, banco y balance. El hito conserva su texto (es contrato
        // del frente de lectura); lo que cambió es el EFECTO: clasifica como novedad, no bloquea.
        progreso?.Report("Validando coherencia multi-ASE...");
        Log.Information("Validando coherencia multi-ASE (Plan 61: divergencia de datos = novedad, no bloqueo)...");
        var erroresCoherencia = _validador.Validar(resultado, leafs);
        foreach (var errorCoherencia in erroresCoherencia)
        {
            Log.Warning("Novedad de coherencia multi-ASE: {Detalle}", errorCoherencia);
        }

        novedades.AddRange(ClasificadorNovedadesEntrega.DesdeCoherencia(erroresCoherencia, solicitud.Periodo));

        // HU-20 (G3-D1/R-G3-3) + D8 (ratificada): el DetRetri calculado se contrasta con el R10 en
        // AMBAS quincenas. Si discrepan, MANDA LA PLANTILLA: la diferencia se reporta como novedad y
        // NUNCA se corrige el cálculo ni se amplía la tolerancia (ver §13.2-C: errata del oráculo).
        progreso?.Report($"Validando DetRetri calculado contra el R10 ({Path.GetFileName(rutaR10)})...");
        Log.Information("Validando DetRetri calculado contra el R10 ({Archivo})...", Path.GetFileName(rutaR10));
        var erroresR10 = _validador.ValidarDetRetriContraR10(leafs, r10);
        foreach (var errorR10 in erroresR10)
        {
            Log.Warning("Novedad del asiento frente al R10: {Detalle}", errorR10);
        }

        novedades.AddRange(ClasificadorNovedadesEntrega.DesdeR10(
            leafs, r10, solicitud.Periodo, Path.GetFileName(rutaR10)));

        // HU-13 (2.7): validaciones cruzadas como ORÁCULO DE LECTURA (D1). Solo si hay reader
        // (modo período completo de la UI): se lee el snapshot de la SALIDA (read-only, caché
        // visible — OpenXML no recalcula) y se evalúan los gates 2.7. Plan 61 (D-NOVEDAD-27): una
        // divergencia NO declarada degrada a novedad con su ficha (celda + Δ + oráculo) — negar el
        // Excel no hace el FALSE menos real; entregarlo lo hace revisable. Sin reader (regresión) =
        // HU-12 puro por construcción (D3).
        var lineasValidaciones = new List<string>();
        ClasificacionDivergenciasDeclaradas? clasificacion = null;
        var rutaReporte = string.Empty;
        var hayDeclaradaReconocida = false;
        if (_validacionOracleReader is not null)
        {
            progreso?.Report("Leyendo oráculo de validaciones cruzadas de la salida (read-only)...");
            Log.Information("Leyendo oráculo de validaciones cruzadas de la salida (read-only)...");
            // Plan 43 (O-4): el oráculo lee el archivo MARCADO (donde el writer lo dejó); el
            // contenido es idéntico al definitivo que se promoverá al cerrar bien.
            var snapshots = _validacionOracleReader.LeerSnapshots(rutaSinValidar, solicitud.Periodo);

            // Plan 42 (T1, R-CLS-1/R-CLS-2/R-CLS-3, D-B/D-F): clasificación en Core. La declarada
            // es la ÚLTIMA puerta: primero todo lo interno; solo si queda SOLO la declarada se pide
            // reconocimiento y, con él, se produce + reporta. Plan 61 (H3): lo NO declarado ya no
            // aborta la entrega — se reporta como novedad y el archivo queda marcado.
            clasificacion = _validador.ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

            if (clasificacion.HayNoDeclaradas)
            {
                Log.Warning(
                    "La validación cruzada (2.7) trae divergencia(s) NO declarada(s): se entregan como novedad. {Detalle}",
                    string.Join("; ", clasificacion.ErroresNoDeclarados));
                novedades.AddRange(ClasificadorNovedadesEntrega.DesdeValidacionCruzada(
                    clasificacion, erroresCoherencia, solicitud.Periodo));

                // Red de seguridad anti-falso-verde: si por cualquier motivo la clasificación declara
                // divergencias no declaradas y no se pudo componer NINGUNA ficha (p. ej. todos sus
                // errores ya venían del gate §2.5), la corrida NO puede salir certificada. Se entrega
                // con una novedad explícita que manda a revisar la hoja de validaciones.
                if (novedades.Count == 0)
                {
                    novedades.Add(new NovedadEntrega(
                        "Validación cruzada de la salida (VALIDACION_TOTAL)",
                        "VALIDACION_TOTAL",
                        0,
                        "Control de la hoja de validaciones cruzadas",
                        null,
                        null,
                        null,
                        "Equipo de la aplicación (control interno del cálculo)",
                        "Revise la hoja VALIDACION_TOTAL del Excel entregado: hay controles internos que no cierran.",
                        "La validación cruzada de la salida reportó diferencias no declaradas que no se pudieron detallar por separado."));
                }
            }

            if (clasificacion.SoloDeclaradas)
            {
                var aceptadas = new HashSet<string>(solicitud.DivergenciasDeclaradasAceptadas, StringComparer.OrdinalIgnoreCase);
                var pendientes = clasificacion.Declaradas
                    .Where(v => v.EntradaId is null || !aceptadas.Contains(v.EntradaId))
                    .ToList();
                if (pendientes.Count > 0)
                {
                    // S3 (D-D): sin acto consciente de ESTA corrida (UI) o sin flag (CLI), la
                    // declarada también bloquea — nunca se cierra "por costumbre". Plan 61: el
                    // flujo Plan 42 se mantiene (la UI/CLI necesitan la clasificación para ofrecer
                    // el reconocimiento), pero AHORA con el Excel y el reporte ya en disco.
                    var detalleDeclarada = string.Join("; ", clasificacion.Declaradas
                        .Select(v => $"[{CodigoError.Validacion}] ASE {v.AseId} · VALIDACION_TOTAL · {v.Celda}: divergencia declarada ({v.EntradaId}) pendiente de reconocimiento explícito en esta corrida (Δ = {v.Delta:0.##})"));
                    Log.Warning("Cierre bloqueado: hay divergencia(s) declarada(s) sin reconocimiento explícito en esta corrida (el Excel marcado y su reporte quedan entregados).");
                    EmitirReporteNovedadesDeBloqueo(
                        solicitud, runId, usuarioAceptacion, canalAceptacion, rutaSinValidar, clasificacion,
                        ClasificadorNovedadesEntrega.DesdeDeclaradasPendientes(clasificacion));
                    throw new CalculoInvalidoException(
                        CodigoError.Validacion,
                        $"La validación cruzada (2.7) falló: {detalleDeclarada}",
                        clasificacion);
                }

                hayDeclaradaReconocida = true;
            }

            // Veredictos por ASE (bloque VALIDACIONES del log/UI, §2.6). Honestidad: se lee el
            // caché visible; el recálculo Excel (Capa B) es acción del usuario (§2.7 A5).
            progreso?.Report("VALIDACIONES por ASE (caché visible de la salida; recálculo Excel = Capa B manual):");
            Log.Information("VALIDACIONES por ASE (caché visible de la salida; recálculo Excel = Capa B manual):");
            foreach (var snapshot in snapshots.OrderBy(s => s.Ase.Id))
            {
                var aseId = snapshot.Ase.Id;
                lineasValidaciones.Add($"ASE {aseId} VALIDACIONES:");
                Log.ForContext("AseId", aseId).Debug("VALIDACIONES del ASE {AseId}", aseId);
                foreach (var empresa in snapshot.PorEmpresa.OrderBy(e => e.Empresa, StringComparer.OrdinalIgnoreCase))
                {
                    lineasValidaciones.Add($"  {empresa.Empresa}: O (Recaudo vs REMUNERACION) = {empresa.DiferenciaO:0.###}; P (INT(O)=0) = {(empresa.VerificacionP ? "TRUE" : "FALSE")}");
                    Log.ForContext("AseId", aseId)
                        .ForContext("Validacion", EmpresaFacturacion.Catalogo.FirstOrDefault(e => e.Nombre == empresa.Empresa)?.HojaValidacion ?? empresa.Empresa)
                        .Debug("Empresa {Empresa}: O (Recaudo vs REMUNERACION) = {O:0.###}; P (INT(O)=0) = {P}", empresa.Empresa, empresa.DiferenciaO, empresa.VerificacionP);
                }

                if (snapshot.DetValiRetri is not null)
                {
                    var maxDiff = snapshot.DetValiRetri.DiferenciasAse.Count == 0
                        ? 0m
                        : snapshot.DetValiRetri.DiferenciasAse.Max(c => Math.Abs(c.Valor));
                    var verifComposicion = snapshot.DetValiRetri.VerificacionesAse.All(v => v.Verificacion);
                    lineasValidaciones.Add($"  DetValiRetri D16:D20 cierra (máx |dif| = {maxDiff:0.###}); verificación D24:D28 = {(verifComposicion ? "TRUE" : "FALSE")}; D29 = {(snapshot.DetValiRetri.VerificacionTotalD29 ? "TRUE" : "FALSE")}");
                    lineasValidaciones.Add($"  DetValiRetri D21 (fila Total) divergencia documentada = {snapshot.DetValiRetri.DiferenciaTotalD21:0.###} (excluida del gate, D6); D9:D14/J9:J14 ≠ ROUND documentado.");
                    Log.ForContext("AseId", aseId)
                        .ForContext("Validacion", "DetValiRetri")
                        .Debug("DetValiRetri D16:D20 cierra (máx |dif| = {MaxDiff:0.###}); verificación D24:D28 = {Verif}; D29 = {D29}", maxDiff, verifComposicion, snapshot.DetValiRetri.VerificacionTotalD29);
                    Log.ForContext("AseId", aseId)
                        .ForContext("Validacion", "DetValiRetri")
                        .Debug("DetValiRetri D21 (fila Total) divergencia documentada = {D21:0.###} (excluida del gate, D6); D9:D14/J9:J14 ≠ ROUND documentado.", snapshot.DetValiRetri.DiferenciaTotalD21);
                }

                lineasValidaciones.Add($"  VALIDACION_TOTAL O9 = {snapshot.ValidacionTotal:0.###}; P9 = {(snapshot.ValidacionTotalOkP ? "TRUE" : "FALSE")}");
                Log.ForContext("AseId", aseId)
                    .ForContext("Validacion", "VALIDACION_TOTAL") // W-3: nombre real de la hoja (constante del mapa oráculo en Infrastructure)
                    .Debug("VALIDACION_TOTAL O9 = {O9:0.###}; P9 = {P9}", snapshot.ValidacionTotal, snapshot.ValidacionTotalOkP);
            }

            foreach (var linea in lineasValidaciones)
            {
                progreso?.Report(linea);
            }
        }

        // Plan 61 (H3.2 / S3.5, §2-D-H3-ii): PROMOCIÓN SOLO CON CERO NOVEDADES. Si el workbook tiene
        // alguna novedad de datos, la corrida TERMINA AQUÍ con la marca "sin validar": nunca toma el
        // nombre limpio, nunca se reporta como OK y siempre deja el reporte de novedades. El Excel
        // fue escrito ANTES de los gates, así que el usuario tiene el artefacto para revisar.
        if (novedades.Count > 0)
        {
            var rutaReporteNovedades = EmitirReporteNovedades(
                solicitud, runId, usuarioAceptacion, canalAceptacion, rutaSinValidar, novedades, clasificacion);

            Log.Warning(
                "Entrega CON NOVEDADES (no certificada): {Cantidad} novedad(es) de datos. Excel marcado: {Excel}. Reporte: {Reporte}. RunId: {RunId}.",
                novedades.Count,
                rutaSinValidar,
                rutaReporteNovedades,
                runId);
            progreso?.Report($"ENTREGADO CON NOVEDADES — NO CERTIFICADO. {novedades.Count} novedad(es) por revisar. Excel: {rutaSinValidar}. Reporte: {rutaReporteNovedades}");

            var entregaConNovedades = ConstruirResultado(
                resultado,
                leafs,
                rutaSinValidar,
                lineasValidaciones,
                clasificacion,
                rutaReporte,
                nuevoEstado: EstadoEntrega.EntregadaConNovedades,
                novedades,
                rutaReporteNovedades);
            InvarianteEntrega.Garantizar(entregaConNovedades);
            return entregaConNovedades;
        }

        // Reconocida y sin novedades de datos: la doble huella es obligatoria (D-E). Sin reporte NO
        // hay ÉXITO (Plan 42 intacto). Plan 61 (D-TXT): cuando SOLO hay declaradas se conserva el
        // nombre anterior del reporte; el de novedades absorbe esa sección cuando coexisten.
        if (hayDeclaradaReconocida)
        {
            if (_reporteDivergencias is null)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Inesperado,
                    "El cierre con divergencia declarada exige el reporte durable y no hay emisor de reporte configurado (sin reporte no hay éxito).");
            }

            rutaReporte = _reporteDivergencias.Emitir(solicitud.RutaSalida, solicitud.Periodo, runId, clasificacion!, usuarioAceptacion, canalAceptacion);
            RegistrarHuellaDeclaradas(clasificacion!, rutaReporte, runId, usuarioAceptacion, canalAceptacion);
            progreso?.Report($"Divergencia(s) declarada(s) reconocida(s) en esta corrida. Reporte: {rutaReporte}");
        }

        // Plan 43 (O-4): la corrida no trae novedades (si las hubiera, habría retornado arriba).
        // Recién ahora el archivo deja la marca "sin validar" y toma el nombre DEFINITIVO.
        if (!string.Equals(Path.GetFullPath(rutaSinValidar), Path.GetFullPath(solicitud.RutaSalida), StringComparison.OrdinalIgnoreCase))
        {
            File.Move(rutaSinValidar, solicitud.RutaSalida, overwrite: true);
        }

        progreso?.Report("Proceso del período completado correctamente.");
        Log.Information("Proceso del período completado correctamente.");

        var entregaCertificada = ConstruirResultado(
            resultado,
            leafs,
            solicitud.RutaSalida,
            lineasValidaciones,
            clasificacion,
            rutaReporte,
            nuevoEstado: EstadoEntrega.Certificada,
            [],
            string.Empty);
        InvarianteEntrega.Garantizar(entregaCertificada);
        return entregaCertificada;
    }

    /// <summary>
    /// Plan 61 (H3.2/H4.1): emite el reporte de novedades y devuelve su ruta. Fail-closed: sin
    /// emisor configurado la corrida NO puede devolver un resultado con novedades (sin reporte no hay
    /// entrega, misma doctrina que el reporte de declaradas).
    /// </summary>
    private string EmitirReporteNovedades(
        SolicitudProcesoPeriodo solicitud,
        Guid runId,
        string usuario,
        string canal,
        string rutaSinValidar,
        IReadOnlyList<NovedadEntrega> novedades,
        ClasificacionDivergenciasDeclaradas? clasificacion)
    {
        if (_reporteNovedades is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Inesperado,
                $"La corrida tiene {novedades.Count} novedad(es) de datos y exige el reporte de novedades, pero no hay emisor de reporte configurado. Se entrega el Excel marcado solo si además queda el reporte que el usuario pueda leer (sin reporte no hay entrega).");
        }

        return _reporteNovedades.Emitir(
            solicitud.RutaSalida,
            rutaSinValidar,
            solicitud.Periodo,
            runId,
            novedades,
            clasificacion,
            usuario,
            canal);
    }

    /// <summary>
    /// Plan 61 (H3.2): bloqueo por divergencia DECLARADA sin reconocer. El flujo Plan 42 se mantiene
    /// (la excepción porta la clasificación para que UI/CLI ofrezcan el reconocimiento), pero el
    /// usuario ya no se queda sin nada: el Excel marcado está escrito y, si hay emisor, se deja el
    /// reporte que explica qué hay que reconocer.
    /// </summary>
    private void EmitirReporteNovedadesDeBloqueo(
        SolicitudProcesoPeriodo solicitud,
        Guid runId,
        string usuario,
        string canal,
        string rutaSinValidar,
        ClasificacionDivergenciasDeclaradas clasificacion,
        IReadOnlyList<NovedadEntrega> novedades)
    {
        if (_reporteNovedades is null)
        {
            // Composición sin emisor (regresión): el bloqueo del Plan 42 sigue siendo la garantía.
            return;
        }

        try
        {
            var ruta = _reporteNovedades.Emitir(
                solicitud.RutaSalida,
                rutaSinValidar,
                solicitud.Periodo,
                runId,
                novedades,
                clasificacion,
                usuario,
                canal);
            Log.Warning("Corrida bloqueada por divergencia declarada sin reconocer: Excel marcado y reporte entregados. Reporte: {Reporte}", ruta);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // El bloqueo (ERR-VALIDACION) es la garantía dura; un fallo al dejar el reporte no debe
            // enmascararlo con otro error distinto.
            Log.Error(ex, "No se pudo emitir el reporte de novedades de la corrida bloqueada.");
        }
    }

    /// <summary>
    /// Plan 61 (H3.1): compone el resultado del período. Aditivo: los campos nuevos llevan sus
    /// defaults cuando la corrida se certifica.
    /// </summary>
    private static ResultadoProcesoPeriodo ConstruirResultado(
        ResultadoRemuneracion resultado,
        IReadOnlyList<WorkbookLeafInputs> leafs,
        string rutaEntregada,
        IReadOnlyList<string> lineasValidaciones,
        ClasificacionDivergenciasDeclaradas? clasificacion,
        string rutaReporteDeclaradas,
        EstadoEntrega nuevoEstado,
        IReadOnlyList<NovedadEntrega> novedades,
        string rutaReporteNovedades) =>
        new()
        {
            Resultado = resultado,
            Leafs = leafs,
            RutaSalida = rutaEntregada,
            Validaciones = lineasValidaciones,
            DivergenciasDeclaradas = clasificacion,
            ReporteDivergenciasDeclaradas = rutaReporteDeclaradas,
            EstadoEntrega = nuevoEstado,
            Novedades = novedades,
            RutaReporteNovedades = rutaReporteNovedades,
            FuentesAusentes = leafs
                .SelectMany(leaf => leaf.Conciliacion.SelectMany(conc => conc.FuentesAusentes)
                    .Select(fuente => $"ASE {leaf.Ase.Id} · {fuente}"))
                .Concat(leafs
                    .SelectMany(leaf => (leaf.ReporteBanco?.Ases ?? Array.Empty<ReporteBancoAseInputs>())
                        .SelectMany(bloque => bloque.FuentesAusentesBanco)
                        .Select(fuente => $"ASE {leaf.Ase.Id} · {fuente}")))
                // Plan 64 (H2.1, §2.4 paso 5): la traza ANTEXTREV:{fuente} del lector viaja también al
                // resultado del período, igual que las ausencias del Plan 50/51 (nunca un 0 silencioso).
                .Concat(leafs
                    .SelectMany(leaf => (leaf.AntExtRev?.FuentesAusentes ?? [])
                        .Select(fuente => $"ASE {leaf.Ase.Id} · {fuente}")))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList(),
            EtiquetasNoReconocidas = leafs
                .SelectMany(leaf => leaf.Conciliacion.SelectMany(conc => conc.EtiquetasNoReconocidas)
                    .Select(etiqueta => $"ASE {leaf.Ase.Id} · {etiqueta}"))
                .Concat(leafs
                    .SelectMany(leaf => (leaf.ReporteBanco?.Ases ?? Array.Empty<ReporteBancoAseInputs>())
                        .SelectMany(bloque => bloque.EtiquetasBancoNoReconocidas)
                        .Select(etiqueta => $"ASE {leaf.Ase.Id} · {etiqueta}")))
                // Plan 64 (H2.1, §2.4 paso 6): la etiqueta de fuente que el catálogo no tabula queda como
                // EVIDENCIA auditable (si además había una ausencia, el lector ya hizo fail-fast A3).
                .Concat(leafs
                    .SelectMany(leaf => (leaf.AntExtRev?.EtiquetasNoReconocidas ?? [])
                        .Select(etiqueta => $"ASE {leaf.Ase.Id} · {etiqueta}")))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
        };


    /// <summary>
    /// Plan 43 (O-3): resuelve el usuario de la aceptación de forma honesta. Usa el valor declarado
    /// por el frontend si llega; si no, el usuario del sistema (<see cref="Environment.UserName"/>)
    /// y, si el entorno no lo expone, <see cref="HuellaAceptacion.UsuarioNoDisponible"/> (nunca un
    /// nombre inventado).
    /// </summary>
    private static string ResolverUsuario(string? usuarioSolicitado) =>
        !string.IsNullOrWhiteSpace(usuarioSolicitado)
            ? usuarioSolicitado
            : string.IsNullOrWhiteSpace(Environment.UserName) ? HuellaAceptacion.UsuarioNoDisponible : Environment.UserName;

    /// <summary>
    /// Plan 42 (T2, R-REP-2, D-E) + Plan 43 (O-3): huella en el log (Serilog) del reconocimiento de
    /// divergencias declaradas en esta corrida. Junto con el reporte durable, es la segunda de las
    /// DOS huellas obligatorias del cierre (declarada ≠ tolerada en silencio). Registra QUIÉN aceptó
    /// (<paramref name="usuario"/>) y POR QUÉ CANAL (<paramref name="canal"/>) como propiedades
    /// estructuradas, además de celda/observado/esperado/Δ/RunId.
    /// </summary>
    private static void RegistrarHuellaDeclaradas(
        ClasificacionDivergenciasDeclaradas clasificacion,
        string rutaReporte,
        Guid runId,
        string usuario,
        string canal)
    {
        Log.ForContext("Usuario", usuario).ForContext("Canal", canal)
            .Information(
                "Cierre con divergencia(s) declarada(s) reconocida(s): {Cantidad} veredicto(s), entradas [{Entradas}]. Aceptado por {Usuario} (canal {Canal}). Reporte durable: {Reporte}. RunId: {RunId}.",
                clasificacion.Declaradas.Count,
                string.Join(", ", clasificacion.EntradasDeclaradas),
                usuario,
                canal,
                rutaReporte,
                runId);

        foreach (var veredicto in clasificacion.Declaradas
                     .GroupBy(v => (v.EntradaId, v.Celda))
                     .Select(g => g.First()))
        {
            Log.ForContext("Usuario", usuario)
                .ForContext("Canal", canal)
                .ForContext("AseId", veredicto.AseId)
                .ForContext("Validacion", "VALIDACION_TOTAL")
                .Information(
                    "Divergencia declarada reconocida (entrada {Entrada}): celda {Celda}; observado {Observado:0.##}; esperado {Esperado:0.##}; Δ {Delta:0.##}.",
                    veredicto.EntradaId,
                    veredicto.Celda,
                    veredicto.MontoObservado,
                    veredicto.MontoEsperado,
                    veredicto.Delta);
        }
    }

    /// <summary>
    /// Plan 21 (T4, R-E-5): asserts de cierre del espejo R1. Las tres invariantes duras T0e
    /// (<c>Componente/Total</c>, <c>Subs/Cont/Total</c> y <c>Total</c> final) deben estar presentes
    /// en la secuencia observada; si falta una, fail-fast que nombra ASE + reporte + fila esperada.
    /// <c>Mes</c> y <c>AFaseo</c> NO son invariantes (la forma la define la fuente del período).
    /// W-7 (auditoría PR3): el listado de faltantes lo aporta el modelo
    /// (<see cref="BloqueEspejoAseInputs.FaltantesInvariantesDeCierre"/>), fuente única compartida
    /// con <c>ExcelDataReaderWorkbookLeafInputReader.LeerEspejoR1</c>.
    /// </summary>
    internal static void ValidarInvariantesEspejoR1(Ase ase, BloqueEspejoAseInputs espejo)
    {
        var faltantes = espejo.FaltantesInvariantesDeCierre();
        if (faltantes.Count > 0)
        {
            // Doctrina fail-fast: el mensaje nombra ASE + reporte + fila esperada (nunca 0/invención).
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"Espejo estructural: la fuente R1 del ASE {ase.Id} ({ase.NombreCorto}) no trae las filas invariantes de cierre en el reporte 'Reporte Componentes R1': {string.Join(", ", faltantes)}.");
        }
    }
}
