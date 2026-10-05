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
        IValidacionOracleReader? validacionOracleReader = null)
    {
        _recaudoReader = recaudoReader ?? throw new ArgumentNullException(nameof(recaudoReader));
        _leafReader = leafReader ?? throw new ArgumentNullException(nameof(leafReader));
        _calculoRemuneracion = calculoRemuneracion ?? throw new ArgumentNullException(nameof(calculoRemuneracion));
        _validador = validador ?? throw new ArgumentNullException(nameof(validador));
        _workbookLeafWriter = workbookLeafWriter ?? throw new ArgumentNullException(nameof(workbookLeafWriter));
        _localizador = localizador ?? throw new ArgumentNullException(nameof(localizador));
        _detRetriR10Reader = detRetriR10Reader ?? throw new ArgumentNullException(nameof(detRetriR10Reader));
        _validacionOracleReader = validacionOracleReader;
        _validadorInsumos = new ValidadorInsumosPeriodo(_localizador);
    }

    public ResultadoProcesoPeriodo Ejecutar(SolicitudProcesoPeriodo solicitud, IProgress<string>? progreso = null)
    {
        ArgumentNullException.ThrowIfNull(solicitud);

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

            progreso?.Report($"ASE {idAse}: leyendo R1, R2 y R4...");
            Log.Information("ASE {AseId}: leyendo R1, R2 y R4...", idAse);
            var r1 = _recaudoReader.LeerR1(rutaR1);
            var r2 = _recaudoReader.LeerR2(rutaR2);
            var r4 = _recaudoReader.LeerR4(rutaR4);

            progreso?.Report($"ASE {idAse}: leyendo inputs leaf del workbook...");
            Log.Information("ASE {AseId}: leyendo inputs leaf del workbook...", idAse);
            var leaf = _leafReader.LeerLeafInputs(ase, solicitud.Periodo, rutaR1, rutaR2, rutaR4);

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
            // HU-20-T0b/G2-D1: el T0 re-verificó el layout R4-por-empresa Q2 y la divergencia
            // PERSISTE (ASE2-Q2 trae ENEL+OCCIDENTE y NO la fila RECIPROCIDAD/"NUEVO ESQUEMA"; el
            // mapa HU-08 congelado para Q1 no la puede resolver). Por el alcance acotado del plan,
            // el levantamiento Q2 cubre SOLO Recaudos; la conciliación por empresa en Q2 queda
            // como follow-up con su propio T0 (NUNCA se reescribe el mapa HU-08 a ciegas).
            // leaf.Conciliacion vacío en Q2 = comportamiento HU-07 puro para validador/writer.
            if (!esQuincena2)
            {
                progreso?.Report($"ASE {idAse}: leyendo conciliación por empresa (R1/R2/R4)...");
                Log.Information("ASE {AseId}: leyendo conciliación por empresa (R1/R2/R4)...", idAse);
                leaf.Conciliacion = _leafReader.LeerConciliacionEmpresas(ase, solicitud.Periodo, rutaR1, rutaR2, rutaR4);
            }
            else
            {
                progreso?.Report($"ASE {idAse}: conciliación por empresa omitida en Q2 (HU-20-T0b: layout R4 divergente persistente; alcance acotado a Recaudos).");
                Log.Warning("ASE {AseId}: conciliación por empresa omitida en Q2 (HU-20-T0b: layout R4 divergente persistente; alcance acotado a Recaudos).", idAse);
            }

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

        progreso?.Report("Validando coherencia multi-ASE...");
        Log.Information("Validando coherencia multi-ASE...");
        var errores = _validador.Validar(resultado, leafs);
        if (errores.Count > 0)
        {
            var detalle = string.Join("; ", errores);
            // HU-15 (S-2): sin prefijo [CÓDIGO] redundante — el código viaja en la excepción
            // (Codigo) y cada error del validador ya lo porta (AgregarError). El catch del
            // frontend agrega [{Codigo}] una sola vez (antes quedaba doble/triple en el log).
            throw new CalculoInvalidoException(CodigoError.Validacion, $"La validación multi-ASE falló: {detalle}");
        }

        // HU-20 (G3-D1/R-G3-3): DetRetri calculado vs R10 (±0.5 post-redondeo) en AMBAS quincenas.
        // Divergencia → error fail-fast con período + archivo + ambos valores; nunca escritura
        // parcial silenciosa.
        progreso?.Report($"Validando DetRetri calculado contra el R10 ({Path.GetFileName(rutaR10)})...");
        Log.Information("Validando DetRetri calculado contra el R10 ({Archivo})...", Path.GetFileName(rutaR10));
        var erroresR10 = _validador.ValidarDetRetriContraR10(leafs, r10);
        if (erroresR10.Count > 0)
        {
            var detalleR10 = string.Join("; ", erroresR10);
            throw new CalculoInvalidoException(
                CodigoError.Validacion,
                $"La validación DetRetri-vs-R10 falló (periodo {solicitud.Periodo.CodigoCompleto}, archivo {Path.GetFileName(rutaR10)}): {detalleR10}");
        }

        progreso?.Report("Generando workbook de salida (una sola escritura)...");
        Log.Information("Generando workbook de salida (una sola escritura)...");
        _workbookLeafWriter.GenerarWorkbook(solicitud.RutaPlantilla, solicitud.RutaSalida, resultado, leafs);

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

        // HU-13 (2.7): validaciones cruzadas como ORÁCULO DE LECTURA (D1). Solo si hay reader
        // (modo período completo de la UI): se lee el snapshot de la SALIDA (read-only, caché
        // visible — OpenXML no recalcula) y se evalúan los gates 2.7 con fail-fast que nombra
        // ASE + validación. Sin reader (regresión) = HU-12 puro por construcción (D3).
        var lineasValidaciones = new List<string>();
        if (_validacionOracleReader is not null)
        {
            progreso?.Report("Leyendo oráculo de validaciones cruzadas de la salida (read-only)...");
            Log.Information("Leyendo oráculo de validaciones cruzadas de la salida (read-only)...");
            var snapshots = _validacionOracleReader.LeerSnapshots(solicitud.RutaSalida, solicitud.Periodo);
            var erroresValidacion = _validador.Validar(resultado, leafs, snapshots);
            if (erroresValidacion.Count > 0)
            {
                var detalleValidacion = string.Join("; ", erroresValidacion);
                // HU-15 (S-2): sin prefijo [CÓDIGO] redundante (ver nota de validación multi-ASE).
                throw new CalculoInvalidoException(CodigoError.Validacion, $"La validación cruzada (2.7) falló: {detalleValidacion}");
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

        progreso?.Report("Proceso del período completado correctamente.");
        Log.Information("Proceso del período completado correctamente.");

        return new ResultadoProcesoPeriodo
        {
            Resultado = resultado,
            Leafs = leafs,
            RutaSalida = solicitud.RutaSalida,
            Validaciones = lineasValidaciones
        };
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
