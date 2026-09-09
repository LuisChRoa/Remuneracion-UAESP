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
    private readonly IValidacionOracleReader? _validacionOracleReader;

    public ProcesadorPeriodo(
        IRecaudoReader recaudoReader,
        IWorkbookLeafInputReader leafReader,
        ICalculoRemuneracion calculoRemuneracion,
        IValidador validador,
        IWorkbookLeafWriter workbookLeafWriter,
        ILocalizadorArchivosAse localizador,
        IValidacionOracleReader? validacionOracleReader = null)
    {
        _recaudoReader = recaudoReader ?? throw new ArgumentNullException(nameof(recaudoReader));
        _leafReader = leafReader ?? throw new ArgumentNullException(nameof(leafReader));
        _calculoRemuneracion = calculoRemuneracion ?? throw new ArgumentNullException(nameof(calculoRemuneracion));
        _validador = validador ?? throw new ArgumentNullException(nameof(validador));
        _workbookLeafWriter = workbookLeafWriter ?? throw new ArgumentNullException(nameof(workbookLeafWriter));
        _localizador = localizador ?? throw new ArgumentNullException(nameof(localizador));
        _validacionOracleReader = validacionOracleReader;
    }

    public ResultadoProcesoPeriodo Ejecutar(SolicitudProcesoPeriodo solicitud, IProgress<string>? progreso = null)
    {
        ArgumentNullException.ThrowIfNull(solicitud);

        // HU-14 (3.2, D5): RunId por ejecución correlaciona todos los eventos de ESTE procesador
        // (ambos modos). LogContext es Serilog core; las propiedades Periodo/Quincena/Modo
        // acompañan al RunId para filtrar el log (CA-6).
        var runId = Guid.NewGuid();
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

            // HU-08 (2.2): conciliación por empresa de facturación de este ASE (fail-fast ASE+empresa).
            // RECORTE HONESTO T0-0.6 (Riesgo 5): en Q2 el layout del R4 por empresa DIVERGE del Q1
            // (ASE2 trae ENEL+OCCIDENTE, no RECIPROCIDAD/"NUEVO ESQUEMA"; el template Q2 tampoco
            // tiene esa fila). El mapa HU-08 está congelado para Q1 y el plan §0.2 prohíbe
            // reescribirlo → la conciliación por empresa y las hojas Recaudo * se omiten en Q2
            // (leaf.Conciliacion/Recaudos vacíos = comportamiento HU-07 puro para validador/writer).
            if (!esQuincena2)
            {
                progreso?.Report($"ASE {idAse}: leyendo conciliación por empresa (R1/R2/R4)...");
                Log.Information("ASE {AseId}: leyendo conciliación por empresa (R1/R2/R4)...", idAse);
                leaf.Conciliacion = _leafReader.LeerConciliacionEmpresas(ase, solicitud.Periodo, rutaR1, rutaR2, rutaR4);
            }
            else
            {
                progreso?.Report($"ASE {idAse}: conciliación por empresa omitida en Q2 (recorte T0-0.6: layout R4 divergente vs Q1).");
                Log.Warning("ASE {AseId}: conciliación por empresa omitida en Q2 (recorte T0-0.6: layout R4 divergente vs Q1).", idAse);
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

                // HU-12 (2.6 ampliada, V0.4): DetRetri-Q2 por ASE — composición CONGELADA
                // ROUND(D104:D108,0) vía DetRetriRounder (origen = Σ visibles leaf Q2 + AJUSTES-SF-T).
                // Fail-fast si el leaf no expone los visibles Q2 (el reader Q2 ya falló aguas arriba).
                leaf.DetRetriQ2 = new DetRetriQ2Inputs
                {
                    Ase = ase,
                    TotalD104 = leaf.R1.TotalOportunoEsperadoPorAse
                        + leaf.R2.TotalOportunoEsperado
                        + leaf.R1.ExtemporaneoEsperadoPorAse
                        + leaf.R4.TotalReversionEsperada
                        + leaf.AjustesSfT.TotalAjustes
                };

                progreso?.Report($"ASE {idAse}: DetRetri esperado post-Excel = {leaf.DetRetriQ2.Detalle:0} (origen D104:D108 = {leaf.DetRetriQ2.TotalD104:0.##}).");
                Log.Information("ASE {AseId}: DetRetri esperado post-Excel = {Detalle:0} (origen D104:D108 = {Total:0.##}).", idAse, leaf.DetRetriQ2.Detalle, leaf.DetRetriQ2.TotalD104);
            }

            datos.Add((ase, r1, r2, r4));
            if (esQuincena2)
            {
                datosConAjustes.Add((ase, r1, r2, r4, leaf.AjustesSfT?.TotalAjustes ?? 0m));
            }

            leafs.Add(leaf);
        }

        // HU-08 (2.2): hojas Recaudo * ← Consolidado/Conciliaciones (T0-0.6). Se leen una vez
        // y se comparten en los 5 leafs; fail-fast nombra la empresa si falta su archivo.
        // RECORTE HONESTO T0-0.6: omitido en Q2 (misma razón que la conciliación por empresa).
        if (!esQuincena2)
        {
            progreso?.Report("Leyendo hojas Recaudo * desde las conciliaciones por empresa...");
            Log.Information("Leyendo hojas Recaudo * desde las conciliaciones por empresa...");
            var recaudos = _leafReader.LeerRecaudosEmpresa(
                empresa => _localizador.BuscarConciliacion(solicitud.CarpetaPeriodo, empresa.PrefijoConciliacion));
            foreach (var leaf in leafs)
            {
                leaf.Recaudos = recaudos;
            }
        }

        progreso?.Report("Calculando consolidados de los 5 ASE...");
        Log.Information("Calculando consolidados de los 5 ASE...");
        var resultado = esQuincena2
            ? _calculoRemuneracion.CalcularConsolidado(solicitud.Periodo, datosConAjustes)
            : _calculoRemuneracion.CalcularConsolidado(solicitud.Periodo, datos);

        progreso?.Report("Validando coherencia multi-ASE...");
        Log.Information("Validando coherencia multi-ASE...");
        var errores = _validador.Validar(resultado, leafs);
        if (errores.Count > 0)
        {
            var detalle = string.Join("; ", errores);
            throw new CalculoInvalidoException(CodigoError.Validacion, $"[{CodigoError.Validacion}] La validación multi-ASE falló: {detalle}");
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
                throw new CalculoInvalidoException(CodigoError.Validacion, $"[{CodigoError.Validacion}] La validación cruzada (2.7) falló: {detalleValidacion}");
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
}