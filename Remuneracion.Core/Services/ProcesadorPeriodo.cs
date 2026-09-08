using Remuneracion.Core.Constants;
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
/// </summary>
public sealed class ProcesadorPeriodo : IProcesadorPeriodo
{
    private readonly IRecaudoReader _recaudoReader;
    private readonly IWorkbookLeafInputReader _leafReader;
    private readonly ICalculoRemuneracion _calculoRemuneracion;
    private readonly IValidador _validador;
    private readonly IWorkbookLeafWriter _workbookLeafWriter;
    private readonly ILocalizadorArchivosAse _localizador;

    public ProcesadorPeriodo(
        IRecaudoReader recaudoReader,
        IWorkbookLeafInputReader leafReader,
        ICalculoRemuneracion calculoRemuneracion,
        IValidador validador,
        IWorkbookLeafWriter workbookLeafWriter,
        ILocalizadorArchivosAse localizador)
    {
        _recaudoReader = recaudoReader ?? throw new ArgumentNullException(nameof(recaudoReader));
        _leafReader = leafReader ?? throw new ArgumentNullException(nameof(leafReader));
        _calculoRemuneracion = calculoRemuneracion ?? throw new ArgumentNullException(nameof(calculoRemuneracion));
        _validador = validador ?? throw new ArgumentNullException(nameof(validador));
        _workbookLeafWriter = workbookLeafWriter ?? throw new ArgumentNullException(nameof(workbookLeafWriter));
        _localizador = localizador ?? throw new ArgumentNullException(nameof(localizador));
    }

    public ResultadoProcesoPeriodo Ejecutar(SolicitudProcesoPeriodo solicitud, IProgress<string>? progreso = null)
    {
        ArgumentNullException.ThrowIfNull(solicitud);

        progreso?.Report("Iniciando proceso del período (5 ASE).");
        progreso?.Report($"Periodo: {solicitud.Periodo.CodigoCompleto}; carpeta: {solicitud.CarpetaPeriodo}");

        if (string.IsNullOrWhiteSpace(solicitud.CarpetaPeriodo))
        {
            throw new ArchivoFuenteNoEncontradoException("Debe indicarse la carpeta del período que contiene las 5 carpetas de ASE.");
        }

        if (string.IsNullOrWhiteSpace(solicitud.RutaPlantilla) || string.IsNullOrWhiteSpace(solicitud.RutaSalida))
        {
            throw new ArchivoFuenteNoEncontradoException("Debe indicarse plantilla y ruta de salida.");
        }

        // Resolver las 5 carpetas de ASE en orden 1..5 (fail-fast si falta alguna).
        var carpetasPeriodo = _localizador.ObtenerCarpetasAse(solicitud.CarpetaPeriodo);
        if (carpetasPeriodo.Length == 0)
        {
            throw new ArchivoFuenteNoEncontradoException(
                $"No se encontraron carpetas de ASE en '{solicitud.CarpetaPeriodo}'. Deben existir 5 carpetas con prefijos {string.Join(", ", CarpetasAse.Prefijos)}.");
        }

        var datos = new List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4)>();
        var leafs = new List<WorkbookLeafInputs>();

        for (var idAse = 1; idAse <= 5; idAse++)
        {
            var ase = CrearAse(idAse);
            var carpetaAse = carpetasPeriodo.FirstOrDefault(c =>
                    Path.GetFileName(c).StartsWith(idAse.ToString(), StringComparison.OrdinalIgnoreCase))
                ?? throw new ArchivoFuenteNoEncontradoException(
                    $"No se encontró la carpeta del ASE {idAse} ({CarpetasAse.Prefijos[idAse - 1]}) en '{solicitud.CarpetaPeriodo}'.");

            progreso?.Report($"ASE {idAse} ({ase.NombreCompleto}): localizando fuentes...");
            var rutaR1 = _localizador.BuscarArchivo(carpetaAse, "Recaudoporcomponente")
                ?? throw new ArchivoFuenteNoEncontradoException($"No se encontró R1 del ASE {idAse} en {carpetaAse}.");
            var rutaR2 = _localizador.BuscarArchivo(carpetaAse, "RerpoteDetalleSaldosaFavor")
                ?? throw new ArchivoFuenteNoEncontradoException($"No se encontró R2 del ASE {idAse} en {carpetaAse}.");
            var rutaR4 = _localizador.BuscarArchivo(carpetaAse, "ReversiónPorComponente")
                ?? _localizador.BuscarArchivo(carpetaAse, "ReversionPorComponente")
                ?? throw new ArchivoFuenteNoEncontradoException($"No se encontró R4 del ASE {idAse} en {carpetaAse}.");

            progreso?.Report($"ASE {idAse}: leyendo R1, R2 y R4...");
            var r1 = _recaudoReader.LeerR1(rutaR1);
            var r2 = _recaudoReader.LeerR2(rutaR2);
            var r4 = _recaudoReader.LeerR4(rutaR4);

            progreso?.Report($"ASE {idAse}: leyendo inputs leaf del workbook...");
            var leaf = _leafReader.LeerLeafInputs(ase, solicitud.Periodo, rutaR1, rutaR2, rutaR4);

            // HU-08 (2.2): conciliación por empresa de facturación de este ASE (fail-fast ASE+empresa).
            progreso?.Report($"ASE {idAse}: leyendo conciliación por empresa (R1/R2/R4)...");
            leaf.Conciliacion = _leafReader.LeerConciliacionEmpresas(ase, solicitud.Periodo, rutaR1, rutaR2, rutaR4);

            // HU-09 (2.3): reporte de recaudo por banco de este ASE (fail-fast ASE+empresa-columna).
            var rutaBanco = _localizador.BuscarReporteBanco(carpetaAse)
                ?? throw new ArchivoFuenteNoEncontradoException(
                    $"No se encontró ReportePagosxBanco del ASE {idAse} en {carpetaAse}.");
            progreso?.Report($"ASE {idAse}: leyendo resumen de recaudo por banco...");
            leaf.ReporteBanco = _leafReader.LeerReporteBanco(ase, solicitud.Periodo, rutaBanco);

            // HU-10 (2.4): balance de subsidios y contribuciones de este ASE (fail-fast ASE).
            // Dos prefijos del locator: base R4-BalanceSubsidioyContribuciones_ + variante
            // -Optimizado_ (V8/T0-0.3); Q2 sigue bloqueada aguas arriba (sin cambios).
            var rutaBalance = _localizador.BuscarBalance(carpetaAse)
                ?? throw new ArchivoFuenteNoEncontradoException(
                    $"No se encontró R4-BalanceSubsidioyContribuciones del ASE {idAse} en {carpetaAse}.");
            progreso?.Report($"ASE {idAse}: leyendo balance de subsidios y contribuciones...");
            leaf.BalanceSc = _leafReader.LeerBalanceSc(ase, solicitud.Periodo, rutaBalance);

            datos.Add((ase, r1, r2, r4));
            leafs.Add(leaf);
        }

        // HU-08 (2.2): hojas Recaudo * ← Consolidado/Conciliaciones (T0-0.6). Se leen una vez
        // y se comparten en los 5 leafs; fail-fast nombra la empresa si falta su archivo.
        progreso?.Report("Leyendo hojas Recaudo * desde las conciliaciones por empresa...");
        var recaudos = _leafReader.LeerRecaudosEmpresa(
            empresa => _localizador.BuscarConciliacion(solicitud.CarpetaPeriodo, empresa.PrefijoConciliacion));
        foreach (var leaf in leafs)
        {
            leaf.Recaudos = recaudos;
        }

        progreso?.Report("Calculando consolidados de los 5 ASE...");
        var resultado = _calculoRemuneracion.CalcularConsolidado(solicitud.Periodo, datos);

        progreso?.Report("Validando coherencia multi-ASE...");
        var errores = _validador.Validar(resultado, leafs);
        if (errores.Count > 0)
        {
            var detalle = string.Join("; ", errores);
            throw new CalculoInvalidoException($"La validación multi-ASE falló: {detalle}");
        }

        progreso?.Report("Generando workbook de salida (una sola escritura)...");
        _workbookLeafWriter.GenerarWorkbook(solicitud.RutaPlantilla, solicitud.RutaSalida, resultado, leafs);

        progreso?.Report("Proceso del período completado correctamente.");

        return new ResultadoProcesoPeriodo
        {
            Resultado = resultado,
            Leafs = leafs,
            RutaSalida = solicitud.RutaSalida
        };
    }

    private static Ase CrearAse(int id)
    {
        var prefijo = CarpetasAse.Prefijos[id - 1];
        var guion = prefijo.IndexOf('-');
        var nombre = guion > 0 ? prefijo[(guion + 1)..] : prefijo;
        return new Ase
        {
            Id = id,
            NombreCorto = nombre.ToUpperInvariant(),
            NombreCompleto = nombre,
            NumeroCarpeta = id
        };
    }
}