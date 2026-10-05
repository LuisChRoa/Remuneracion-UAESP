using Serilog;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Services;

/// <summary>
/// Orquestador del caso de uso de remuneración para un único ASE.
/// HU-14 (3.1 + 3.2): cada fail-fast porta un <see cref="CodigoError"/> del catálogo (D1);
/// RunId por ejecución vía <c>Serilog.Context.LogContext</c> correlaciona los eventos (D5) y
/// cada paso se registra con nivel y propiedades estructuradas (CA-6).
/// </summary>
public sealed class ProcesadorRemuneracion : IProcesadorRemuneracion
{
    private readonly IRecaudoReader _recaudoReader;
    private readonly IWorkbookLeafInputReader _leafReader;
    private readonly ICalculoRemuneracion _calculoRemuneracion;
    private readonly IValidador _validador;
    private readonly IWorkbookLeafWriter _workbookLeafWriter;

    public ProcesadorRemuneracion(
        IRecaudoReader recaudoReader,
        IWorkbookLeafInputReader leafReader,
        ICalculoRemuneracion calculoRemuneracion,
        IValidador validador,
        IWorkbookLeafWriter workbookLeafWriter)
    {
        _recaudoReader = recaudoReader ?? throw new ArgumentNullException(nameof(recaudoReader));
        _leafReader = leafReader ?? throw new ArgumentNullException(nameof(leafReader));
        _calculoRemuneracion = calculoRemuneracion ?? throw new ArgumentNullException(nameof(calculoRemuneracion));
        _validador = validador ?? throw new ArgumentNullException(nameof(validador));
        _workbookLeafWriter = workbookLeafWriter ?? throw new ArgumentNullException(nameof(workbookLeafWriter));
    }

    public ResultadoProcesoAse Ejecutar(SolicitudProcesoAse solicitud, IProgress<string>? progreso = null)
    {
        ArgumentNullException.ThrowIfNull(solicitud);

        // HU-14 (3.2, D5): RunId por ejecución correlaciona todos los eventos del procesador.
        // HU-15 (W-2.1, D4): si el frontend (UI/CLI) inyectó un RunId en la solicitud, se respeta
        // (un solo Guid correlaciona UI → procesador → writer); null = genera uno (compat HU-14).
        var runId = solicitud.RunId ?? Guid.NewGuid();
        using var _runIdScope = Serilog.Context.LogContext.PushProperty("RunId", runId);
        using var _periodoScope = Serilog.Context.LogContext.PushProperty("Periodo", solicitud.Periodo.CodigoCompleto);
        using var _quincenaScope = Serilog.Context.LogContext.PushProperty("Quincena", solicitud.Periodo.NumeroQuincena);
        using var _modoScope = Serilog.Context.LogContext.PushProperty("Modo", $"ASE {solicitud.Ase.Id}");

        progreso?.Report("Iniciando ejecución del procesador real.");
        Log.Information("Iniciando ejecución del procesador real.");
        progreso?.Report($"Periodo: {solicitud.Periodo.CodigoCompleto}; ASE: {solicitud.Ase.Id} - {solicitud.Ase.NombreCompleto}");
        Log.Information("Periodo: {Periodo}; ASE {AseId} ({Nombre})", solicitud.Periodo.CodigoCompleto, solicitud.Ase.Id, solicitud.Ase.NombreCompleto);

        if (string.IsNullOrWhiteSpace(solicitud.RutaR1) || string.IsNullOrWhiteSpace(solicitud.RutaR2) || string.IsNullOrWhiteSpace(solicitud.RutaR4))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.FuenteNoEncontrada, "Debe indicarse R1, R2 y R4 para ejecutar el proceso.");
        }

        if (string.IsNullOrWhiteSpace(solicitud.RutaPlantilla) || string.IsNullOrWhiteSpace(solicitud.RutaSalida))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.Plantilla, "Debe indicarse plantilla y ruta de salida.");
        }

        // Plan 26 (T2, R-F-5/S5): guardrail de existencia de R1/R2/R4 con el mismo código y
        // formateador del preflight de período, ANTES de leer nada. El flujo single-ASE NO consume
        // banco/balance/conciliaciones/R10/saldos (§V7): solo verifica estas 3 rutas explícitas.
        // Plan 27 (T2, D-F): además de existir, cada ruta se firma-chequea con el mismo helper;
        // un archivo no-Excel se reporta como ítem administrativo (nunca se intenta leer).
        var faltantes = new List<InsumoFaltante>();
        AgregarSiNoDisponible(solicitud.RutaR1, solicitud.Ase.Id, ReporteInsumoAse.R1, faltantes);
        AgregarSiNoDisponible(solicitud.RutaR2, solicitud.Ase.Id, ReporteInsumoAse.R2, faltantes);
        AgregarSiNoDisponible(solicitud.RutaR4, solicitud.Ase.Id, ReporteInsumoAse.R4, faltantes);

        if (faltantes.Count > 0)
        {
            Log.Warning(
                "Preflight single-ASE: faltan o son inválidos {CantidadInsumos} insumo(s); no se inicia el procesamiento.",
                faltantes.Count);
            throw new ArchivoFuenteNoEncontradoException(
                CodigoError.FuenteNoEncontrada,
                FormateadorInsumosFaltantes.MensajeAse(solicitud.Periodo, solicitud.Ase, faltantes));
        }

        progreso?.Report("Leyendo R1, R2 y R4...");
        Log.Information("Leyendo R1, R2 y R4...");
        var r1 = _recaudoReader.LeerR1(solicitud.RutaR1);
        var r2 = _recaudoReader.LeerR2(solicitud.RutaR2);
        var r4 = _recaudoReader.LeerR4(solicitud.RutaR4);

        progreso?.Report("Calculando consolidado del ASE...");
        Log.Information("Calculando consolidado del ASE...");
        var resultado = _calculoRemuneracion.CalcularConsolidado(solicitud.Periodo, [(solicitud.Ase, r1, r2, r4)]);

        progreso?.Report("Leyendo inputs leaf del workbook...");
        Log.Information("Leyendo inputs leaf del workbook...");
        var leaf = _leafReader.LeerLeafInputs(solicitud.Ase, solicitud.Periodo, solicitud.RutaR1, solicitud.RutaR2, solicitud.RutaR4);

        progreso?.Report("Validando coherencia básica...");
        Log.Information("Validando coherencia básica...");
        var errores = _validador.Validar(resultado, leaf);
        if (errores.Count > 0)
        {
            var detalle = string.Join("; ", errores);
            throw new CalculoInvalidoException(CodigoError.Validacion, $"La validación básica falló: {detalle}");
        }

        progreso?.Report("Generando workbook de salida...");
        Log.Information("Generando workbook de salida...");
        _workbookLeafWriter.GenerarWorkbook(solicitud.RutaPlantilla, solicitud.RutaSalida, resultado, leaf);

        progreso?.Report("Proceso completado correctamente.");
        Log.Information("Proceso completado correctamente.");

        return new ResultadoProcesoAse
        {
            Resultado = resultado,
            Leaf = leaf,
            RutaSalida = solicitud.RutaSalida
        };
    }

    /// <summary>
    /// Plan 27 (T2, D-F): si la ruta no existe, agrega el ítem de faltante vigente (Plan 26); si
    /// existe pero su firma no es de Excel, agrega el ítem administrativo de archivo inválido
    /// (D-E). Un archivo no-Excel equivale a un insumo no utilizable: mismo código y formateador.
    /// </summary>
    private static void AgregarSiNoDisponible(
        string ruta,
        int aseId,
        ReporteInsumoAse reporte,
        List<InsumoFaltante> faltantes)
    {
        if (!File.Exists(ruta))
        {
            faltantes.Add(FormateadorInsumosFaltantes.ReporteAseRuta(aseId, ruta, reporte));
            return;
        }

        var resultado = InspectorFirmaExcel.Inspeccionar(ruta);
        if (resultado.Firma == FirmaExcel.Valida)
        {
            return;
        }

        Log.Warning(
            "Preflight single-ASE: el archivo {Archivo} no tiene firma de Excel ({Firma}); primeros bytes {PrimerosBytes}.",
            ruta,
            resultado.Firma,
            resultado.PrimerosBytes.Length == 0 ? "(vacío)" : Convert.ToHexString(resultado.PrimerosBytes));

        faltantes.Add(resultado.Firma == FirmaExcel.NoExcelPareceCopiaWeb
            ? FormateadorInsumosFaltantes.ArchivoNoEsExcelPareceWeb($"ASE {aseId}", ruta, $"la ruta \"{ruta}\"")
            : FormateadorInsumosFaltantes.ArchivoNoEsExcelFormatoDesconocido($"ASE {aseId}", ruta, $"la ruta \"{ruta}\""));
    }
}
