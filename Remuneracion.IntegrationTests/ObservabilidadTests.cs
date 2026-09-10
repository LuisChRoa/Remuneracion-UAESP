using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-14 (3.2): observabilidad con sink Serilog en memoria propio (sin paquetes nuevos:
/// Serilog core ya está en el árbol). Verifica Req 4 y Req 5: niveles por evento
/// (Information hitos / Debug detalle), propiedades buscables (RunId/AseId/Validacion),
/// W-3 (Validacion = nombre real, nunca "cruzada") y correlación RunId inicio→fin de una
/// ejecución real de período (CA-6).
///
/// La colección desactiva paralelismo para no contaminar el sink con otras suites que usan
/// el Log.Logger global; además los asserts filtran por RunId de la propia ejecución.
/// </summary>
[Collection("Observabilidad")]
public sealed class ObservabilidadTests : IDisposable
{
    private readonly CapturaEventos _sink = new();
    private readonly Logger _logger;

    public ObservabilidadTests()
    {
        _logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .Enrich.FromLogContext() // D5: LogContext (RunId) adjunta las propiedades a cada evento
            .WriteTo.Sink(_sink)
            .CreateLogger();
        Log.Logger = _logger;
    }

    public void Dispose()
    {
        Log.CloseAndFlush();
        Log.Logger = Logger.None;
    }

    [Fact]
    public void EjecucionPeriodo_RunIdCorrelacionaInicioAFin_YAseId()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-obs-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var procesador = CrearProcesadorConOracle();
        procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.Periodo(),
            CarpetaPeriodo = Insumos.CarpetaPeriodo,
            RutaPlantilla = Insumos.Plantilla,
            RutaSalida = salida
        });

        var inicio = _sink.Eventos.First(e => e.MessageTemplate.Text.Contains("Iniciando proceso del período (5 ASE).", StringComparison.Ordinal));
        var runId = RunIdDe(inicio);
        Assert.NotEqual(Guid.Empty, runId);

        // El evento de fin comparte el MISMO RunId (correlación inicio→fin, Req 4).
        var fin = _sink.Eventos.First(e =>
            e.MessageTemplate.Text.Contains("Proceso del período completado correctamente.", StringComparison.Ordinal)
            && RunIdDe(e) == runId);
        Assert.NotNull(fin);

        // Un evento por ASE (AseId como propiedad) comparte el RunId de la ejecución.
        var porAse = _sink.Eventos.First(e =>
            e.Properties.TryGetValue("AseId", out var _) && RunIdDe(e) == runId);
        Assert.NotNull(porAse);
    }

    [Fact]
    public void BloqueValidaciones_PropiedadValidacion_NombreReal_ConAseId()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-obs-w3-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var procesador = CrearProcesadorConOracle();
        procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.Periodo(),
            CarpetaPeriodo = Insumos.CarpetaPeriodo,
            RutaPlantilla = Insumos.Plantilla,
            RutaSalida = salida
        });

        // W-3: la propiedad Validacion lleva el nombre REAL (VALIDACION_ENEL), no "cruzada".
        var eventoEnel = _sink.Eventos.First(e =>
            e.Properties.TryGetValue("Validacion", out var v)
            && v is ScalarValue { Value: string texto }
            && texto == "VALIDACION_ENEL");
        Assert.NotNull(eventoEnel);
        Assert.True(eventoEnel.Properties.TryGetValue("AseId", out var aseIdProp), "El evento de empresa debe portar AseId como propiedad.");
        Assert.Equal(1, ((ScalarValue)aseIdProp).Value);

        Assert.DoesNotContain(_sink.Eventos, e =>
            e.Properties.TryGetValue("Validacion", out var v)
            && v is ScalarValue { Value: string texto }
            && texto == "cruzada");

        Assert.Contains(_sink.Eventos, e =>
            e.Properties.TryGetValue("Validacion", out var v)
            && v is ScalarValue { Value: string texto }
            && texto == "VALIDACION_TOTAL");
        Assert.Contains(_sink.Eventos, e =>
            e.Properties.TryGetValue("Validacion", out var v)
            && v is ScalarValue { Value: string texto }
            && texto == "DetValiRetri");
    }

    [Fact]
    public void Niveles_HitosInformation_DetalleDebug()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-obs-niv-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var procesador = CrearProcesadorConOracle();
        procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.Periodo(),
            CarpetaPeriodo = Insumos.CarpetaPeriodo,
            RutaPlantilla = Insumos.Plantilla,
            RutaSalida = salida
        });

        // Hitos = Information (§2.1).
        Assert.Equal(LogEventLevel.Information, _sink.Eventos.First(e => e.MessageTemplate.Text.Contains("Iniciando proceso del período (5 ASE).", StringComparison.Ordinal)).Level);
        Assert.Equal(LogEventLevel.Information, _sink.Eventos.First(e => e.MessageTemplate.Text.Contains("Proceso del período completado correctamente.", StringComparison.Ordinal)).Level);

        // Detalle por empresa/validación = Debug (nunca Information), Req 4.
        var detalle = _sink.Eventos.First(e =>
            e.Properties.TryGetValue("Validacion", out var v)
            && v is ScalarValue { Value: string texto }
            && texto == "VALIDACION_ENEL");
        Assert.Equal(LogEventLevel.Debug, detalle.Level);
    }

    private static Guid RunIdDe(LogEvent evento) =>
        evento.Properties.TryGetValue("RunId", out var v) && v is ScalarValue { Value: Guid guid }
            ? guid
            : throw new Xunit.Sdk.XunitException("El evento no porta RunId (Guid).");

    private static ProcesadorPeriodo CrearProcesadorConOracle() =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ValidacionOracleReader());

    private sealed class CapturaEventos : ILogEventSink
    {
        public List<LogEvent> Eventos { get; } = [];

        public void Emit(LogEvent logEvent) => Eventos.Add(logEvent);
    }
}

[CollectionDefinition("Observabilidad", DisableParallelization = true)]
public sealed class ObservabilidadCollection
{
}
