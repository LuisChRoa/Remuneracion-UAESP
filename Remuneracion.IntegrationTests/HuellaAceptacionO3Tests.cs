using System.IO;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.WinForms;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 43 (O-3): la huella de quién aceptó (usuario del sistema) y por qué canal (UI o CLI) queda
/// registrada en el REPORTE durable y en el LOG, además de celda, observado, esperado, Δ, dueño,
/// evidencia, fecha, condición de retiro y RunId. Cuando el dato no llega, se resuelve de forma
/// honesta (nunca se inventa). Colección no paralela: la prueba de log usa el Log.Logger global.
/// </summary>
[Collection("Cli")]
public sealed class HuellaAceptacionO3Tests : IDisposable
{
    public void Dispose()
    {
        Log.CloseAndFlush();
        Log.Logger = Logger.None;
    }

    [Fact]
    public void O3_Reporte_ConstruirTexto_RegistraUsuarioYCanal()
    {
        var texto = ReporteDivergenciasDeclaradasWriter.ConstruirTexto(
            @"C:\salidas\Remuneración 202607-2 Total.xlsx",
            Periodo.Parse("2026072"),
            Guid.NewGuid(),
            ClasificacionDeclarada(),
            usuario: "tester-o3",
            canal: HuellaAceptacion.CanalCli);

        Assert.Contains("Aceptado por: tester-o3 (canal: CLI).", texto, StringComparison.Ordinal);
    }

    [Fact]
    public void O3_Reporte_SinDatos_UsaFallbackHonesto()
    {
        var texto = ReporteDivergenciasDeclaradasWriter.ConstruirTexto(
            @"C:\salidas\Remuneración 202607-2 Total.xlsx",
            Periodo.Parse("2026072"),
            Guid.NewGuid(),
            ClasificacionDeclarada());

        // Sin canal declarado → "No especificado"; sin usuario → el del sistema (Environment.UserName).
        Assert.Contains(
            $"Aceptado por: {Environment.UserName} (canal: {HuellaAceptacion.CanalIndefinido}).",
            texto,
            StringComparison.Ordinal);
    }

    [Fact]
    public void O3_Emitir_EscribeReporteConUsuarioYCanal()
    {
        var dir = Path.Combine(Path.GetTempPath(), "o3-reporte-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var salida = Path.Combine(dir, Periodo.Parse("2026072").NombreArchivo);

            var ruta = new ReporteDivergenciasDeclaradasWriter().Emitir(
                salida,
                Periodo.Parse("2026072"),
                Guid.NewGuid(),
                ClasificacionDeclarada(),
                usuario: "tester-o3",
                canal: HuellaAceptacion.CanalCli);

            Assert.Contains(
                "Aceptado por: tester-o3 (canal: CLI).",
                File.ReadAllText(ruta),
                StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void O3_Log_Huella_RegistraUsuarioYCanalComoPropiedadesEstructuradas()
    {
        var sink = new CapturaEventos();
        Log.Logger = new LoggerConfiguration().MinimumLevel.Debug().WriteTo.Sink(sink).CreateLogger();

        var dir = Path.Combine(Path.GetTempPath(), "o3-log-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var salida = Path.Combine(dir, Insumos.PeriodoQ2().NombreArchivo);
            var (_, periodo, _) = ComposicionServicios.Crear();

            periodo.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
                RutaPlantilla = Insumos.PlantillaOficial(),
                RutaSalida = salida,
                DivergenciasDeclaradasAceptadas = ["OCCIDENTE-ASE5"],
                Canal = HuellaAceptacion.CanalCli,
                UsuarioAceptacion = "tester-o3"
            });

            var evento = sink.Eventos.FirstOrDefault(e =>
                e.Properties.ContainsKey("Usuario") && e.Properties.ContainsKey("Canal"));

            Assert.NotNull(evento);
            Assert.Equal("tester-o3", ((ScalarValue)evento!.Properties["Usuario"]).Value);
            Assert.Equal(HuellaAceptacion.CanalCli, ((ScalarValue)evento.Properties["Canal"]).Value);
            var mensaje = evento.RenderMessage();
            Assert.Contains("Aceptado por", mensaje, StringComparison.Ordinal);
            Assert.Contains("tester-o3", mensaje, StringComparison.Ordinal);
            Assert.Contains(HuellaAceptacion.CanalCli, mensaje, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    private static ClasificacionDivergenciasDeclaradas ClasificacionDeclarada()
    {
        var veredictos = new List<VeredictoDivergenciaValidacion>();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            foreach (var celda in new[] { "O25", "D34", "F34" })
            {
                veredictos.Add(new VeredictoDivergenciaValidacion
                {
                    AseId = aseId,
                    Celda = celda,
                    Declarada = true,
                    EntradaId = "OCCIDENTE-ASE5",
                    MontoObservado = 72441184838.33m,
                    MontoEsperado = 72440767318.71m
                });
            }
        }

        return new ClasificacionDivergenciasDeclaradas { Veredictos = veredictos };
    }

    private sealed class CapturaEventos : ILogEventSink
    {
        public List<LogEvent> Eventos { get; } = [];

        public void Emit(LogEvent logEvent) => Eventos.Add(logEvent);
    }
}
