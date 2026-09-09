using Serilog;

namespace Remuneracion.Cli;

internal static class Program
{
    /// <summary>
    /// HU-15 (3.3, D1): punto de entrada del CLI. El ÚNICO <c>Environment.Exit</c> del PRODUCTO
    /// (V2) vive AQUÍ, SIEMPRE después de <c>Log.CloseAndFlush()</c> (Exit no ejecuta finally;
    /// el flush evita perder el evento final del log).
    /// </summary>
    private static void Main(string[] args)
    {
        var codigo = EjecutorCli.Ejecutar(args);
        Log.CloseAndFlush();
        Environment.Exit(codigo);
    }

    /// <summary>
    /// HU-15 (D8): configuración Serilog del CLI — MISMO rolling/template que el CANÓNICO
    /// <c>Remuneracion.WinForms.Form1.ConfigurarSerilog</c> (HU-15 S-3: incluye {Properties}).
    /// Core no referencia Serilog.Sinks.File (ADR HU-14) y no vale un helper compartido para
    /// 12 líneas; la paridad CLI↔UI (HU-15 §2.7) detecta cualquier divergencia.
    /// Internal + InternalsVisibleTo para los tests in-process (D3).
    /// </summary>
    internal static void ConfigurarSerilog()
    {
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .Enrich.FromLogContext() // D5: LogContext (RunId/Periodo/Modo) adjunta las propiedades a cada evento
            .WriteTo.File(
                "remuneracion_log_.txt",
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 30,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] (RunId={RunId} Periodo={Periodo} AseId={AseId} Hoja={Hoja} Validacion={Validacion}) {Message:lj}{NewLine}{Exception}{Properties}{NewLine}")
            .CreateLogger();
    }
}