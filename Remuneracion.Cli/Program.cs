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
        try
        {
            var codigo = EjecutorCli.Ejecutar(args);
            Log.CloseAndFlush();
            Environment.Exit(codigo);
        }
        catch (Exception ex)
        {
            // HU-17 (S-1 HU-15): red de seguridad externa — cualquier excepción que escape del
            // contrato (p. ej. ConfigurarSerilog con el sink bloqueado) jamás termina con un
            // código no-contract ni pierde el evento final: se registra, se hace flush y se sale
            // con 4 (ERR-INESPERADO). El contrato 0-5 sigue viviendo en EjecutorCli.
            Log.Error(ex, "[{Codigo}] Error inesperado fuera del contrato: {Mensaje}", Remuneracion.Core.Errors.CodigoError.Inesperado, ex.Message);
            Log.CloseAndFlush();
            Environment.Exit(Remuneracion.Core.Errors.CodigosSalida.Inesperado);
        }
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
        // HU-17 (S-2 HU-15): dispose del logger previo antes de reconfigurar — evita fuga de
        // sinks (el archivo queda abierto) en reconfiguraciones (tests in-process y doble uso).
        Log.CloseAndFlush();
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
