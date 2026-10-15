using System.IO;
using Remuneracion.Cli;
using Remuneracion.Core.Errors;
using Serilog;
using Serilog.Core;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 42 (T3, R-CLI-1, D-G): cierre no interactivo con reconocimiento explícito por id de entrada
/// (<c>--aceptar-divergencia-declarada</c>) y código de salida DISTINTO de <c>Ok = 0</c>
/// (<see cref="CodigosSalida.DivergenciaDeclaradaReconocida"/>). Sin el flag, la declarada también
/// bloquea (FAIL <see cref="CodigosSalida.Validacion"/>). La salida textual enumera lo mismo que la
/// UI (celda, Δ, dueño, reporte emitido). Colección no paralela: el CLI reconfigura el Log.Logger.
/// </summary>
[Collection("Cli")]
public sealed class DivergenciaDeclaradaCliTests : IDisposable
{
    private readonly StringWriter _stdout = new();
    private readonly StringWriter _stderr = new();

    public void Dispose()
    {
        Log.CloseAndFlush();
        Log.Logger = Logger.None;
        _stdout.Dispose();
        _stderr.Dispose();
    }

    [Fact]
    public void T3_Flag_ParseaElIdDeLaEntrada()
    {
        var opciones = OpcionesCli.Parse(
        [
            "--periodo", "2026072", "--carpeta", @"C:\periodo", "--plantilla", @"C:\oficial.xlsx", "--salida", @"C:\salida",
            "--aceptar-divergencia-declarada", "OCCIDENTE-ASE5"
        ]);

        Assert.Equal(["OCCIDENTE-ASE5"], opciones.DivergenciasDeclaradasAceptadas);
    }

    [Fact]
    public void T3_Flag_EsRepetible()
    {
        var opciones = OpcionesCli.Parse(
        [
            "--periodo", "2026072", "--carpeta", @"C:\periodo", "--plantilla", @"C:\oficial.xlsx", "--salida", @"C:\salida",
            "--aceptar-divergencia-declarada", "A", "--aceptar-divergencia-declarada", "B"
        ]);

        Assert.Equal(["A", "B"], opciones.DivergenciasDeclaradasAceptadas);
    }

    [Fact]
    public void T3_Flag_SinValor_ErrorDeUso()
    {
        Assert.Throws<UsoCliException>(() => OpcionesCli.Parse(
            ["--periodo", "2026072", "--carpeta", "c", "--plantilla", "p", "--salida", "s", "--aceptar-divergencia-declarada"]));
    }

    [Fact]
    public void T3_Ayuda_MencionaElFlagYElCodigoDistinto()
    {
        var uso = OpcionesCli.Uso();
        Assert.Contains("--aceptar-divergencia-declarada", uso, StringComparison.Ordinal);
        Assert.Contains("divergencia declarada reconocida", uso, StringComparison.Ordinal);
        Assert.Contains("6", uso, StringComparison.Ordinal);
    }

    [Fact]
    public void T3_Q2_ConFlag_CierraConCodigoDistintoYReporte()
    {
        var salidaDir = DirectorioTemp("p42-cli-declarada");

        var codigo = EjecutorCli.Ejecutar(
            ["--periodo", "2026072", "--carpeta", Insumos.CarpetaPeriodoQ2, "--plantilla", Insumos.PlantillaOficial(),
             "--salida", salidaDir, "--cinco-ase", "--sobrescribir",
             "--aceptar-divergencia-declarada", "OCCIDENTE-ASE5"],
            _stdout, _stderr);

        Assert.Equal(6, codigo);
        Assert.Equal(CodigosSalida.DivergenciaDeclaradaReconocida, codigo);

        var salida = Path.Combine(salidaDir, Insumos.PeriodoQ2().NombreArchivo);
        Assert.True(File.Exists(salida), "El cierre con reconocimiento debe producir el archivo.");

        var reporte = Path.Combine(salidaDir, "Remuneración 202607-2 Total - Divergencias declaradas.txt");
        Assert.True(File.Exists(reporte), "El cierre con reconocimiento debe emitir el reporte durable.");
        Assert.Contains("REPORTE DE DIVERGENCIAS DECLARADAS", File.ReadAllText(reporte), StringComparison.Ordinal);

        var stdout = _stdout.ToString();
        Assert.Contains("OCCIDENTE-ASE5", stdout, StringComparison.Ordinal);
        Assert.Contains("VALIDACION_TOTAL", stdout, StringComparison.Ordinal);
        Assert.Contains("417", stdout, StringComparison.Ordinal); // Δ
        Assert.Contains("UAESP", stdout, StringComparison.Ordinal);
        Assert.Contains("Reporte emitido", stdout, StringComparison.Ordinal);
        Assert.Contains("codigo=6", stdout, StringComparison.Ordinal);
    }

    [Fact]
    public void T3_Q2_SinFlag_BloqueaConValidacion()
    {
        var salidaDir = DirectorioTemp("p42-cli-sinflag");

        var codigo = EjecutorCli.Ejecutar(
            ["--periodo", "2026072", "--carpeta", Insumos.CarpetaPeriodoQ2, "--plantilla", Insumos.PlantillaOficial(),
             "--salida", salidaDir, "--cinco-ase", "--sobrescribir"],
            _stdout, _stderr);

        Assert.Equal(CodigosSalida.Validacion, codigo);
        var stderr = _stderr.ToString();
        Assert.Contains("OCCIDENTE-ASE5", stderr, StringComparison.Ordinal);
        Assert.Contains("VALIDACION_TOTAL", stderr, StringComparison.Ordinal);
        Assert.Contains("--aceptar-divergencia-declarada", stderr, StringComparison.Ordinal);
        Assert.Contains("RESULTADO ERROR codigo=ERR-VALIDACION salida=1", _stdout.ToString(), StringComparison.Ordinal);
    }

    private static string DirectorioTemp(string etiqueta)
    {
        var dir = Path.Combine(Path.GetTempPath(), $"remuneracion-{etiqueta}-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }
}
