using System.IO;
using Remuneracion.Cli;
using Remuneracion.Core.Errors;
using Serilog;
using Serilog.Core;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-15 (Req 3, D3): matriz de códigos de salida 0-5 del CLI ejecutado IN-PROCESS
/// (<c>EjecutorCli.Ejecutar</c>, nunca subproceso ni <c>Environment.Exit</c>) con stdout/stderr
/// capturados. Colección no paralela: el CLI reconfigura el <c>Log.Logger</c> global (D8).
/// </summary>
[Collection("Cli")]
public sealed class CodigosSalidaCliTests : IDisposable
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

    private int Ejecutar(params string[] args) => EjecutorCli.Ejecutar(args, _stdout, _stderr);

    // ── 0: ayuda y ejecución correcta ─────────────────────────────────────────────────────────

    [Fact]
    public void Ayuda_Help_Devuelve0_UsoAStdout()
    {
        var codigo = Ejecutar("--help");
        Assert.Equal(CodigosSalida.Ok, codigo);
        Assert.Contains("Uso:", _stdout.ToString(), StringComparison.Ordinal);
        Assert.Contains("--periodo", _stdout.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Ayuda_H_Devuelve0()
    {
        Assert.Equal(CodigosSalida.Ok, Ejecutar("-h"));
    }

    [Fact]
    public void Ejecucion_5Ase_Q1_Devuelve0_ConArchivoYResultadoOk()
    {
        var salidaDir = DirectorioTemp("cli-ok");
        var codigo = Ejecutar(
            "--periodo", "2026071",
            "--carpeta", Insumos.CarpetaPeriodo,
            "--plantilla", Insumos.Plantilla,
            "--salida", salidaDir,
            "--cinco-ase",
            "--sobrescribir");

        Assert.Equal(CodigosSalida.Ok, codigo);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);
        Assert.True(File.Exists(salida), "Debe existir la salida certificada.");
        var stdout = _stdout.ToString();
        Assert.Contains("RESULTADO OK codigo=0", stdout, StringComparison.Ordinal);
        Assert.Contains("runId=", stdout, StringComparison.Ordinal);
        Assert.Contains(salida, stdout, StringComparison.Ordinal);
    }

    // ── 4: errores de uso (G6: stderr + ayuda, sin códigos nuevos) ────────────────────────────

    [Theory]
    [InlineData("--periodo", "2026")]
    [InlineData("--periodo", "2026070")]
    public void Uso_PeriodoMalformado_Devuelve4_ConUsoAStderr(string flag, string valor)
    {
        var codigo = Ejecutar(flag, valor, "--carpeta", "c", "--plantilla", "p", "--salida", "s");
        Assert.Equal(CodigosSalida.Inesperado, codigo);
        var stderr = _stderr.ToString();
        Assert.Contains("Uso:", stderr, StringComparison.Ordinal);
        Assert.Contains("--help", stderr, StringComparison.Ordinal);
        Assert.DoesNotContain("RESULTADO", _stdout.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Uso_PeriodoSinQuincena_Devuelve4()
    {
        Assert.Equal(CodigosSalida.Inesperado, Ejecutar("--periodo", "202607", "--carpeta", "c", "--plantilla", "p", "--salida", "s"));
        Assert.Contains("--quincena", _stderr.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Uso_Ase6_Devuelve4()
    {
        Assert.Equal(CodigosSalida.Inesperado, Ejecutar("--periodo", "2026071", "--carpeta", "c", "--plantilla", "p", "--salida", "s", "--ase", "6"));
    }

    [Fact]
    public void Uso_AseYCincoAse_Devuelve4()
    {
        Assert.Equal(CodigosSalida.Inesperado, Ejecutar("--periodo", "2026071", "--carpeta", "c", "--plantilla", "p", "--salida", "s", "--ase", "2", "--cinco-ase"));
    }

    // ── 2: fuente/plantilla ───────────────────────────────────────────────────────────────────

    [Fact]
    public void FuenteNoEncontrada_CarpetaAseIncompleta_Devuelve2_NombraAse()
    {
        // Carpeta con una sola subcarpeta ASE sin fuentes → fail-fast nombra el ASE 1 y su R1.
        var carpetaPeriodo = DirectorioTemp("cli-fuente");
        Directory.CreateDirectory(Path.Combine(carpetaPeriodo, "1-Promoambiental"));
        var salidaDir = DirectorioTemp("cli-fuente-salida");

        var codigo = Ejecutar("--periodo", "2026071", "--carpeta", carpetaPeriodo, "--plantilla", Insumos.Plantilla, "--salida", salidaDir);

        Assert.Equal(CodigosSalida.FuenteOPlantilla, codigo);
        var stderr = _stderr.ToString();
        Assert.Contains("[ERR-FUENTE-NO-ENCONTRADA]", stderr, StringComparison.Ordinal);
        Assert.Contains("ASE 1", stderr, StringComparison.Ordinal);
        Assert.Contains("RESULTADO ERROR codigo=ERR-FUENTE-NO-ENCONTRADA salida=2", _stdout.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void PlantillaIgualSalida_Devuelve2_ConLogError()
    {
        var carpetaSalida = DirectorioTemp("cli-inplace");
        var plantillaCopia = Path.Combine(carpetaSalida, Insumos.Periodo().NombreArchivo);
        File.Copy(Insumos.Plantilla, plantillaCopia, overwrite: true);

        var codigo = Ejecutar("--periodo", "2026071", "--carpeta", Insumos.CarpetaPeriodo, "--plantilla", plantillaCopia, "--salida", carpetaSalida);

        Assert.Equal(CodigosSalida.FuenteOPlantilla, codigo);
        Assert.Contains("[ERR-PLANTILLA]", _stderr.ToString(), StringComparison.Ordinal);
        Assert.Contains("RESULTADO ERROR codigo=ERR-PLANTILLA salida=2", _stdout.ToString(), StringComparison.Ordinal);
    }

    // ── 5: cancelación no-interactiva (sin prompt; Req 1) ─────────────────────────────────────

    [Fact]
    public void SalidaExistente_SinSobrescribir_Devuelve5_WarnCancelado()
    {
        var salidaDir = DirectorioTemp("cli-cancelado");
        File.WriteAllText(Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo), "placeholder");

        var codigo = Ejecutar("--periodo", "2026071", "--carpeta", Insumos.CarpetaPeriodo, "--plantilla", Insumos.Plantilla, "--salida", salidaDir);

        Assert.Equal(CodigosSalida.CanceladoPorUsuario, codigo);
        Assert.Contains("[WARN-CANCELADO]", _stderr.ToString(), StringComparison.Ordinal);
        Assert.Contains("RESULTADO ERROR codigo=WARN-CANCELADO salida=5", _stdout.ToString(), StringComparison.Ordinal);
    }

    // ── 3: escritura bloqueada → ERR-ESCRITURA con inner ──────────────────────────────────────

    [Fact]
    public void EscrituraBloqueada_Devuelve3_ConInnerException()
    {
        var salidaDir = DirectorioTemp("cli-escritura");
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);
        using (var bloqueo = new FileStream(salida, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
        {
            var codigo = Ejecutar(
                "--periodo", "2026071",
                "--carpeta", Insumos.CarpetaPeriodo,
                "--plantilla", Insumos.Plantilla,
                "--salida", salidaDir,
                "--ase", "1",
                "--sobrescribir");

            Assert.Equal(CodigosSalida.Escritura, codigo);
            Assert.Contains("[ERR-ESCRITURA]", _stderr.ToString(), StringComparison.Ordinal);
            Assert.Contains("RESULTADO ERROR codigo=ERR-ESCRITURA salida=3", _stdout.ToString(), StringComparison.Ordinal);
        }
    }

    // ── 1: validación (red de seguridad Q2 del modo 1-ASE, sin tocar cálculo) ─────────────────

    [Fact]
    public void Quincena2_Modo1Ase_Devuelve1_Validacion()
    {
        // El modo 1-ASE usa el overload sin ajustes (red de seguridad HU-11): en Q2 falla con
        // ERR-VALIDACION (salida 1) exactamente igual que Form1.EjecutarModoUnAse. El path Q2
        // completo (AJUSTES-SF-T/DetRetri) vive en el modo período (ParidadCliTests).
        var salidaDir = DirectorioTemp("cli-q2-single");
        var codigo = Ejecutar(
            "--periodo", "2026072",
            "--carpeta", Insumos.CarpetaPeriodoQ2,
            "--plantilla", Insumos.PlantillaQ2,
            "--salida", salidaDir,
            "--ase", "1");

        Assert.Equal(CodigosSalida.Validacion, codigo);
        Assert.Contains("[ERR-VALIDACION]", _stderr.ToString(), StringComparison.Ordinal);
        Assert.Contains("RESULTADO ERROR codigo=ERR-VALIDACION salida=1", _stdout.ToString(), StringComparison.Ordinal);
    }

    private static string DirectorioTemp(string etiqueta)
    {
        var dir = Path.Combine(Path.GetTempPath(), $"remuneracion-{etiqueta}-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }
}

[CollectionDefinition("Cli", DisableParallelization = true)]
public sealed class CliCollection
{
}