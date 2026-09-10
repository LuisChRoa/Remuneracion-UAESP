using Remuneracion.Cli;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-15 (Req 1-2): parseo PURO de <see cref="OpcionesCli"/> (sin I/O): matriz de flags §2.1,
/// defaults (5 ASE), exclusión de modo, quincena obligatoria si el período no la trae y los
/// errores de uso → <see cref="UsoCliException"/> (el ejecutor los traduce a stderr + salida 4).
/// </summary>
public sealed class OpcionesCliTests
{
    private static string[] BaseValida() =>
        ["--periodo", "2026071", "--carpeta", @"C:\periodo", "--plantilla", @"C:\plantilla.xlsx", "--salida", @"C:\salida"];

    [Fact]
    public void Ayuda_Help_True_IgnoraElResto()
    {
        var opciones = OpcionesCli.Parse(["--periodo", "basura", "--help"]);
        Assert.True(opciones.Ayuda);
    }

    [Fact]
    public void Ayuda_Help_True_Gana_ConArgsInvalidosTrasEl()
    {
        // HU-15 §2.1 (W-1): --help presente gana con ayuda y SIN error de uso aunque haya
        // args inválidos DESPUÉS de él (p. ej. `--help --ase 9`): el early-return corta
        // antes de validar el resto → Ayuda y sin UsoCliException (que se traduciría a 4).
        var opciones = OpcionesCli.Parse(["--help", "--ase", "9"]);
        Assert.True(opciones.Ayuda);
    }

    [Fact]
    public void Ayuda_H_True()
    {
        Assert.True(OpcionesCli.Parse(["-h"]).Ayuda);
    }

    [Fact]
    public void Defaults_SinAseNiCincoAse_Es5Ase()
    {
        var opciones = OpcionesCli.Parse(BaseValida());
        Assert.Null(opciones.AseId);
        Assert.True(opciones.CincoAse);
        Assert.False(opciones.Ayuda);
        Assert.False(opciones.Sobrescribir);
    }

    [Fact]
    public void PeriodoConQuincena_ResuelveCodigoCompleto()
    {
        var opciones = OpcionesCli.Parse(BaseValida());
        Assert.Equal("2026071", opciones.PeriodoCodigoCompleto);
        Assert.Equal("202607", opciones.Periodo.CodigoAAAAMM);
        Assert.Equal(1, opciones.Periodo.NumeroQuincena);
    }

    [Fact]
    public void PeriodoAaaamm_QuincenaComplementaria_Obligatoria()
    {
        var opciones = OpcionesCli.Parse(["--periodo", "202607", "--quincena", "2", "--carpeta", "c", "--plantilla", "p", "--salida", "s"]);
        Assert.Equal("2026072", opciones.PeriodoCodigoCompleto);
        Assert.Equal(2, opciones.Periodo.NumeroQuincena);
    }

    [Fact]
    public void PeriodoConQuincena_QuincenaCoincidente_EsComplementoValido()
    {
        var opciones = OpcionesCli.Parse(["--periodo", "2026071", "--quincena", "1", "--carpeta", "c", "--plantilla", "p", "--salida", "s"]);
        Assert.Equal("2026071", opciones.PeriodoCodigoCompleto);
    }

    [Fact]
    public void AseExplicito_Modo1Ase()
    {
        var opciones = OpcionesCli.Parse([.. BaseValida(), "--ase", "3"]);
        Assert.Equal(3, opciones.AseId);
        Assert.False(opciones.CincoAse);
        Assert.Equal("ASE 3", opciones.ModoTexto);
    }

    [Fact]
    public void CincoAseExplicito_Modo5Ase()
    {
        var opciones = OpcionesCli.Parse([.. BaseValida(), "--cinco-ase"]);
        Assert.True(opciones.CincoAse);
        Assert.Null(opciones.AseId);
        Assert.Equal("5 ASE", opciones.ModoTexto);
    }

    [Fact]
    public void Sobrescribir_True()
    {
        Assert.True(OpcionesCli.Parse([.. BaseValida(), "--sobrescribir"]).Sobrescribir);
    }

    // ── Errores de uso (→ stderr + ayuda + salida 4, G6) ─────────────────────────────────────

    [Theory]
    [InlineData("2026")]
    [InlineData("2026070")]
    [InlineData("20260712")]
    [InlineData("20260A1")]
    public void Uso_PeriodoMalformado_LanzaUso(string periodo)
    {
        var ex = Assert.Throws<UsoCliException>(() => OpcionesCli.Parse(["--periodo", periodo, "--quincena", "1", "--carpeta", "c", "--plantilla", "p", "--salida", "s"]));
        Assert.Contains("período", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Uso_PeriodoAaaammSinQuincena_LanzaUso()
    {
        var ex = Assert.Throws<UsoCliException>(() => OpcionesCli.Parse(["--periodo", "202607", "--carpeta", "c", "--plantilla", "p", "--salida", "s"]));
        Assert.Contains("--quincena", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Uso_QuincenaDifiereDelPeriodo_LanzaUso()
    {
        var ex = Assert.Throws<UsoCliException>(() => OpcionesCli.Parse(["--periodo", "2026071", "--quincena", "2", "--carpeta", "c", "--plantilla", "p", "--salida", "s"]));
        Assert.Contains("difiere", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Uso_QuincenaInvalida_LanzaUso()
    {
        var ex = Assert.Throws<UsoCliException>(() => OpcionesCli.Parse(["--periodo", "202607", "--quincena", "3", "--carpeta", "c", "--plantilla", "p", "--salida", "s"]));
        Assert.Contains("--quincena", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Uso_AseFueraDeRango_LanzaUso()
    {
        var ex = Assert.Throws<UsoCliException>(() => OpcionesCli.Parse([.. BaseValida(), "--ase", "6"]));
        Assert.Contains("1 y 5", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Uso_AseNoNumerico_LanzaUso()
    {
        Assert.Throws<UsoCliException>(() => OpcionesCli.Parse([.. BaseValida(), "--ase", "tres"]));
    }

    [Fact]
    public void Uso_AseYCincoAse_MutuamenteExcluyentes()
    {
        Assert.Throws<UsoCliException>(() => OpcionesCli.Parse([.. BaseValida(), "--ase", "2", "--cinco-ase"]));
        Assert.Throws<UsoCliException>(() => OpcionesCli.Parse([.. BaseValida(), "--cinco-ase", "--ase", "2"]));
    }

    [Fact]
    public void Uso_FlagDesconocido_LanzaUso()
    {
        var ex = Assert.Throws<UsoCliException>(() => OpcionesCli.Parse([.. BaseValida(), "--nada"]));
        Assert.Contains("--nada", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Uso_FaltanFlagsObligatorios_LanzanUso()
    {
        Assert.Throws<UsoCliException>(() => OpcionesCli.Parse(["--carpeta", "c", "--plantilla", "p", "--salida", "s"]));
        Assert.Throws<UsoCliException>(() => OpcionesCli.Parse(["--periodo", "2026071", "--plantilla", "p", "--salida", "s"]));
        Assert.Throws<UsoCliException>(() => OpcionesCli.Parse(["--periodo", "2026071", "--carpeta", "c", "--salida", "s"]));
        Assert.Throws<UsoCliException>(() => OpcionesCli.Parse(["--periodo", "2026071", "--carpeta", "c", "--plantilla", "p"]));
    }

    [Fact]
    public void Uso_FlagSinValor_LanzaUso()
    {
        Assert.Throws<UsoCliException>(() => OpcionesCli.Parse(["--periodo"]));
        Assert.Throws<UsoCliException>(() => OpcionesCli.Parse([.. BaseValida(), "--ase"]));
    }

    [Fact]
    public void Uso_TextoTieneLosFlagsDelContrato()
    {
        var uso = OpcionesCli.Uso();
        Assert.Contains("--periodo", uso, StringComparison.Ordinal);
        Assert.Contains("--carpeta", uso, StringComparison.Ordinal);
        Assert.Contains("--plantilla", uso, StringComparison.Ordinal);
        Assert.Contains("--salida", uso, StringComparison.Ordinal);
        Assert.Contains("--ase", uso, StringComparison.Ordinal);
        Assert.Contains("--cinco-ase", uso, StringComparison.Ordinal);
        Assert.Contains("--sobrescribir", uso, StringComparison.Ordinal);
        Assert.Contains("--help", uso, StringComparison.Ordinal);
    }
}
