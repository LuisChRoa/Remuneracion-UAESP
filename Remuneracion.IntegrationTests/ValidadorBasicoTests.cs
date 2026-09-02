using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Xunit;

namespace Remuneracion.IntegrationTests;

public sealed class ValidadorBasicoTests
{
    private const decimal Tolerancia = 0.5m;

    [Fact]
    public void Validar_Q1HappyPath_ListaVacia()
    {
        var (resultado, leaf) = CrearCasoValido();
        var errores = new ValidadorBasico().Validar(resultado, leaf);
        Assert.Empty(errores);
    }

    [Fact]
    public void Validar_AjustesDistintoDeCero_Bloquea()
    {
        var (resultado, leaf) = CrearCasoValido();
        resultado.Consolidados[0].AjustesSfT = 10m;
        var errores = new ValidadorBasico().Validar(resultado, leaf);
        Assert.Contains(errores, e => e.Contains("AjustesSfT", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_MismatchR2_Bloquea()
    {
        var (resultado, leaf) = CrearCasoValido();
        leaf.R2.E15 += 10m;
        var errores = new ValidadorBasico().Validar(resultado, leaf);
        Assert.Contains(errores, e => e.Contains("R2", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_NoComparaTotOptHu02ContraF46()
    {
        var (resultado, leaf) = CrearCasoValido();
        resultado.Consolidados[0].TotOpt = 19556118465.99m;
        Assert.True(Math.Abs(leaf.R1.TotalOportunoEsperado - resultado.Consolidados[0].TotOpt) > Tolerancia);

        var errores = new ValidadorBasico().Validar(resultado, leaf);
        Assert.Empty(errores);
        Assert.DoesNotContain(errores, e => e.Contains("F46", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(errores, e => e.Contains("TotOpt", StringComparison.OrdinalIgnoreCase));
    }

    private static (ResultadoRemuneracion Resultado, WorkbookLeafInputs Leaf) CrearCasoValido()
    {
        var ase = new Ase { Id = 1, NombreCorto = "PROMOAMBIENTAL", NombreCompleto = "Promoambiental", NumeroCarpeta = 1 };
        var consolidado = new ConsolidadoAse
        {
            Ase = ase,
            TotOpt = 19556118465.99m,
            R2TotalOportuno = 54216385.68m,
            Extemp = 19549786950.62m,
            ReversionR4 = -12054255.65m,
            AjustesSfT = 0m
        };
        var resultado = new ResultadoRemuneracion
        {
            Periodo = new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 1 },
            Consolidados = [consolidado],
            Exitoso = true
        };
        var leaf = new WorkbookLeafInputs
        {
            Ase = ase,
            Periodo = resultado.Periodo,
            R1 = new WorkbookLeafInputsR1
            {
                F25 = 19549786950.62m,
                F41 = -2827260776.81m,
                L25 = 18193739.24m,
                F30 = 5341504.63m,
                F10 = 6331515.37m,
                L10 = 0m
            },
            R2 = new WorkbookLeafInputsR2
            {
                E15 = 56353887.23m,
                E26 = -2080239.85m,
                K15 = 57261.7m
            },
            R4 = new WorkbookLeafInputsR4
            {
                D9 = -12054255.65m,
                P9 = 0m
            }
        };
        return (resultado, leaf);
    }
}
