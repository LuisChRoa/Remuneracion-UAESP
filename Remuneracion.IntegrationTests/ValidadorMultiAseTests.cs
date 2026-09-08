using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-07 Req 4 / plan §2.5: validador multi-ASE con matcheo estricto por Ase.Id (in-memory, sin Excel).
/// </summary>
public sealed class ValidadorMultiAseTests
{
    private const decimal Tolerancia = 0.5m;

    [Fact]
    public void Validar_CincoAseHappyPath_ListaVacia()
    {
        var (resultado, leafs) = CrearCasoValido();
        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Empty(errores);
    }

    [Fact]
    public void Validar_ModoPeriodoConUnSoloAse_ErrorExplícito()
    {
        var (resultado, leafs) = CrearCasoValido();
        var errores = new ValidadorBasico().Validar(resultado, [leafs[0]]);
        Assert.Contains(errores, e => e.Contains("exige exactamente 5", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(errores, e => e.Contains("single-ASE", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_FaltanAse_ErrorQueNombraLaCantidad()
    {
        var (resultado, leafs) = CrearCasoValido();
        var errores = new ValidadorBasico().Validar(resultado, leafs.Take(3).ToList());
        Assert.Contains(errores, e => e.Contains("3", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_LeafDeAse3SinConsolidado_ErrorQueNombraElAse()
    {
        var (resultado, leafs) = CrearCasoValido();
        resultado.Consolidados.RemoveAt(2); // quita ASE3
        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("ASE 3", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_MismatchR4SoloEnAse3_ErrorQueNombraElAse3()
    {
        var (resultado, leafs) = CrearCasoValido();
        leafs[2].R4.D9 += 50m; // rompe solo ASE3
        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("ASE 3", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(errores, e => e.Contains("ASE 4", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_MismatchR2SoloEnAse4_ErrorQueNombraElAse4()
    {
        var (resultado, leafs) = CrearCasoValido();
        leafs[3].R2.E26 += 100m; // rompe solo ASE4
        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("ASE 4", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(errores, e => e.Contains("ASE 1", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_AjustesSfTDistintoDeCero_Bloquea()
    {
        var (resultado, leafs) = CrearCasoValido();
        resultado.Consolidados[1].AjustesSfT = 10m;
        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("AjustesSfT", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_Regla5_GranTotalEsSumaDeTotalAse_SiempreConsistente()
    {
        // §2.5 regla 5: GranTotal == Σ TotalAse (±0.5). GranTotal es computed => estructuralmente
        // consistente; el test fija que no genera falso positivo.
        var (resultado, leafs) = CrearCasoValido();
        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Empty(errores);
        Assert.Equal(resultado.GranTotal, resultado.Consolidados.Sum(c => c.TotalAse));
    }

    [Fact]
    public void Validar_LeafsConAseDuplicado_ErrorPorDuplicado()
    {
        var (resultado, leafs) = CrearCasoValido();
        var duplicados = leafs.Append(leafs[0]).ToList();
        var errores = new ValidadorBasico().Validar(resultado, duplicados);
        Assert.Contains(errores, e => e.Contains("duplicado", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_MatcheoEstrictoEliminaFallback_ConsolidadoAse3NuncaValidaContraLeafAse1()
    {
        // Regresión del fallback FirstOrDefault: si ASE1 no tiene consolidado, el gate NO debe
        // validar ASE1 contra ASE2/ASE3. Con matcheo estricto el error nombra ASE 1.
        var (resultado, leafs) = CrearCasoValido();
        resultado.Consolidados.RemoveAt(0); // quita ASE1
        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("ASE 1", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_NoComparaTotOptHu02ContraVisiblesR1_PorAse()
    {
        // Prohibido: TotOpt HU-02 vs visibles R1. Cambiar TotOpt no debe generar error.
        var (resultado, leafs) = CrearCasoValido();
        resultado.Consolidados[0].TotOpt = 999_999_999_999m;
        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.DoesNotContain(errores, e => e.Contains("TotOpt", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(errores, e => e.Contains("F46", StringComparison.OrdinalIgnoreCase));
    }

    private static (ResultadoRemuneracion Resultado, List<WorkbookLeafInputs> Leafs) CrearCasoValido()
    {
        // Valores golden Q1 §2.1 (cache del golden — referencia).
        var valores = new[]
        {
            (1, 16704332434.57m, 54216385.68m, 11673020m, -12054255.65m),
            (2, 12157780441.19m, 79400801.26m, 0m, -9889189.72m),
            (3, 10101988514.82m, 31111803.75m, 10198723.07m, -16103442.89m),
            (4, 10552409503.83m, 17236000.33m, 5419780m, -21288908.57m),
            (5, 8519310329.28m, 30774154.23m, 0m, -6421274.68m)
        };

        var consolidados = valores.Select(v =>
        {
            var ase = Insumos.Ase(v.Item1);
            return new ConsolidadoAse
            {
                Ase = ase,
                TotOpt = v.Item2 + 100m, // TotOpt HU-02 ≠ visible; no debe afectar gates
                R2TotalOportuno = v.Item3,
                Extemp = v.Item2 + 100m, // F25-equiv no es el visible; usamos el F25 fuente (≈Extemp HU-02 distinto del visible)
                ReversionR4 = v.Item5,
                AjustesSfT = 0m
            };
        }).ToList();

        var resultado = new ResultadoRemuneracion
        {
            Periodo = Insumos.Periodo(),
            Consolidados = consolidados,
            Exitoso = true
        };

        var leafs = valores.Select(v =>
        {
            var ase = Insumos.Ase(v.Item1);
            var r1 = new WorkbookLeafInputsR1
            {
                F25 = v.Item2 + 100m, // primera fila Mes/Total col F = Extemp HU-02 (no el visible)
                F41 = -100m,
                L25 = 0m,
                CeldasPorAse = new Dictionary<string, decimal> { ["F25"] = v.Item2 + 100m },
                TotalOportunoEsperadoPorAse = v.Item2,
                ExtemporaneoEsperadoPorAse = v.Item4
            };
            var r2 = new WorkbookLeafInputsR2
            {
                E15 = v.Item3,
                E26 = 0m,
                K15 = 0m,
                CeldasPorAse = new Dictionary<string, decimal> { ["E15"] = v.Item3 }
            };
            var r4 = new WorkbookLeafInputsR4
            {
                D9 = v.Item5,
                P9 = 0m,
                CeldasPorAse = new Dictionary<string, decimal> { ["D9"] = v.Item5 }
            };
            return new WorkbookLeafInputs { Ase = ase, Periodo = resultado.Periodo, R1 = r1, R2 = r2, R4 = r4 };
        }).ToList();

        return (resultado, leafs);
    }
}