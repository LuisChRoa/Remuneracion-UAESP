using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-08 (2.2): gate Σ empresas = visible de bloque por ASE y hoja (R1/R2/R4) in-memory.
/// Valores golden Q1 (tabla §2.1 + T0-0.5). Ceros legítimos: EAAB-CL todo 0 en Q1 es válido.
/// El mismatch nombra ASE y empresa.
/// </summary>
public sealed class ConciliacionEmpresaTests
{
    private const decimal Tolerancia = 0.5m;

    [Fact]
    public void Validar_ConciliacionQ1_SumaEmpresasIgualaVisibleDeBloque_CincoAse()
    {
        var (resultado, leafs) = CrearCasoValido();
        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Empty(errores);
    }

    [Fact]
    public void Validar_MismatchOCCIDENTEEnAse3_NombraAseYEmpresa()
    {
        var (resultado, leafs) = CrearCasoValido();

        // Rompe SOLO la conciliación R1 de OCCIDENTE en ASE3 (visible 453277723.14 → 453277723.64).
        var occidenteAse3 = leafs[2].Conciliacion.Single(c => c.Empresa.Id == 4);
        occidenteAse3.VisibleR1 += 1000m;

        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("ASE 3", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(errores, e => e.Contains("OccidenteDirecta", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_EaabCiudadLimpiaTodoCero_EsLegitimoYNoFalla()
    {
        var (resultado, leafs) = CrearCasoValido();

        // EAAB-CL (empresa 5) tiene visibles 0 en Q1 (V8): el gate debe permitir ceros explícitos.
        foreach (var leaf in leafs)
        {
            var eaabCl = leaf.Conciliacion.Single(c => c.Empresa.Id == 5);
            Assert.Equal(0m, eaabCl.VisibleR1);
            Assert.Equal(0m, eaabCl.VisibleR2);
            Assert.Equal(0m, eaabCl.VisibleR4);
        }

        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs));
    }

    [Fact]
    public void Validar_VisibleIndebidoNoCeroEnEmpresaCero_RompeElGate()
    {
        var (resultado, leafs) = CrearCasoValido();

        // ENERBIT (empresa 3) no tiene recaudo en Q1 salvo ASE5. Un visible indebido en ASE1
        // debe romper el gate Σ R1 (añade 1.000.000 sin contrapartida en el bloque).
        var enerbitAse1 = leafs[0].Conciliacion.Single(c => c.Empresa.Id == 3);
        enerbitAse1.VisibleR1 = 1_000_000m;

        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Contains(errores, e => e.Contains("Σ visibles de empresas en R1", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(errores, e => e.Contains("ASE 1", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Validar_ListaConciliacionVacia_ComportamientoHu07Intacto()
    {
        var (resultado, leafs) = CrearCasoValido();
        foreach (var leaf in leafs)
        {
            leaf.Conciliacion = [];
        }

        // Lista vacía = HU-07 puro: no hay gate 2.2 y la validación multi-ASE sigue verde.
        Assert.Empty(new ValidadorBasico().Validar(resultado, leafs));
    }

    [Fact]
    public void Validar_GranTotalSigueSiendoSumaDeTotalAse_NoLoTocaLaConciliacion()
    {
        var (resultado, leafs) = CrearCasoValido();
        var errores = new ValidadorBasico().Validar(resultado, leafs);
        Assert.Empty(errores);
        Assert.Equal(resultado.GranTotal, resultado.Consolidados.Sum(c => c.TotalAse));
    }

    /// <summary>
    /// Construye el caso Q1 con visibles por empresa de la tabla §2.1 (R1) + T0-0.5 (R2/R4).
    /// Los visibles de bloque se derivan de la Σ empresas (gate honesto, nunca agregados HU-02).
    /// </summary>
    internal static (ResultadoRemuneracion Resultado, List<WorkbookLeafInputs> Leafs) CrearCasoValido()
    {
        // Visibles por empresa (R1, R2, R4) — golden Q1.
        var visibles = new (int EmpresaId, decimal R1, decimal R2, decimal R4)[5][];
        visibles[0] = // ASE1
        [
            (1, 0m, 0m, 0m), (2, 16180195504.29m, 51905326.66m, 0m), (3, 0m, 0m, 0m),
            (4, 524136930.28m, 2311059.02m, -12054255.65m), (5, 0m, 0m, 0m)
        ];
        visibles[1] = // ASE2
        [
            (1, 13570060m, 32822845.35m, -19778.72m), (2, 11762458070.29m, 34136584.16m, -9869411m),
            (3, 0m, 0m, 0m), (4, 381752310.90m, 12441371.75m, 0m), (5, 0m, 0m, 0m)
        ];
        visibles[2] = // ASE3
        [
            (1, 0m, 0m, 0m), (2, 9648710791.68m, 30434743.30m, -5571961m),
            (3, 0m, 0m, 0m), (4, 453277723.14m, 677060.45m, -10531481.89m), (5, 0m, 0m, 0m)
        ];
        visibles[3] = // ASE4
        [
            (1, 7264980m, 4804828.26m, -464370m), (2, 10318956340.32m, 10429855.41m, -15968186.61m),
            (3, 0m, 0m, 0m), (4, 226188183.51m, 2001316.66m, -4856351.96m), (5, 0m, 0m, 0m)
        ];
        visibles[4] = // ASE5
        [
            (1, 0m, 0m, 0m), (2, 8229736739.09m, 30118495.22m, -5470988m),
            (3, 126372775.04m, 0m, 0m), (4, 163200815.15m, 655659.01m, -950286.68m), (5, 0m, 0m, 0m)
        ];

        var consolidados = new List<ConsolidadoAse>();
        var leafs = new List<WorkbookLeafInputs>();
        for (var i = 0; i < 5; i++)
        {
            var ase = Insumos.Ase(i + 1);
            var porAse = visibles[i];
            var totOpt = porAse.Sum(v => v.R1);
            var r2 = porAse.Sum(v => v.R2);
            var r4 = porAse.Sum(v => v.R4);

            var r1Model = new WorkbookLeafInputsR1
            {
                F25 = totOpt + 100m, // F25-equiv no es el visible; no afecta el gate 2.2
                F41 = -100m,
                L25 = 0m,
                CeldasPorAse = new Dictionary<string, decimal> { ["F25"] = totOpt + 100m },
                TotalOportunoEsperadoPorAse = totOpt,
                ExtemporaneoEsperadoPorAse = 0m
            };
            var r2Model = new WorkbookLeafInputsR2
            {
                E15 = r2, E26 = 0m, K15 = 0m,
                CeldasPorAse = new Dictionary<string, decimal> { ["E15"] = r2 }
            };
            var r4Model = new WorkbookLeafInputsR4
            {
                D9 = r4, P9 = 0m,
                CeldasPorAse = new Dictionary<string, decimal> { ["D9"] = r4 }
            };

            var conciliacion = porAse.Select(v =>
            {
                var empresa = EmpresaFacturacion.Obtener(v.EmpresaId);
                return new ConciliacionEmpresaInputs
                {
                    Empresa = empresa,
                    Ase = ase,
                    VisibleR1 = v.R1,
                    VisibleR2 = v.R2,
                    VisibleR4 = v.R4
                };
            }).ToList();

            leafs.Add(new WorkbookLeafInputs
            {
                Ase = ase,
                Periodo = Insumos.Periodo(),
                R1 = r1Model,
                R2 = r2Model,
                R4 = r4Model,
                Conciliacion = conciliacion
            });

            consolidados.Add(new ConsolidadoAse
            {
                Ase = ase,
                TotOpt = totOpt + 100m, // TotOpt HU-02 ≠ visible; no debe afectar gates
                R2TotalOportuno = r2,
                Extemp = totOpt + 100m,
                ReversionR4 = r4,
                AjustesSfT = 0m
            });
        }

        return (new ResultadoRemuneracion
        {
            Periodo = Insumos.Periodo(),
            Consolidados = consolidados,
            Exitoso = true
        }, leafs);
    }
}
