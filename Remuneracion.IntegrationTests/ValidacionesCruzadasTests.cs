using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-13 (2.7): gates de dominio de validaciones cruzadas (Plan 13 §5.2) con snapshots
/// in-memory — el validador NUNCA abre .xlsx (D1). Semántica congelada por T0.
/// </summary>
public sealed class ValidacionesCruzadasTests
{
    [Fact]
    public void SnapshotVerde_TodosLosGates_SinErrores()
    {
        var (resultado, leafs) = CrearCasoValido();
        var snapshots = CrearSnapshotsValidos();

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        Assert.Empty(errores);
    }

    [Fact]
    public void O_FueraDeTolerancia_FallaNombrandoAseYEmpresa()
    {
        var (resultado, leafs) = CrearCasoValido();
        var snapshots = CrearSnapshotsValidos();
        // ASE3 · OCCIDENTE: O = 2.0 (fuera de ±0.5).
        snapshots[2].PorEmpresa.Single(e => e.Empresa == "OccidenteDirecta").DiferenciaO = 2.0m;

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        var error = Assert.Single(errores);
        Assert.Contains("ASE 3", error);
        Assert.Contains("OccidenteDirecta", error);
    }

    [Fact]
    public void P_Falso_FallaAunqueO_EsteEnTolerancia()
    {
        var (resultado, leafs) = CrearCasoValido();
        var snapshots = CrearSnapshotsValidos();
        // ASE1 · Enel: O = 0.3 (en tolerancia) pero P = false → falla exacto.
        snapshots[0].PorEmpresa.Single(e => e.Empresa == "Enel").DiferenciaO = 0.3m;
        snapshots[0].PorEmpresa.Single(e => e.Empresa == "Enel").VerificacionP = false;

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        var error = Assert.Single(errores);
        Assert.Contains("ASE 1", error);
        Assert.Contains("Enel", error);
        Assert.Contains("P (INT(O)=0)", error);
    }

    [Fact]
    public void DetValiRetri_D18_FueraDeTolerancia_FallaNombrandoAse3YCelda()
    {
        var (resultado, leafs) = CrearCasoValido();
        var snapshots = CrearSnapshotsValidos();
        // D18 = fila del ASE3 (15+3) con 1.2 → error que nombra ASE3 + D18.
        snapshots[2].DetValiRetri!.DiferenciasAse.Single().Celda = "D18";
        snapshots[2].DetValiRetri!.DiferenciasAse.Single().Valor = 1.2m;

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        var error = Assert.Single(errores);
        Assert.Contains("ASE 3", error);
        Assert.Contains("D18", error);
    }

    [Fact]
    public void DetValiRetri_VerificacionFalsa_FallaNombrandoAseYCelda()
    {
        var (resultado, leafs) = CrearCasoValido();
        var snapshots = CrearSnapshotsValidos();
        snapshots[4].DetValiRetri!.VerificacionesAse.Single().Celda = "D28";
        snapshots[4].DetValiRetri!.VerificacionesAse.Single().Verificacion = false;

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        var error = Assert.Single(errores);
        Assert.Contains("ASE 5", error);
        Assert.Contains("D28", error);
    }

    [Fact]
    public void ValidacionTotal_FueraDeTolerancia_Falla()
    {
        var (resultado, leafs) = CrearCasoValido();
        var snapshots = CrearSnapshotsValidos();
        snapshots[1].ValidacionTotal = 3.0m;

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        var error = Assert.Single(errores);
        Assert.Contains("VALIDACION_TOTAL", error);
    }

    [Fact]
    public void D21_Total_DivergenciaDocumentada_NoFalla()
    {
        // D6: la fila Total D21 diverge en los goldens (Q1 −0.62, Q2 −1.71) y está EXCLUIDA del
        // gate: un valor grande en DiferenciaTotalD21 NO produce error.
        var (resultado, leafs) = CrearCasoValido();
        var snapshots = CrearSnapshotsValidos();
        foreach (var snapshot in snapshots)
        {
            snapshot.DetValiRetri!.DiferenciaTotalD21 = 5.0m;
        }

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        Assert.Empty(errores);
    }

    [Fact]
    public void SinSnapshots_OverloadTresParametros_Hu12Puro_SinErrores()
    {
        // A7: sin snapshot (= reader ausente en regresión) el path HU-12 puro no se toca.
        var (resultado, leafs) = CrearCasoValido();

        var errores = new ValidadorBasico().Validar(resultado, leafs);

        Assert.Empty(errores);
    }

    [Fact]
    public void SnapshotConAseDesconocido_MatcheoEstricto_FallaNombrandoAse()
    {
        var (resultado, leafs) = CrearCasoValido();
        var snapshots = CrearSnapshotsValidos();
        snapshots[0].Ase.Id = 99;

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        Assert.Contains(errores, e => e.Contains("ASE 99", StringComparison.Ordinal));
    }

    // ── Fixtures in-memory (dominio puro; todos los valores 0 → gates HU-07..HU-12 verdes) ─────

    internal static (ResultadoRemuneracion Resultado, List<WorkbookLeafInputs> Leafs) CrearCasoValido()
    {
        var periodo = Insumos.Periodo();
        var resultado = new ResultadoRemuneracion { Periodo = periodo };
        var leafs = new List<WorkbookLeafInputs>(5);

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var ase = Insumos.Ase(aseId);
            resultado.Consolidados.Add(new ConsolidadoAse { Ase = ase });
            leafs.Add(new WorkbookLeafInputs
            {
                Ase = ase,
                Periodo = periodo,
                R1 = new WorkbookLeafInputsR1 { F25 = 0m }
            });
        }

        return (resultado, leafs);
    }

    internal static List<ValidacionCruzadaSnapshot> CrearSnapshotsValidos()
    {
        var snapshots = new List<ValidacionCruzadaSnapshot>(5);
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var porEmpresa = EmpresaFacturacion.Catalogo
                .OrderBy(e => e.Id)
                .Select(e => new ValidacionEmpresaSnapshot
                {
                    Empresa = e.Nombre,
                    AseId = aseId,
                    DiferenciaO = 0m,
                    VerificacionP = true
                })
                .ToList();

            snapshots.Add(new ValidacionCruzadaSnapshot
            {
                Ase = Insumos.Ase(aseId),
                PorEmpresa = porEmpresa,
                DetValiRetri = new DetValiRetriSnapshot
                {
                    Ase = Insumos.Ase(aseId),
                    DiferenciasAse =
                    [
                        new DetValiRetriCeldaValor { Celda = $"D{15 + aseId}", Valor = 0m }
                    ],
                    VerificacionesAse =
                    [
                        new DetValiRetriCeldaVerificacion { Celda = $"D{23 + aseId}", Verificacion = true }
                    ],
                    DiferenciaTotalD21 = 0m,
                    VerificacionTotalD29 = true
                },
                ValidacionTotal = 0m,
                ValidacionTotalOkP = true
            });
        }

        return snapshots;
    }
}
