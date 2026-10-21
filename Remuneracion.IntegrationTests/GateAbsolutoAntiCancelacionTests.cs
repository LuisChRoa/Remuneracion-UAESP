using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 49 (A2, §2.4): test de MUTACIÓN anti-cancelación. Prueba la clase de error que dejó el
/// falso verde de julio-2 (T0 §6.3): una pérdida común a las DOS derivaciones del mismo workbook es
/// invisible al gate 2.7 (<c>D34</c> sigue TRUE porque <c>D33</c> y <c>D109</c> bajan lo mismo).
///
/// El fixture es SINTÉTICO y determinista (sin disco, sin Excel/COM): reutiliza los helpers
/// in-memory de <see cref="ValidacionesCruzadasTests"/> y ejercita la implementación REAL de
/// <see cref="ValidadorBasico"/> (patrón Plan 46 review W-1).
/// </summary>
public sealed class GateAbsolutoAntiCancelacionTests
{
    private const int AseExtemp = 5;

    /// <summary>Pérdida exacta del EXTEMPORÁNEO del ASE5 en agosto-2 (T0 §2.1).</summary>
    private const decimal PerdidaExtempAgosto = 1_249_810m;

    [Fact]
    public void Caso1_CancelacionSimulada_D34Verde_GateAbsolutoFallaConDeltaExacto()
    {
        // Escenario = firma exacta del falso verde:
        //  · D34 = TRUE entregado (el gate viejo pasa; no hay evidencia evaluable en el snapshot).
        //  · CONSOLIDADO!D51 resuelto = 0 (la pérdida entra también en el lado "empresas" y cancela).
        //  · Dominio: ConsolidadoAse.Extemp del ASE5 = 1.249.810 (el dato correcto que SÍ existe).
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var consolidadoAse5 = resultado.Consolidados.Single(c => c.Ase.Id == AseExtemp);
        consolidadoAse5.Extemp = PerdidaExtempAgosto;
        var leafAse5 = leafs.Single(l => l.Ase.Id == AseExtemp);
        leafAse5.R1.F25 = PerdidaExtempAgosto; // gates previos siguen verdes
        leafAse5.R1.ExtemporaneoEsperadoPorAse = PerdidaExtempAgosto; // agregado del dominio por firma

        var snapshots = ValidacionesCruzadasTests.CrearSnapshotsValidos();
        var snapshotAse5 = snapshots.Single(s => s.Ase.Id == AseExtemp);
        snapshotAse5.SubBloquesValidacionTotal = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase)
        {
            ["C15"] = true,
            ["D25"] = true,
            ["O25"] = true,
            ["D34"] = true, // el gate 2.7 NO ve la pérdida (ambos operandos bajan juntos)
            ["F34"] = true
        };
        snapshotAse5.ExtempResuelto = 0m; // D51 resuelto desde lo escrito

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        // Con el código anterior al Plan 49 este assert fallaba (el gate no existía): rojo→verde honesto.
        var error = Assert.Single(errores);
        Assert.Contains("[ERR-VALIDACION]", error, StringComparison.Ordinal);
        Assert.Contains("ASE 5", error, StringComparison.Ordinal);
        Assert.Contains("CONSOLIDADO_TOTAL RECAUDO", error, StringComparison.Ordinal);
        Assert.Contains("D51", error, StringComparison.Ordinal);
        Assert.Contains("1249810", error, StringComparison.Ordinal);
        Assert.Contains("0.5", error, StringComparison.Ordinal);
        Assert.Contains("Plan 49 A1", error, StringComparison.Ordinal);

        // La clasificación hereda el gate SIN cambios: la pérdida no es una divergencia declarada
        // (D-F: la no declarada manda aunque hubiera una declarada reconocida).
        var clasificacion = new ValidadorBasico().ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);
        Assert.True(clasificacion.HayNoDeclaradas);
        Assert.False(clasificacion.SoloDeclaradas);
        Assert.Contains(clasificacion.ErroresNoDeclarados, e => e.Contains("Plan 49 A1", StringComparison.Ordinal));
    }

    [Fact]
    public void Caso2_ResueltoIgualAlDominio_CeroErroresA1()
    {
        // Caso sano (protege contra un gate que falle siempre): resuelto == dominio ±0.5.
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var consolidadoAse5 = resultado.Consolidados.Single(c => c.Ase.Id == AseExtemp);
        consolidadoAse5.Extemp = PerdidaExtempAgosto;
        var leafAse5 = leafs.Single(l => l.Ase.Id == AseExtemp);
        leafAse5.R1.F25 = PerdidaExtempAgosto;
        leafAse5.R1.ExtemporaneoEsperadoPorAse = PerdidaExtempAgosto;

        var snapshots = ValidacionesCruzadasTests.CrearSnapshotsValidos();
        snapshots.Single(s => s.Ase.Id == AseExtemp).ExtempResuelto = PerdidaExtempAgosto;

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        Assert.Empty(errores);
        Assert.DoesNotContain(errores, e => e.Contains("Plan 49 A1", StringComparison.Ordinal));
    }
}
