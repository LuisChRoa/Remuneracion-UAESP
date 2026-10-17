using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 42 (T1, R-CLS-1/R-CLS-2/R-CLS-3, D-B/D-F): la clasificación Declarada/NoDeclarada vive en
/// el validador de Core (una sola decisión). Lo declarado cita la entrada del catálogo; lo NO
/// declarado conserva el criterio estricto E1-1..E1-5 (mismo nombre accionable). La declarada es la
/// ÚLTIMA puerta: una no declarada coexistente bloquea aunque la declarada exista.
///
/// Plan 46 (H2, decisión del Ingeniero 2026-10-17): la entrada REAL <c>OCCIDENTE-ASE5</c> quedó
/// RETIRADA (la divergencia se resolvió). Estos tests de clasificación usan el FIXTURE SINTÉTICO
/// (<see cref="FixtureDivergenciaSintetica"/>, entrada vigente SINTETICA-C15 con Δ = 1) para seguir
/// cubriendo la clasificación declarada sin depender de la divergencia retirada.
/// </summary>
public sealed class ClasificacionDivergenciasDeclaradasTests
{
    private const decimal Izq = 72497949968m;
    private const decimal Der = 72497949967m; // Δ = 1 (fixture sintético SINTETICA-C15)

    [Fact]
    public void T1_SinteticaC15_SeClasificaDeclaradaConReferencia()
    {
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var snapshots = SnapshotsConSinteticaDeclarada();

        var clasificacion = FixtureDivergenciaSintetica.CrearValidador()
            .ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

        Assert.Equal(5, clasificacion.Declaradas.Count); // 5 ASE × C15
        Assert.Empty(clasificacion.NoDeclaradas);
        Assert.Empty(clasificacion.ErroresNoDeclarados);
        Assert.True(clasificacion.SoloDeclaradas);
        Assert.False(clasificacion.HayNoDeclaradas);
        Assert.Equal([FixtureDivergenciaSintetica.EntradaId], clasificacion.EntradasDeclaradas);

        var veredicto = clasificacion.Declaradas.First(v => v.Celda == FixtureDivergenciaSintetica.Celda);
        Assert.True(veredicto.Declarada);
        Assert.Equal(FixtureDivergenciaSintetica.EntradaId, veredicto.EntradaId);
        Assert.Equal(Izq, veredicto.MontoObservado);
        Assert.Equal(Der, veredicto.MontoEsperado);
        Assert.Equal(FixtureDivergenciaSintetica.Delta, veredicto.Delta);
    }

    [Fact]
    public void T1_DivergenciaInyectadaFueraDeCatalogo_FallaDuroNoDeclarada()
    {
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var snapshots = SnapshotsConSinteticaDeclarada();

        // Divergencia NO declarada inyectada: D25 FALSE con Δ fuera del catálogo (el sintético solo
        // declara C15 con Δ = 1).
        foreach (var snapshot in snapshots)
        {
            var evidencia = new Dictionary<string, SubBloqueValidacionTotalEvidencia>(snapshot.SubBloquesValidacionTotalEvidencia, StringComparer.OrdinalIgnoreCase)
            {
                ["D25"] = new SubBloqueValidacionTotalEvidencia
                {
                    Celda = "D25",
                    TextoFormula = FormasValidacionTotal.TextoEsperado("D25")!,
                    Computable = true,
                    OperandoIzquierdo = 1000m,
                    OperandoDerecho = 1m
                }
            };
            var booleanos = new Dictionary<string, bool>(snapshot.SubBloquesValidacionTotal, StringComparer.OrdinalIgnoreCase) { ["D25"] = false };
            snapshot.SubBloquesValidacionTotalEvidencia = evidencia;
            snapshot.SubBloquesValidacionTotal = booleanos;
        }

        var clasificacion = FixtureDivergenciaSintetica.CrearValidador()
            .ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

        Assert.True(clasificacion.HayNoDeclaradas); // D-F: la no declarada manda
        Assert.False(clasificacion.SoloDeclaradas);
        Assert.Contains(clasificacion.NoDeclaradas, v => v.Celda == "D25" && !v.Declarada);
        Assert.Contains(clasificacion.ErroresNoDeclarados, e => e.Contains("VALIDACION_TOTAL · D25", StringComparison.Ordinal));

        // Las declaradas siguen reconocidas (pero no alcanzan a cerrar por la no declarada).
        Assert.Equal(5, clasificacion.Declaradas.Count);
    }

    [Fact]
    public void T1_CadenaNoResoluble_NuncaSeDeclara()
    {
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var snapshots = SnapshotsConSinteticaDeclarada();

        // C15 no evaluable desde lo escrito (E1-1): aunque su Δ "coincida", NO se declara.
        foreach (var snapshot in snapshots)
        {
            snapshot.SubBloquesValidacionTotalEvidencia[FixtureDivergenciaSintetica.Celda].Computable = false;
        }

        var clasificacion = FixtureDivergenciaSintetica.CrearValidador()
            .ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

        Assert.Contains(clasificacion.NoDeclaradas, v => v.Celda == FixtureDivergenciaSintetica.Celda);
        Assert.True(clasificacion.HayNoDeclaradas);
    }

    [Fact]
    public void T1_SinEvidencia_NoSeDeclara()
    {
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var snapshots = ValidacionesCruzadasTests.CrearSnapshotsValidos();
        foreach (var snapshot in snapshots)
        {
            // Snapshot sin evidencia: el booleano FALSE no es clasificable → no declarada.
            snapshot.SubBloquesValidacionTotal = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase) { ["C15"] = false };
        }

        var clasificacion = FixtureDivergenciaSintetica.CrearValidador()
            .ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

        Assert.True(clasificacion.HayNoDeclaradas);
        Assert.All(clasificacion.NoDeclaradas, v => Assert.Null(v.EntradaId));
    }

    [Fact]
    public void T1_SinDivergencias_ClasificacionVacia()
    {
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var snapshots = ValidacionesCruzadasTests.CrearSnapshotsValidos();

        var clasificacion = FixtureDivergenciaSintetica.CrearValidador()
            .ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

        Assert.True(clasificacion.SinDivergencias);
        Assert.False(clasificacion.SoloDeclaradas);
        Assert.False(clasificacion.HayNoDeclaradas);
        Assert.Empty(clasificacion.Veredictos);
    }

    private static List<ValidacionCruzadaSnapshot> SnapshotsConSinteticaDeclarada()
    {
        var snapshots = ValidacionesCruzadasTests.CrearSnapshotsValidos();
        foreach (var snapshot in snapshots)
        {
            var evidencia = new Dictionary<string, SubBloqueValidacionTotalEvidencia>(StringComparer.OrdinalIgnoreCase);
            var booleanos = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            var celda = FixtureDivergenciaSintetica.Celda;
            evidencia[celda] = new SubBloqueValidacionTotalEvidencia
            {
                Celda = celda,
                TextoFormula = FormasValidacionTotal.TextoEsperado(celda)!,
                Computable = true,
                OperandoIzquierdo = Izq,
                OperandoDerecho = Der
            };
            booleanos[celda] = false;

            snapshot.SubBloquesValidacionTotalEvidencia = evidencia;
            snapshot.SubBloquesValidacionTotal = booleanos;
        }

        return snapshots;
    }
}
