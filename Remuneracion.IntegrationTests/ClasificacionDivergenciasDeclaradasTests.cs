using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 42 (T1, R-CLS-1/R-CLS-2/R-CLS-3, D-B/D-F): la clasificación Declarada/NoDeclarada vive en
/// el validador de Core (una sola decisión). Lo declarado cita la entrada del catálogo; lo NO
/// declarado conserva el criterio estricto E1-1..E1-5 (mismo nombre accionable). La declarada es la
/// ÚLTIMA puerta: una no declarada coexistente bloquea aunque la declarada exista.
/// </summary>
public sealed class ClasificacionDivergenciasDeclaradasTests
{
    private const decimal Izq = 72441184838.33m;
    private const decimal Der = 72440767318.71m; // Δ = 417.519,62 (OCCIDENTE ASE5)

    [Fact]
    public void T1_OccidenteAse5_SeClasificaDeclaradaConReferencia()
    {
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var snapshots = SnapshotsConOccidenteDeclarado();

        var clasificacion = new ValidadorBasico().ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

        Assert.Equal(15, clasificacion.Declaradas.Count); // 5 ASE × O25/D34/F34
        Assert.Empty(clasificacion.NoDeclaradas);
        Assert.Empty(clasificacion.ErroresNoDeclarados);
        Assert.True(clasificacion.SoloDeclaradas);
        Assert.False(clasificacion.HayNoDeclaradas);
        Assert.Equal(["OCCIDENTE-ASE5"], clasificacion.EntradasDeclaradas);

        var veredicto = clasificacion.Declaradas.First(v => v.Celda == "O25");
        Assert.True(veredicto.Declarada);
        Assert.Equal("OCCIDENTE-ASE5", veredicto.EntradaId);
        Assert.Equal(Izq, veredicto.MontoObservado);
        Assert.Equal(Der, veredicto.MontoEsperado);
        Assert.Equal(417519.62m, veredicto.Delta);
    }

    [Fact]
    public void T1_DivergenciaInyectadaFueraDeCatalogo_FallaDuroNoDeclarada()
    {
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var snapshots = SnapshotsConOccidenteDeclarado();

        // Divergencia NO declarada inyectada: C15 (recaudo) FALSE con Δ fuera del catálogo.
        foreach (var snapshot in snapshots)
        {
            var evidencia = new Dictionary<string, SubBloqueValidacionTotalEvidencia>(snapshot.SubBloquesValidacionTotalEvidencia, StringComparer.OrdinalIgnoreCase)
            {
                ["C15"] = new SubBloqueValidacionTotalEvidencia
                {
                    Celda = "C15",
                    TextoFormula = FormasValidacionTotal.TextoEsperado("C15")!,
                    Computable = true,
                    OperandoIzquierdo = 1000m,
                    OperandoDerecho = 1m
                }
            };
            var booleanos = new Dictionary<string, bool>(snapshot.SubBloquesValidacionTotal, StringComparer.OrdinalIgnoreCase) { ["C15"] = false };
            snapshot.SubBloquesValidacionTotalEvidencia = evidencia;
            snapshot.SubBloquesValidacionTotal = booleanos;
        }

        var clasificacion = new ValidadorBasico().ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

        Assert.True(clasificacion.HayNoDeclaradas); // D-F: la no declarada manda
        Assert.False(clasificacion.SoloDeclaradas);
        Assert.Contains(clasificacion.NoDeclaradas, v => v.Celda == "C15" && !v.Declarada);
        Assert.Contains(clasificacion.ErroresNoDeclarados, e => e.Contains("VALIDACION_TOTAL · C15", StringComparison.Ordinal));

        // Las declaradas siguen reconocidas (pero no alcanzan a cerrar por la no declarada).
        Assert.Equal(15, clasificacion.Declaradas.Count);
    }

    [Fact]
    public void T1_CadenaNoResoluble_NuncaSeDeclara()
    {
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var snapshots = SnapshotsConOccidenteDeclarado();

        // O25 no evaluable desde lo escrito (E1-1): aunque su Δ "coincida", NO se declara.
        foreach (var snapshot in snapshots)
        {
            snapshot.SubBloquesValidacionTotalEvidencia["O25"].Computable = false;
        }

        var clasificacion = new ValidadorBasico().ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

        Assert.Contains(clasificacion.NoDeclaradas, v => v.Celda == "O25");
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
            snapshot.SubBloquesValidacionTotal = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase) { ["O25"] = false };
        }

        var clasificacion = new ValidadorBasico().ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

        Assert.True(clasificacion.HayNoDeclaradas);
        Assert.All(clasificacion.NoDeclaradas, v => Assert.Null(v.EntradaId));
    }

    [Fact]
    public void T1_SinDivergencias_ClasificacionVacia()
    {
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var snapshots = ValidacionesCruzadasTests.CrearSnapshotsValidos();

        var clasificacion = new ValidadorBasico().ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

        Assert.True(clasificacion.SinDivergencias);
        Assert.False(clasificacion.SoloDeclaradas);
        Assert.False(clasificacion.HayNoDeclaradas);
        Assert.Empty(clasificacion.Veredictos);
    }

    private static List<ValidacionCruzadaSnapshot> SnapshotsConOccidenteDeclarado()
    {
        var snapshots = ValidacionesCruzadasTests.CrearSnapshotsValidos();
        foreach (var snapshot in snapshots)
        {
            var evidencia = new Dictionary<string, SubBloqueValidacionTotalEvidencia>(StringComparer.OrdinalIgnoreCase);
            var booleanos = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            foreach (var celda in new[] { "O25", "D34", "F34" })
            {
                evidencia[celda] = new SubBloqueValidacionTotalEvidencia
                {
                    Celda = celda,
                    TextoFormula = FormasValidacionTotal.TextoEsperado(celda)!,
                    Computable = true,
                    OperandoIzquierdo = Izq,
                    OperandoDerecho = Der
                };
                booleanos[celda] = false;
            }

            snapshot.SubBloquesValidacionTotalEvidencia = evidencia;
            snapshot.SubBloquesValidacionTotal = booleanos;
        }

        return snapshots;
    }
}
