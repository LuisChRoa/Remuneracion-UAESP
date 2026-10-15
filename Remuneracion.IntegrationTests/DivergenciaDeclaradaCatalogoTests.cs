using Remuneracion.Core.Models;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 42 (T1, R-CAT-1/R-CAT-2, D-C): catálogo de divergencias declaradas como DATOS citados
/// (patrón <see cref="SinonimosEmpresaR1"/>) + matcher puro. Sin los 6 campos no hay entrada;
/// añadir una divergencia futura = añadir una FILA citada (sin código nuevo).
/// </summary>
public sealed class DivergenciaDeclaradaCatalogoTests
{
    [Fact]
    public void T1_EntradaOccidenteAse5_TieneLos6CamposCitados()
    {
        var entrada = Assert.Single(DivergenciasDeclaradasCatalogo.Entradas);

        Assert.Equal("OCCIDENTE-ASE5", entrada.Id);
        Assert.True(entrada.Completa, "La entrada OCCIDENTE ASE5 debe traer sus 6 campos citados.");
        Assert.Equal(new[] { "O25", "D34", "F34" }, entrada.Celdas);
        Assert.Equal(417519.62m, entrada.Delta);
        Assert.Contains("UAESP", entrada.Dueno, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(entrada.OrigenMonto));
        Assert.False(string.IsNullOrWhiteSpace(entrada.CitaEvidencia));
        Assert.Equal("2026-10-14", entrada.Fecha);
        Assert.False(string.IsNullOrWhiteSpace(entrada.CondicionRetiro));
    }

    [Theory]
    [InlineData("O25")]
    [InlineData("D34")]
    [InlineData("F34")]
    public void T1_Matcher_ReconoceLasTresCeldasConElDeltaCitado(string celda)
    {
        // Δ observado en la corrida viva Q2 = |72441184838,33 − 72440767318,71| = 417.519,62.
        var entrada = DivergenciasDeclaradasCatalogo.Buscar(celda, 72441184838.33m, 72440767318.71m);

        Assert.NotNull(entrada);
        Assert.Equal("OCCIDENTE-ASE5", entrada!.Id);
    }

    [Fact]
    public void T1_Matcher_DeltaDistinto_NoAplica()
    {
        // Mismo celda pero Δ = 1 (defecto interno, no la divergencia citada) → NO declarable.
        Assert.Null(DivergenciasDeclaradasCatalogo.Buscar("O25", 100m, 99m));
    }

    [Fact]
    public void T1_Matcher_CeldaFueraDeLaEntrada_NoAplica()
    {
        Assert.Null(DivergenciasDeclaradasCatalogo.Buscar("C15", 72441184838.33m, 72440767318.71m));
    }

    [Fact]
    public void T1_Matcher_EntradaIncompleta_NoAplica()
    {
        var incompleta = new DivergenciasDeclaradasCatalogo.Entrada(
            Id: "SIN-CITA",
            Celdas: ["O25"],
            Delta: 417519.62m,
            OrigenMonto: "origen",
            Dueno: "UAESP",
            CitaEvidencia: string.Empty, // sin cita no hay fila (R-CAT-2)
            Fecha: string.Empty,
            CondicionRetiro: string.Empty);

        Assert.False(incompleta.Completa);
        Assert.Null(DivergenciasDeclaradasCatalogo.Buscar([incompleta], "O25", 72441184838.33m, 72440767318.71m));
    }

    [Fact]
    public void T1_CatalogoAbierto_AgregarFilaFutura_NoRequiereCodigo()
    {
        var futura = new DivergenciasDeclaradasCatalogo.Entrada(
            Id: "FUTURA",
            Celdas: ["Z99"],
            Delta: 1234.5m,
            OrigenMonto: "origen citado",
            Dueno: "tercero externo",
            CitaEvidencia: "evidencia citada",
            Fecha: "2026-01-01",
            CondicionRetiro: "se retira cuando el dueño resuelva");

        Assert.True(futura.Completa);
        Assert.NotNull(DivergenciasDeclaradasCatalogo.Buscar([futura], "Z99", 2000m, 765.5m));
        // El matcher sobre el catálogo vigente NO conoce la fila futura (catálogo abierto, sin ramas).
        Assert.Null(DivergenciasDeclaradasCatalogo.Buscar("Z99", 2000m, 765.5m));
    }
}
