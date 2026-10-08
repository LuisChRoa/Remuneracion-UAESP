using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 29 (T2, Unidad R — SOLO LECTURA): el lector de detalle R2/R4 extrae la matriz por
/// componente de las fuentes REALES (julio-Q2 <c>Docs/Prueba Julio-2/Insumos</c> y agosto-Q2
/// <c>Docs/Prueba2/Insumos</c>), resolviendo las filas por LABEL (A–D) y las columnas por
/// ENCABEZADO de componente, tolerante a la deriva de la malla entre períodos (veredicto T0b).
///
/// Cubre: S3 (detalle R2 ASE1-julio: Vlr Servicio/Total = 104754634.94), S4 (detalle R4 ASE1-julio:
/// Vlr Servicio = -15794348.21 y OCCIDENTE Vlr Servicio = -762167.27), la invariante de
/// <c>Especiales</c> (ausente en julio → 0 explícito; presente en agosto con su valor), la
/// cobertura de los 5 ASE en ambos períodos Q2, y que la lectura NO toca los agregados existentes
/// (<c>LeerR2</c>/<c>LeerR4</c> y las fórmulas esperadas del leaf). CERO escritura.
/// </summary>
public sealed class DetalleR2R4Tests
{
    private const decimal Tol = 0.5m;

    private static ExcelDataReaderWorkbookLeafInputReader Reader() => new();

    // ── S3: detalle R2 julio ASE1 ────────────────────────────────────────────────────────────

    [Fact]
    public void LeerDetalleR2_JulioAse1_VlrServicioTotalEsFuente()
    {
        var detalle = Reader().LeerDetalleR2(Insumos.Ase(1), Insumos.R2JulioQ2(1));

        Assert.Equal(1, detalle.Ase.Id);
        Assert.False(detalle.TieneColumnaEspeciales); // julio: la fuente no trae 'Especiales'
        Assert.NotEmpty(detalle.Componentes);

        var fila = detalle.Filas.First(f => f.Firma == "|||Vlr Servicio");
        Assert.InRange(fila.Valor("Total")!.Value - 104754634.94m, -Tol, Tol);
        // Componente TDF (fuente F5 = 6975298.69): resuelto por ENCABEZADO, no por índice fijo.
        Assert.InRange(fila.Valor("Componente TDF")!.Value - 6975298.69m, -Tol, Tol);
        // K-destino (Especiales) = 0 explícito (invariante: columna ausente en julio/ASE4).
        Assert.Equal(0m, fila.Valor("Especiales"));
    }

    // ── S4: detalle R4 julio ASE1 ────────────────────────────────────────────────────────────

    [Fact]
    public void LeerDetalleR4_JulioAse1_VlrServicioTotalesSonFuente()
    {
        var detalle = Reader().LeerDetalleR4(Insumos.Ase(1), Insumos.R4JulioQ2(1));

        Assert.Equal(1, detalle.Ase.Id);
        Assert.NotEmpty(detalle.Componentes);

        var filasVlrServicio = detalle.Filas.Where(f => f.C == "Vlr Servicio").ToList();
        Assert.True(filasVlrServicio.Count >= 2, "ASE1 debe traer Vlr Servicio del bloque principal y de OCCIDENTE.");

        // D3-destino (fuente r5) y D9-destino (fuente r11 OCCIDENTE).
        Assert.InRange(filasVlrServicio[0].Valor("Total")!.Value - (-15794348.21m), -Tol, Tol);
        Assert.InRange(filasVlrServicio[1].Valor("Total")!.Value - (-762167.27m), -Tol, Tol);

        // Cierre del bloque: fila A='Total'.
        Assert.Contains(detalle.Filas, f => f.A == "Total");
    }

    // ── Cobertura: 5 ASE × {julio-Q2, agosto-Q2} ─────────────────────────────────────────────

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void LeerDetalleR2_TodosLosAseQ2_TienenVlrServicioYComponentes(int aseId)
    {
        var reader = Reader();
        foreach (var ruta in new[] { Insumos.R2JulioQ2(aseId), Insumos.R2AgostoQ2(aseId) })
        {
            var detalle = reader.LeerDetalleR2(Insumos.Ase(aseId), ruta);
            Assert.Equal(aseId, detalle.Ase.Id);
            Assert.NotEmpty(detalle.Componentes);
            Assert.NotEmpty(detalle.Filas);
            Assert.Contains(detalle.Filas, f => f.D == "Vlr Servicio");
            Assert.Contains(detalle.Filas, f => f.A == "Total");
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void LeerDetalleR4_TodosLosAseQ2_TienenVlrServicioYCierre(int aseId)
    {
        var reader = Reader();
        foreach (var ruta in new[] { Insumos.R4JulioQ2(aseId), Insumos.R4AgostoQ2(aseId) })
        {
            var detalle = reader.LeerDetalleR4(Insumos.Ase(aseId), ruta);
            Assert.Equal(aseId, detalle.Ase.Id);
            Assert.NotEmpty(detalle.Componentes);
            Assert.NotEmpty(detalle.Filas);
            Assert.Contains(detalle.Filas, f => f.C == "Vlr Servicio");
            Assert.Contains(detalle.Filas, f => f.A == "Total");
        }
    }

    // ── Invariante Especiales ────────────────────────────────────────────────────────────────

    [Fact]
    public void LeerDetalleR2_AgostoAse1_TieneEspecialesConValorDeFuente()
    {
        var detalle = Reader().LeerDetalleR2(Insumos.Ase(1), Insumos.R2AgostoQ2(1));

        Assert.True(detalle.TieneColumnaEspeciales); // agosto: la fuente trae 'Especiales' en K

        var fila = detalle.Filas.First(f => f.Firma == "|||Vlr Servicio");
        Assert.InRange(fila.Valor("Especiales")!.Value - 144941.68m, -Tol, Tol);
    }

    // ── Conexión con el flujo (leaf) ─────────────────────────────────────────────────────────

    [Fact]
    public void LeerLeafInputs_Q2_PueblaDetalleYNoAlteraAgregados()
    {
        var leaf = Reader().LeerLeafInputs(
            Insumos.Ase(1), new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 },
            Insumos.R1JulioQ2(1), Insumos.R2JulioQ2(1), Insumos.R4JulioQ2(1));

        Assert.NotNull(leaf.DetalleR2);
        Assert.NotNull(leaf.DetalleR4);

        // Agregados existentes intactos (fórmulas visibles del workbook).
        Assert.Equal(leaf.R2.E15 + leaf.R2.E26 - leaf.R2.K15, leaf.R2.TotalOportunoEsperado);
        Assert.Equal(leaf.R4.D9 - leaf.R4.P9, leaf.R4.TotalReversionEsperada);
    }

    [Fact]
    public void LeerLeafInputs_Q1_NoPueblaDetalle()
    {
        var leaf = Reader().LeerLeafInputs(
            Insumos.Ase(1), Insumos.Periodo(), Insumos.R1(1), Insumos.R2(1), Insumos.R4(1));

        Assert.Null(leaf.DetalleR2);
        Assert.Null(leaf.DetalleR4);
    }
}
