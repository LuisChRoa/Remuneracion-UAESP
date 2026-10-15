using System.Text;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 42 (T2, R-REP-1, D-E): reporte durable de divergencias declaradas, legible por el
/// administrativo (lenguaje de negocio, sin jerga de código) con el contenido mandatorio: celda,
/// observado vs esperado, Δ, dueño, cita de evidencia, fecha y qué se reconoció.
/// </summary>
public sealed class ReporteDivergenciasDeclaradasTests
{
    private const decimal Izq = 72441184838.33m;
    private const decimal Der = 72440767318.71m;

    [Fact]
    public void T2_Reporte_ContieneElContenidoMandatorio()
    {
        var texto = ReporteDivergenciasDeclaradasWriter.ConstruirTexto(
            @"C:\salidas\Remuneración 202607-2 Total.xlsx",
            Periodo.Parse("2026072"),
            Guid.Parse("11111111-2222-3333-4444-555555555555"),
            ClasificacionDeclarada());

        Assert.Contains("REPORTE DE DIVERGENCIAS DECLARADAS", texto, StringComparison.Ordinal);
        Assert.Contains("2026072", texto, StringComparison.Ordinal);
        Assert.Contains("segunda quincena", texto, StringComparison.Ordinal);

        // Celda + observado vs esperado + Δ (formato de negocio con punto de miles y coma decimal).
        Assert.Contains("celda O25", texto, StringComparison.Ordinal);
        Assert.Contains("celda D34", texto, StringComparison.Ordinal);
        Assert.Contains("celda F34", texto, StringComparison.Ordinal);
        Assert.Contains("72.441.184.838,33", texto, StringComparison.Ordinal);
        Assert.Contains("72.440.767.318,71", texto, StringComparison.Ordinal);
        Assert.Contains("417.519,62", texto, StringComparison.Ordinal);

        // Dueño + evidencia + fecha + condición de retiro.
        Assert.Contains("UAESP", texto, StringComparison.Ordinal);
        Assert.Contains("T0-Plan41-MiniT0-OccidenteAse5", texto, StringComparison.Ordinal);
        Assert.Contains("2026-10-14", texto, StringComparison.Ordinal);
        Assert.Contains("Se retira cuando", texto, StringComparison.Ordinal);

        // Qué se reconoció (aceptación por corrida, no persistente).
        Assert.Contains("Aceptación", texto, StringComparison.Ordinal);
        Assert.Contains("ESTA corrida", texto, StringComparison.Ordinal);
    }

    [Fact]
    public void T2_Reporte_SinJergaDeCodigo()
    {
        var texto = ReporteDivergenciasDeclaradasWriter.ConstruirTexto(
            @"C:\salidas\Remuneración 202607-2 Total.xlsx",
            Periodo.Parse("2026072"),
            Guid.NewGuid(),
            ClasificacionDeclarada());

        // Doctrina Plan 27 D-E: prohibido el lenguaje de código en el texto visible.
        Assert.DoesNotContain("ERR-", texto, StringComparison.Ordinal);
        Assert.DoesNotContain("<f>", texto, StringComparison.Ordinal);
        Assert.DoesNotContain("</f>", texto, StringComparison.Ordinal);
        Assert.DoesNotContain("SeGatea", texto, StringComparison.Ordinal);
        Assert.DoesNotContain("si=", texto, StringComparison.Ordinal);
        Assert.DoesNotContain("CodigoError", texto, StringComparison.Ordinal);
        Assert.DoesNotContain("Excepcion", texto, StringComparison.Ordinal);
    }

    [Fact]
    public void T2_Emitir_EscribeReporteJuntoALaSalida()
    {
        var dir = Path.Combine(Path.GetTempPath(), "p42-reporte-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var salida = Path.Combine(dir, Periodo.Parse("2026072").NombreArchivo);
            File.WriteAllText(salida, "contenido simulado");

            var ruta = new ReporteDivergenciasDeclaradasWriter().Emitir(salida, Periodo.Parse("2026072"), Guid.NewGuid(), ClasificacionDeclarada());

            Assert.True(File.Exists(ruta));
            Assert.Equal(dir, Path.GetDirectoryName(ruta));
            Assert.Contains("Divergencias declaradas", Path.GetFileName(ruta), StringComparison.Ordinal);
            Assert.Contains("REPORTE DE DIVERGENCIAS DECLARADAS", File.ReadAllText(ruta), StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    private static ClasificacionDivergenciasDeclaradas ClasificacionDeclarada()
    {
        var veredictos = new List<VeredictoDivergenciaValidacion>();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            foreach (var celda in new[] { "O25", "D34", "F34" })
            {
                veredictos.Add(new VeredictoDivergenciaValidacion
                {
                    AseId = aseId,
                    Celda = celda,
                    Declarada = true,
                    EntradaId = "OCCIDENTE-ASE5",
                    MontoObservado = Izq,
                    MontoEsperado = Der
                });
            }
        }

        return new ClasificacionDivergenciasDeclaradas { Veredictos = veredictos };
    }
}
