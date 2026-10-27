using System.IO;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.WinForms;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 43 (O-1): la composición ÚNICA del host (<see cref="ComposicionServicios"/>, compilada por
/// archivo vinculado desde el proyecto WinForms) incluye SIEMPRE el oráculo de validaciones. La vía
/// del constructor por defecto de <c>Form1</c> y la de <c>Program.cs</c> son la MISMA unidad, así
/// que ninguna puede cerrar en verde sin validar. Se prueba en ambos sentidos:
///   (i)  la vía por defecto NO cierra en verde: la corrida Q2 sin reconocimiento BLOQUEA con la
///        clasificación (el oráculo corrió; sin él el run habría completado sin excepción);
///   (ii) la vía de producción sigue funcionando: la misma composición cierra con reconocimiento.
/// </summary>
public sealed class ComposicionServiciosTests
{
    [Fact]
    public void O1_ComposicionDelHost_NoCierraEnVerdeSinValidar()
    {
        var dir = DirectorioTemp("o1-sin-validar");
        var salida = Path.Combine(dir, Insumos.PeriodoQ2().NombreArchivo);
        try
        {
            // Plan 46 (H2): la divergencia real quedó resuelta; la composición ÚNICA del host se
            // ejercita con el FIXTURE SINTÉTICO (SINTETICA-C15, Δ = 1) inyectado por la costura.
            var (_, periodo, _) = ComposicionServicios.Crear(
                FixtureDivergenciaSintetica.CrearValidador(),
                FixtureDivergenciaSintetica.CrearLector());

            var ex = Assert.Throws<CalculoInvalidoException>(() => periodo.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
                RutaPlantilla = Insumos.PlantillaHistoricaV2(),
                RutaSalida = salida
            }));

            // La clasificación viaja ⇒ la validación (W-1) CORRIÓ. Sin oráculo, el run completaría
            // sin excepción: la vía no puede cerrar en verde sin validar.
            Assert.NotNull(ex.Clasificacion);
            Assert.True(ex.Clasificacion!.SoloDeclaradas);
            Assert.Equal([FixtureDivergenciaSintetica.EntradaId], ex.Clasificacion.EntradasDeclaradas);
        }
        finally
        {
            Limpiar(dir);
        }
    }

    [Fact]
    public void O1_ComposicionDelHost_LaViaDeProduccionCierraConReconocimiento()
    {
        var dir = DirectorioTemp("o1-produccion");
        var salida = Path.Combine(dir, Insumos.PeriodoQ2().NombreArchivo);
        try
        {
            var (_, periodo, _) = ComposicionServicios.Crear(
                FixtureDivergenciaSintetica.CrearValidador(),
                FixtureDivergenciaSintetica.CrearLector());

            var resultado = periodo.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
                RutaPlantilla = Insumos.PlantillaHistoricaV2(),
                RutaSalida = salida,
                DivergenciasDeclaradasAceptadas = [FixtureDivergenciaSintetica.EntradaId]
            });

            Assert.True(File.Exists(salida), "La vía de producción debe producir el archivo definitivo.");
            Assert.False(string.IsNullOrWhiteSpace(resultado.ReporteDivergenciasDeclaradas));
            Assert.True(File.Exists(resultado.ReporteDivergenciasDeclaradas));
        }
        finally
        {
            Limpiar(dir);
        }
    }

    private static string DirectorioTemp(string etiqueta)
    {
        var dir = Path.Combine(Path.GetTempPath(), $"{etiqueta}-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void Limpiar(string dir)
    {
        try
        {
            Directory.Delete(dir, recursive: true);
        }
        catch
        {
            // best-effort
        }
    }
}
