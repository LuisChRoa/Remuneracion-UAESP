using System.IO;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.WinForms;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 43 (O-4): una corrida que NO cierra bien NUNCA produce el nombre definitivo. El writer deja
/// el archivo con la marca determinista "sin validar"; solo una corrida cerrada bien lo promueve al
/// nombre limpio. Se prueba en ambos sentidos y, además, que un fallo ANTES de escribir no inventa
/// archivo alguno. Usa la composición ÚNICA del host (oráculo + reporte).
/// </summary>
public sealed class SalidaSinValidarO4Tests
{
    [Fact]
    public void O4_CorridaQueNoCierraBien_NoProduceNombreDefinitivo()
    {
        var dir = DirectorioTemp("o4-falla");
        var salida = Path.Combine(dir, Insumos.PeriodoQ2().NombreArchivo);
        try
        {
            // Plan 46 (H2): la divergencia real quedó resuelta; el escenario "no cierra bien" se
            // ejercita con el FIXTURE SINTÉTICO (SINTETICA-C15, Δ = 1) vía la composición del host.
            var (_, periodo, _) = ComposicionServicios.Crear(
                FixtureDivergenciaSintetica.CrearValidador(),
                FixtureDivergenciaSintetica.CrearLector());

            // Q2 sin reconocimiento: la divergencia declarada BLOQUEA después de escribir el workbook.
            Assert.Throws<CalculoInvalidoException>(() => periodo.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
                RutaPlantilla = Insumos.PlantillaHistoricaV2(),
                RutaSalida = salida
            }));

            // El nombre DEFINITIVO NO existe; el archivo queda con la marca visible y determinista.
            var marcada = NombresSalida.RutaSinValidar(salida);
            Assert.False(File.Exists(salida), "Una corrida no cerrada no debe producir el nombre definitivo.");
            Assert.True(File.Exists(marcada), "El archivo de la corrida no cerrada debe quedar marcado como 'sin validar'.");

            // Determinismo: misma entrada → misma ruta marcada (sufijo antes de la extensión).
            Assert.Equal(
                Path.Combine(dir, "Remuneración 202607-2 Total (sin validar).xlsx"),
                marcada,
                ignoreCase: false);
            Assert.Equal(marcada, NombresSalida.RutaSinValidar(salida));
        }
        finally
        {
            Limpiar(dir);
        }
    }

    [Fact]
    public void O4_CorridaCerradaBien_ProduceNombreLimpio()
    {
        var dir = DirectorioTemp("o4-cierra");
        var salida = Path.Combine(dir, Insumos.PeriodoQ2().NombreArchivo);
        try
        {
            var (_, periodo, _) = ComposicionServicios.Crear();

            periodo.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
                RutaPlantilla = Insumos.PlantillaHistoricaV2(),
                RutaSalida = salida,
                DivergenciasDeclaradasAceptadas = ["OCCIDENTE-ASE5"]
            });

            Assert.True(File.Exists(salida), "Una corrida cerrada bien debe producir el nombre limpio.");
            Assert.False(
                File.Exists(NombresSalida.RutaSinValidar(salida)),
                "Tras cerrar bien, no debe quedar el archivo marcado como 'sin validar'.");
        }
        finally
        {
            Limpiar(dir);
        }
    }

    [Fact]
    public void O4_FalloAntesDeEscribir_NoInventaArchivo()
    {
        var dir = DirectorioTemp("o4-pre-escritura");
        var salida = Path.Combine(dir, Insumos.Periodo().NombreArchivo);
        try
        {
            var (_, periodo, _) = ComposicionServicios.Crear();

            // Plan 56 (H2 — desviación documentada): el escenario "fallo ANTES de escribir" se ejerce
            // con una PLANTILLA AUSENTE (ERR-PLANTILLA al abrir el workbook de salida, antes de tocar
            // disco), no con el defecto ERR-SIN-SLOT de Q1 que este plan corrige (§2-D2/D3). La
            // garantía probada es la misma: un fallo previo a la escritura no deja ni el nombre
            // definitivo ni el marcado "sin validar".
            var plantillaAusente = Path.Combine(dir, "no-existe-plantilla.xlsx");
            var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() => periodo.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.Periodo(),
                CarpetaPeriodo = Insumos.CarpetaPeriodo,
                RutaPlantilla = plantillaAusente,
                RutaSalida = salida
            }));

            // El código exacto depende del primer punto que abre la plantilla (preflight del lector
            // leaf vs. escritor); ambos son fallos ESTRUCTURALES previos a la escritura.
            Assert.Contains(ex.Codigo, new[] { CodigoError.Plantilla, CodigoError.FuenteNoEncontrada });
            Assert.False(File.Exists(salida), "Un fallo antes de escribir no debe dejar el nombre definitivo.");
            Assert.False(
                File.Exists(NombresSalida.RutaSinValidar(salida)),
                "Un fallo antes de escribir no debe inventar un archivo 'sin validar'.");
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
