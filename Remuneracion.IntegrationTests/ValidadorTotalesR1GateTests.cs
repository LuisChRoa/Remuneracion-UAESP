using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 31 (T1, pieza b+c — R-G-1 / S1): gate workbook-vs-dominio de los totales visibles R1.
/// T2 (recomposición por firma en el pase final del mutador) cierra el rojo intencional de T1.
///
/// Corre el flujo 5-ASE end-to-end contra la plantilla en ceros del período (julio:
/// <c>Docs/Plantilla_Remuneracion.xlsx</c>; agosto: base <c>Plantilla_Remuneracion_2026082.xlsx</c>,
/// Plan 30/T3) hacia una ruta TEMPORAL con insumos REALES, y luego valida con
/// <see cref="ValidadorTotalesR1Workbook"/> contra los leafs (dominio por firma):
///   - Julio-2026072 (geometría = plantilla): el gate CIERRA (identidad, D-E).
///   - Agosto-2026082 (la fuente recorta en la cabeza): tras T2 el gate CIERRA: los visibles
///     recompuestos se evalúan contra el dominio por firma ±0.5.
///
/// Sin Excel/COM; el gate usa ZIP+XML BCL. Tolerancia ±0.5.
/// </summary>
public sealed class ValidadorTotalesR1GateTests
{
    private const string RutaPlantillaJulio = "Docs";

    // S1 (R-G-1): julio-2026072 -> gate VERDE contra la salida base de julio.
    [Fact]
    public void S1_Julio2026072_GateVerde()
    {
        using var salida = new SalidaTemporal("remuneracion-gate-r1-julio-");
        var resultado = Ejecutar(
            new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 },
            Insumos.CarpetaInsumosJulioQ2,
            Path.Combine(Insumos.Raiz(), RutaPlantillaJulio, "Plantilla_Remuneracion.xlsx"),
            salida.Ruta);

        var gate = new ValidadorTotalesR1Workbook();
        var validacion = gate.Validar(salida.Ruta, resultado.Leafs);

        Assert.True(validacion.EsValido, validacion.Mensaje);
    }

    // S1 (R-G-1): agosto-2026082 -> gate VERDE tras T2. El pase final del mutador recompone por
    // firma las fórmulas de los visibles (TOT_OPT/TDF/EXTEMP) con las filas reales del período, de
    // modo que su evaluación coincide con el dominio por firma ±0.5. (En T1 este test era el rojo
    // intencional que reproducía el anclaje viejo; T2 lo cierra.)
    [Fact]
    public void S1_Agosto2026082_GateVerde()
    {
        using var salida = new SalidaTemporal("remuneracion-gate-r1-agosto-");
        var resultado = Ejecutar(
            new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 },
            Insumos.CarpetaInsumosAgosto,
            Insumos.PlantillaAgosto2026082,
            salida.Ruta);

        var gate = new ValidadorTotalesR1Workbook();
        var validacion = gate.Validar(salida.Ruta, resultado.Leafs);

        Assert.True(validacion.EsValido, validacion.Mensaje);
    }

    private static ResultadoProcesoPeriodo Ejecutar(Periodo periodo, string carpetaInsumos, string rutaPlantilla, string rutaSalida)
    {
        var procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader());

        return procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = periodo,
            CarpetaPeriodo = carpetaInsumos,
            RutaPlantilla = rutaPlantilla,
            RutaSalida = rutaSalida
        });
    }

    private sealed class SalidaTemporal : IDisposable
    {
        private readonly string _dir;

        public SalidaTemporal(string prefijo)
        {
            _dir = Path.Combine(Path.GetTempPath(), prefijo + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
            Ruta = Path.Combine(_dir, "salida.xlsx");
        }

        public string Ruta { get; }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_dir))
                {
                    Directory.Delete(_dir, recursive: true);
                }
            }
            catch
            {
                // best-effort
            }
        }
    }
}
