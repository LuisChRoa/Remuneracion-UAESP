using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 32 (T1, pieza c — R-G-1 / S1): gate-v2 ESTRUCTURAL del INTERIOR de <c>Reporte Componentes R1</c>.
///
/// Corre el flujo 5-ASE end-to-end contra la plantilla en ceros del período hacia una ruta TEMPORAL con
/// insumos REALES, y valida con
/// <see cref="ValidadorTotalesR1Workbook.ExigirFormaRecompuestaInterior"/> contra el fixture congelado
/// del manual (<see cref="InterioresR1Esperados"/>):
///   - Julio-2026072 (geometría = plantilla): el gate CIERRA (identidad, D-E).
///   - Agosto-2026082 (la fuente recorta en la cabeza): el gate queda en ROJO en T1 (el interior conserva
///     el anclaje julio); T2 lo cierra al recomponer el interior por firma.
///
/// Sin Excel/COM; el gate usa ZIP+XML BCL. Tolerancia ±0.5 (el gate interior es textual, sin valores).
/// </summary>
public sealed class ValidadorInteriorR1GateTests
{
    private const string RutaPlantillaJulio = "Docs";

    // S1 (R-G-1): julio-2026072 -> gate interior VERDE (identidad). Julio no trae clases L, así que el
    // fixture es 100% R-1 y el pase de julio reproduce byte-idénticas las <f>.
    [Fact]
    public void S1_Julio2026072_FormaInteriorVerde()
    {
        using var salida = new SalidaTemporal("remuneracion-gate-interior-julio-");
        Ejecutar(
            new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 },
            Insumos.CarpetaInsumosJulioQ2,
            Path.Combine(Insumos.Raiz(), RutaPlantillaJulio, "Plantilla_Remuneracion.xlsx"),
            salida.Ruta);

        var gate = new ValidadorTotalesR1Workbook();
        gate.ExigirFormaRecompuestaInterior(salida.Ruta, Esperados(InterioresR1Esperados.Julio));
    }

    // S1 (R-G-1): agosto-2026082 -> gate interior EN ROJO hasta T2. Este test aserta el estado OBJETIVO
    // (verde): es el rojo intencional de T1 que reproduce el anclaje julio del interior. T2 lo cierra al
    // recomponer el interior por firma (texto = manual).
    [Fact]
    public void S1_Agosto2026082_FormaInteriorVerde()
    {
        using var salida = new SalidaTemporal("remuneracion-gate-interior-agosto-");
        Ejecutar(
            new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 },
            Insumos.CarpetaInsumosAgosto,
            Insumos.PlantillaAgosto2026082,
            salida.Ruta);

        var gate = new ValidadorTotalesR1Workbook();
        var ex = Record.Exception(() => gate.ExigirFormaRecompuestaInterior(salida.Ruta, Esperados(InterioresR1Esperados.Agosto)));

        Assert.True(
            ex is null,
            "Gate interior R1 de agosto en ROJO (intencional en T1: el interior conserva el anclaje julio; T2 lo cierra). Detalle: " + ex?.Message);
    }

    private static IReadOnlyList<InteriorR1Esperado> Esperados(IReadOnlyList<InterioresR1Esperados.InteriorR1> fixture) =>
        fixture.Select(v => new InteriorR1Esperado(v.AseId, v.Celda, v.Etiqueta, v.Formula, v.ValorLiteral)).ToList();

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
