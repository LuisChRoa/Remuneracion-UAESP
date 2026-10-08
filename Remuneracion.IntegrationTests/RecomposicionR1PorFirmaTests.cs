using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 31 (T2, R-B-1..R-B-4 / S2/S4/S5): recomposición por firma de los totales visibles R1.
///
/// Corre el flujo 5-ASE end-to-end con insumos REALES y valida:
///   - S2 (texto-vs-manual agosto): las <c>&lt;f&gt;</c> recompuestas de agosto == texto transcrito
///     del manual (fixture E2 <see cref="TotalesR1Esperados.Agosto"/>, congelado en T1), 5 ASE
///     incl. ASE5-3Mes (`F548+F525+F494-L494-L525`).
///   - S5 (ASE3-agosto 0-Aplic): EXTEMP visible = literal 0 (sin <c>&lt;f&gt;</c>) y ASE5-3Mes
///     compuesto con 3 términos.
///   - S4 (julio-identidad, D-E): el pase final reproduce byte-idénticas las <c>&lt;f&gt;</c>
///     canónicas de julio (0-diff R1 vs la plantilla origen) — invariante de reversión.
///
/// Sin Excel/COM; lectura ZIP+XML BCL. Tolerancia ±0.5.
/// </summary>
public sealed class RecomposicionR1PorFirmaTests
{
    [Fact]
    public void S2_Agosto_RecompuestaPorFirma_TextoIgualAlManual()
    {
        using var salida = new SalidaTemporal("remuneracion-recomp-r1-agosto-");
        Ejecutar(
            new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 },
            Insumos.CarpetaInsumosAgosto,
            Insumos.PlantillaAgosto2026082,
            salida.Ruta);

        var celdas = ValidadorTotalesR1Workbook.LeerCeldas(salida.Ruta, TotalesR1Esperados.Hoja);
        foreach (var visible in TotalesR1Esperados.Agosto)
        {
            Assert.True(celdas.TryGetValue(visible.Celda, out var celda), $"Falta el visible {visible.Celda} (ASE {visible.AseId}, {visible.Etiqueta}) en la salida de agosto.");
            if (visible.Formula is null)
            {
                Assert.Null(celda!.Formula);
                Assert.Equal(0m, celda.Numero);
            }
            else
            {
                Assert.Equal(visible.Formula, celda!.Formula);
            }
        }
    }

    [Fact]
    public void S5_Agosto_Ase3SinAplicExtempCero_Ase5TresTerminos()
    {
        using var salida = new SalidaTemporal("remuneracion-recomp-r1-s5-");
        Ejecutar(
            new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 },
            Insumos.CarpetaInsumosAgosto,
            Insumos.PlantillaAgosto2026082,
            salida.Ruta);

        var celdas = ValidadorTotalesR1Workbook.LeerCeldas(salida.Ruta, TotalesR1Esperados.Hoja);

        // S5: ASE3-agosto no trae filas Aplicacion → EXTEMP = literal 0 (no fórmula vacía).
        var extempAse3 = celdas["F327"];
        Assert.Null(extempAse3.Formula);
        Assert.Equal(0m, extempAse3.Numero);

        // ASE5-agosto: 3 filas Mes (494/525/548) → TOT_OPT/TDF compuestos con 3 términos.
        Assert.Equal("F548+F525+F494-L494-L525", celdas["F554"].Formula);
        Assert.Equal("G525+G494", celdas["G554"].Formula);
    }

    [Fact]
    public void S4_Julio_IdentidadByteIdenticaYTextoManual()
    {
        using var salida = new SalidaTemporal("remuneracion-recomp-r1-julio-");
        var plantilla = Path.Combine(Insumos.Raiz(), "Docs", "Plantilla_Remuneracion.xlsx");
        Ejecutar(
            new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 },
            Insumos.CarpetaInsumosJulioQ2,
            plantilla,
            salida.Ruta);

        var celdasSalida = ValidadorTotalesR1Workbook.LeerCeldas(salida.Ruta, TotalesR1Esperados.Hoja);
        var celdasOrigen = ValidadorTotalesR1Workbook.LeerCeldas(plantilla, TotalesR1Esperados.Hoja);

        // D-E: toda <f> de R1 de la salida es byte-idéntica a la de la plantilla origen (Δ=0).
        foreach (var (referencia, celdaOrigen) in celdasOrigen)
        {
            if (celdaOrigen.Formula is null)
            {
                continue;
            }

            Assert.True(celdasSalida.TryGetValue(referencia, out var celdaSalida), $"Falta {referencia} en la salida de julio.");
            Assert.Equal(celdaOrigen.Formula, celdaSalida!.Formula);
        }

        // Y coincide con el texto del manual de julio (E2 congelado).
        foreach (var visible in TotalesR1Esperados.Julio)
        {
            Assert.True(celdasSalida.TryGetValue(visible.Celda, out var celda), $"Falta el visible {visible.Celda} (ASE {visible.AseId}) en la salida de julio.");
            Assert.Equal(visible.Formula, celda!.Formula);
        }
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
