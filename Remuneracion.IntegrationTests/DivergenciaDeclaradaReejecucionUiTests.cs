using System.IO;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Remuneracion.WinForms;
using Xunit;

namespace Remuneracion.IntegrationTests;

// Plan 42 (T3-UI, R-RE-RUN-UI/R-UI-1, D-D/D-F): superficie de reconocimiento explícito de la UI.
// El presentador puro (ReconocimientoDivergencia.cs, archivo vinculado desde el proyecto
// WinForms: la misma unidad que compila la UI) más el flujo exacto que ejecuta Form1:
// pasada 1 sin acto → captura con SoloDeclaradas → ids de la clasificación → pasada 2 que
// cierra con reporte. La pasada 2 end-to-end corre contra el OFICIAL Q2 a temp FRESCA.
public sealed class DivergenciaDeclaradaReejecucionUiTests : IDisposable
{
    public void Dispose() => Serilog.Log.CloseAndFlush();

    // ── Puerta de ofrecimiento (D-F): SOLO SoloDeclaradas abre la superficie ─────────────

    [Fact]
    public void Ui_SoloDeclaradas_OfreceReconocimiento()
    {
        var ex = new CalculoInvalidoException(
            CodigoError.Validacion, "bloqueo", ClasificacionSoloDeclarada());

        Assert.True(ReconocimientoDivergencia.OfreceReconocimiento(ex));
    }

    [Fact]
    public void Ui_HayNoDeclaradas_NoOfreceReconocimiento()
    {
        // Mixta (D-F: la declarada es la última puerta): FAIL duro, sin superficie.
        var veredictos = QuinceVeredictosDeclarados().ToList();
        veredictos.Add(new VeredictoDivergenciaValidacion
        {
            AseId = 1, Celda = "C15", Declarada = false, EntradaId = null,
            MontoObservado = 1m, MontoEsperado = 2m
        });
        var ex = new CalculoInvalidoException(
            CodigoError.Validacion, "bloqueo",
            new ClasificacionDivergenciasDeclaradas { Veredictos = veredictos });

        Assert.False(ReconocimientoDivergencia.OfreceReconocimiento(ex));
    }

    [Fact]
    public void Ui_BloqueoSinClasificacion_NoOfreceReconocimiento()
    {
        // El FAIL por no-declaradas no porta clasificación (mismo camino que HayNoDeclaradas).
        var ex = new CalculoInvalidoException(CodigoError.Validacion, "bloqueo");

        Assert.False(ReconocimientoDivergencia.OfreceReconocimiento(ex));
        Assert.False(ReconocimientoDivergencia.OfreceReconocimiento(null));
    }

    [Fact]
    public void Ui_SinDivergencias_NoOfreceReconocimiento()
    {
        var ex = new CalculoInvalidoException(
            CodigoError.Validacion, "bloqueo",
            new ClasificacionDivergenciasDeclaradas());

        Assert.False(ReconocimientoDivergencia.OfreceReconocimiento(ex));
    }

    // ── Deduplicación: 15 veredictos (5 ASE × 3 celdas) → una fila por divergencia ─────────

    [Fact]
    public void Ui_QuinceVeredictos_AgrupaEnTresFilasPorCelda()
    {
        var filas = ReconocimientoDivergencia.AgruparPorDivergencia(ClasificacionSoloDeclarada());

        Assert.Equal(3, filas.Count);
        Assert.Equal(["D34", "F34", "O25"], filas.Select(v => v.Celda).OrderBy(c => c, StringComparer.Ordinal));
        Assert.All(filas, f => Assert.Equal("OCCIDENTE-ASE5", f.EntradaId));
    }

    [Fact]
    public void Ui_IdsAceptados_SonLasEntradasDeclaradasDeLaClasificacion()
    {
        Assert.Equal(
            ["OCCIDENTE-ASE5"],
            ReconocimientoDivergencia.IdsAceptados(ClasificacionSoloDeclarada()));
    }

    // ── Formato de negocio + copy administrativo sin jerga (§2.3, doctrina Plan 27 D-E) ──────

    [Fact]
    public void Ui_FormatoNegocio_MilesConPuntoYComaDecimal()
    {
        Assert.Equal("417.519,62", ReconocimientoDivergencia.FormatearMontoNegocio(417519.62m));
        Assert.Equal("0,00", ReconocimientoDivergencia.FormatearMontoNegocio(0m));
    }

    [Fact]
    public void Ui_Copy_NombraQueDuenoYPideSinJergaDeCodigo()
    {
        var fila = ReconocimientoDivergencia.AgruparPorDivergencia(ClasificacionSoloDeclarada()).First(v => v.Celda == "O25");
        var entrada = ReconocimientoDivergencia.BuscarEntrada("OCCIDENTE-ASE5");

        var que = ReconocimientoDivergencia.TextoQue(fila);
        var deQuien = ReconocimientoDivergencia.TextoDeQuienEs(entrada);
        var aceptacion = ReconocimientoDivergencia.TextoAceptacion("OCCIDENTE-ASE5", fila.Delta, entrada?.Dueno);

        Assert.Contains("O25", que, StringComparison.Ordinal);
        Assert.Contains("417.519,62", que, StringComparison.Ordinal);
        Assert.Contains("UAESP", deQuien, StringComparison.Ordinal);
        Assert.Contains("OCCIDENTE-ASE5", aceptacion, StringComparison.Ordinal);
        Assert.Contains("417.519,62", aceptacion, StringComparison.Ordinal);
        Assert.Contains("UAESP", aceptacion, StringComparison.Ordinal);

        foreach (var texto in new[] { que, deQuien, aceptacion })
        {
            Assert.DoesNotContain("ERR-", texto, StringComparison.Ordinal);
            Assert.DoesNotContain("si=", texto, StringComparison.Ordinal);
            Assert.DoesNotContain("<f>", texto, StringComparison.Ordinal);
            Assert.DoesNotContain("SeGatea", texto, StringComparison.Ordinal);
        }
    }

    // ── Flujo en dos pasadas (R-RE-RUN-UI): el baseline exige que sin acto LANCE ────────────

    [Fact]
    public void Ui_ReejecucionEnDosPasadas_PrimeraLanzaYSegundaCierraConReporte()
    {
        var dir = Path.Combine(Path.GetTempPath(), "p42-t3ui-rejecutar-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var salida = Path.Combine(dir, Insumos.PeriodoQ2().NombreArchivo);
            // Misma composición que Program.cs (con emisor de reporte: 9.º argumento).
            var procesador = new ProcesadorPeriodo(
                new ExcelDataReaderRecaudoReader(),
                new ExcelDataReaderWorkbookLeafInputReader(),
                new CalculoRemuneracion(),
                new ValidadorBasico(),
                new OpenXmlPlantillaWriter(),
                new ArchivoFuenteLocator(),
                new ExcelDataReaderDetRetriR10Reader(),
                new ValidacionOracleReader(),
                new ReporteDivergenciasDeclaradasWriter());

            // Pasada 1 (como la UI: sin aceptaciones) → LANZA con SoloDeclaradas.
            var bloqueo = Assert.Throws<CalculoInvalidoException>(() => procesador.Ejecutar(
                new SolicitudProcesoPeriodo
                {
                    Periodo = Insumos.PeriodoQ2(),
                    CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
                    RutaPlantilla = Insumos.PlantillaOficial(),
                    RutaSalida = salida
                }));
            Assert.True(ReconocimientoDivergencia.OfreceReconocimiento(bloqueo));

            // Acto consciente: los ids salen DE LA CLASIFICACIÓN (nada quemado, nada persistente).
            var aceptadas = ReconocimientoDivergencia.IdsAceptados(bloqueo.Clasificacion!);

            // Pasada 2 (como la UI tras "Aceptar y cerrar") → CIERRA con archivo + reporte.
            var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
                RutaPlantilla = Insumos.PlantillaOficial(),
                RutaSalida = salida,
                DivergenciasDeclaradasAceptadas = aceptadas
            });

            Assert.True(File.Exists(salida));
            Assert.False(string.IsNullOrWhiteSpace(resultado.ReporteDivergenciasDeclaradas));
            Assert.True(File.Exists(resultado.ReporteDivergenciasDeclaradas));
        }
        finally
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

    private static ClasificacionDivergenciasDeclaradas ClasificacionSoloDeclarada() =>
        new() { Veredictos = QuinceVeredictosDeclarados().ToList() };

    // 5 ASE × 3 celdas del mismo par declarado, Δ = 417.519,62 en los tres.
    private static IEnumerable<VeredictoDivergenciaValidacion> QuinceVeredictosDeclarados()
    {
        foreach (var ase in Enumerable.Range(1, 5))
        {
            foreach (var celda in new[] { "O25", "D34", "F34" })
            {
                yield return new VeredictoDivergenciaValidacion
                {
                    AseId = ase,
                    Celda = celda,
                    Declarada = true,
                    EntradaId = "OCCIDENTE-ASE5",
                    MontoObservado = 1000000m,
                    MontoEsperado = 1000000m - 417519.62m
                };
            }
        }
    }
}
