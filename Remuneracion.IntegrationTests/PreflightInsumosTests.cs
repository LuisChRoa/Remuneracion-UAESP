using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 26 (WU-1, T1+T2): preflight de insumos del período con insumos REALES (nunca fixtures
/// sintéticas). Cubre:
///   - R-F-1/R-F-2/R-F-3 (S1): Agosto-2 2026082 → UN error con los 5 archivos de conciliación
///     enumerados de una vez, en lenguaje administrativo, sin iniciar el procesamiento.
///   - R-F-4 (S2/S3): períodos completos Q1/Q2 → CERO faltantes; Q1 no exige reportes de Q2.
///   - R-F-2 (S4/S6): agregación de faltantes múltiples y carpeta parcial.
///   - R-F-5 (S5): guardrail single-ASE con el mismo código y formateador.
///   - S7: el mensaje no contiene jerga técnica y sí nombres reconocibles.
/// </summary>
public sealed class PreflightInsumosTests
{
    // ── S1 / R-F-1 / R-F-2 / R-F-3: el caso real Agosto-2 ───────────────────────────────────────

    [Fact]
    public void Preflight_AgostoSinConciliaciones_ListaLosCincoDeUnaVez_YNoInicia()
    {
        // Plan 27: Agosto-2 YA trae Conciliaciones/ (corregidas por el usuario); el escenario
        // "carpeta ausente" del Plan 26 se reproduce sobre una copia temporal (insumo real).
        var carpeta = CopiarPeriodo(Insumos.CarpetaInsumosAgosto);
        Directory.Delete(Path.Combine(carpeta, "Conciliaciones"), recursive: true);
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-preflight-2026082-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, "Remuneracion2026082.xlsx");

        try
        {
            var procesador = CrearProcesadorPeriodo();
            var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
                procesador.Ejecutar(new SolicitudProcesoPeriodo
                {
                    Periodo = new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 },
                    CarpetaPeriodo = carpeta,
                    RutaPlantilla = Insumos.PlantillaQ2,
                    RutaSalida = salida
                }));

            // Código idéntico al existente (D-B): compatibilidad de guía de pantalla/CLI.
            Assert.Equal(CodigoError.FuenteNoEncontrada, ex.Codigo);

            // UN solo error que lista los 5 archivos de conciliación DE UNA VEZ (D-F), no el primero.
            Assert.Contains("2026082", ex.Message, StringComparison.Ordinal);
            Assert.Contains("No se procesó ningún ASE", ex.Message, StringComparison.Ordinal);
            Assert.Contains("Conciliaciones", ex.Message, StringComparison.Ordinal);
            Assert.Contains("Reciprocidad EAAB", ex.Message, StringComparison.Ordinal);
            Assert.Contains("ENEL", ex.Message, StringComparison.Ordinal);
            Assert.Contains("ENERBIT", ex.Message, StringComparison.Ordinal);
            Assert.Contains("Occidente Directa", ex.Message, StringComparison.Ordinal);
            Assert.Contains("Otros", ex.Message, StringComparison.Ordinal);

            // Cero procesamiento: no hay salida creada (el flujo no inicia).
            Assert.False(File.Exists(salida), "El preflight debe abortar antes de leer/escribir cualquier workbook.");
        }
        finally
        {
            BorrarCarpeta(carpeta);
            BorrarCarpeta(salidaDir);
        }
    }

    // ── S2 / S3 / R-F-4: períodos completos → cero faltantes ────────────────────────────────────

    [Fact]
    public void Preflight_PeriodosCompletosQ1YQ2_SinFaltantes()
    {
        var validador = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator());

        Assert.Empty(validador.Validar(Insumos.CarpetaPeriodo, Insumos.Periodo()));
        Assert.Empty(validador.Validar(Insumos.CarpetaPeriodoQ2, Insumos.PeriodoQ2()));
    }

    [Fact]
    public void Preflight_Q1_NoExigeReportesExclusivosDeQ2()
    {
        // Las fuentes reales de Q1 no traen saldos-notas ni retribución-negativa: si el preflight
        // las exigiera en Q1, la lista no estaría vacía. Gobierno por Periodo.NumeroQuincena (D-D).
        var validador = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator());
        var faltantes = validador.Validar(Insumos.CarpetaPeriodo, Insumos.Periodo());

        Assert.Empty(faltantes);
        Assert.DoesNotContain(faltantes, f => f.QueFalta.Contains("Saldos a favor", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(faltantes, f => f.QueFalta.Contains("Retribución negativa", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Preflight_Q2_SiExigeSaldosNotasYRetribucionNegativa()
    {
        var carpeta = CopiarPeriodo(Insumos.CarpetaPeriodoQ2);
        try
        {
            BorrarPorPrefijo(Path.Combine(carpeta, "3-Ciudad Limpia"), "SaldosaFavorAplicadosPorNotas");

            var faltantes = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator()).Validar(carpeta, Insumos.PeriodoQ2());

            Assert.Single(faltantes);
            Assert.Equal("ASE 3", faltantes[0].Alcance);
            Assert.Contains("Saldos a favor aplicados por notas", faltantes[0].QueFalta, StringComparison.Ordinal);
        }
        finally
        {
            BorrarCarpeta(carpeta);
        }
    }

    // ── S4 / S6: agregación múltiple y carpeta parcial ──────────────────────────────────────────

    [Fact]
    public void Preflight_FaltantesMultiples_ListaCompletaEnOrdenEstable()
    {
        // Copia real de Q2; se eliminan 3 insumos de ámbitos distintos (ASE1, ASE2 y período).
        var carpeta = CopiarPeriodo(Insumos.CarpetaPeriodoQ2);
        try
        {
            BorrarPorPrefijo(Path.Combine(carpeta, "1-Promoambiental"), "Recaudoporcomponente");
            BorrarPorPrefijo(Path.Combine(carpeta, "2-Lime"), "ReportePagosxBanco");
            BorrarPorPrefijo(carpeta, "R10_");

            var faltantes = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator()).Validar(carpeta, Insumos.PeriodoQ2());

            Assert.Equal(3, faltantes.Count);
            // Orden estable: ASE 1..5, luego período (D-F).
            Assert.Equal("ASE 1", faltantes[0].Alcance);
            Assert.Contains("Recaudo por componente (R1)", faltantes[0].QueFalta, StringComparison.Ordinal);
            Assert.Equal("ASE 2", faltantes[1].Alcance);
            Assert.Contains("Reporte de recaudo por banco", faltantes[1].QueFalta, StringComparison.Ordinal);
            Assert.Equal("Período", faltantes[2].Alcance);
            Assert.Contains("R10_Remuneracion_2026072", faltantes[2].QueFalta, StringComparison.Ordinal);
        }
        finally
        {
            BorrarCarpeta(carpeta);
        }
    }

    [Fact]
    public void Preflight_ConciliacionesConCuatroDeCinco_NombraLaEmpresaFaltante()
    {
        var carpeta = CopiarPeriodo(Insumos.CarpetaPeriodoQ2);
        try
        {
            BorrarPorPrefijo(Path.Combine(carpeta, "Conciliaciones"), "Conjunta ENERBIT");

            var faltantes = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator()).Validar(carpeta, Insumos.PeriodoQ2());

            Assert.Single(faltantes);
            Assert.Contains("ENERBIT", faltantes[0].QueFalta, StringComparison.Ordinal);
            Assert.Contains("Conciliaciones", faltantes[0].DondeDebeIr, StringComparison.Ordinal);
        }
        finally
        {
            BorrarCarpeta(carpeta);
        }
    }

    // ── S5 / R-F-5: guardrail single-ASE ────────────────────────────────────────────────────────

    [Fact]
    public void Preflight_SingleAse_RutaInexistente_MismoCodigoYReporteNombrado()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-preflight-1ase-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, "salida.xlsx");

        var procesador = new ProcesadorRemuneracion(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter());

        var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoAse
            {
                Ase = Insumos.Ase(1),
                Periodo = Insumos.Periodo(),
                RutaR1 = Insumos.R1(1),
                RutaR2 = Insumos.R2(1),
                RutaR4 = Path.Combine(Path.GetTempPath(), "no-existe-" + Guid.NewGuid().ToString("N") + ".xlsx"),
                RutaPlantilla = Insumos.Plantilla,
                RutaSalida = salida
            }));

        Assert.Equal(CodigoError.FuenteNoEncontrada, ex.Codigo);
        Assert.Contains("Reversión por componente (R4)", ex.Message, StringComparison.Ordinal);
        Assert.Contains("No se procesó el ASE 1", ex.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(salida), "El guardrail single-ASE no debe leer ni escribir.");
    }

    // ── S7 / R-F-2: lenguaje administrativo, sin jerga ──────────────────────────────────────────

    [Fact]
    public void Preflight_Mensaje_SinJergaTecnica_ConNombresReconocibles()
    {
        // Plan 27: Agosto-2 ya trae Conciliaciones/; se reproduce el escenario incompleto en temp
        // para que el mensaje siga ejercitando el ítem de la carpeta ausente con nombres reales.
        var carpeta = CopiarPeriodo(Insumos.CarpetaInsumosAgosto);
        Directory.Delete(Path.Combine(carpeta, "Conciliaciones"), recursive: true);
        try
        {
            var periodo = new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 };
            var faltantes = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator())
                .Validar(carpeta, periodo);
            var mensaje = FormateadorInsumosFaltantes.Mensaje(periodo, faltantes);

            foreach (var prohibido in new[] { "prefijo", "matcher", "finder", "TopDirectoryOnly", "ValidadorInsumosPeriodo", "ArchivoFuenteLocator", "runtime" })
            {
                Assert.DoesNotContain(prohibido, mensaje, StringComparison.OrdinalIgnoreCase);
            }

            Assert.Contains("Conciliaciones", mensaje, StringComparison.Ordinal);
            Assert.Contains("vuelva a ejecutar", mensaje, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            BorrarCarpeta(carpeta);
        }
    }

    // ── Helpers ─────────────────────────────────────────────────────────────────────────────────

    private static ProcesadorPeriodo CrearProcesadorPeriodo() =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader());

    /// <summary>Copia recursiva de solo xlsx (mismo patrón que las regresiones existentes).</summary>
    private static string CopiarPeriodo(string origen)
    {
        var destino = Path.Combine(Path.GetTempPath(), "remuneracion-preflight-copia-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(destino);
        foreach (var archivo in Directory.EnumerateFiles(origen, "*.xlsx", SearchOption.AllDirectories))
        {
            var relativo = Path.GetRelativePath(origen, archivo);
            var destinoArchivo = Path.Combine(destino, relativo);
            Directory.CreateDirectory(Path.GetDirectoryName(destinoArchivo)!);
            File.Copy(archivo, destinoArchivo);
        }

        return destino;
    }

    private static void BorrarPorPrefijo(string carpeta, string prefijo)
    {
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"El fixture debía traer '{prefijo}' en {carpeta}.");
        File.Delete(archivo!);
    }

    private static void BorrarCarpeta(string ruta)
    {
        try
        {
            if (Directory.Exists(ruta))
            {
                Directory.Delete(ruta, recursive: true);
            }
        }
        catch
        {
            // best-effort
        }
    }
}
