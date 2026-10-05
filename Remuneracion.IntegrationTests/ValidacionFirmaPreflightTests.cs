using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 27 (T3, Unidad B por firma + Unidad A por nombre de hoja) con insumos REALES:
///   - R-F-1/S5: las 15 conciliaciones reales tienen firma válida y Prueba2 corregida no trae
///     faltantes ni inválidos.
///   - R-F-3/S3: un MHTML en temp (patrón real "Saved by Blink") produce UN
///     <see cref="ArchivoFuenteNoEncontradoException"/> previo a todo I/O Excel, sin salida.
///   - R-F-4/S4: bytes desconocidos → variante "no es un Excel válido" sin "página web".
///   - R-F-5/S6: guardrail single-ASE con el mismo código y formateador.
///   - R-F-2/S7: sin jerga técnica en el mensaje; el detalle va al log.
///   - R-A-1/S1: <c>RESUMEN MES</c> por nombre en la conciliación multi-hoja de agosto.
///   - R-R-3: Prueba2 corregida supera la lectura de conciliaciones.
/// </summary>
public sealed class ValidacionFirmaPreflightTests
{
    private const string PrefijoRecip = "Conjunta Recip";

    [Fact]
    public void Preflight_QuinceConciliacionesReales_TienenFirmaValida()
    {
        var archivos = Insumos.ConciliacionesReales();
        Assert.Equal(15, archivos.Count);
        foreach (var archivo in archivos)
        {
            Assert.Equal(FirmaExcel.Valida, InspectorFirmaExcel.Inspeccionar(archivo).Firma);
        }
    }

    [Fact]
    public void Preflight_Prueba2Corregida_SinFaltantesNiInvalidos()
    {
        // S5: con las conciliaciones corregidas por el usuario, la lista queda vacía.
        var faltantes = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator())
            .Validar(Insumos.CarpetaInsumosAgosto, Insumos.PeriodoAgosto());
        Assert.Empty(faltantes);
    }

    [Fact]
    public void Preflight_MhtmlEnConciliacion_ListaItemWeb_YNoInicia()
    {
        var carpeta = CopiarPeriodo(Insumos.CarpetaPeriodoQ2);
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-firma-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, "salida.xlsx");
        var progreso = new ListaProgreso();
        try
        {
            var recip = Insumos.Conciliacion(carpeta, PrefijoRecip);
            File.WriteAllText(recip, InspectorFirmaExcelTests.MhtmlBlink());

            // El validador reporta UN ítem administrativo de copia web (S3).
            var faltantes = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator()).Validar(carpeta, Insumos.PeriodoQ2());
            var item = Assert.Single(faltantes);
            Assert.Equal("Período", item.Alcance);
            Assert.Contains("no es un Excel válido", item.QueFalta, StringComparison.Ordinal);
            Assert.Contains("parece una copia de una página web", item.QueFalta, StringComparison.Ordinal);
            Assert.Contains(Path.GetFileName(recip), item.QueFalta, StringComparison.Ordinal);

            // El procesador aborta con UN error antes de leer ningún workbook (R-F-3): cero salida.
            var procesador = CrearProcesadorPeriodo();
            var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
                procesador.Ejecutar(new SolicitudProcesoPeriodo
                {
                    Periodo = Insumos.PeriodoQ2(),
                    CarpetaPeriodo = carpeta,
                    RutaPlantilla = Insumos.PlantillaQ2,
                    RutaSalida = salida
                }, progreso));

            Assert.Equal(CodigoError.FuenteNoEncontrada, ex.Codigo);
            Assert.Contains(Path.GetFileName(recip), ex.Message, StringComparison.Ordinal);
            Assert.Contains("parece una copia de una página web", ex.Message, StringComparison.Ordinal);
            Assert.False(File.Exists(salida), "El preflight de firma debe abortar antes de leer/escribir.");

            // Se quedó en el preflight: nunca anunció la lectura de conciliaciones.
            Assert.Contains(progreso.Lineas, l => l.Contains("Verificando insumos", StringComparison.Ordinal));
            Assert.DoesNotContain(progreso.Lineas, l => l.Contains("Leyendo hojas Recaudo", StringComparison.Ordinal));
        }
        finally
        {
            BorrarCarpeta(carpeta);
            BorrarCarpeta(salidaDir);
        }
    }

    [Fact]
    public void Preflight_BytesDesconocidos_ListaItemSinPaginaWeb()
    {
        var carpeta = CopiarPeriodo(Insumos.CarpetaPeriodoQ2);
        try
        {
            var recip = Insumos.Conciliacion(carpeta, PrefijoRecip);
            File.WriteAllText(recip, "NO_ES_UN_EXCEL_NI_UNA_PAGINA");

            var item = Assert.Single(new ValidadorInsumosPeriodo(new ArchivoFuenteLocator()).Validar(carpeta, Insumos.PeriodoQ2()));
            Assert.Equal("Período", item.Alcance);
            Assert.Contains("no es un Excel válido", item.QueFalta, StringComparison.Ordinal);
            Assert.DoesNotContain("página web", item.QueFalta, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            BorrarCarpeta(carpeta);
        }
    }

    [Fact]
    public void Preflight_SingleAse_ArchivoWeb_MismoCodigoYMensajeAse()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "remuneracion-firma-1ase-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var r2 = Path.Combine(tempDir, "R2-web.xlsx");
        var salida = Path.Combine(tempDir, "salida.xlsx");
        var progreso = new ListaProgreso();
        try
        {
            File.WriteAllText(r2, InspectorFirmaExcelTests.MhtmlBlink());

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
                    RutaR2 = r2,
                    RutaR4 = Insumos.R4(1),
                    RutaPlantilla = Insumos.Plantilla,
                    RutaSalida = salida
                }, progreso));

            Assert.Equal(CodigoError.FuenteNoEncontrada, ex.Codigo);
            Assert.Contains("No se procesó el ASE 1", ex.Message, StringComparison.Ordinal);
            Assert.Contains("no es un Excel válido", ex.Message, StringComparison.Ordinal);
            Assert.Contains("parece una copia de una página web", ex.Message, StringComparison.Ordinal);
            Assert.False(File.Exists(salida));
            Assert.DoesNotContain(progreso.Lineas, l => l.Contains("Leyendo R1, R2 y R4", StringComparison.Ordinal));
        }
        finally
        {
            BorrarCarpeta(tempDir);
        }
    }

    [Fact]
    public void Preflight_MensajeConItemsInvalidos_SinJergaTecnica()
    {
        var carpeta = CopiarPeriodo(Insumos.CarpetaPeriodoQ2);
        try
        {
            File.WriteAllText(Insumos.Conciliacion(carpeta, PrefijoRecip), InspectorFirmaExcelTests.MhtmlBlink());

            var faltantes = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator()).Validar(carpeta, Insumos.PeriodoQ2());
            var mensaje = FormateadorInsumosFaltantes.Mensaje(Insumos.PeriodoQ2(), faltantes);

            foreach (var prohibido in new[] { "magic", "MIME", "firma binaria", "header", "MHTML", "0x" })
            {
                Assert.DoesNotContain(prohibido, mensaje, StringComparison.OrdinalIgnoreCase);
            }

            Assert.Contains("no es un Excel válido", mensaje, StringComparison.Ordinal);
            Assert.Contains("vuelva a ejecutar", mensaje, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            BorrarCarpeta(carpeta);
        }
    }

    [Fact]
    public void LecturaResumenMesPorNombre_RecipAgosto_LeiTresBloquesYColF()
    {
        // Unidad A (S1/R-A-1): la hoja RESUMEN MES es la 3ª de la conciliación de agosto; leerla
        // por nombre devuelve los 3 bloques (OPORTUNO/EXTEMP/TOTAL) con los valores de la col F.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var recaudos = reader.LeerRecaudosEmpresa(
            Insumos.PeriodoAgosto(),
            empresa => new ArchivoFuenteLocator().BuscarConciliacion(Insumos.CarpetaInsumosAgosto, empresa.PrefijoConciliacion));

        var recip = recaudos.Single(r => r.Empresa.Id == 1);

        // Testigo de sesión: ASE2 OPORTUNO col F = 18.262.430 (fila 4 del bloque OPORTUNO).
        Assert.Equal(18262430m, recip.Celdas["F4"]);

        // Los 3 bloques se leyeron (totales de cada bloque, col F).
        Assert.Equal(26878280m, recip.TotalOportuno);      // F9
        Assert.Equal(1027290m, recip.TotalExtemporaneo);   // F18
        Assert.Equal(27905570m, recip.Total);              // F27
    }

    [Fact]
    public void FlujoPrueba2Corregida_AvanzaMasAlaDeConciliaciones()
    {
        // R-R-3/S5: con las conciliaciones corregidas por el usuario, el flujo supera la lectura
        // de conciliaciones (puede fallar más adelante por otra causa del período; se reporta).
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-avance-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, "salida.xlsx");
        var progreso = new ListaProgreso();
        try
        {
            var procesador = CrearProcesadorPeriodo();
            var fallo = Record.Exception(() =>
                procesador.Ejecutar(new SolicitudProcesoPeriodo
                {
                    Periodo = Insumos.PeriodoAgosto(),
                    CarpetaPeriodo = Insumos.CarpetaInsumosAgosto,
                    RutaPlantilla = Insumos.PlantillaQ2,
                    RutaSalida = salida
                }, progreso));

            // El preflight no debe ser el que corta (puede fallar más adelante por otra causa).
            Assert.False(fallo is ArchivoFuenteNoEncontradoException,
                $"El preflight de insumos no debe fallar en Prueba2 corregida: {fallo}");
            // Llegó más allá de la lectura de conciliaciones: anunció el cálculo del consolidado.
            Assert.Contains(progreso.Lineas, l => l.Contains("Leyendo hojas Recaudo", StringComparison.Ordinal));
            Assert.Contains(progreso.Lineas, l => l.Contains("Calculando consolidados", StringComparison.Ordinal));
        }
        finally
        {
            BorrarCarpeta(salidaDir);
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

    private static string CopiarPeriodo(string origen)
    {
        var destino = Path.Combine(Path.GetTempPath(), "remuneracion-firma-copia-" + Guid.NewGuid().ToString("N"));
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

    /// <summary>Colector determinista de progreso (evita el post asíncrono de <c>Progress&lt;T&gt;</c>).</summary>
    private sealed class ListaProgreso : IProgress<string>
    {
        public List<string> Lineas { get; } = [];

        public void Report(string value) => Lineas.Add(value);
    }
}
