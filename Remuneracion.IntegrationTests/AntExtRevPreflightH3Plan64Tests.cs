using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 64 (H3.4, REQ-08 — CA-04 / D-07): preflight del insumo <c>RecaudosReversados_*</c>.
///
/// <para><b>Qué se prueba y con qué oráculo.</b> El runtime lee el recaudo reversado para el bloque
/// <c>ANT EXT-REV</c> en AMBAS quincenas, así que el preflight del período debe exigirlo en ambas
/// (superficie = lo que el runtime resuelve, Plan 26 D-D). Los casos se construyen sobre
/// <b>copias temporales de los insumos reales</b> (nunca fixtures sintéticas): al insumo real se le
/// quita o se le corrompe el archivo y se mide el efecto.</para>
///
/// <list type="bullet">
/// <item>Período sin el insumo ⇒ UN solo error agregado (<c>ERR-FUENTE-NO-ENCONTRADA</c>), antes de
/// abrir ningún workbook y sin salida: el preflight no hace un error por ASE ni arranca el loop.</item>
/// <item>Archivo presente pero que no es un Excel (copia de página web) ⇒ ítem administrativo de la
/// misma lista, con la variante "parece una copia de una página web".</item>
/// <item>El mensaje administrativo dice QUÉ falta, DÓNDE debe ir (carpeta del ASE + inicio de nombre)
/// y QUÉ HACER, sin jerga técnica (Plan 26 D-E).</item>
/// <item>El finder dedicado (<see cref="ILocalizadorArchivosAse.BuscarRecaudosReversados"/>) resuelve
/// el mismo archivo que el runtime y normaliza mayúsculas/diacríticos (prefijo
/// <c>recaudosreversados</c>, sin fechas).</item>
/// </list>
///
/// <para>Single-ASE queda fuera por decisión D-01: <c>ProcesadorRemuneracion</c> no exige este
/// insumo y su guardrail sigue siendo R1/R2/R4 (Plan 26 §V7).</para>
/// </summary>
public sealed class AntExtRevPreflightH3Plan64Tests
{
    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia"
    ];

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Rutas explícitas (prohibido Insumos.Raiz() en lo nuevo — Plan 64 §2.1)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static string RaizRepositorio()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new InvalidOperationException($"No se encontró AGENTS.md partiendo de '{AppContext.BaseDirectory}'.");
    }

    private static string Ruta(params string[] partes) => Path.Combine([RaizRepositorio(), .. partes]);

    private static string PlantillaV4() =>
        Ruta("Docs", "Plantilla estandarizada", "v4", "Formato Remuneracion 2026xx-x Total v4.xlsx");

    private static string CarpetaInsumos(string carpetaPeriodo) => Ruta("Docs", carpetaPeriodo, "Insumos");

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H3.4 (1) — período sin el insumo: UN error agregado, antes de leer, sin salida. Ambas quincenas.
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// CA-04: en Q1 y en Q2, si falta el <c>RecaudosReversados_*.xlsx</c> de un ASE, el preflight
    /// enumera UN solo ítem (no uno por ASE ni uno por archivo), el procesador aborta con
    /// <c>ERR-FUENTE-NO-ENCONTRADA</c> ANTES de localizar/leer fuentes y no queda salida escrita.
    /// </summary>
    [Theory]
    [InlineData("2026091", "Prueba Septiembre 1", 2)]   // Q1
    [InlineData("2026092", "Prueba Septiembre-2", 5)]   // Q2
    public void Preflight_SinElInsumo_UnErrorAgregadoAntesDeLeer_SinSalida(
        string codigo, string carpetaPeriodo, int aseId)
    {
        var periodo = Periodo.Parse(codigo);
        var carpeta = CopiarPeriodo(CarpetaInsumos(carpetaPeriodo));
        var salidaDir = NuevoDirectorio();
        var salida = Path.Combine(salidaDir, periodo.NombreArchivo);
        var progreso = new ListaProgreso();

        try
        {
            var carpetaAse = Path.Combine(carpeta, CarpetasAse[aseId - 1]);
            BorrarPorPrefijo(carpetaAse, "RecaudosReversados");

            // (a) El preflight enumera UN único faltante: el insumo del ASE intervenido.
            var faltantes = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator()).Validar(carpeta, periodo);
            var item = Assert.Single(faltantes);
            Assert.Equal($"ASE {aseId}", item.Alcance);
            Assert.Contains("Recaudos reversados por aplicar (ANT EXT-REV)", item.QueFalta, StringComparison.Ordinal);

            // (b) El procesador aborta con UN error agregado y en lenguaje administrativo.
            var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
                CrearProcesadorPeriodo().Ejecutar(
                    new SolicitudProcesoPeriodo
                    {
                        Periodo = periodo,
                        CarpetaPeriodo = carpeta,
                        RutaPlantilla = PlantillaV4(),
                        RutaSalida = salida
                    },
                    progreso));

            Assert.Equal(CodigoError.FuenteNoEncontrada, ex.Codigo);
            Assert.Contains(codigo, ex.Message, StringComparison.Ordinal);
            Assert.Contains("No se procesó ningún ASE", ex.Message, StringComparison.Ordinal);
            Assert.Contains("Recaudos reversados por aplicar (ANT EXT-REV)", ex.Message, StringComparison.Ordinal);
            Assert.Contains($"ASE {aseId}", ex.Message, StringComparison.Ordinal);
            Assert.Contains(Path.GetFileName(carpetaAse), ex.Message, StringComparison.Ordinal);
            Assert.Contains("RecaudosReversados_", ex.Message, StringComparison.Ordinal);

            // (c) Abortó EN el preflight: anunció la verificación y nunca llegó a resolver fuentes.
            Assert.Contains(progreso.Lineas, linea => linea.Contains("Verificando insumos", StringComparison.Ordinal));
            Assert.DoesNotContain(progreso.Lineas, linea => linea.Contains("localizando fuentes", StringComparison.Ordinal));
            Assert.DoesNotContain(progreso.Lineas, linea => linea.Contains("leyendo R1", StringComparison.Ordinal));

            // (d) Cero salida: no se abrió ni se creó ningún workbook.
            Assert.False(File.Exists(salida), "El preflight no debe dejar salida ante un insumo faltante.");
            Assert.Empty(Directory.EnumerateFiles(salidaDir, "*.xlsx", SearchOption.TopDirectoryOnly));

            // (e) El hallazgo corresponde al insumo que se quitó: el archivo ya no está en la carpeta
            //     del ASE (los demás insumos siguen ahí y no generan ítem).
            Assert.Empty(Directory.EnumerateFiles(carpetaAse, "RecaudosReversados*.xlsx", SearchOption.TopDirectoryOnly));
        }
        finally
        {
            BorrarCarpeta(carpeta);
            BorrarCarpeta(salidaDir);
        }
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H3.4 (2) — firma inválida: variante administrativa (Plan 27), mismo camino que los demás insumos
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Plan 27 (D-B) aplicado al insumo nuevo: el archivo existe pero su firma es de una página web
    /// (MHTML renombrado a <c>.xlsx</c>) ⇒ un ítem administrativo —no técnico— en la misma lista
    /// numerada, y el procesador aborta antes de leer, sin salida.
    /// </summary>
    [Fact]
    public void Preflight_FirmaInvalidaDelInsumo_VarianteAdministrativa_SinJerga()
    {
        const string codigo = "2026091";
        const int aseId = 3;
        var periodo = Periodo.Parse(codigo);
        var carpeta = CopiarPeriodo(CarpetaInsumos("Prueba Septiembre 1"));
        var salidaDir = NuevoDirectorio();
        var salida = Path.Combine(salidaDir, periodo.NombreArchivo);

        try
        {
            var carpetaAse = Path.Combine(carpeta, CarpetasAse[aseId - 1]);
            var insumo = BuscarPorPrefijo(carpetaAse, "RecaudosReversados");
            File.WriteAllText(insumo, InspectorFirmaExcelTests.MhtmlBlink());

            var faltantes = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator()).Validar(carpeta, periodo);
            var item = Assert.Single(faltantes);
            Assert.Equal($"ASE {aseId}", item.Alcance);
            Assert.Contains("no es un Excel válido", item.QueFalta, StringComparison.Ordinal);
            Assert.Contains("parece una copia de una página web", item.QueFalta, StringComparison.Ordinal);
            Assert.Contains(Path.GetFileName(insumo), item.QueFalta, StringComparison.Ordinal);
            Assert.Contains("vuelva a ejecutar", item.QueHacer, StringComparison.OrdinalIgnoreCase);

            // El detalle técnico (firma + bytes) NO viaja al usuario.
            var mensaje = FormateadorInsumosFaltantes.Mensaje(periodo, faltantes);
            foreach (var prohibido in JergaProhibida)
            {
                Assert.DoesNotContain(prohibido, mensaje, StringComparison.OrdinalIgnoreCase);
            }

            var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
                CrearProcesadorPeriodo().Ejecutar(new SolicitudProcesoPeriodo
                {
                    Periodo = periodo,
                    CarpetaPeriodo = carpeta,
                    RutaPlantilla = PlantillaV4(),
                    RutaSalida = salida
                }));

            Assert.Equal(CodigoError.FuenteNoEncontrada, ex.Codigo);
            Assert.Contains(Path.GetFileName(insumo), ex.Message, StringComparison.Ordinal);
            Assert.Contains("parece una copia de una página web", ex.Message, StringComparison.Ordinal);
            Assert.False(File.Exists(salida), "El preflight de firma debe abortar antes de leer/escribir.");
        }
        finally
        {
            BorrarCarpeta(carpeta);
            BorrarCarpeta(salidaDir);
        }
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H3.4 (3) — mensaje administrativo: QUÉ / DÓNDE / QUÉ HACER, sin jerga (Plan 26 D-E)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// El nombre administrativo del insumo nuevo no inventa jerga ni obliga a conocer el mecanismo:
    /// dice qué falta (el recaudo reversado que alimenta ANT EXT-REV), dónde debe ir (la carpeta del
    /// ASE) y qué hacer. Prohibido "prefijo", "finder", "matcher", "runtime", nombres de clase o de
    /// códigos internos.
    /// </summary>
    [Fact]
    public void Mensaje_Administrativo_DiceQueDondeYQueHacer_SinJerga()
    {
        const int aseId = 4;
        var periodo = Periodo.Parse("2026092");
        var carpeta = CopiarPeriodo(CarpetaInsumos("Prueba Septiembre-2"));

        try
        {
            var carpetaAse = Path.Combine(carpeta, CarpetasAse[aseId - 1]);
            BorrarPorPrefijo(carpetaAse, "RecaudosReversados");

            var faltantes = new ValidadorInsumosPeriodo(new ArchivoFuenteLocator()).Validar(carpeta, periodo);
            Assert.Single(faltantes);
            var mensaje = FormateadorInsumosFaltantes.Mensaje(periodo, faltantes);

            foreach (var prohibido in JergaProhibida)
            {
                Assert.DoesNotContain(prohibido, mensaje, StringComparison.OrdinalIgnoreCase);
            }

            // QUÉ falta, con el nombre que usa el área administrativa (y el bloque al que alimenta).
            Assert.Contains("Falta Recaudos reversados por aplicar (ANT EXT-REV)", mensaje, StringComparison.Ordinal);
            Assert.Contains($"del ASE {aseId}", mensaje, StringComparison.Ordinal);
            // DÓNDE: la carpeta del ASE dentro del período + inicio de nombre esperado.
            Assert.Contains($"\"{Path.GetFileName(carpetaAse)}\"", mensaje, StringComparison.Ordinal);
            Assert.Contains("dentro de la carpeta del período", mensaje, StringComparison.Ordinal);
            Assert.Contains("\"RecaudosReversados_\"", mensaje, StringComparison.Ordinal);
            // QUÉ HACER.
            Assert.Contains("vuelva a ejecutar", mensaje, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            BorrarCarpeta(carpeta);
        }
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // H3.1 — finder dedicado: mismo archivo que el runtime, prefijo normalizado sin fechas
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Paridad preflight↔runtime (Plan 26 R-F-6): <c>BuscarRecaudosReversados</c> resuelve EL MISMO
    /// archivo que el insumo real del ASE, ignora el caso y los diacríticos y no depende del rango de
    /// fechas del nombre; devuelve <c>null</c> cuando el insumo no está (ese <c>null</c> es el ítem
    /// del preflight, no un fallo silencioso).
    /// </summary>
    [Fact]
    public void Finder_BuscarRecaudosReversados_MismoArchivoQueElRuntime_YNormalizaElNombre()
    {
        var locator = new ArchivoFuenteLocator();
        var carpetaAse = Path.Combine(CarpetaInsumos("Prueba Septiembre 1"), CarpetasAse[0]);

        // (a) El insumo real del ASE se resuelve, y es el mismo archivo que resolvía el prefijo
        //     exacto que el runtime usaba de forma provisional (misma semántica, finder dedicado).
        var real = locator.BuscarRecaudosReversados(carpetaAse);
        Assert.NotNull(real);
        Assert.True(File.Exists(real));
        Assert.Equal(locator.BuscarArchivo(carpetaAse, "RecaudosReversados"), real);

        // (b) Prefijo normalizado: un nombre en MAYÚSCULAS con otro rango de fechas se resuelve igual
        //     (prohibido quemar nombre exacto, rango de fechas o un literal con acento frágil).
        var variante = NuevoDirectorio();
        var sinElInsumo = NuevoDirectorio();
        try
        {
            var copia = Path.Combine(variante, "RECAUDOSREVERSADOS_to_date01012026_to_date15012026___999.xlsx");
            File.Copy(real!, copia);

            Assert.Equal(copia, locator.BuscarRecaudosReversados(variante));

            // (c) Carpeta sin el insumo ⇒ null (el preflight lo convierte en ítem administrativo).
            Assert.Null(locator.BuscarRecaudosReversados(sinElInsumo));
            Assert.Null(locator.BuscarRecaudosReversados(Path.Combine(sinElInsumo, "no-existe")));
        }
        finally
        {
            BorrarCarpeta(variante);
            BorrarCarpeta(sinElInsumo);
        }
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Helpers
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>Jerga prohibida en todo texto administrativo (doctrina Plan 26 D-E).</summary>
    private static readonly string[] JergaProhibida =
    [
        "prefijo", "matcher", "finder", "TopDirectoryOnly", "ValidadorInsumosPeriodo",
        "ArchivoFuenteLocator", "runtime", "firma binaria", "MHTML", "magic", "0x"
    ];

    private static ProcesadorPeriodo CrearProcesadorPeriodo() =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader());

    private static string NuevoDirectorio()
    {
        var dir = Path.Combine(Path.GetTempPath(), "p64-h3-preflight-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    /// <summary>Copia recursiva de los insumos reales (xlsx) del período a una ruta temporal fresca.</summary>
    private static string CopiarPeriodo(string origen)
    {
        var destino = NuevoDirectorio();
        foreach (var archivo in Directory.EnumerateFiles(origen, "*.xlsx", SearchOption.AllDirectories))
        {
            var relativo = Path.GetRelativePath(origen, archivo);
            var objetivo = Path.Combine(destino, relativo);
            Directory.CreateDirectory(Path.GetDirectoryName(objetivo)!);
            File.Copy(archivo, objetivo);
        }

        return destino;
    }

    private static string BuscarPorPrefijo(string carpeta, string prefijo)
    {
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f)
                .StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"El insumo real debía traer '{prefijo}' en {carpeta}.");
        return archivo!;
    }

    private static void BorrarPorPrefijo(string carpeta, string prefijo)
    {
        var archivo = BuscarPorPrefijo(carpeta, prefijo);
        File.Delete(archivo);
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
            // best-effort (el fixture corre en temp)
        }
    }

    /// <summary>Colector determinista de progreso (evita el post asíncrono de <c>Progress&lt;T&gt;</c>).</summary>
    private sealed class ListaProgreso : IProgress<string>
    {
        public List<string> Lineas { get; } = [];

        public void Report(string value) => Lineas.Add(value);
    }
}
