using System.Globalization;
using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 28 (Unidad S + Unidad F): integridad estructural del workbook de salida y sello de
/// fechas del período.
///
/// Unidad S (D-A/D-B): la salida se genera SIN <c>xl/calcChain.xml</c> y con
/// <c>fullCalcOnLoad="1"</c> en <c>xl/workbook.xml</c>, en todos los caminos de guardado. T0
/// demostró 1003 entradas de calcChain inconsistentes en la salida de agosto (el espejo R1 mueve
/// filas y la cadena de la plantilla queda stale).
///
/// Unidad F (D-C/D-D/D-E): <c>CONSOLIDADO_TOTAL RECAUDO</c> G7/K7 se sellan con las fechas del
/// período leídas del R10 (G7=Desde, J7=Hasta). T0 fijó la columna real de "Fecha Hasta" del
/// template = K7 (J7 es el rótulo). Sin fechas legibles en el R10 → fail-fast; G7/K7 fórmula →
/// ERR-PLANTILLA.
///
/// SOLO insumos reales; las copias negativas van a temp (Docs/ nunca se toca). Verificación
/// estructural por parser (ZIP + XML), sin Excel.
/// </summary>
public sealed class IntegridadCalcChainYSelloFechasTests
{
    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";
    private const string EntradaCalcChain = "xl/calcChain.xml";

    // ── S1/S3: agosto regenerado desde la plantilla en ceros ────────────────────────────────
    [Fact]
    public void SalidaAgosto_DesdePlantillaEnCeros_SinCalcChainYFullCalcOnLoad_YFechasDeAgosto()
    {
        var carpetaPeriodo = PrepararPeriodoAgosto();
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-plan28-agosto-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, "Remuneracion2026082.xlsx");
        try
        {
            var resultado = CrearProcesador().Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoAgosto(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.PlantillaAgosto2026082,
                RutaSalida = salida
            });

            Assert.True(File.Exists(salida), "El flujo 5-ASE de agosto debe certificar la salida.");
            Assert.Equal(5, resultado.Leafs.Count);

            // Unidad S: sin cálculo de cadena stale + recálculo al abrir (verificado por ZIP/XML).
            var estructura = InspeccionarEstructura(salida);
            Assert.False(estructura.ContieneCalcChain, "La salida no debe contener xl/calcChain.xml.");
            Assert.True(estructura.FullCalcOnLoad, "xl/workbook.xml debe traer fullCalcOnLoad=\"1\".");
            Assert.True(estructura.CalculationChainPartAusente, "El WorkbookPart no debe exponer CalculationChainPart.");

            // Unidad F: fechas de AGOSTO (no las de julio de la plantilla).
            Assert.Equal(new DateTime(2026, 8, 16), LeerFechaSalida(salida, "G7"));
            Assert.Equal(new DateTime(2026, 8, 31), LeerFechaSalida(salida, "K7"));

            // Plan 31 (T3, S6/R-S-2): sello de proceso desde el R10 (D6/D7) en DetRetri/DetValiRetri.
            var hojaDetRetri = NombresHojaPeriodo.DetRetri(Insumos.PeriodoAgosto());
            var hojaDetVali = NombresHojaPeriodo.DetValiRetri(Insumos.PeriodoAgosto());
            Assert.Equal("02/09/2026", LeerCeldaTexto(salida, hojaDetRetri, "D6"));
            Assert.Equal("07:42 AM", LeerCeldaTexto(salida, hojaDetRetri, "D7"));
            Assert.Equal("02/09/2026", LeerCeldaTexto(salida, hojaDetVali, "D6"));
            Assert.Equal("07:42 AM", LeerCeldaTexto(salida, hojaDetVali, "D7"));

            // R-S-3 (no-duplicación): el sello NO reescribe CONSOLIDADO D6 (ya venía 46267 en la base).
            Assert.Equal(
                TestHelpers.LeerCeldaNumerica(Insumos.PlantillaAgosto2026082, HojaConsolidado, "D6"),
                TestHelpers.LeerCeldaNumerica(salida, HojaConsolidado, "D6"));
        }
        finally
        {
            BorrarArbol(carpetaPeriodo);
            BorrarArbol(salidaDir);
        }
    }

    // ── S2/S4: julio Q2 intacto (sello idempotente respecto de la plantilla) ─────────────────
    [Fact]
    public void SalidaJulioQ2_SinCalcChainYFullCalcOnLoad_FechasJulioIntactas()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-plan28-julioq2-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.PeriodoQ2().NombreArchivo);
        try
        {
            CrearProcesador().Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
                RutaPlantilla = Insumos.PlantillaQ2,
                RutaSalida = salida
            });

            var estructura = InspeccionarEstructura(salida);
            Assert.False(estructura.ContieneCalcChain);
            Assert.True(estructura.FullCalcOnLoad);
            // T0: la plantilla Q2 trae 46219/46234 (16/07–31/07) → el sello es idéntico (Δ=0).
            Assert.Equal(new DateTime(2026, 7, 16), LeerFechaSalida(salida, "G7"));
            Assert.Equal(new DateTime(2026, 7, 31), LeerFechaSalida(salida, "K7"));

            // Plan 31 (T3, S6/R-S-2): sello de proceso desde el R10 Q2 (04/08/2026 10:15 AM).
            var hojaDetRetri = NombresHojaPeriodo.DetRetri(Insumos.PeriodoQ2());
            var hojaDetVali = NombresHojaPeriodo.DetValiRetri(Insumos.PeriodoQ2());
            Assert.Equal("04/08/2026", LeerCeldaTexto(salida, hojaDetRetri, "D6"));
            Assert.Equal("10:15 AM", LeerCeldaTexto(salida, hojaDetRetri, "D7"));
            Assert.Equal("04/08/2026", LeerCeldaTexto(salida, hojaDetVali, "D6"));
            Assert.Equal("10:15 AM", LeerCeldaTexto(salida, hojaDetVali, "D7"));

            // R-S-3 (no-duplicación): CONSOLIDADO D6 intacto respecto de la plantilla (46238).
            Assert.Equal(
                TestHelpers.LeerCeldaNumerica(Insumos.PlantillaQ2, HojaConsolidado, "D6"),
                TestHelpers.LeerCeldaNumerica(salida, HojaConsolidado, "D6"));
        }
        finally
        {
            BorrarArbol(salidaDir);
        }
    }

    // ── S2/R-R-4: julio Q1 intacto ───────────────────────────────────────────────────────────
    [Fact]
    public void SalidaJulioQ1_SinCalcChainYFullCalcOnLoad_FechasJulioIntactas()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-plan28-julioq1-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);
        try
        {
            CrearProcesador().Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.Periodo(),
                CarpetaPeriodo = Insumos.CarpetaPeriodo,
                RutaPlantilla = Insumos.Plantilla,
                RutaSalida = salida
            });

            var estructura = InspeccionarEstructura(salida);
            Assert.False(estructura.ContieneCalcChain);
            Assert.True(estructura.FullCalcOnLoad);
            Assert.Equal(new DateTime(2026, 7, 1), LeerFechaSalida(salida, "G7"));
            Assert.Equal(new DateTime(2026, 7, 15), LeerFechaSalida(salida, "K7"));
        }
        finally
        {
            BorrarArbol(salidaDir);
        }
    }

    // ── S5/D-E: R10 sin fechas legibles → fail-fast ──────────────────────────────────────────
    [Fact]
    public void R10SinFechasLegibles_LanzaFailFastNombrandoPeriodoArchivoYCelda_SinSalida()
    {
        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-plan28-r10sinfecha-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodoQ2, carpetaPeriodo);
        var rutaR10 = Path.Combine(carpetaPeriodo, "R10_Remuneracion_2026072.xlsx");
        Assert.True(File.Exists(rutaR10));
        var hojaDetRetriQ2 = WorkbookLeafCellMapDetRetri.HojaDetRetri(Insumos.PeriodoQ2());
        VaciarCelda(rutaR10, hojaDetRetriQ2, "G7");
        VaciarCelda(rutaR10, hojaDetRetriQ2, "J7");

        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-plan28-r10sinfecha-salida-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.PeriodoQ2().NombreArchivo);
        try
        {
            var ex = Assert.Throws<CalculoInvalidoException>(() =>
                CrearProcesador().Ejecutar(new SolicitudProcesoPeriodo
                {
                    Periodo = Insumos.PeriodoQ2(),
                    CarpetaPeriodo = carpetaPeriodo,
                    RutaPlantilla = Insumos.PlantillaQ2,
                    RutaSalida = salida
                }));

            Assert.Contains("2026072", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("R10_Remuneracion_2026072.xlsx", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("G7", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("J7", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(salida), "Sin fechas del R10 no debe existir salida certificada.");
        }
        finally
        {
            BorrarArbol(carpetaPeriodo);
            BorrarArbol(salidaDir);
        }
    }

    // ── S6/D-D: G7 fórmula en la plantilla → ERR-PLANTILLA nombrando hoja+celda ──────────────
    [Fact]
    public void PlantillaConG7Formula_ElSelloLanzaErrPlantillaNombrandoCelda_SinEscribir()
    {
        var copia = Path.Combine(Path.GetTempPath(), "remuneracion-plan28-plantillaf-" + Guid.NewGuid().ToString("N"), "plantilla.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(copia)!);
        File.Copy(Insumos.PlantillaQ2, copia);
        ConvertirEnFormula(copia, HojaConsolidado, "G7", "=1+1");

        var resultado = new ResultadoRemuneracion
        {
            Periodo = Insumos.PeriodoQ2(),
            FechaDesde = new DateTime(2026, 8, 16),
            FechaHasta = new DateTime(2026, 8, 31)
        };

        try
        {
            using var documento = SpreadsheetDocument.Open(copia, true);
            var workbookPart = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var ex = Assert.Throws<CalculoInvalidoException>(() =>
                OpenXmlPlantillaWriter.EscribirFechasPeriodo(workbookPart, resultado));

            Assert.Equal(CodigoError.Plantilla, ex.Codigo);
            Assert.Contains(HojaConsolidado, ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("G7", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            BorrarArbol(Path.GetDirectoryName(copia)!);
        }
    }

    // ── S1/R-S-1: salida single-ASE (overload de un solo ASE) ───────────────────────────────
    [Fact]
    public void SalidaSingleAse_SinCalcChainYFullCalcOnLoad_NoSellaFechas()
    {
        var ase = AseFactory.DesdeId(1);
        var periodo = Insumos.Periodo();
        var lector = new ExcelDataReaderRecaudoReader();
        var r1 = lector.LeerR1(Insumos.R1(1));
        var r2 = lector.LeerR2(Insumos.R2(1));
        var r4 = lector.LeerR4(Insumos.R4(1));
        var resultado = new CalculoRemuneracion().CalcularConsolidado(periodo, [(ase, r1, r2, r4)]);
        var leaf = new ExcelDataReaderWorkbookLeafInputReader().LeerLeafInputs(ase, periodo, Insumos.R1(1), Insumos.R2(1), Insumos.R4(1));

        // El single-ASE no lee R10 → no sella: G7/K7 deben quedar idénticas a la plantilla.
        var g7Plantilla = TestHelpers.LeerCeldaNumerica(Insumos.Plantilla, HojaConsolidado, "G7");
        var k7Plantilla = TestHelpers.LeerCeldaNumerica(Insumos.Plantilla, HojaConsolidado, "K7");

        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-plan28-single-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, "salida.xlsx");
        try
        {
            new OpenXmlPlantillaWriter().GenerarWorkbook(Insumos.Plantilla, salida, resultado, leaf);

            var estructura = InspeccionarEstructura(salida);
            Assert.False(estructura.ContieneCalcChain, "La salida single-ASE no debe contener xl/calcChain.xml.");
            Assert.True(estructura.FullCalcOnLoad, "xl/workbook.xml debe traer fullCalcOnLoad=\"1\".");
            Assert.True(estructura.CalculationChainPartAusente);
            Assert.Equal(g7Plantilla, TestHelpers.LeerCeldaNumerica(salida, HojaConsolidado, "G7"));
            Assert.Equal(k7Plantilla, TestHelpers.LeerCeldaNumerica(salida, HojaConsolidado, "K7"));
        }
        finally
        {
            BorrarArbol(salidaDir);
        }
    }

    // ── Plan 31 (T3): sello de proceso D6/D7 (origen congelado por mini-T0 R-S-1) ────────────

    /// <summary>
    /// S6/R-S-2: <c>EscribirFechasPeriodo</c> sella el proceso (D6 fecha + D7 hora) en las 4 celdas
    /// de <c>DetRetri*/DetValiRetri*</c> como TEXTO (representación del manual/base/R10; numFmt 49),
    /// y NO duplica <c>CONSOLIDADO!D6</c> (R-S-3).
    /// </summary>
    [Fact]
    public void SelloProceso_DetRetriYDetValiRetri_D6D7_TextoSinTocarConsolidadoD6()
    {
        var copia = Path.Combine(Path.GetTempPath(), "remuneracion-plan31-sello-" + Guid.NewGuid().ToString("N"), "plantilla.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(copia)!);
        File.Copy(Insumos.PlantillaAgosto2026082, copia);

        var resultado = new ResultadoRemuneracion
        {
            Periodo = Insumos.PeriodoAgosto(),
            FechaDesde = new DateTime(2026, 8, 16),
            FechaHasta = new DateTime(2026, 8, 31),
            FechaProceso = new DateTime(2026, 9, 2),
            HoraProceso = new TimeSpan(7, 42, 0)
        };

        try
        {
            var consolidadoD6Antes = TestHelpers.LeerCeldaNumerica(copia, HojaConsolidado, "D6");
            using (var documento = SpreadsheetDocument.Open(copia, true))
            {
                var workbookPart = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
                OpenXmlPlantillaWriter.EscribirFechasPeriodo(workbookPart, resultado);
            }

            var hojaDetRetri = NombresHojaPeriodo.DetRetri(Insumos.PeriodoAgosto());
            var hojaDetVali = NombresHojaPeriodo.DetValiRetri(Insumos.PeriodoAgosto());
            Assert.Equal("02/09/2026", LeerCeldaTexto(copia, hojaDetRetri, "D6"));
            Assert.Equal("07:42 AM", LeerCeldaTexto(copia, hojaDetRetri, "D7"));
            Assert.Equal("02/09/2026", LeerCeldaTexto(copia, hojaDetVali, "D6"));
            Assert.Equal("07:42 AM", LeerCeldaTexto(copia, hojaDetVali, "D7"));

            // R-S-3: no-duplicación de CONSOLIDADO D6 y guard anti-fórmula intacto (D6 sigue texto).
            Assert.Equal(consolidadoD6Antes, TestHelpers.LeerCeldaNumerica(copia, HojaConsolidado, "D6"));
            Assert.False(TestHelpers.CeldaEsFormula(copia, hojaDetRetri, "D6"));
        }
        finally
        {
            BorrarArbol(Path.GetDirectoryName(copia)!);
        }
    }

    /// <summary>
    /// R-S-2: si la celda destino del sello es fórmula, se lanza <c>ERR-PLANTILLA</c> nombrando
    /// hoja+celda (nunca se sobrescribe una fórmula).
    /// </summary>
    [Fact]
    public void SelloProceso_CeldaD6Formula_LanzaErrPlantillaSinSobrescribir()
    {
        var copia = Path.Combine(Path.GetTempPath(), "remuneracion-plan31-sellof-" + Guid.NewGuid().ToString("N"), "plantilla.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(copia)!);
        File.Copy(Insumos.PlantillaAgosto2026082, copia);
        var hojaDetRetri = NombresHojaPeriodo.DetRetri(Insumos.PeriodoAgosto());
        ConvertirEnFormula(copia, hojaDetRetri, "D6", "=1+1");

        var resultado = new ResultadoRemuneracion
        {
            Periodo = Insumos.PeriodoAgosto(),
            FechaProceso = new DateTime(2026, 9, 2),
            HoraProceso = new TimeSpan(7, 42, 0)
        };

        try
        {
            using var documento = SpreadsheetDocument.Open(copia, true);
            var workbookPart = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var ex = Assert.Throws<CalculoInvalidoException>(() =>
                OpenXmlPlantillaWriter.EscribirFechasPeriodo(workbookPart, resultado));

            Assert.Equal(CodigoError.Plantilla, ex.Codigo);
            Assert.Contains(hojaDetRetri, ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("D6", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            BorrarArbol(Path.GetDirectoryName(copia)!);
        }
    }

    /// <summary>
    /// S6/D-F: R10 sin "Fecha de Proceso" (D6) ni "Hora" (D7) legibles → <c>ERR-FORMATO-FUENTE</c>
    /// nombrando período + archivo + celda, sin salida parcial.
    /// </summary>
    [Fact]
    public void R10SinFechaProceso_LanzaFailFastNombrandoD6D7_SinSalida()
    {
        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-plan31-r10sinproceso-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodoQ2, carpetaPeriodo);
        var rutaR10 = Path.Combine(carpetaPeriodo, "R10_Remuneracion_2026072.xlsx");
        Assert.True(File.Exists(rutaR10));
        var hojaDetRetriQ2 = WorkbookLeafCellMapDetRetri.HojaDetRetri(Insumos.PeriodoQ2());
        VaciarCelda(rutaR10, hojaDetRetriQ2, "D6");
        VaciarCelda(rutaR10, hojaDetRetriQ2, "D7");

        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-plan31-r10sinproceso-salida-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.PeriodoQ2().NombreArchivo);
        try
        {
            var ex = Assert.Throws<CalculoInvalidoException>(() =>
                CrearProcesador().Ejecutar(new SolicitudProcesoPeriodo
                {
                    Periodo = Insumos.PeriodoQ2(),
                    CarpetaPeriodo = carpetaPeriodo,
                    RutaPlantilla = Insumos.PlantillaQ2,
                    RutaSalida = salida
                }));

            Assert.Equal(CodigoError.FormatoFuente, ex.Codigo);
            Assert.Contains("2026072", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("D6", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("D7", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(salida), "Sin sello de proceso legible no debe existir salida certificada.");
        }
        finally
        {
            BorrarArbol(carpetaPeriodo);
            BorrarArbol(salidaDir);
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────────────────

    private static ProcesadorPeriodo CrearProcesador() =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader());

    /// <summary>Estructura del ZIP/XML de la salida (verificación sin Excel).</summary>
    private sealed record EstructuraSalida(
        bool ContieneCalcChain,
        bool FullCalcOnLoad,
        bool CalculationChainPartAusente);

    private static EstructuraSalida InspeccionarEstructura(string ruta)
    {
        bool contiene;
        bool full;
        using (var zip = ZipFile.OpenRead(ruta))
        {
            contiene = zip.Entries.Any(e => string.Equals(e.FullName, EntradaCalcChain, StringComparison.OrdinalIgnoreCase));
            var wb = zip.GetEntry("xl/workbook.xml")
                ?? throw new InvalidOperationException("La salida no tiene xl/workbook.xml.");
            using var lector = new StreamReader(wb.Open());
            var xml = lector.ReadToEnd();
            full = xml.Contains("fullCalcOnLoad=\"1\"", StringComparison.Ordinal);
        }

        bool partAusente;
        using (var documento = SpreadsheetDocument.Open(ruta, false))
        {
            var wp = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            partAusente = wp.CalculationChainPart is null;
        }

        return new EstructuraSalida(contiene, full, partAusente);
    }

    private static DateTime LeerFechaSalida(string ruta, string celda)
    {
        var serial = TestHelpers.LeerCeldaNumerica(ruta, HojaConsolidado, celda);
        Assert.True(serial > 0, $"La celda {HojaConsolidado}!{celda} debe traer un serial de fecha; se leyó {serial}.");
        return DateTime.FromOADate((double)serial).Date;
    }

    /// <summary>
    /// Plan 31 (T3): lee el texto de una celda resolviendo inline string o shared string (el sello de
    /// proceso D6/D7 se escribe como texto, no como serial).
    /// </summary>
    private static string LeerCeldaTexto(string ruta, string hoja, string celda)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var worksheet = Hoja(documento, hoja);
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        if (objCelda.DataType?.Value == CellValues.InlineString)
        {
            return objCelda.InlineString?.InnerText ?? string.Empty;
        }

        if (objCelda.DataType?.Value == CellValues.SharedString)
        {
            var indice = int.Parse(objCelda.CellValue!.InnerText, CultureInfo.InvariantCulture);
            var sst = documento.WorkbookPart!.SharedStringTablePart!.SharedStringTable!;
            return sst.ElementAt(indice).InnerText;
        }

        return objCelda.CellValue?.InnerText ?? string.Empty;
    }

    private static string PrepararPeriodoAgosto()
    {
        var destino = Path.Combine(Path.GetTempPath(), "remuneracion-plan28-agosto-insumos-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(destino);

        var origen = Insumos.CarpetaInsumosAgosto;
        foreach (var archivo in Directory.EnumerateFiles(origen, "*.xlsx", SearchOption.AllDirectories))
        {
            var relativo = Path.GetRelativePath(origen, archivo);
            var destinoArchivo = Path.Combine(destino, relativo);
            Directory.CreateDirectory(Path.GetDirectoryName(destinoArchivo)!);
            File.Copy(archivo, destinoArchivo);
        }

        var conciliacionesDestino = Path.Combine(destino, "Conciliaciones");
        Directory.CreateDirectory(conciliacionesDestino);
        foreach (var archivo in Directory.EnumerateFiles(Insumos.CarpetaConciliacionesQ2, "*.xlsx", SearchOption.TopDirectoryOnly))
        {
            File.Copy(archivo, Path.Combine(conciliacionesDestino, Path.GetFileName(archivo)));
        }

        return destino;
    }

    private static void CopiarArbol(string origen, string destino)
    {
        Directory.CreateDirectory(destino);
        foreach (var archivo in Directory.EnumerateFiles(origen, "*.xlsx", SearchOption.AllDirectories))
        {
            var relativo = Path.GetRelativePath(origen, archivo);
            var destinoArchivo = Path.Combine(destino, relativo);
            Directory.CreateDirectory(Path.GetDirectoryName(destinoArchivo)!);
            File.Copy(archivo, destinoArchivo);
        }
    }

    private static void VaciarCelda(string ruta, string hoja, string celda)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var worksheet = Hoja(documento, hoja);
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        objCelda.CellFormula?.Remove();
        objCelda.DataType = CellValues.String;
        objCelda.CellValue = new CellValue(string.Empty);
        worksheet.Save();
    }

    private static void ConvertirEnFormula(string ruta, string hoja, string celda, string formula)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var worksheet = Hoja(documento, hoja);
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        objCelda.CellValue?.Remove();
        objCelda.DataType = null;
        objCelda.CellFormula = new CellFormula(formula);
        worksheet.Save();
    }

    private static Worksheet Hoja(SpreadsheetDocument documento, string nombre)
    {
        var workbookPart = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet.");
    }

    private static void BorrarArbol(string ruta)
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
