using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Cli;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-17 (§2.5, G4): tests dedicados de los micro-fixes de saneamiento — cada ítem de la tabla
/// de disposición sale con test verde o nota documentada. Regla de hierro del plan: NINGUNO
/// cambia valores, gates, mapas, goldens ni tolerancia ±0.5; lo que rompa la red = NEEDS_CONTEXT
/// con rollback del micro-fix (nunca maquillaje).
/// </summary>
[Collection("MicroFixes")]
public sealed class MicroFixesHu17Tests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    // ── S-1 HU-16: el barrido de L numéricas deriva el límite de la DIMENSIÓN (no >700) ──────

    [Fact]
    public void S1_Hu16_EnumerarCeldasL_DerivaLimiteDeDimension_FilaFantasma()
    {
        var ruta = CrearWorkbookConDimension("A1:N750", filaDentro: 730, filaFuera: 780);
        try
        {
            var celdas = TestHelpers.EnumerarCeldaLNumericas(ruta, "Hoja1").ToList();

            // L730 (fila 730 > 700 pero DENTRO de la dimensión A1:N750) DEBE leerse: el corte
            // mágico "row>700" habría saltado una fila legítima del barrido A8 stale-guard.
            Assert.Contains(celdas, c => c.Celda == "L730" && c.Valor == 42m);

            // L780 (fila fantasma FUERA de la dimensión) se ignora.
            Assert.DoesNotContain(celdas, c => c.Celda == "L780");
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    // ── S-2 HU-16: helpers OpenXML compartidos (TestHelpers) sin duplicado ──────────────────

    [Fact]
    public void S2_Hu16_TestHelpers_Compartidos_InterventoriaYGoldenLosUsan()
    {
        // Evidencia de refactor: los helpers viven una sola vez y funcionan contra el golden real.
        Assert.True(TestHelpers.CeldaEsFormula(Insumos.Plantilla, "INTERVENTORIA", "K31"), "K31 debió ser fórmula SUM (golden Q1).");
        Assert.Equal(378371975m, TestHelpers.LeerCeldaNumerica(Insumos.Plantilla, "INTERVENTORIA", "K26"));
        Assert.Contains(TestHelpers.EnumerarCeldaLNumericas(Insumos.Plantilla, "Reporte Componentes R1"), c => c.Celda == "L25");
    }

    // ── S-3 HU-16: texto en celda de gate ≠ 0 — fail-fast que nombra la celda ────────────────

    [Fact]
    public void S3_Hu16_Writer_TextoEnLCeldaId_FailFastNombraCelda()
    {
        var dir = DirectorioTemp("s3-texto-L26");
        var plantilla = Path.Combine(dir, "plantilla-texto-L26.xlsx");
        File.Copy(Insumos.Plantilla, plantilla, overwrite: true);
        PonerTextoEnCelda(plantilla, "INTERVENTORIA", "L26", "ASE1");

        var ex = Assert.Throws<CalculoInvalidoException>(() => EjecutarProcesadorQ1(plantilla, Path.Combine(dir, Insumos.Periodo().NombreArchivo)));
        Assert.Contains("INTERVENTORIA", ex.Message);
        Assert.Contains("L26", ex.Message);
        Assert.Contains("NO numérico", ex.Message);
    }

    // ── S-4 HU-16: lectura INTERVENTORIA por ASE = Debug; hitos siguen Information ───────────

    [Fact]
    public void S4_Hu16_InterventoriaLecturaPorAse_Debug_HitoInformation()
    {
        var sink = new CapturaEventos();
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .Enrich.FromLogContext()
            .WriteTo.Sink(sink)
            .CreateLogger();
        try
        {
            var salidaDir = DirectorioTemp("s4-interventoria");
            var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);
            EjecutarProcesadorQ1(Insumos.Plantilla, salida);

            // Lectura K/M/N por ASE → Debug (antes Information; era el "triple log D2b").
            var lectura = sink.Eventos.First(e => e.MessageTemplate.Text.Contains("ASE {AseId}: INTERVENTORIA (insumo externo declarado) K=", StringComparison.Ordinal));
            Assert.Equal(LogEventLevel.Debug, lectura.Level);

            // L-Especiales por ASE → Debug.
            var lMenores = sink.Eventos.First(e => e.MessageTemplate.Text.Contains("L-Especiales menores = {Count}", StringComparison.Ordinal));
            Assert.Equal(LogEventLevel.Debug, lMenores.Level);

            // Hito de declaración (read-loop) → Information (doctrina HU-14 D4: hitos a Information).
            var hito = sink.Eventos.First(e => e.MessageTemplate.Text.Contains("INTERVENTORIA = insumo externo declarado — hoja intacta (bloque anual", StringComparison.Ordinal));
            Assert.Equal(LogEventLevel.Information, hito.Level);
        }
        finally
        {
            Log.CloseAndFlush();
            Log.Logger = Logger.None;
        }
    }

    // ── W-2.2 HU-14: cultura es-CO — la línea mostrada y el evento estructurado coinciden ─────

    [Fact]
    public void W22_Hu14_Validaciones_CulturaEsCo_LineaYEventoCoherentes()
    {
        var culturaPrevia = CultureInfo.CurrentCulture;
        var uiPrevia = CultureInfo.CurrentUICulture;
        var sink = new CapturaEventos();
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .Enrich.FromLogContext()
            .WriteTo.Sink(sink)
            .CreateLogger();
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("es-CO");
            CultureInfo.CurrentUICulture = new CultureInfo("es-CO");

            var salidaDir = DirectorioTemp("w22-esco");
            var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);
            var resultado = EjecutarProcesadorPeriodoConOracle(Insumos.Plantilla, salida);

            var esCo = CultureInfo.GetCultureInfo("es-CO");
            foreach (var linea in resultado.Validaciones)
            {
                var match = Regex.Match(linea, @"^  (.+?): O \(Recaudo vs REMUNERACION\) = ([-0-9.,]+); P \(INT\(O\)=0\) = (TRUE|FALSE)$");
                if (!match.Success)
                {
                    continue;
                }

                // La O mostrada se formatea con la cultura de corrida (es-CO): "0,016" con coma.
                var oLinea = decimal.Parse(match.Groups[2].Value, esCo);
                var evento = sink.Eventos.First(e => e.MessageTemplate.Text.Contains("Empresa {Empresa}: O (Recaudo vs REMUNERACION)", StringComparison.Ordinal)
                                                     && e.Properties.TryGetValue("Empresa", out var em)
                                                     && em is ScalarValue { Value: string s } && s == match.Groups[1].Value);
                var oEvento = (decimal)((ScalarValue)evento.Properties["O"]).Value!;
                Assert.Equal(oLinea, oEvento);
            }
        }
        finally
        {
            CultureInfo.CurrentCulture = culturaPrevia;
            CultureInfo.CurrentUICulture = uiPrevia;
            Log.CloseAndFlush();
            Log.Logger = Logger.None;
        }
    }

    // ── S-4 HU-14: selección ASE inválida = ArgumentException, nunca tipo archivo-inexistente ─

    [Fact]
    public void S4_Hu14_SeleccionAseInvalida_TipoArgumentException()
    {
        // Form1.ParseAse lanza ArgumentException para selección inválida (antes
        // ArchivoFuenteNoEncontradoException). El contrato de dominio de la factoría usa el mismo
        // tipo base para id fuera de 1..5 (ArgumentOutOfRangeException ⊆ ArgumentException).
        Assert.ThrowsAny<ArgumentException>(() => AseFactory.DesdeId(6));
        Assert.ThrowsAny<ArgumentException>(() => AseFactory.DesdeId(0));
        Assert.ThrowsAny<ArgumentException>(() => AseFactory.DesdeId(-1));
    }

    // ── S-1 HU-15: Main con red externa — arg inválido extremo nunca sale del contrato ────────

    [Fact]
    public void S1_Hu15_MainCatchExterno_ArgInvalidoExtremo_Exit4_SinCrash()
    {
        var exe = RutaExeCli();
        Assert.True(File.Exists(exe), $"Falta el ejecutable del CLI: {exe}");

        var trabajo = DirectorioTemp("s1-subproceso");
        var psi = new ProcessStartInfo(exe, "--periodo 2026 --carpeta c --plantilla p --salida s")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            WorkingDirectory = trabajo
        };

        using var proceso = Process.Start(psi)!;
        var stdout = proceso.StandardOutput.ReadToEnd();
        var stderr = proceso.StandardError.ReadToEnd();
        Assert.True(proceso.WaitForExit(60_000), "El CLI debe terminar (red externa S-1).");
        Assert.Equal(4, proceso.ExitCode); // uso inválido → 4 (nunca un código no-contract)
        Assert.Contains("Uso:", stderr);
        Assert.DoesNotContain("RESULTADO", stdout); // errores de uso no imprimen RESULTADO
    }

    // ── S-2 HU-15: reconfiguración Serilog sin fuga de sinks ─────────────────────────────────

    [Fact]
    public void S2_Hu15_ConfigurarSerilog_DobleConfiguracion_NoLanzaYSigueEscribiendo()
    {
        var rutaLogs = AppContext.BaseDirectory;
        try
        {
            // Doble configuración: antes, el logger previo quedaba sin dispose (sink abierto).
            Program.ConfigurarSerilog();
            Program.ConfigurarSerilog();
            Log.Information("evento-tras-doble-configuracion-HU17");
            Log.CloseAndFlush();

            var archivo = Directory.EnumerateFiles(rutaLogs, "remuneracion_log_*.txt")
                .OrderByDescending(f => f)
                .FirstOrDefault();
            Assert.NotNull(archivo);
            Assert.Contains("evento-tras-doble-configuracion-HU17", File.ReadAllText(archivo!));
        }
        finally
        {
            Log.CloseAndFlush();
            Log.Logger = Logger.None;
        }
    }

    // ── L1 HU-12: bloque sin datos (Q1: AjustesSfT null) NO se limpia ────────────────────────

    [Fact]
    public void L1_Hu12_BloqueVacio_NoSeLimpia_OutputIgualTemplate()
    {
        var salidaDir = DirectorioTemp("l1-bloque-vacio");
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);
        EjecutarProcesadorQ1(Insumos.Plantilla, salida);

        // En Q1 AjustesSfT es null → el writer NO escribe la hoja AJUSTES-SF-T: las celdas
        // operando de la salida quedan IGUALES a la plantilla (nunca se "limpian" a 0).
        foreach (var celda in new[] { "D47", "D48", "D49", "D50", "D51" })
        {
            Assert.Equal(
                TestHelpers.LeerCeldaNumerica(Insumos.Plantilla, "AJUSTES - SF-T", celda),
                TestHelpers.LeerCeldaNumerica(salida, "AJUSTES - SF-T", celda));
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────────────────────

    private static ResultadoProcesoPeriodo EjecutarProcesadorQ1(string plantilla, string salida) =>
        EjecutarProcesadorPeriodoConOracle(plantilla, salida);

    private static ResultadoProcesoPeriodo EjecutarProcesadorPeriodoConOracle(string plantilla, string salida)
    {
        var procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ValidacionOracleReader());
        return procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.Periodo(),
            CarpetaPeriodo = Insumos.CarpetaPeriodo,
            RutaPlantilla = plantilla,
            RutaSalida = salida
        });
    }

    /// <summary>Crea un workbook con dimensión fija y dos filas L: una dentro (filaDentro) y otra fuera (filaFuera).</summary>
    private static string CrearWorkbookConDimension(string dimension, uint filaDentro, uint filaFuera)
    {
        var ruta = Path.Combine(Path.GetTempPath(), "remuneracion-dim-" + Guid.NewGuid().ToString("N") + ".xlsx");
        using (var doc = SpreadsheetDocument.Create(ruta, SpreadsheetDocumentType.Workbook))
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var sheetPart = wbPart.AddNewPart<WorksheetPart>();
            sheetPart.Worksheet = new Worksheet(
                new SheetDimension { Reference = dimension },
                new SheetData(
                    new Row(
                        new Cell { CellReference = "L" + filaDentro, CellValue = new CellValue("42") })
                    { RowIndex = filaDentro },
                    new Row(
                        new Cell { CellReference = "L" + filaFuera, CellValue = new CellValue("9") })
                    { RowIndex = filaFuera }));
            wbPart.Workbook.Sheets = new Sheets(
                new Sheet { Id = wbPart.GetIdOfPart(sheetPart), SheetId = 1, Name = "Hoja1" });
        }

        return ruta;
    }

    /// <summary>Pone texto (no numérico) en el CellValue de una celda existente (para el gate S-3).</summary>
    private static void PonerTextoEnCelda(string ruta, string hoja, string celda, string texto)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet!;
        var cell = ws.Descendants<Cell>().First(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        cell.CellValue = new CellValue(texto);
        ws.Save();
    }

    private static string RutaExeCli()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")))
            {
                return Path.Combine(dir.FullName, "Remuneracion.Cli", "bin", "Debug", "net10.0", "Remuneracion.Cli.exe");
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio.");
    }

    private static string DirectorioTemp(string etiqueta)
    {
        var dir = Path.Combine(Path.GetTempPath(), $"remuneracion-{etiqueta}-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private sealed class CapturaEventos : ILogEventSink
    {
        public List<LogEvent> Eventos { get; } = [];

        public void Emit(LogEvent logEvent) => Eventos.Add(logEvent);
    }
}

[CollectionDefinition("MicroFixes", DisableParallelization = true)]
public sealed class MicroFixesCollection
{
}
