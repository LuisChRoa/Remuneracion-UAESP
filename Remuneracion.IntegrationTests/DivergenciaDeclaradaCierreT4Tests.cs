using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
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
/// Plan 42 (T4, R-G-2/R-CLS-2, D-F): ambos sentidos end-to-end desde la PLANTILLA OFICIAL, Q2
/// (2026072, 5 ASE), a rutas temp FRESCAS.
///   (i)  declarada reconocida → el run COMPLETA (archivo + reporte + doble huella);
///   (ii) no declarada inyectada → sigue BLOQUEANDO (fail duro) aunque la declarada esté reconocida.
/// Colección no paralela: la huella de log usa el Log.Logger global.
/// </summary>
[Collection("Cli")]
public sealed class DivergenciaDeclaradaCierreT4Tests : IDisposable
{
    public void Dispose()
    {
        Log.CloseAndFlush();
        Log.Logger = Logger.None;
    }

    [Fact]
    public void T4_DeclaradaReconocida_ElRunCompletaConArchivoReporteYLog()
    {
        var sink = new CapturaEventos();
        Log.Logger = new LoggerConfiguration().MinimumLevel.Debug().WriteTo.Sink(sink).CreateLogger();

        var dir = DirectorioTemp("p42-t4-declarada");
        try
        {
            var salida = Path.Combine(dir, Insumos.PeriodoQ2().NombreArchivo);
            // Plan 46 (H2): la divergencia real quedó resuelta; el mecanismo se ejercita con el
            // FIXTURE SINTÉTICO (SINTETICA-C15, Δ = 1) — el emisor del reporte es el real.
            var procesador = FixtureDivergenciaSintetica.CrearProcesador();

            var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
                RutaPlantilla = Insumos.PlantillaHistoricaV2(),
                RutaSalida = salida,
                DivergenciasDeclaradasAceptadas = [FixtureDivergenciaSintetica.EntradaId]
            });

            // El run COMPLETA (no lanza): archivo producido.
            Assert.True(File.Exists(salida), "El run con reconocimiento debe producir el archivo.");

            // La clasificación viaja en el resultado (aditivo) para que UI/CLI la presenten.
            Assert.NotNull(resultado.DivergenciasDeclaradas);
            Assert.True(resultado.DivergenciasDeclaradas!.SoloDeclaradas);
            Assert.Equal([FixtureDivergenciaSintetica.EntradaId], resultado.DivergenciasDeclaradas.EntradasDeclaradas);

            // Reporte durable junto a la salida (sin reporte no hay ÉXITO).
            Assert.False(string.IsNullOrWhiteSpace(resultado.ReporteDivergenciasDeclaradas));
            Assert.True(File.Exists(resultado.ReporteDivergenciasDeclaradas));
            Assert.Contains("REPORTE DE DIVERGENCIAS DECLARADAS", File.ReadAllText(resultado.ReporteDivergenciasDeclaradas), StringComparison.Ordinal);

            // Huella en el log (segunda huella obligatoria): entrada, Δ, corrida, aceptación.
            var mensajes = sink.Eventos.Select(e => e.RenderMessage()).ToList();
            Assert.Contains(mensajes, m => m.Contains(FixtureDivergenciaSintetica.EntradaId, StringComparison.Ordinal));
            Assert.Contains(mensajes, m => m.Contains("Divergencia declarada reconocida", StringComparison.Ordinal) && m.Contains("Δ 1.", StringComparison.Ordinal));
        }
        finally
        {
            Limpiar(dir);
        }
    }

    [Fact]
    public void T4_NoDeclaradaInyectada_SigueBloqueandoAunqueLaDeclaradaEsteReconocida()
    {
        var dir = DirectorioTemp("p42-t4-mixta");
        try
        {
            var salida = Path.Combine(dir, Insumos.PeriodoQ2().NombreArchivo);
            // Inyecta una divergencia NO declarada (C15 recaudo) tras la escritura del writer real.
            var procesador = CrearProcesador(escritor: new EscritorConDivergenciaInyectada(), reporte: new ReporteDivergenciasDeclaradasWriter());

            var ex = Assert.Throws<CalculoInvalidoException>(() => procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
                RutaPlantilla = Insumos.PlantillaHistoricaV2(),
                RutaSalida = salida,
                DivergenciasDeclaradasAceptadas = ["OCCIDENTE-ASE5"] // la declarada SÍ está reconocida
            }));

            // D-F: la no declarada manda; FAIL duro nombrando la celda inyectada.
            Assert.Contains("[ERR-VALIDACION]", ex.Message, StringComparison.Ordinal);
            Assert.Contains("VALIDACION_TOTAL · C15", ex.Message, StringComparison.Ordinal);
            // La clasificación NO viaja (es un fallo no declarado, no un cierre pendiente).
            Assert.Null(ex.Clasificacion);
        }
        finally
        {
            Limpiar(dir);
        }
    }

    private static ProcesadorPeriodo CrearProcesador(
        IWorkbookLeafWriter? escritor = null,
        IReporteDivergenciasDeclaradas? reporte = null) =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            escritor ?? new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader(),
            new ValidacionOracleReader(),
            reporte);

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

    /// <summary>
    /// Decorador de test: delega en el writer real y, tras escribir, corrompe el LITERAL
    /// <c>'Valida - Control Recaudo'!D7</c> (control de recaudo que alimenta <c>E7 → F10 → C14</c>)
    /// para provocar un FALSE REAL en <c>C15</c> = (<c>C14 = C9</c>). Plan 48: en la v2 <c>F10</c> es
    /// FORMULA y el evaluador ignora su cache; el literal inyectable es <c>D7</c>.
    /// </summary>
    private sealed class EscritorConDivergenciaInyectada : IWorkbookLeafWriter
    {
        private readonly OpenXmlPlantillaWriter _inner = new();

        public void GenerarWorkbook(string rutaPlantillaOrigen, string rutaSalida, ResultadoRemuneracion resultado, WorkbookLeafInputs leafInputs)
        {
            _inner.GenerarWorkbook(rutaPlantillaOrigen, rutaSalida, resultado, leafInputs);
            Inyectar(rutaSalida);
        }

        public void GenerarWorkbook(string rutaPlantillaOrigen, string rutaSalida, ResultadoRemuneracion resultado, IReadOnlyList<WorkbookLeafInputs> leafInputs)
        {
            _inner.GenerarWorkbook(rutaPlantillaOrigen, rutaSalida, resultado, leafInputs);
            Inyectar(rutaSalida);
        }

        private static void Inyectar(string ruta)
        {
            using var workbook = SpreadsheetDocument.Open(ruta, true);
            var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var sheet = workbookPart.Workbook!.Descendants<Sheet>()
                .First(s => string.Equals(s.Name?.Value, "Valida - Control Recaudo", StringComparison.OrdinalIgnoreCase));
            var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
                ?? throw new InvalidOperationException("Worksheet null");
            var celda = worksheet.Descendants<Cell>()
                .First(c => string.Equals(c.CellReference?.Value, "D7", StringComparison.OrdinalIgnoreCase));
            celda.CellValue = new CellValue("1");
            celda.DataType = CellValues.Number;
            worksheet.Save();
        }
    }

    private sealed class CapturaEventos : ILogEventSink
    {
        public List<LogEvent> Eventos { get; } = [];

        public void Emit(LogEvent logEvent) => Eventos.Add(logEvent);
    }
}
