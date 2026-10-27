using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Cli;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-15 (Req 7, §2.7): PARIDAD CLI↔vía directa — mismos valores, NO nuevo oráculo/golden.
/// Ante mismos insumos, <c>EjecutorCli.Ejecutar</c> (in-process, sin Exit) y el procesador
/// directo (como lo usa Form1) producen celdas escritas iguales en ambos workbooks y snapshots
/// del oráculo iguales (comparados por valores vía reader, nunca por bytes: los metadatos
/// OpenXML llevan timestamps). Casos: 5-ASE Q1 (camino principal), 1-ASE Q1 (ASE 3) y
/// 5-ASE Q2 (cubre AJUSTES-SF-T + DetRetri en Q2 sin golden nuevo).
/// Colección no paralela: el CLI reconfigura el <c>Log.Logger</c> global (D8).
/// </summary>
[Collection("Cli")]
public sealed class ParidadCliTests : IDisposable
{
    private readonly StringWriter _stdout = new();
    private readonly StringWriter _stderr = new();

    public void Dispose()
    {
        Log.CloseAndFlush();
        Log.Logger = Logger.None;
        _stdout.Dispose();
        _stderr.Dispose();
    }

    // ── 5-ASE Q1: camino principal ─────────────────────────────────────────────────────────────

    /// <summary>
    /// Plan 59 (§15, 2026-10-27): bases VIGENTES (v4) de las pruebas de PARIDAD 5-ASE. El anclaje a la
    /// v2 que fijó el H5 se apoyaba en el typo de plantilla <c>R1!F328</c> (−M213 donde debe ser −L213),
    /// que impedía cerrar julio-1 y julio-2 (Δ 22.921,61 y Δ 1.101.678,21 respectivamente) y dejaba
    /// estos casos rojos por una causa de PLANTILLA, no de paridad. La v4 corrige la fórmula y ambas
    /// corridas cierran con <c>RESULTADO OK</c> (§14.3; oráculos: R10 de julio-1 y el administrativo de
    /// julio-2). El objeto del test no cambia: CLI y vía directa producen los MISMOS valores. No se
    /// relajó ninguna aserción: <c>CodigosSalida.Ok</c> se exige igual que antes.
    /// </summary>
    private static readonly string PlantillaVigenteQ1 = Insumos.BaseDelPeriodoV4("2026071");

    private static readonly string PlantillaVigenteQ2 = Insumos.BaseDelPeriodoV4("2026072");
    [Fact]
    public void Paridad_5Ase_Q1_CliVsDirecto_ValoresYSnapshotsIguales()
    {
        var dirCli = DirectorioTemp("paridad-cli-q1");
        var dirDirecto = DirectorioTemp("paridad-directo-q1");

        // Vía CLI in-process.
        var codigoCli = EjecutorCli.Ejecutar(
            ["--periodo", "2026071", "--carpeta", Insumos.CarpetaPeriodo, "--plantilla", PlantillaVigenteQ1, "--salida", dirCli, "--cinco-ase", "--sobrescribir"],
            _stdout, _stderr);
        Assert.Equal(CodigosSalida.Ok, codigoCli);

        // Vía directa (procesador, como lo usa Form1/Program.cs).
        var resultadoDirecto = CrearProcesadorPeriodo().Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.Periodo(),
            CarpetaPeriodo = Insumos.CarpetaPeriodo,
            RutaPlantilla = PlantillaVigenteQ1,
            RutaSalida = Path.Combine(dirDirecto, Insumos.Periodo().NombreArchivo)
        });

        var salidaCli = Path.Combine(dirCli, Insumos.Periodo().NombreArchivo);
        var salidaDirecta = resultadoDirecto.RutaSalida;
        Assert.True(File.Exists(salidaCli));
        Assert.True(File.Exists(salidaDirecta));

        // 1) MISMOS VALORES en TODO el libro: la comprobación de paridad se hace celda por celda
        //    sobre las 40 hojas, SIN mapa de coordenadas. (Plan 59 §15, 2026-10-27: la verificación
        //    anterior recorría `WorkbookLeafCellMapPorAse` + `RolsQ2`, tablas LEGACY con la geometría
        //    de la v2 — con la plantilla vigente esas direcciones ya no describen los roles y el
        //    fixture enrojecía por su propia foto, no por la paridad. La comparación libro↔libro es
        //    más fuerte: cubre TODAS las celdas escritas, no un subconjunto congelado.)
        AssertMismoContenidoPorValor(salidaCli, salidaDirecta);

        // 2) Resultado del dominio directo: GranTotal + 5 totales por ASE (el mismo cálculo que
        //    el CLI; la igualdad de celdas escritas arriba ata la vía CLI al mismo Resultado).
        Assert.Equal(5, resultadoDirecto.Resultado.Consolidados.Count);
        var granTotal = resultadoDirecto.Resultado.GranTotal;
        Assert.True(granTotal > 0m);
        Assert.Equal(granTotal, resultadoDirecto.Resultado.Consolidados.Sum(c => c.TotalAse));

        // 3) Snapshots del oráculo de ambas salidas iguales (sin golden nuevo).
        AssertSnapshotsIguales(salidaCli, salidaDirecta, Insumos.Periodo());

        // 4) Req 4: el runId impreso por el CLI correlaciona los eventos del log (inicio→fin).
        // El CLI escribe el log (D8) en el CWD de la corrida (bin de tests) con el patrón
        // diario "remuneracion_log_YYYYMMDD.txt" (mismo prefijo que Form1). El sink de archivo
        // mantiene el archivo abierto → se cierra el logger antes de leerlo.
        var runId = RunIdDe(_stdout.ToString());
        Log.CloseAndFlush();
        Log.Logger = Logger.None;
        var logRuta = Directory.EnumerateFiles(AppContext.BaseDirectory, "remuneracion_log_*.txt", SearchOption.TopDirectoryOnly)
            .OrderByDescending(f => f)
            .FirstOrDefault();
        Assert.True(logRuta is not null, "El CLI debe escribir el log (D8) en el directorio de la corrida.");
        var lineas = File.ReadAllLines(logRuta!);
        var conRunId = lineas.Count(l => l.Contains(runId.ToString(), StringComparison.Ordinal));
        Assert.True(conRunId >= 2, $"El log debe tener al menos el inicio y el fin con el RunId {runId}; encontradas {conRunId}.");
    }

    // ── 1-ASE Q1 (ASE 3): mismas fuentes que Form1.EjecutarModoUnAse ───────────────────────────

    [Fact]
    public void Paridad_1Ase_Q1_Ase3_CliVsDirecto_ValoresIguales()
    {
        var dirCli = DirectorioTemp("paridad-cli-1ase");
        var dirDirecto = DirectorioTemp("paridad-directo-1ase");

        var codigoCli = EjecutorCli.Ejecutar(
            ["--periodo", "2026071", "--carpeta", Insumos.CarpetaPeriodo, "--plantilla", Insumos.Plantilla, "--salida", dirCli, "--ase", "3"],
            _stdout, _stderr);
        Assert.Equal(CodigosSalida.Ok, codigoCli);

        var resultadoDirecto = EjecutarUnAseDirecto(3, Insumos.Periodo(), Path.Combine(dirDirecto, Insumos.Periodo().NombreArchivo));

        var salidaCli = Path.Combine(dirCli, Insumos.Periodo().NombreArchivo);
        var salidaDirecta = resultadoDirecto.RutaSalida;
        Assert.True(File.Exists(salidaCli));
        Assert.True(File.Exists(salidaDirecta));

        // Celdas leaf del mapa single-ASE (F25/F41/L25/F30/F10/L10, E15/E26/K15, D9/P9).
        var leaf = resultadoDirecto.Leaf;
        var celdas = new (string Hoja, string Celda, decimal Valor)[]
        {
            (WorkbookLeafCellMap.HojaR1, "F25", leaf.R1.F25),
            (WorkbookLeafCellMap.HojaR1, "F41", leaf.R1.F41),
            (WorkbookLeafCellMap.HojaR1, "L25", leaf.R1.L25),
            (WorkbookLeafCellMap.HojaR1, "F30", leaf.R1.F30),
            (WorkbookLeafCellMap.HojaR1, "F10", leaf.R1.F10),
            (WorkbookLeafCellMap.HojaR1, "L10", leaf.R1.L10),
            (WorkbookLeafCellMap.HojaR2, "E15", leaf.R2.E15),
            (WorkbookLeafCellMap.HojaR2, "E26", leaf.R2.E26),
            (WorkbookLeafCellMap.HojaR2, "K15", leaf.R2.K15),
            (WorkbookLeafCellMap.HojaR4, "D9", leaf.R4.D9),
            (WorkbookLeafCellMap.HojaR4, "P9", leaf.R4.P9)
        };
        foreach (var (hoja, celda, valor) in celdas)
        {
            Assert.Equal(valor, LeerCelda(salidaCli, hoja, celda));
            Assert.Equal(valor, LeerCelda(salidaDirecta, hoja, celda));
        }

        Assert.Equal(resultadoDirecto.Resultado.GranTotal, resultadoDirecto.Resultado.Consolidados.Single().TotalAse);
        AssertSnapshotsIguales(salidaCli, salidaDirecta, Insumos.Periodo());
    }

    // ── 5-ASE Q2: cubre AJUSTES-SF-T + DetRetri en Q2 (sin golden nuevo) ───────────────────────

    [Fact]
    public void Paridad_5Ase_Q2_CliVsDirecto_IncluyeAjustesYDetRetri()
    {
        var dirCli = DirectorioTemp("paridad-cli-q2");
        var dirDirecto = DirectorioTemp("paridad-directo-q2");

        var codigoCli = EjecutorCli.Ejecutar(
            ["--periodo", "2026072", "--carpeta", Insumos.CarpetaPeriodoQ2, "--plantilla", PlantillaVigenteQ2, "--salida", dirCli, "--cinco-ase", "--sobrescribir"],
            _stdout, _stderr);
        Assert.Equal(CodigosSalida.Ok, codigoCli);

        var resultadoDirecto = CrearProcesadorPeriodo().Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.PeriodoQ2(),
            CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
            RutaPlantilla = PlantillaVigenteQ2,
            RutaSalida = Path.Combine(dirDirecto, Insumos.PeriodoQ2().NombreArchivo)
        });

        var salidaCli = Path.Combine(dirCli, Insumos.PeriodoQ2().NombreArchivo);
        var salidaDirecta = resultadoDirecto.RutaSalida;
        Assert.True(File.Exists(salidaCli));
        Assert.True(File.Exists(salidaDirecta));

        // El path Q2 incluye AJUSTES-SF-T y DetRetri: los leafs del directo los traen.
        Assert.All(resultadoDirecto.Leafs, l => Assert.NotNull(l.AjustesSfT));
        Assert.All(resultadoDirecto.Leafs, l => Assert.NotNull(l.DetRetriQ2));

        // 1) MISMOS VALORES en TODO el libro (celda por celda, 40 hojas, sin mapa de coordenadas).
        //    Plan 59 §15: la verificación anterior recorría `WorkbookLeafCellMapQ2` + `RolsQ2`
        //    (foto de la v2); con la plantilla vigente esas direcciones ya no describen los roles,
        //    así que el rojo era del fixture, no de la paridad. La comparación libro↔libro cubre
        //    todas las celdas escritas (R1/R2/R4 Q2, CONSOLIDADO, DetRetri, detalle R2/R4, recaudos).
        AssertMismoContenidoPorValor(salidaCli, salidaDirecta);

        foreach (var leaf in resultadoDirecto.Leafs)
        {
            // AJUSTES-SF-T: operandos SALDOS POR NOTA y RETRIBUCION NEGATIVA idénticos. Las celdas
            // vienen RESUELTAS por el reader (dict celda→valor), no de un mapa de coordenadas; solo
            // el nombre de la hoja sale del catálogo (dato del negocio, `SALDOS POR NOTA`).
            var ajustes = leaf.AjustesSfT!;
            foreach (var (celda, valor) in ajustes.SaldosNotas.Celdas)
            {
                Assert.Equal(valor, LeerCelda(salidaCli, WorkbookLeafCellMapAjustesSfT.HojaSaldosNotas, celda));
                Assert.Equal(valor, LeerCelda(salidaDirecta, WorkbookLeafCellMapAjustesSfT.HojaSaldosNotas, celda));
            }

            foreach (var (celda, valor) in ajustes.RetribucionNegativa.Celdas)
            {
                Assert.Equal(valor, LeerCelda(salidaCli, WorkbookLeafCellMapAjustesSfT.HojaRetribucionNegativa, celda));
                Assert.Equal(valor, LeerCelda(salidaDirecta, WorkbookLeafCellMapAjustesSfT.HojaRetribucionNegativa, celda));
            }

            // DetRetri-Q2: D9:D13 = ROUND(D104:D108) y D14 = ROUND(Σ).
            var celdaDetRetri = WorkbookLeafCellMapQ2.ObtenerDetRetriDestino(leaf.Ase.Id);
            Assert.Equal(leaf.DetRetriQ2!.Detalle, LeerCelda(salidaCli, WorkbookLeafCellMapQ2.HojaDetRetri(Insumos.PeriodoQ2()), celdaDetRetri));
            Assert.Equal(leaf.DetRetriQ2.Detalle, LeerCelda(salidaDirecta, WorkbookLeafCellMapQ2.HojaDetRetri(Insumos.PeriodoQ2()), celdaDetRetri));
        }

        var totalD104 = resultadoDirecto.Leafs.Sum(l => l.DetRetriQ2!.TotalD104);
        var d14Esperado = DetRetriRounder.Round(totalD104);
        Assert.Equal(d14Esperado, LeerCelda(salidaCli, WorkbookLeafCellMapQ2.HojaDetRetri(Insumos.PeriodoQ2()), WorkbookLeafCellMapQ2.DetRetriTotal));
        Assert.Equal(d14Esperado, LeerCelda(salidaDirecta, WorkbookLeafCellMapQ2.HojaDetRetri(Insumos.PeriodoQ2()), WorkbookLeafCellMapQ2.DetRetriTotal));

        AssertSnapshotsIguales(salidaCli, salidaDirecta, Insumos.PeriodoQ2());
    }

    // ── W-2.1 / Req 4: RunId inyectado respetado por el procesador ─────────────────────────────

    [Fact]
    public void RunId_Inyectado_ElProcesadorLoRespeta_UnSoloGuid()
    {
        var sink = new CapturaEventos();
        var logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .Enrich.FromLogContext()
            .WriteTo.Sink(sink)
            .CreateLogger();
        Log.Logger = logger;

        var guidFijo = Guid.Parse("11111111-2222-3333-4444-555555555555");
        var salida = Path.Combine(DirectorioTemp("runid"), Insumos.Periodo().NombreArchivo);
        var resultado = EjecutarUnAseDirectoConRunId(1, Insumos.Periodo(), salida, guidFijo);

        Assert.NotNull(resultado);
        var inicio = sink.Eventos.First(e => e.MessageTemplate.Text.Contains("Iniciando ejecución del procesador real.", StringComparison.Ordinal));
        var runIdEvento = inicio.Properties.TryGetValue("RunId", out var v) && v is ScalarValue { Value: Guid g } ? g : Guid.Empty;
        Assert.Equal(guidFijo, runIdEvento); // NO se genera un segundo Guid
    }

    [Fact]
    public void RunId_Nulo_ElProcesadorGeneraUnoDistintoPorEjecucion()
    {
        var sink = new CapturaEventos();
        var logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .Enrich.FromLogContext()
            .WriteTo.Sink(sink)
            .CreateLogger();
        Log.Logger = logger;

        var salida1 = Path.Combine(DirectorioTemp("runid-null-1"), Insumos.Periodo().NombreArchivo);
        var salida2 = Path.Combine(DirectorioTemp("runid-null-2"), Insumos.Periodo().NombreArchivo);
        EjecutarUnAseDirecto(1, Insumos.Periodo(), salida1);
        EjecutarUnAseDirecto(1, Insumos.Periodo(), salida2);

        var runIds = sink.Eventos
            .Where(e => e.MessageTemplate.Text.Contains("Iniciando ejecución del procesador real.", StringComparison.Ordinal))
            .Select(e => e.Properties["RunId"])
            .Cast<ScalarValue>()
            .Select(s => (Guid)s.Value!)
            .ToList();
        Assert.Equal(2, runIds.Count);
        Assert.NotEqual(runIds[0], runIds[1]);
    }

    // ── Helpers ────────────────────────────────────────────────────────────────────────────────

    private static ProcesadorPeriodo CrearProcesadorPeriodo() =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader(),
            new ValidacionOracleReader());

    private static ResultadoProcesoAse EjecutarUnAseDirecto(int aseId, Periodo periodo, string salida) =>
        EjecutarUnAseDirectoConRunId(aseId, periodo, salida, null);

    private static ResultadoProcesoAse EjecutarUnAseDirectoConRunId(int aseId, Periodo periodo, string salida, Guid? runId)
    {
        var ase = AseFactory.DesdeId(aseId);
        var locator = new ArchivoFuenteLocator();
        var carpetaAse = locator.ObtenerCarpetasAse(periodo.NumeroQuincena == 2 ? Insumos.CarpetaPeriodoQ2 : Insumos.CarpetaPeriodo)
            .First(c => Path.GetFileName(c).StartsWith(aseId.ToString(), StringComparison.OrdinalIgnoreCase));
        var r1 = locator.BuscarArchivo(carpetaAse, "Recaudoporcomponente")!;
        var r2 = locator.BuscarArchivo(carpetaAse, "RerpoteDetalleSaldosaFavor")!;
        var r4 = locator.BuscarArchivo(carpetaAse, "ReversiónPorComponente")
            ?? locator.BuscarArchivo(carpetaAse, "ReversionPorComponente")!;
        var plantilla = periodo.NumeroQuincena == 2 ? Insumos.PlantillaQ2 : Insumos.Plantilla;

        return new ProcesadorRemuneracion(
                new ExcelDataReaderRecaudoReader(),
                new ExcelDataReaderWorkbookLeafInputReader(),
                new CalculoRemuneracion(),
                new ValidadorBasico(),
                new OpenXmlPlantillaWriter())
            .Ejecutar(new SolicitudProcesoAse
            {
                Ase = ase,
                Periodo = periodo,
                RutaR1 = r1,
                RutaR2 = r2,
                RutaR4 = r4,
                RutaPlantilla = plantilla,
                RutaSalida = salida,
                RunId = runId
            });
    }

    /// <summary>
    /// Plan 59 (§15): PARIDAD sin coordenadas. Recorre las 40 hojas de los dos libros y compara el
    /// valor de cada celda (numérico con tolerancia ±0,5 —tolerancia del proyecto—, texto exacto).
    /// No usa ningún mapa de celdas ni de roles: la plantilla vigente puede mover la geometría sin
    /// que este gate se vuelva una foto. Un libro con una celda de más, de menos o con otro valor
    /// enrojece.
    /// </summary>
    private static void AssertMismoContenidoPorValor(string rutaA, string rutaB)
    {
        using var libroA = SpreadsheetDocument.Open(rutaA, false);
        using var libroB = SpreadsheetDocument.Open(rutaB, false);

        static Dictionary<string, WorksheetPart> Hojas(SpreadsheetDocument libro)
        {
            var parte = libro.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            return parte.Workbook!.Descendants<Sheet>().ToDictionary(
                s => s.Name!.Value!,
                s => (WorksheetPart)parte.GetPartById(s.Id!),
                StringComparer.Ordinal);
        }

        static Dictionary<string, string> Valores(WorksheetPart parte, SharedStringTable? compartidas)
        {
            var mapa = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var celda in parte.Worksheet!.Descendants<Cell>())
            {
                if (celda.CellValue is null)
                {
                    continue;
                }

                var texto = celda.CellValue.InnerText;
                if (celda.DataType?.Value == CellValues.SharedString && compartidas is not null)
                {
                    texto = compartidas.ElementAt(int.Parse(texto, System.Globalization.CultureInfo.InvariantCulture)).InnerText;
                }

                mapa[celda.CellReference!.Value!] = texto;
            }

            return mapa;
        }

        var hojasA = Hojas(libroA);
        var hojasB = Hojas(libroB);
        var compartidasA = libroA.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
        var compartidasB = libroB.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
        Assert.Equal(hojasA.Keys.OrderBy(k => k, StringComparer.Ordinal), hojasB.Keys.OrderBy(k => k, StringComparer.Ordinal));

        foreach (var (nombre, parteA) in hojasA)
        {
            var valoresA = Valores(parteA, compartidasA);
            var valoresB = Valores(hojasB[nombre], compartidasB);
            var diferencias = new List<string>();

            foreach (var (referencia, valorA) in valoresA.OrderBy(kv => kv.Key, StringComparer.Ordinal))
            {
                if (!valoresB.TryGetValue(referencia, out var valorB))
                {
                    diferencias.Add($"{nombre}!{referencia}: existe en la vía directa ({valorA}) y falta en la del CLI.");
                    continue;
                }

                var esNumA = decimal.TryParse(valorA, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var numA);
                var esNumB = decimal.TryParse(valorB, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var numB);
                if (esNumA && esNumB)
                {
                    if (Math.Abs(numA - numB) > Insumos.Tolerancia)
                    {
                        diferencias.Add($"{nombre}!{referencia}: CLI={numB} vs directo={numA}.");
                    }
                }
                else if (!string.Equals(valorA, valorB, StringComparison.Ordinal))
                {
                    diferencias.Add($"{nombre}!{referencia}: CLI='{valorB}' vs directo='{valorA}'.");
                }
            }

            foreach (var referencia in valoresB.Keys.Where(k => !valoresA.ContainsKey(k)))
            {
                diferencias.Add($"{nombre}!{referencia}: existe en la del CLI ({valoresB[referencia]}) y falta en la directa.");
            }

            Assert.True(
                diferencias.Count == 0,
                $"{nombre}: {diferencias.Count} celdas distintas entre CLI y vía directa:{Environment.NewLine}" +
                string.Join(Environment.NewLine, diferencias.Take(20)));
        }
    }

    private static void AssertSnapshotsIguales(string rutaA, string rutaB, Periodo periodo)
    {
        var snapshotsA = new ValidacionOracleReader().LeerSnapshots(rutaA, periodo);
        var snapshotsB = new ValidacionOracleReader().LeerSnapshots(rutaB, periodo);
        Assert.Equal(snapshotsA.Count, snapshotsB.Count);

        foreach (var (a, b) in snapshotsA.OrderBy(s => s.Ase.Id).Zip(snapshotsB.OrderBy(s => s.Ase.Id)))
        {
            Assert.Equal(a.Ase.Id, b.Ase.Id);
            Assert.Equal(a.ValidacionTotal, b.ValidacionTotal);
            Assert.Equal(a.ValidacionTotalOkP, b.ValidacionTotalOkP);
            Assert.Equal(a.SubBloquesValidacionTotal.OrderBy(kv => kv.Key), b.SubBloquesValidacionTotal.OrderBy(kv => kv.Key));
            foreach (var (empA, empB) in a.PorEmpresa.OrderBy(e => e.Empresa, StringComparer.OrdinalIgnoreCase)
                         .Zip(b.PorEmpresa.OrderBy(e => e.Empresa, StringComparer.OrdinalIgnoreCase)))
            {
                Assert.Equal(empA.Empresa, empB.Empresa);
                Assert.Equal(empA.DiferenciaO, empB.DiferenciaO);
                Assert.Equal(empA.VerificacionP, empB.VerificacionP);
            }

            Assert.Equal(a.DetValiRetri is null, b.DetValiRetri is null);
            if (a.DetValiRetri is not null && b.DetValiRetri is not null)
            {
                Assert.Equal(a.DetValiRetri.DiferenciasAse.Select(d => (d.Celda, d.Valor)), b.DetValiRetri.DiferenciasAse.Select(d => (d.Celda, d.Valor)));
                Assert.Equal(a.DetValiRetri.VerificacionesAse.Select(v => (v.Celda, v.Verificacion)), b.DetValiRetri.VerificacionesAse.Select(v => (v.Celda, v.Verificacion)));
                Assert.Equal(a.DetValiRetri.VerificacionTotalD29, b.DetValiRetri.VerificacionTotalD29);
            }
        }
    }

    private static Guid RunIdDe(string stdout)
    {
        var match = Regex.Match(stdout, @"runId=([0-9a-fA-F-]{36})");
        Assert.True(match.Success, "La línea RESULTADO OK debe imprimir el runId.");
        return Guid.Parse(match.Groups[1].Value);
    }

    private static decimal LeerCelda(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var cell = worksheetPart.Worksheet!.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        Assert.True(cell is not null, $"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}");
        Assert.NotNull(cell!.CellValue);
        var texto = cell.CellValue!.InnerText;
        if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString)
        {
            var shared = workbookPart.SharedStringTablePart!.SharedStringTable!.ElementAt(int.Parse(texto));
            texto = shared.InnerText;
        }

        return decimal.Parse(texto, System.Globalization.CultureInfo.InvariantCulture);
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
