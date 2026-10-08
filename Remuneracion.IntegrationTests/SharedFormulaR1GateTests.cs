using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 33 (T1/T2, R-G-1..R-G-4 / R-B-1): gate TDD rojo-primero del invariante shared de
/// <c>Reporte Componentes R1</c>.
///
/// Corpus:
///   - <b>regen fresca julio-2026072</b> a temp (E5, pipeline in-process, nunca sobre una salida
///     existente): PRE-FIX en ROJO (33 masters, huérfanos si=0/si=20 — E4); POST-FIX en VERDE (35
///     masters, G53/G468 preservados con t/ref/si y texto idéntico — T2).
///   - <b>Resultado1</b> pre-commit (control sano): VERDE siempre.
///   - <b>salida agosto existente sana</b> (workbook-wide, E7): VERDE.
///
/// Nota de entorno (2026-10-08): los insumos vigentes viven en <c>Docs/Prueba Julio-1</c>,
/// <c>Docs/Prueba Julio-2</c> y <c>Docs/Prueba Agosto-2</c>. Este gate resuelve la raíz por
/// <c>AGENTS.md</c> y usa las rutas REALES vigentes. La regen fresca de agosto NO es
/// posible hoy: el pipeline aborta antes en <c>F199 [SUB_EMP] [Mes0]</c> (sujeto del T0/T4 de la Fase 2).
/// </summary>
public sealed class SharedFormulaR1GateTests
{
    private const string Hoja = SharedFormulaR1Esperados.Hoja;

    // ── S1 (TDD rojo-primero): regen fresca julio ────────────────────────────────────────────

    [Fact]
    public void S1_RegenFrescaJulio_MaestrosSharedPreservados()
    {
        using var salida = new SalidaTemporal("p33-shared-julio-");
        Ejecutar(
            new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 },
            InsumosJulio(),
            PlantillaJulio(),
            salida.Ruta);

        SharedFormulaR1Gate.ExigirInvariante(salida.Ruta, Hoja);

        var lectura = SharedFormulaR1Gate.Leer(salida.Ruta, Hoja);
        Assert.Equal(SharedFormulaR1Esperados.TotalMaestros, lectura.Maestros.Count);
        Assert.Equal(SharedFormulaR1Esperados.TotalSeguidoras, lectura.Seguidoras.Count);
        AssertMaster(lectura, "G53", 0, "G53:AP53", "G32+G12");
        AssertMaster(lectura, "G468", 20, "G468:AO468", "G422+G384");
    }

    // ── S2 (control sano): Resultado1 pre-commit ─────────────────────────────────────────────

    [Fact]
    public void S2_ControlResultado1_InvarianteVerde()
    {
        var ruta = Resultado1();

        SharedFormulaR1Gate.ExigirInvariante(ruta, Hoja);

        var lectura = SharedFormulaR1Gate.Leer(ruta, Hoja);
        Assert.Equal(SharedFormulaR1Esperados.TotalMaestros, lectura.Maestros.Count);
        Assert.Equal(SharedFormulaR1Esperados.TotalSeguidoras, lectura.Seguidoras.Count);
        AssertMaster(lectura, "G53", 0, "G53:AP53", "G32+G12");
        AssertMaster(lectura, "G468", 20, "G468:AO468", "G422+G384");
    }

    // ── S3 (fixture-drift): el mapa congelado coincide con el disco ──────────────────────────

    [Fact]
    public void S3_FixtureMapaShared_CoincideConPlantillaYResultado1()
    {
        AssertFixture(PlantillaJulio());
        AssertFixture(Resultado1());
    }

    // ── S4 (caso agosto, E7): invariante workbook-wide sobre la salida agosto sana ───────────

    [Fact]
    public void S4_AgostoSalidaSana_InvarianteVerde()
    {
        var ruta = AgostoResultadoSano();

        SharedFormulaR1Gate.ExigirInvariante(ruta, Hoja);
        var lectura = SharedFormulaR1Gate.Leer(ruta, Hoja);
        Assert.Equal(SharedFormulaR1Esperados.TotalMaestros, lectura.Maestros.Count);
        Assert.Equal(SharedFormulaR1Esperados.TotalSeguidoras, lectura.Seguidoras.Count);
    }

    // ── T2 (R-B-1): unitarios del helper preserve-and-rewrite por rama ───────────────────────

    [Fact]
    public void T2_ConCellFormulaPrevia_MutaTextoYPreservaAtributosShared()
    {
        var celda = new Cell { CellReference = "G53" };
        celda.CellFormula = new CellFormula("G32+G12")
        {
            FormulaType = CellFormulaValues.Shared,
            Reference = "G53:AP53",
            SharedIndex = 0
        };

        OpenXmlEspejoR1Mutador.EscribirFormulaPreservando(celda, "G32+G12");

        Assert.NotNull(celda.CellFormula);
        Assert.Equal("G32+G12", celda.CellFormula!.Text);
        Assert.Equal(CellFormulaValues.Shared, celda.CellFormula.FormulaType!.Value);
        Assert.Equal("G53:AP53", celda.CellFormula.Reference!.Value);
        Assert.Equal(0u, celda.CellFormula.SharedIndex!.Value);
    }

    [Fact]
    public void T2_SinCellFormulaPrevia_CreaFormulaPlanaSinAtributosShared()
    {
        var celda = new Cell { CellReference = "F53" };

        OpenXmlEspejoR1Mutador.EscribirFormulaPreservando(celda, "F32+F12");

        Assert.NotNull(celda.CellFormula);
        Assert.Equal("F32+F12", celda.CellFormula!.Text);
        Assert.Null(celda.CellFormula.FormulaType);
        Assert.Null(celda.CellFormula.Reference);
        Assert.Null(celda.CellFormula.SharedIndex);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────────────────

    private static void AssertMaster(SharedFormulaR1Gate.LecturaShared lectura, string celda, uint si, string rango, string texto)
    {
        Assert.True(lectura.Maestros.TryGetValue(celda, out var master),
            $"Falta el master shared {celda} (si={si}, ref={rango}) — la seguidora quedaría huérfana.");
        Assert.Equal(si, master!.Si);
        Assert.Equal(rango, master.Rango);
        Assert.Equal(texto, master.Texto);
    }

    private static void AssertFixture(string ruta)
    {
        var lectura = SharedFormulaR1Gate.Leer(ruta, Hoja);
        var esperados = SharedFormulaR1Esperados.Maestros.ToDictionary(m => m.Celda, StringComparer.OrdinalIgnoreCase);

        Assert.Equal(SharedFormulaR1Esperados.TotalMaestros, lectura.Maestros.Count);
        Assert.Equal(SharedFormulaR1Esperados.TotalSeguidoras, lectura.Seguidoras.Count);
        Assert.Equal(SharedFormulaR1Esperados.TotalMaestros, esperados.Count);

        foreach (var (celda, real) in lectura.Maestros)
        {
            Assert.True(esperados.TryGetValue(celda, out var esperado), $"El fixture no congela el master {celda} del disco (drift).");
            Assert.Equal(esperado!.Si, real.Si);
            Assert.Equal(esperado.Rango, real.Rango);
            Assert.Equal(esperado.Texto, real.Texto);
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

    private static string Raiz()
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

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }

    private static string PlantillaJulio() =>
        Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Plantilla_Remuneracion.xlsx");

    private static string InsumosJulio() =>
        Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Insumos");

    private static string Resultado1() =>
        PrimerXlsx(Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Resultado", "Resultado1"));

    private static string AgostoResultadoSano() =>
        PrimerXlsx(Path.Combine(Raiz(), "Docs", "Prueba Agosto-2", "Resultado", "Resultado2"));

    private static string PrimerXlsx(string carpeta)
    {
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta de evidencia: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly).OrderBy(f => f).FirstOrDefault();
        Assert.True(archivo is not null, $"Falta un .xlsx en {carpeta}");
        return archivo!;
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
