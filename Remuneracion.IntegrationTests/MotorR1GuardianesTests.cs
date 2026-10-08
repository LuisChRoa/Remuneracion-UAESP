using DocumentFormat.OpenXml.Packaging;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 35 (T2, R-B-1/R-B-5 parcial — S2): gate de los GUARDIANES explícitos del motor R1-interior.
///
/// Corre el pase del espejo (<c>AjustarEnWorkbook</c>: dimensionalización 5→1 + recomposición de
/// visibles + recomposición del interior) sobre una copia TEMP FRESCA de la base canónica de agosto
/// <c>Docs/Prueba Agosto-2/Plantilla_Remuneracion_2026082.xlsx</c> con las fuentes R1 REALES de agosto,
/// y exige el texto de <c>&lt;f&gt;</c> del manual del administrativo (fixture congelado
/// <see cref="InterioresR1Esperados.Agosto"/>) en:
///   - ASE4 <c>F461/F463/F466/F468/F476/F478</c> (empresa no-última: el defecto H2/F463),
///   - ASE2 <c>F199</c> (sinonímia RECIPROCIDAD↔NUEVO ESQUEMA, Plan 34) y <c>F204/F206</c>
///     (miscomposición silenciosa: 3 términos con quirk <c>--L</c> y 2 términos).
///
/// Rutas EXPLÍCITAS (estilo Plan 33/34): no se usa <see cref="Insumos.Raiz"/> ni <c>Docs/Insumos</c>
/// (borrados por el reorg <c>a867706</c>). La salida se lee con ZIP+XML BCL (sin Excel/COM); PROHIBIDO
/// usar cachés <c>&lt;v&gt;</c> como oráculo. Tolerancia de valores ±0.5 (aquí el gate es textual).
/// </summary>
public sealed class MotorR1GuardianesTests
{
    private const string HojaR1 = "Reporte Componentes R1";

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia"
    ];

    /// <summary>ASE4 (Bogotá Limpia): empresa-dato que no cierra su sección (causa F463/H2).</summary>
    [Theory]
    [InlineData(4, "F461")]
    [InlineData(4, "F463")]
    [InlineData(4, "F466")]
    [InlineData(4, "F468")]
    [InlineData(4, "F476")]
    [InlineData(4, "F478")]
    public void Ase4_GuardianesManua_TextoIgualAlManual(int aseId, string celda)
    {
        AssertCelda(aseId, celda);
    }

    /// <summary>ASE2 (LIME): sinonímia Plan 34 y miscomposición silenciosa F204/F206.</summary>
    [Theory]
    [InlineData(2, "F199")]
    [InlineData(2, "F204")]
    [InlineData(2, "F206")]
    public void Ase2_GuardianesManua_TextoIgualAlManual(int aseId, string celda)
    {
        AssertCelda(aseId, celda);
    }

    /// <summary>F463 (EXT_INT) compone FÓRMULA real (no literal) — H1 descartada para ASE4 (T2).</summary>
    [Fact]
    public void Ase4_F463_ComponeFormulaNoLiteral()
    {
        using var salida = EjecutarAgosto();
        var celdas = ValidadorTotalesR1Workbook.LeerCeldas(salida.Ruta, HojaR1);

        Assert.True(celdas.TryGetValue("F463", out var celda), "Falta F463 en la salida.");
        Assert.NotNull(celda!.Formula);
        Assert.Equal(Esperado(4, "F463"), celda.Formula);
    }

    private static void AssertCelda(int aseId, string celda)
    {
        using var salida = EjecutarAgosto();
        var celdas = ValidadorTotalesR1Workbook.LeerCeldas(salida.Ruta, HojaR1);

        Assert.True(celdas.TryGetValue(celda, out var real), $"Falta {celda} (ASE {aseId}) en la salida de agosto.");
        var esperado = Esperado(aseId, celda);
        Assert.Equal(esperado, real!.Formula);
    }

    private static string Esperado(int aseId, string celda)
    {
        var interior = InterioresR1Esperados.Agosto
            .FirstOrDefault(e => e.AseId == aseId && string.Equals(e.Celda, celda, StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(interior);
        Assert.NotNull(interior!.Formula);
        return interior.Formula!;
    }

    private static SalidaTemporal EjecutarAgosto()
    {
        var baseCanonica = Path.Combine(Raiz(), "Docs", "Prueba Agosto-2", "Plantilla_Remuneracion_2026082.xlsx");
        var salida = new SalidaTemporal(baseCanonica);

        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var bloques = new List<BloqueEspejoAseInputs>(5);
        for (var ase = 1; ase <= 5; ase++)
        {
            bloques.Add(reader.LeerEspejoR1(Insumos.Ase(ase), R1Agosto(ase)));
        }

        using var documento = SpreadsheetDocument.Open(salida.Ruta, true);
        OpenXmlEspejoR1Mutador.AjustarEnWorkbook(documento.WorkbookPart!, bloques);
        documento.WorkbookPart!.Workbook?.Save();

        return salida;
    }

    private static string R1Agosto(int aseId)
    {
        var carpeta = Path.Combine(Raiz(), "Docs", "Prueba Agosto-2", "Insumos", CarpetasAse[aseId - 1]);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta de insumos de agosto: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f)
                .StartsWith("Recaudoporcomponente", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta la fuente R1 (Recaudoporcomponente_*) del ASE {aseId} en {carpeta}");
        return archivo!;
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

    private sealed class SalidaTemporal : IDisposable
    {
        private readonly string _dir;

        public SalidaTemporal(string origen)
        {
            _dir = Path.Combine(Path.GetTempPath(), "p35-guardianes-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
            Ruta = Path.Combine(_dir, "salida.xlsx");
            File.Copy(origen, Ruta, overwrite: true);
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
