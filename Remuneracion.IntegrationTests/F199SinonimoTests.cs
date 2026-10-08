using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 34 (T1/T2, R-G-1/R-G-2 + R-B-1..R-B-5): gate TDD rojo-primero del aborto
/// <c>ASE 2 Reporte Componentes R1!F199 [SUB_EMP]: faltan las anclas [Mes0]</c> y de la equivalencia
/// de sinónimos de empresa del interior R1.
///
/// Causa adoptada (Plan 33-T0, H1 CONFIRMADA): la celda del template <c>C199</c> trae el rótulo legado
/// <c>RECIPROCIDAD</c> (marca histórica de EAAB) mientras la fuente R1 de agosto trae el nombre vigente
/// <c>NUEVO ESQUEMA</c> (<c>Detalle de plantilla</c>: «NUEVO ESQUEMA = recaudo EAAB Reciprocidad»).
/// El filtro por nombre del interior (<c>R1FirmaInterior.EsDatoEmpresa(fila, empresa)</c>) no halla filas
/// → <c>MesOportuno</c> vacío → el throw aborta el pipeline.
///
/// Método: reproducción in-process (lectura ZIP+XML BCL, sin Excel/COM) — se abren los 5 bloques R1 de la
/// fuente REAL de agosto con <c>LeerEspejoR1</c> y se corre <c>AjustarEnWorkbook</c> sobre la base
/// canónica <c>2026082</c> copiada a una ruta TEMP FRESCA. Pre-fix el pase aborta con el throw exacto
/// (ROJO de T1); post-fix compone <c>F199</c> con el texto del manual (VERDE de T2).
///
/// Rutas por EXPLICITA construcción (estilo Plan 33): NO se usa <see cref="Insumos.Raiz"/> (exige
/// <c>Docs/Insumos</c>, borrado por el reorg <c>a867706</c>) ni ninguna resolución de raíz de fixtures
/// pre-existente. PROHIBIDO usar cachés <c>&lt;v&gt;</c> como oráculo.
/// </summary>
public sealed class F199SinonimoTests
{
    private const string HojaR1 = "Reporte Componentes R1";

    /// <summary>Texto de <c>&lt;f&gt;</c> que el manual del administrativo trae en <c>F199</c> (oráculo).</summary>
    private const string FormulaManualF199 = "F140+F114-L114";

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental",
        "2-Lime",
        "3-Ciudad Limpia",
        "4-Bogota Limpia",
        "5-Área Limpia"
    ];

    // ── S1 (TDD rojo-primero): el pase interior compone F199 sin abortar ────────────────────────

    [Fact]
    public void S1_Agosto_Ase2_F199_SeComponeDesdeFuenteReal()
    {
        using var copia = new CopiaTemporal(PlantillaAgosto2026082());
        var bloques = LeerBloquesAgosto();

        string? textoF199;
        Exception? aborto = null;
        using (var documento = SpreadsheetDocument.Open(copia.Ruta, true))
        {
            var workbookPart = documento.WorkbookPart
                ?? throw new InvalidOperationException("El workbook de la base canónica no tiene WorkbookPart.");

            // Pre-fix (T1): aquí se lanzaba `...ASE 2 ...F199 [SUB_EMP]: faltan las anclas [Mes0]...`
            // (ROJO reproducido). Post-fix (T2): ASE2 F199 se compone; si el pase avanza y aborta en
            // OTRA celda (R-MAPA-OCULTO: p. ej. ASE4 F463), ese aborto es follow-up declarado, no
            // regresión de T5. F199 ya quedó escrito en el DOM antes de ese throw.
            try
            {
                OpenXmlEspejoR1Mutador.AjustarEnWorkbook(workbookPart, bloques);
            }
            catch (Exception ex)
            {
                aborto = ex;
            }

            textoF199 = LeerFormulaR1(workbookPart, "F199");
        }

        // R-B-3 / S2: F199 compuesto desde las filas reales de NUEVO ESQUEMA, texto = manual.
        Assert.Equal(FormulaManualF199, textoF199);

        // R-G-1 / S1: el aborto F199 [SUB_EMP] [Mes0] desapareció con la equivalencia.
        var sigueAbortandoF199 = aborto is not null
            && aborto.Message.Contains("F199", StringComparison.Ordinal)
            && aborto.Message.Contains("[SUB_EMP]", StringComparison.Ordinal);
        Assert.False(
            sigueAbortandoF199,
            $"El aborto F199 [SUB_EMP] [Mes0] debe desaparecer con la equivalencia (post-fix). Aborto observado: {aborto?.Message}");
    }

    // ── S3 (no-atribución cruzada): la equivalencia no arrastra otras empresas ──────────────────

    [Fact]
    public void S3_Equivalencia_NoAtrastraOtrasEmpresas()
    {
        var bloque = LeerBloquesAgosto()[1]; // ASE2 (LIME)

        Assert.True(
            bloque.Filas.Any(f => R1FirmaInterior.EsDatoEmpresa(f, "RECIPROCIDAD")),
            "MesOportuno(RECIPROCIDAD) debe hallar las filas de la fuente rotuladas NUEVO ESQUEMA (S2).");

        // ENEL no debe matchear las filas de NUEVO ESQUEMA ni OCCIDENTE (matching exacto salvo sinónimos).
        Assert.False(
            bloque.Filas.Any(f => R1FirmaInterior.EsDatoEmpresa(f, "ENEL")
                && !string.Equals(f.C, "ENEL", StringComparison.OrdinalIgnoreCase)),
            "ENEL solo debe matchear filas ENEL (la equivalencia es de la clase EAAB, no de todo el bloque).");
    }

    // ── R-B-1: tabla de sinonimia como DATOS citados + matcher puro ─────────────────────────────

    [Fact]
    public void R_B_1_TablaSinonimos_ContieneExactamenteLasClasesCitadas()
    {
        var clase = Assert.Single(SinonimosEmpresaR1.Clases);
        Assert.True(
            clase.Miembros.SetEquals(["RECIPROCIDAD", "NUEVO ESQUEMA"]),
            "La clase inicial debe ser exactamente {RECIPROCIDAD, NUEVO ESQUEMA} (EAAB Reciprocidad).");
        Assert.False(
            string.IsNullOrWhiteSpace(clase.Cita),
            "Cada fila de la tabla exige su cita de negocio (sin cita no hay fila — R-SINONIMO-ABUSO).");
    }

    [Theory]
    [InlineData("RECIPROCIDAD", "NUEVO ESQUEMA", true)]
    [InlineData("NUEVO ESQUEMA", "RECIPROCIDAD", true)]
    [InlineData("reciprocidad", "nuevo esquema", true)]
    [InlineData("RECIPROCIDAD", "NUEVO ESQUÉMA", true)]
    [InlineData("RECIPROCIDAD", "RECIPROCIDAD", true)]
    [InlineData("ENEL", "ENEL", true)]
    [InlineData("ENEL", "OCCIDENTE", false)]
    [InlineData("ENEL", "NUEVO ESQUEMA", false)]
    [InlineData("OCCIDENTE", "RECIPROCIDAD", false)]
    [InlineData("ENERBIT", "CIUDAD LIMPIA-ACUEDUCTO", false)]
    [InlineData("EMPRESA FUTURA", "EMPRESA FUTURA", true)]
    [InlineData("EMPRESA FUTURA", "OTRA EMPRESA", false)]
    [InlineData("", "RECIPROCIDAD", false)]
    public void R_B_1_SonMismaEmpresa(string a, string b, bool esperado)
    {
        Assert.Equal(esperado, SinonimosEmpresaR1.SonMismaEmpresa(a, b));
    }

    // ── R-B-2: el overload EsDatoEmpresa(fila, empresa) aprende la equivalencia ─────────────────

    [Fact]
    public void R_B_2_EsDatoEmpresa_AprendeSinonimoLegadoVigente()
    {
        var legado = new FilaEspejoR1 { C = "RECIPROCIDAD", D = "Total" };
        var vigente = new FilaEspejoR1 { C = "NUEVO ESQUEMA", D = "Total" };

        Assert.True(R1FirmaInterior.EsDatoEmpresa(vigente, "RECIPROCIDAD"));
        Assert.True(R1FirmaInterior.EsDatoEmpresa(legado, "NUEVO ESQUEMA"));

        var enel = new FilaEspejoR1 { C = "ENEL", D = "Total" };
        Assert.True(R1FirmaInterior.EsDatoEmpresa(enel, "ENEL"));
        Assert.False(R1FirmaInterior.EsDatoEmpresa(enel, "OCCIDENTE"));

        // Una fila con B no-vacía no es fila-dato de empresa (firma intacta).
        var conB = new FilaEspejoR1 { B = "Mes", C = "NUEVO ESQUEMA", D = "Total" };
        Assert.False(R1FirmaInterior.EsDatoEmpresa(conB, "RECIPROCIDAD"));
    }

    // ── S5 (D-G): fail-fast enriquecido con rótulo-template + etiquetas fuente observadas ───────

    [Fact]
    public void S5_FailFast_EnriquecidoConAmbosLados()
    {
        using var copia = new CopiaTemporal(PlantillaAgosto2026082());

        // Se renombra la empresa de la fuente del ASE2 a un rótulo NO tabulado: el sub-bloque queda
        // en 0 filas y el throw debe nombrar el rótulo del template y las etiquetas observadas.
        var bloques = LeerBloquesAgosto().ToList();
        var ase2 = bloques[1];
        bloques[1] = new BloqueEspejoAseInputs
        {
            Ase = ase2.Ase,
            Encabezados = ase2.Encabezados,
            TieneColumnaEspeciales = ase2.TieneColumnaEspeciales,
            Filas = ase2.Filas
                .Select(f => string.Equals(f.C, "NUEVO ESQUEMA", StringComparison.OrdinalIgnoreCase)
                    ? new FilaEspejoR1
                    {
                        A = f.A,
                        B = f.B,
                        C = "EMPRESA FUTURA X",
                        D = f.D,
                        E = f.E,
                        ValoresPorColumna = f.ValoresPorColumna
                    }
                    : f)
                .ToList()
        };

        var excepcion = Assert.Throws<Remuneracion.Core.Exceptions.CalculoInvalidoException>(() =>
        {
            using var documento = SpreadsheetDocument.Open(copia.Ruta, true);
            OpenXmlEspejoR1Mutador.AjustarEnWorkbook(documento.WorkbookPart!, bloques);
        });

        Assert.Contains("F199 [SUB_EMP]", excepcion.Message, StringComparison.Ordinal);
        Assert.Contains("Mes0", excepcion.Message, StringComparison.Ordinal);
        Assert.Contains("rótulo template: 'RECIPROCIDAD'", excepcion.Message, StringComparison.Ordinal);
        Assert.Contains("EMPRESA FUTURA X", excepcion.Message, StringComparison.Ordinal);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────────────────

    private static IReadOnlyList<BloqueEspejoAseInputs> LeerBloquesAgosto()
    {
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var bloques = new List<BloqueEspejoAseInputs>(5);
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            bloques.Add(reader.LeerEspejoR1(AseFactory.DesdeId(aseId), R1Agosto(aseId)));
        }

        return bloques;
    }

    private static string PlantillaAgosto2026082() =>
        Path.Combine(Raiz(), "Docs", "Prueba Agosto-2", "Plantilla_Remuneracion_2026082.xlsx");

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

    private static string? LeerFormulaR1(WorkbookPart workbookPart, string referencia)
    {
        var hoja = workbookPart.Workbook?.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, HojaR1, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"El workbook no tiene la hoja '{HojaR1}'.");
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(hoja.Id!);
        var worksheet = worksheetPart.Worksheet
            ?? throw new InvalidOperationException($"La hoja '{HojaR1}' no tiene Worksheet.");
        var celda = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        return celda?.CellFormula?.Text;
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

    private sealed class CopiaTemporal : IDisposable
    {
        private readonly string _dir;

        public CopiaTemporal(string origen)
        {
            _dir = Path.Combine(Path.GetTempPath(), "p34-f199-" + Guid.NewGuid().ToString("N"));
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
