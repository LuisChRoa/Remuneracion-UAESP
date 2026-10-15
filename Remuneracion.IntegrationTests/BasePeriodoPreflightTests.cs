using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
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
/// Plan 36 (T3, R-B-5 / S8 — §2.3-A / D-G): base por procedimiento + preflight de base.
///
/// La plantilla estandarizada única trae internamente el período base <c>2026092</c> (su filename
/// dice <c>2026xx-x</c>): es, por construcción, la base canónica de ese período (procedimiento
/// §2.3-A: los tokens de período se reescriben UNA vez offline con
/// <c>Herramientas/PreparadorBasePeriodo --plantilla-unica</c>; el runtime NUNCA los toca).
///
/// S8 exige: (a) base del período consistente en tokens + <c>N3</c> + hojas; (b) runtime con base
/// EQUIVOCADA → fail-fast ANTES de escribir. Se verifica sobre los bytes del disco (BCL
/// ZIP+XML + el preflight de producción <see cref="ValidadorBasePeriodo"/>); rutas EXPLÍCITAS
/// (nunca <c>Insumos.Raiz()</c> ni la raíz <c>Docs/Insumos</c>).
/// </summary>
public sealed class BasePeriodoPreflightTests
{
    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";

    private static readonly Regex ReferenciaHojaPeriodo = new(
        "(?:DetRetri|DetValiRetri)(?<periodo>\\d{7})",
        RegexOptions.Compiled);

    private static readonly Regex FormulaRegex = new(
        "<f(?<attrs>[^>]*?)>(?<texto>[^<]*)</f>",
        RegexOptions.Singleline | RegexOptions.Compiled);

    // S8 (a): la base del período 2026092 pasa el preflight de producción (N3 + hojas + sellos +
    // variantes de quincena). Plan 40 (T3): la base se PRODUCE con Phase 0 desde el oficial (el
    // oficial crudo ya no es una base preparada: sus refs de 'Recaudo *' se normalizan por quincena).
    [Fact]
    public void S8_BaseDelPeriodo_PreflightVerde()
    {
        var periodo = Periodo.Parse("2026092");
        var salida = Path.Combine(Path.GetTempPath(), "p40-s8-base-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            PreparadorBasePipeline.Ejecutar(PlantillaUnica(), periodo, salida);

            using var documento = SpreadsheetDocument.Open(salida, false);
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");

            ValidadorBasePeriodo.Exigir(parte, periodo);
        }
        finally
        {
            Borrar(salida);
        }
    }

    // S8 (b): la MISMA base usada con otro período → fail-fast antes de escribir (nunca salida
    // sellada con la base de otro período).
    [Fact]
    public void S8_BaseDeOtroPeriodo_FailFastAntesDeEscribir()
    {
        using var documento = SpreadsheetDocument.Open(PlantillaUnica(), false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");

        var ex = Assert.Throws<CalculoInvalidoException>(() => ValidadorBasePeriodo.Exigir(parte, Periodo.Parse("2026082")));

        Assert.Contains("2026082", ex.Message, StringComparison.Ordinal);
        Assert.Contains("no corresponde al período", ex.Message, StringComparison.Ordinal);
    }

    // S8 (a) evidencia de TOKENS: la base trae exactamente las 11 <f> del período de su contrato y
    // 0 referencias a hojas DetRetri/DetValiRetri de OTRO período (0 stale).
    [Fact]
    public void S8_TokensDePeriodo_BaseConsistente()
    {
        using var zip = ZipFile.OpenRead(PlantillaUnica());
        var propias = 0;
        var stale = new List<string>();

        foreach (var entrada in zip.Entries)
        {
            if (!Regex.IsMatch(entrada.FullName, "^xl/worksheets/sheet\\d+\\.xml$"))
            {
                continue;
            }

            using var stream = entrada.Open();
            using var lector = new StreamReader(stream);
            var texto = lector.ReadToEnd();
            foreach (Match formula in FormulaRegex.Matches(texto))
            {
                foreach (Match referencia in ReferenciaHojaPeriodo.Matches(formula.Groups["texto"].Value))
                {
                    if (string.Equals(referencia.Groups["periodo"].Value, "2026092", StringComparison.Ordinal))
                    {
                        propias++;
                    }
                    else
                    {
                        stale.Add($"{entrada.FullName}!{referencia.Value}");
                    }
                }
            }
        }

        Assert.Empty(stale);
        Assert.Equal(11, propias);
    }

    // S8 (a) evidencia de METADATOS: N3 = período de la base, con los sellos base presentes.
    [Fact]
    public void S8_MetadatosBase_N3YSellos()
    {
        using var documento = SpreadsheetDocument.Open(PlantillaUnica(), false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = ((WorksheetPart)parte.GetPartById(
            parte.Workbook!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                .First(s => string.Equals(s.Name?.Value, HojaConsolidado, StringComparison.OrdinalIgnoreCase)).Id!)).Worksheet!;

        string Valor(string referencia) => hoja.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
            .First(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase))
            .CellValue?.InnerText?.Trim() ?? string.Empty;

        Assert.Equal("2026092", Valor("N3"));
        Assert.False(string.IsNullOrWhiteSpace(Valor("G7")));
        Assert.False(string.IsNullOrWhiteSpace(Valor("K7")));
        Assert.False(string.IsNullOrWhiteSpace(Valor("D6")));
    }

    private static string PlantillaUnica()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            var candidata = Path.Combine(dir.FullName, "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");
            if (File.Exists(candidata))
            {
                return candidata;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (plantilla estandarizada única).");
    }

    // ── Plan 39 (T2, R-B-2/R-B-3/R-B-4 — S3/S4/S5): Phase 0 + R-INTERV-Q + preflight por quincena ──

    // S3 (R-INTERV-Q): Phase 0 desde el OFICIAL produce la variante L de la quincena y el preflight
    // de producción queda VERDE sobre la copia auto-preparada (red ciega), para AMBAS quincenas.
    [Theory]
    [InlineData("2026071", "N")]
    [InlineData("2026072", "M")]
    public void S3_Phase0DesdeElOficial_PreflightVerdeYVarianteL(string codigo, string columna)
    {
        var periodo = Periodo.Parse(codigo);
        var salida = Path.Combine(Path.GetTempPath(), "p39-phase0-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            PreparadorBasePipeline.Ejecutar(PlantillaUnica(), periodo, salida);

            using var documento = SpreadsheetDocument.Open(salida, false);
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");

            // Preflight de producción verde (hojas + N3 + sellos + tokens + variante L).
            ValidadorBasePeriodo.Exigir(parte, periodo);

            var consolidado = Hoja(parte, HojaConsolidado);
            for (var i = 0; i < 5; i++)
            {
                Assert.Equal($"INTERVENTORIA!{columna}{26 + i}", FormulaDe(consolidado, $"L{9 + i}"));
            }

            Assert.Equal(codigo, ValorDe(consolidado, "N3"));
        }
        finally
        {
            Borrar(salida);
        }
    }

    // S3 (tercera forma): una copia cuya L9:L13 no es ni M- ni N-variante → ERR-PLANTILLA nombrando
    // hoja + celda + variante esperada. Cero escritura parcial.
    [Fact]
    public void S3_TerceraFormaL9_FailFastNombraHojaCeldaYVariante()
    {
        var copia = Path.Combine(Path.GetTempPath(), "p39-tercera-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            File.Copy(PlantillaUnica(), copia, overwrite: false);
            FijarFormula(copia, HojaConsolidado, "L9", "INTERVENTORIA!Z26");

            using var documento = SpreadsheetDocument.Open(copia, false);
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var ex = Assert.Throws<CalculoInvalidoException>(() => ValidadorBasePeriodo.Exigir(parte, Periodo.Parse("2026092")));

            Assert.Equal(CodigoError.Plantilla, ex.Codigo);
            Assert.Contains("L9", ex.Message, StringComparison.Ordinal);
            Assert.Contains("INTERVENTORIA!M26", ex.Message, StringComparison.Ordinal);
            Assert.Contains("R-INTERV-Q", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            Borrar(copia);
        }
    }

    // S4 (in-place denegado): Phase 0 y el writer rechazan la ruta del oficial como salida ANTES de
    // tocar un byte; el oficial queda intacto (hash idéntico).
    [Fact]
    public void S4_InPlaceDenegado_OficialIntacto()
    {
        var oficial = PlantillaUnica();
        var hashAntes = Sha256(oficial);

        // (a) Phase 0 endurecido: jamás escribe in-place sobre su entrada.
        Assert.Throws<CalculoInvalidoException>(() =>
            PreparadorBasePipeline.Ejecutar(oficial, Periodo.Parse("2026072"), oficial));

        // (b) el writer rechaza salida == plantilla (ERR-PLANTILLA) antes de copiar/Phase 0.
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            new OpenXmlPlantillaWriter().GenerarWorkbook(oficial, oficial, new ResultadoRemuneracion(), new WorkbookLeafInputs()));
        Assert.Equal(CodigoError.Plantilla, ex.Codigo);

        Assert.Equal(hashAntes, Sha256(oficial));
    }

    // S5 (quincena por dominio, end-to-end desde el OFICIAL): el oficial es el input DIRECTO; Phase 0
    // prepara la copia de trabajo (3 hojas + 11 tokens + N3 + columna L) antes del preflight y de
    // cualquier escritura de negocio. La quincena del run (no el contenido) gobierna L9:L13.
    [Theory]
    [InlineData("2026071", "Docs/Prueba Julio-1/Insumos", "N")]
    [InlineData("2026072", "Docs/Prueba Julio-2/Insumos", "M")]
    public void S5_OficialDirecto_EndToEnd_L9L13PorQuincena(string codigo, string carpetaRel, string columna)
    {
        var oficial = PlantillaUnica();
        var hashAntes = Sha256(oficial);
        var periodo = Periodo.Parse(codigo);
        var salidaDir = Path.Combine(Path.GetTempPath(), "p39-e2e-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, "salida.xlsx");
        try
        {
            new ProcesadorPeriodo(
                    new ExcelDataReaderRecaudoReader(),
                    new ExcelDataReaderWorkbookLeafInputReader(),
                    new CalculoRemuneracion(),
                    new ValidadorBasico(),
                    new OpenXmlPlantillaWriter(),
                    new ArchivoFuenteLocator(),
                    new ExcelDataReaderDetRetriR10Reader())
                .Ejecutar(new SolicitudProcesoPeriodo
                {
                    Periodo = periodo,
                    CarpetaPeriodo = Path.Combine(Raiz(), carpetaRel),
                    RutaPlantilla = oficial,
                    RutaSalida = salida
                });

            Assert.True(File.Exists(salida), $"El flujo 5-ASE {codigo} desde el OFICIAL debe certificar la salida.");

            using var documento = SpreadsheetDocument.Open(salida, false);
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var consolidado = Hoja(parte, HojaConsolidado);
            for (var i = 0; i < 5; i++)
            {
                Assert.Equal($"INTERVENTORIA!{columna}{26 + i}", FormulaDe(consolidado, $"L{9 + i}"));
            }

            Assert.Equal(codigo, ValorDe(consolidado, "N3"));
            Assert.Equal(hashAntes, Sha256(oficial));
        }
        finally
        {
            Borrar(salidaDir);
        }
    }

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (plantilla estandarizada única).");
    }

    private static string Sha256(string ruta) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(ruta)));

    private static Worksheet Hoja(WorkbookPart parte, string nombre)
    {
        var sheet = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!;
    }

    private static Cell? Celda(Worksheet hoja, string referencia) =>
        hoja.Descendants<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));

    private static string? FormulaDe(Worksheet hoja, string referencia) => Celda(hoja, referencia)?.CellFormula?.Text;

    private static string ValorDe(Worksheet hoja, string referencia) => Celda(hoja, referencia)?.CellValue?.InnerText?.Trim() ?? string.Empty;

    private static void FijarFormula(string ruta, string hojaNombre, string referencia, string formula)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = Hoja(parte, hojaNombre);
        var celda = Celda(hoja, referencia) ?? throw new InvalidOperationException($"Falta {hojaNombre}!{referencia}");
        celda.CellFormula = new CellFormula(formula);
        documento.Save();
    }

    private static void Borrar(string ruta)
    {
        try
        {
            if (Directory.Exists(ruta))
            {
                Directory.Delete(ruta, recursive: true);
            }
            else if (File.Exists(ruta))
            {
                File.Delete(ruta);
            }
        }
        catch
        {
            // best-effort
        }
    }
}
