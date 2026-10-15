using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 41 (T3, R-S31-1 — D-G): regla honesta de <c>INTERVENTORIA!S26:S31</c> (auto-chequeo de la
/// tabla J). El oficial trae el maestro <c>R26 = J15</c> (la tabla anual = los valores de la 2.ª
/// quincena), mientras <c>Q26:Q31 = 'CONSOLIDADO_TOTAL RECAUDO'!L104:L109</c> sigue la selección
/// <c>L9:L13</c> por quincena (R-INTERV-Q). En Q1, <c>Q26 = N26 = 189185987 ≠ R26 = J15 = 189185988</c>
/// ⇒ <c>S26</c> FALSE real. Phase 0 reescribe el maestro por quincena (Q2 → <c>J15</c>; Q1 → <c>N26</c>)
/// — la columna la decide el DOMINIO, nunca el contenido. Prohibido sellar TRUE cosmético.
/// </summary>
public sealed class TablaJS26S31T3Tests
{
    private const string PartWorkbook = "xl/workbook.xml";
    private static readonly Regex CeldaRegex = new(
        "<(?<p>[A-Za-z_][A-Za-z0-9_]*:)?c r=\"(?<ref>[A-Z]+\\d+)\"(?<attrs>[^>]*?)(?:/>|>(?<inner>.*?)</(?:[A-Za-z_][A-Za-z0-9_]*:)?c>)",
        RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex FormulaRegex = new(
        "<(?:[A-Za-z_][A-Za-z0-9_]*:)?f(?<attrs>[^>]*?)>(?<texto>[^<]*)</(?:[A-Za-z_][A-Za-z0-9_]*:)?f>",
        RegexOptions.Singleline | RegexOptions.Compiled);

    private static string PlantillaOficial() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    [Theory]
    [InlineData("2026071", "N26")]
    [InlineData("2026072", "J15")]
    public void T3_Phase0_ReescribeTablaJPorQuincena(string codigo, string maestroEsperado)
    {
        var periodo = Periodo.Parse(codigo);
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "p41-t3-" + Guid.NewGuid().ToString("N"))).FullName;
        var salida = Path.Combine(dir, periodo.NombreArchivo);
        try
        {
            PreparadorBasePipeline.Ejecutar(PlantillaOficial(), periodo, salida);

            using var zip = ZipFile.OpenRead(salida);
            var interv = LeerHoja(zip, "INTERVENTORIA");
            Assert.Equal(maestroEsperado, FormulaDe(interv, "R26"));

            // Preflight de producción verde (incluye la tabla J de la quincena del run).
            using var documento = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(salida, false);
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            ValidadorBasePeriodo.Exigir(parte, periodo);
        }
        finally
        {
            Borrar(dir);
        }
    }

    [Fact]
    public void T3_TablaJDeOtraQuincena_FailFastAntesDeEscribir()
    {
        // Una base Q1 con el maestro R26 de la tabla J en la variante Q2 (J15) → ERR-PLANTILLA que
        // nombra R26 y la variante esperada (N26). Nunca se escribe sobre una base con la tabla de la
        // otra quincena.
        var periodo = Periodo.Parse("2026071");
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "p41-t3-bad-" + Guid.NewGuid().ToString("N"))).FullName;
        var salida = Path.Combine(dir, periodo.NombreArchivo);
        try
        {
            PreparadorBasePipeline.Ejecutar(PlantillaOficial(), periodo, salida);
            FijarFormula(salida, "INTERVENTORIA", "R26", "J15");

            using var documento = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(salida, false);
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var ex = Assert.Throws<CalculoInvalidoException>(() => ValidadorBasePeriodo.Exigir(parte, periodo));

            Assert.Contains("R26", ex.Message, StringComparison.Ordinal);
            Assert.Contains("N26", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            Borrar(dir);
        }
    }

    [Fact]
    public void T3_Q1Honesto_NoSeSellaTrueCosmetico()
    {
        // S26 = (Q26 = R26). Q26 sigue 'CONSOLIDADO_TOTAL RECAUDO'!L104 → L9 = INTERVENTORIA!N26 en
        // Q1. El oficial (sin preparar) trae R26 = J15 (Q2) → S26 sería FALSE. Phase 0 lo reescribe a
        // N26 (la tabla de Q1) — regla por DOMINIO, no sellado: el texto-<f> cambia, el <v> del
        // chequeo no se fuerza.
        var periodo = Periodo.Parse("2026071");
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "p41-t3-hon-" + Guid.NewGuid().ToString("N"))).FullName;
        var salida = Path.Combine(dir, periodo.NombreArchivo);
        try
        {
            PreparadorBasePipeline.Ejecutar(PlantillaOficial(), periodo, salida);

            using var zip = ZipFile.OpenRead(salida);

            // La forma del chequeo es invariante (Q=R); solo el maestro R26 cambia de columna.
            var interv = LeerHoja(zip, "INTERVENTORIA");
            Assert.Equal("Q26=R26", FormulaDe(interv, "S26"));
            Assert.Equal("N26", FormulaDe(interv, "R26"));

            // Q26 referencia el CONSOLIDADO!L104 (que en Q1 sigue L9 = INTERVENTORIA!N26).
            var consolidado = LeerHoja(zip, "CONSOLIDADO_TOTAL RECAUDO");
            Assert.Equal("INTERVENTORIA!N26", FormulaDe(consolidado, "L9"));
        }
        finally
        {
            Borrar(dir);
        }
    }

    // ── BCL helpers ──────────────────────────────────────────────────────────────────────────────

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

    private static string LeerHoja(ZipArchive zip, string nombreHoja)
    {
        var workbook = LeerParte(zip, PartWorkbook);
        var rels = LeerParte(zip, "xl/_rels/workbook.xml.rels");
        var sheet = Regex.Match(workbook, $"<(?:[A-Za-z_][A-Za-z0-9_]*:)?sheet name=\"{Regex.Escape(nombreHoja)}\"[^>]*r:id=\"(?<r>[^\"]+)\"");
        Assert.True(sheet.Success, $"La plantilla no trae la hoja '{nombreHoja}'.");
        var rid = sheet.Groups["r"].Value;
        var rel = Regex.Match(rels, $"<Relationship[^>]*Id=\"{Regex.Escape(rid)}\"[^>]*Target=\"(?<t>[^\"]+)\"");
        Assert.True(rel.Success, $"No se resolvió '{nombreHoja}'.");
        var target = rel.Groups["t"].Value;
        return LeerParte(zip, target.StartsWith('/') ? target[1..] : "xl/" + target);
    }

    private static string LeerParte(ZipArchive zip, string nombre)
    {
        var entrada = zip.GetEntry(nombre) ?? throw new InvalidOperationException($"Falta {nombre} en el zip.");
        using var stream = entrada.Open();
        using var lector = new StreamReader(stream, Encoding.UTF8);
        return lector.ReadToEnd();
    }

    private static string FormulaDe(string xml, string celda)
    {
        foreach (Match match in CeldaRegex.Matches(xml))
        {
            if (!string.Equals(match.Groups["ref"].Value, celda, StringComparison.Ordinal))
            {
                continue;
            }

            if (!match.Groups["inner"].Success)
            {
                return string.Empty;
            }

            var f = FormulaRegex.Match(match.Groups["inner"].Value);
            return f.Success ? f.Groups["texto"].Value : string.Empty;
        }

        return string.Empty;
    }

    private static void Borrar(string ruta)
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

    private static void FijarFormula(string ruta, string hojaNombre, string referencia, string formula)
    {
        using var documento = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = parte.Workbook!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .First(s => string.Equals(s.Name?.Value, hojaNombre, StringComparison.OrdinalIgnoreCase));
        var ws = ((DocumentFormat.OpenXml.Packaging.WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!;
        var celda = ws.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
            .First(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        celda.CellFormula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula(formula) { FormulaType = DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Shared, SharedIndex = 0, Reference = "R26:R31" };
        ws.Save();
        documento.Save();
    }
}
