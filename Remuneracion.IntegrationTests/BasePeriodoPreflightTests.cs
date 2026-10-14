using System.IO.Compression;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
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

    // S8 (a): la base del período 2026092 pasa el preflight de producción (N3 + hojas + sellos).
    [Fact]
    public void S8_BaseDelPeriodo_PreflightVerde()
    {
        using var documento = SpreadsheetDocument.Open(PlantillaUnica(), false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");

        ValidadorBasePeriodo.Exigir(parte, Periodo.Parse("2026092"));
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
}
