using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-17 (S-2 HU-16): helpers OpenXML compartidos por los tests (antes duplicados en
/// <c>InterventoriaTests</c> y <c>GoldenInterventoriaTests</c> — ~90 líneas repetidas).
/// Solo tests; cero impacto en producción. Incluye el fix S-1 HU-16: el límite de filas de
/// <see cref="EnumerarCeldaLNumericas"/> se deriva de la DIMENSIÓN real de la hoja, no de un
/// número mágico (700), para que una fila legítima más allá de 700 (dentro de la dimensión)
/// nunca quede fuera del barrido A8 stale-guard.
/// </summary>
internal static class TestHelpers
{
    /// <summary>
    /// Enumera las celdas de la columna L con valor numérico (no fórmula, no texto) de la hoja.
    /// El tope de filas se deriva de <c>SheetDimension</c> (S-1): una fila fantasma fuera de la
    /// dimensión se ignora; una fila legítima dentro de la dimensión (aunque supere 700) se lee.
    /// </summary>
    public static IEnumerable<(string Celda, decimal Valor)> EnumerarCeldaLNumericas(string ruta, string hoja)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var workbookXml = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook null");
        var sheet = workbookXml.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet!;
        var limite = LimiteFilaDesdeDimension(ws);
        foreach (var row in ws.Descendants<Row>())
        {
            if (row.RowIndex is null || row.RowIndex.Value > limite)
            {
                continue;
            }

            foreach (var cell in row.Elements<Cell>())
            {
                var referencia = cell.CellReference?.Value ?? "";
                if (referencia.Length < 2 || char.ToUpperInvariant(referencia[0]) != 'L')
                {
                    continue;
                }

                if (cell.CellFormula is not null || cell.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
                {
                    continue;
                }

                // Celdas de texto (shared string / inline) NO son L-menores numéricas (ej. L5).
                if (cell.DataType is not null
                    && (cell.DataType.Value == CellValues.SharedString || cell.DataType.Value == CellValues.InlineString))
                {
                    continue;
                }

                if (decimal.TryParse(cell.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor))
                {
                    yield return (referencia, valor);
                }
            }
        }
    }

    /// <summary>
    /// S-1 HU-16 (HU-17): último número de fila declarado por <c>SheetDimension</c> (p. ej.
    /// "A1:N560" → 560). Sin dimensión → sin tope (enumera todo; Excel siempre la escribe).
    /// </summary>
    private static int LimiteFilaDesdeDimension(Worksheet ws)
    {
        var referencia = ws.SheetDimension?.Reference?.Value;
        if (string.IsNullOrWhiteSpace(referencia))
        {
            return int.MaxValue;
        }

        var ultima = referencia.Split(':').LastOrDefault();
        if (string.IsNullOrWhiteSpace(ultima))
        {
            return int.MaxValue;
        }

        var fila = new string(ultima.Where(char.IsAsciiDigit).ToArray());
        return int.TryParse(fila, NumberStyles.None, CultureInfo.InvariantCulture, out var numero) ? numero : int.MaxValue;
    }

    public static bool CeldaEsFormula(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var workbookXml = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook null");
        var sheet = workbookXml.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet!;
        var cell = ws.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        return cell?.CellFormula is not null;
    }

    public static decimal LeerCeldaNumerica(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var workbookXml = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook null");
        var sheet = workbookXml.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet!;
        var cell = ws.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        if (cell?.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
        {
            return 0m;
        }

        return decimal.TryParse(cell.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor)
            ? valor
            : 0m;
    }
}
