using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

var path = @"C:\Users\lechaparro.ASEINGES\OneDrive - kumo2\Proyectos\AAA\2026\08\Automatización\Docs\Insumos\Remuneracion 202607-1 Total.xlsx";
using var doc = SpreadsheetDocument.Open(path, false);
var wb = doc.WorkbookPart!.Workbook;
Console.WriteLine("SHEETS:");
foreach (var sh in wb.Sheets!.Cast<Sheet>())
{
    Console.WriteLine($"- {sh.Name} (id={sh.Id})");
    var wsPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sh.Id!);
    var ws = wsPart.Worksheet;
    foreach (var row in ws.Descendants<Row>().Where(r => r.RowIndex is not null && (r.RowIndex.Value >= 1 && r.RowIndex.Value <= 20 || r.RowIndex.Value >= 40 && r.RowIndex.Value <= 60 || r.RowIndex.Value >= 90 && r.RowIndex.Value <= 110)))
    {
        var cells = row.Elements<Cell>().ToList();
        if (cells.Count == 0) continue;
        Console.WriteLine($" Row {row.RowIndex.Value}");
        foreach (var c in cells)
        {
            var v = c.CellValue?.InnerText ?? "";
            Console.WriteLine($"  {c.CellReference}: type={c.DataType} val={v} form={(c.CellFormula is null ? "" : c.CellFormula.Text)}");
        }
    }
}
