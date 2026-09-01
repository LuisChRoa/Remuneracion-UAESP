using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

string workbookPath = @"C:\Users\lechaparro.ASEINGES\OneDrive - kumo2\Proyectos\AAA\2026\08\Automatización\Docs\Insumos\Remuneracion 202607-1 Total.xlsx";
using var doc = SpreadsheetDocument.Open(workbookPath, false);
var workbook = doc.WorkbookPart!.Workbook;
var sheetNames = workbook.Sheets!.OfType<Sheet>().Select(s => s.Name!.Value).ToList();
foreach (var sheetName in sheetNames)
{
    if (sheetName.Contains("CONSOLIDADO") || sheetName.Contains("Reporte") || sheetName.Contains("Anticipos") || sheetName.Contains("Reversion") || sheetName.Contains("AJUSTES"))
    {
        Console.WriteLine("## " + sheetName);
        var sheet = workbook.Sheets!.OfType<Sheet>().First(s => s.Name!.Value == sheetName);
        var wsPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id!);
        var rows = wsPart.Worksheet.Descendants<Row>().Where(r => r.RowIndex != null && (r.RowIndex.Value >= 1 && r.RowIndex.Value <= 25 || r.RowIndex.Value >= 40 && r.RowIndex.Value <= 70)).ToList();
        foreach(var row in rows)
        {
            var values = new List<string>();
            foreach (var cell in row.Elements<Cell>())
            {
                string text = "";
                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                {
                    var idx = int.Parse(cell.CellValue!.InnerText);
                    var ss = doc.WorkbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(idx);
                    if (ss != null) text = string.Join("", ss.Elements<Text>().Select(t => t.Text));
                }
                else if (cell.CellValue != null)
                {
                    text = cell.CellValue.InnerText;
                }
                if (cell.CellFormula != null) text = "FORMULA:" + cell.CellFormula.Text;
                values.Add($"{cell.CellReference}:{text}");
            }
            Console.WriteLine(string.Join(" | ", values));
        }
    }
}
