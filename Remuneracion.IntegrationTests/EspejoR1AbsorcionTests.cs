using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 21 — PR 4 (T5): EVIDENCIA de absorción del path legado rol/ocurrencia
/// (<c>LeerLEspecialesMenores</c> + <c>WorkbookLeafCellMapInterventoria.LMenoresPorAse/Q2</c>)
/// por el espejo estructural R1, en los 15 escenarios (5 ASE × Q1-julio / Q2-julio / agosto).
///
/// El mapa legado escribía celdas L del template resueltas por rol+ocurrencia sobre la fuente R1.
/// La columna L del template tiene encabezado <c>SERVICIO ESPECIALES</c> (verificado en disco: fila
/// de headers del bloque). El espejo escribe esa misma columna por ENCABEZADO en cada fila de la
/// secuencia fuente. La absorción es total si:
///   (a) julio: toda celda L del mapa legado cae DENTRO del bloque espejo del template destino
///       (rango [nameRow, totalRow]) — y el flujo real con espejo ya cierra contra golden ±0.5
///       (red de <see cref="GoldenInterventoriaTests"/>);
///   (b) agosto: el espejo escribe la columna L de TODA la secuencia fuente del período (no queda
///       ninguna fila del bloque con L residual del template).
///
/// Solo insumos REALES de <c>Docs/Insumos</c> y <c>Docs/Prueba2/Insumos</c>; no hay fixtures
/// sintéticas ni golden inventado de agosto.
/// </summary>
public sealed class EspejoR1AbsorcionTests
{
    private static readonly string[] NombresAse = ["PROMOAMBIENTAL", "LIME", "CIUDAD LIMPIA", "BOGOTA LIMPIA", "AREA LIMPIA"];
    private static readonly int[] ConteoAgosto = [42, 66, 37, 79, 60];

    /// <summary>
    /// (a) Julio (Q1 + Q2): toda celda L del mapa legado del período cae dentro del bloque espejo
    /// del template destino (misma fila está en [nameRow, totalRow]). Si alguna cayera fuera, el
    /// espejo no la tocaría y el retiro sería incorrecto.
    /// </summary>
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void Julio_CeldasDelMapaLegado_CaenDentroDelBloqueEspejo(int quincena)
    {
        var template = quincena == 2 ? Insumos.PlantillaQ2 : Insumos.Plantilla;
        var mapa = MapaLMenoresEvidenciaT0.PorQuincena(quincena);

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var (nameRow, totalRow) = RangoBloque(template, NombresAse[aseId - 1]);
            Assert.True(nameRow > 0 && totalRow > nameRow, $"ASE {aseId}: no se localizó el bloque en el template.");

            foreach (var celda in mapa[aseId])
            {
                Assert.StartsWith("L", celda, StringComparison.OrdinalIgnoreCase);
                var fila = FilaDe(celda);
                Assert.True(
                    fila >= nameRow && fila <= totalRow,
                    $"Q{quincena} ASE {aseId}: la celda L del mapa legado {celda} (fila {fila}) queda FUERA del bloque espejo [{nameRow}, {totalRow}] — no absorbida.");
            }
        }
    }

    /// <summary>
    /// (b) Agosto-Prueba2: el espejo gobierna la columna L de TODA la secuencia fuente del período.
    /// Se aplica el espejo de agosto a la plantilla Q2 y, para cada ASE, se contrasta la celda L de
    /// cada fila del bloque resultante contra el valor <c>SERVICIO ESPECIALES</c> de la fila fuente
    /// correspondiente (por posición de secuencia). Ninguna fila conserva L residual del template.
    /// </summary>
    [Fact]
    public void Agosto_EspejoGobiernaLaColumnaLDeTodaLaSecuencia()
    {
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var bloques = new List<BloqueEspejoAseInputs>();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var bloque = reader.LeerEspejoR1(Insumos.Ase(aseId), AgostoR1(aseId));
            Assert.Equal(ConteoAgosto[aseId - 1], bloque.TotalFilas);
            bloques.Add(bloque);
        }

        var salida = Path.Combine(Path.GetTempPath(), $"remuneracion-espejo21-absorcion-{Guid.NewGuid():N}.xlsx");
        try
        {
            new OpenXmlPlantillaWriter().EscribirEspejoR1(Insumos.GoldenQ2, salida, bloques);

            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var bloque = bloques[aseId - 1];
                var filasL = LeerColumnaLDelBloque(salida, NombresAse[aseId - 1]);
                Assert.Equal(bloque.Filas.Count, filasL.Count);

                for (var i = 0; i < bloque.Filas.Count; i++)
                {
                    var esperado = bloque.Filas[i].Valor("SERVICIO ESPECIALES");
                    var escrito = filasL[i];
                    Assert.True(
                        IgualesTolerancia(escrito, esperado),
                        $"Agosto ASE {aseId}: la celda L de la fila {i} del bloque = {escrito?.ToString(CultureInfo.InvariantCulture) ?? "vacío"}, y la fuente trae SERVICIO ESPECIALES = {esperado?.ToString(CultureInfo.InvariantCulture) ?? "vacío"} (el espejo debe gobernar TODA la columna L).");
                }
            }
        }
        finally
        {
            try
            {
                if (File.Exists(salida))
                {
                    File.Delete(salida);
                }
            }
            catch
            {
                // best-effort
            }
        }
    }

    private static string AgostoR1(int aseId)
    {
        var carpeta = Path.Combine(Insumos.Raiz(), "Docs", "Prueba2", "Insumos");
        var dir = Directory.EnumerateDirectories(carpeta)
            .First(d => Path.GetFileName(d).StartsWith($"{aseId}-", StringComparison.OrdinalIgnoreCase));
        return Directory.EnumerateFiles(dir, "Recaudoporcomponente*.xlsx", SearchOption.TopDirectoryOnly).First();
    }

    /// <summary>
    /// Rango [nameRow, totalRow] del bloque del ASE en el template (misma delimitación del espejo:
    /// fila de nombre B=&lt;ASE&gt; → primera fila Total final A='Total' B vacío).
    /// </summary>
    private static (int NameRow, int TotalRow) RangoBloque(string ruta, string nombreAse)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = Hoja(workbookPart, "Reporte Componentes R1");
        var filas = hoja.Descendants<Row>().OrderBy(r => r.RowIndex?.Value ?? 0).ToList();

        var nameRow = 0;
        var totalRow = 0;
        foreach (var fila in filas)
        {
            var b = Normalizar(Texto(workbookPart, fila, "B"));
            if (nameRow == 0 && b == Normalizar(nombreAse))
            {
                nameRow = (int)(fila.RowIndex?.Value ?? 0);
                continue;
            }

            if (nameRow > 0 && EsTotalFinal(workbookPart, fila))
            {
                totalRow = (int)(fila.RowIndex?.Value ?? 0);
                break;
            }
        }

        return (nameRow, totalRow);
    }

    /// <summary>
    /// Columna L de las filas del bloque del ASE en la salida, en orden de fila (solo filas con
    /// firma A–E no vacía, igual criterio que el espejo).
    /// </summary>
    private static List<decimal?> LeerColumnaLDelBloque(string ruta, string nombreAse)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = Hoja(workbookPart, "Reporte Componentes R1");
        var filas = hoja.Descendants<Row>().OrderBy(r => r.RowIndex?.Value ?? 0).ToList();

        var nameRow = 0;
        var totalRow = 0;
        foreach (var fila in filas)
        {
            var b = Normalizar(Texto(workbookPart, fila, "B"));
            if (nameRow == 0 && b == Normalizar(nombreAse))
            {
                nameRow = (int)(fila.RowIndex?.Value ?? 0);
                continue;
            }

            if (nameRow > 0 && EsTotalFinal(workbookPart, fila))
            {
                totalRow = (int)(fila.RowIndex?.Value ?? 0);
                break;
            }
        }

        var resultado = new List<decimal?>();
        foreach (var fila in filas)
        {
            var idx = (int)(fila.RowIndex?.Value ?? 0);
            if (idx <= nameRow || idx > totalRow)
            {
                continue;
            }

            if (!new[] { "A", "B", "C", "D", "E" }.Any(c => !string.IsNullOrWhiteSpace(Texto(workbookPart, fila, c))))
            {
                continue;
            }

            resultado.Add(ValorNumerico(workbookPart, fila, "L"));
        }

        return resultado;
    }

    private static decimal? ValorNumerico(WorkbookPart workbookPart, Row fila, string columna)
    {
        var indice = ((uint?)fila.RowIndex?.Value ?? 0).ToString(CultureInfo.InvariantCulture);
        var celda = fila.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, columna + indice, StringComparison.OrdinalIgnoreCase));
        if (celda?.CellValue is null || celda.CellFormula is not null)
        {
            return null;
        }

        return decimal.TryParse(celda.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var numero)
            ? numero
            : null;
    }

    private static bool IgualesTolerancia(decimal? a, decimal? b)
    {
        if (a is null && b is null)
        {
            return true;
        }

        var va = a ?? 0m;
        var vb = b ?? 0m;
        return Math.Abs(va - vb) <= Insumos.Tolerancia;
    }

    private static int FilaDe(string celda) =>
        int.Parse(new string(celda.Where(char.IsAsciiDigit).ToArray()), CultureInfo.InvariantCulture);

    private static bool EsTotalFinal(WorkbookPart workbookPart, Row fila) =>
        string.Equals(Texto(workbookPart, fila, "A"), "Total", StringComparison.OrdinalIgnoreCase)
        && string.IsNullOrWhiteSpace(Texto(workbookPart, fila, "B"));

    private static string Texto(WorkbookPart workbookPart, Row fila, string columna)
    {
        var indice = ((uint?)fila.RowIndex?.Value ?? 0).ToString(CultureInfo.InvariantCulture);
        var celda = fila.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, columna + indice, StringComparison.OrdinalIgnoreCase));
        if (celda is null)
        {
            return string.Empty;
        }

        if (celda.InlineString is not null)
        {
            return celda.InlineString.InnerText ?? string.Empty;
        }

        if (celda.DataType?.Value == CellValues.SharedString && celda.CellValue is not null
            && int.TryParse(celda.CellValue.Text, out var i))
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable?.ElementAt(i).InnerText ?? string.Empty;
        }

        return celda.CellValue?.InnerText ?? string.Empty;
    }

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet!;
    }

    private static string Normalizar(string texto)
    {
        var d = texto.Normalize(System.Text.NormalizationForm.FormD);
        var sb = new System.Text.StringBuilder();
        foreach (var ch in d)
        {
            if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch) != System.Globalization.UnicodeCategory.NonSpacingMark)
            {
                sb.Append(char.ToUpperInvariant(ch));
            }
        }

        return sb.ToString();
    }
}
