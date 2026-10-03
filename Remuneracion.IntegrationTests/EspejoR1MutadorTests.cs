using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;
using Xunit.Abstractions;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 21 — PR 3 (T3): capacidad de escritura espejo R1 en <see cref="OpenXmlPlantillaWriter"/>.
///
/// Evidencia: <c>plans/21 - T0 Evidencia.md</c> §2 (conteos), §4 (fórmulas), §6 (invariantes).
/// Solo se usan insumos REALES de <c>Docs/Insumos</c> y <c>Docs/Prueba2/Insumos</c>: los goldens
/// de julio para el caso Δ=0 (bit-compatible) y la fuente de agosto contra la plantilla Q2 de
/// julio para el caso Δ≠0 (dimensionado + invariantes). NO se inventan fixtures sintéticas.
/// </summary>
public sealed class EspejoR1MutadorTests
{
    private static readonly int[] ConteoQ2 = [45, 75, 43, 73, 52];
    private static readonly int[] ConteoAgosto = [42, 66, 37, 79, 60];
    private static readonly string[] NombresAse = ["PROMOAMBIENTAL", "LIME", "CIUDAD LIMPIA", "BOGOTA LIMPIA", "AREA LIMPIA"];

    private readonly ITestOutputHelper _output;

    public EspejoR1MutadorTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public void Q2_Julio_DeltaCero_BloquesIntactosEInvariantes()
    {
        var bloques = LeerBloques(Insumos.R1Q2, Insumos.PeriodoQ2()).ToArray();
        var salida = SalidaTemporal("q2");
        try
        {
            new OpenXmlPlantillaWriter().EscribirEspejoR1(Insumos.GoldenQ2, salida, bloques);

            Assert.True(File.Exists(salida));
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var (nameRow, totalRow, filas) = MedirBloque(salida, NombresAse[aseId - 1]);
                Assert.True(nameRow > 0 && totalRow > nameRow, $"ASE {aseId}: anclas del bloque no encontradas.");
                Assert.Equal(ConteoQ2[aseId - 1], filas);
                Assert.True(bloques[aseId - 1].Filas.Count == filas);
            }
        }
        finally
        {
            Borrar(salida);
        }
    }

    [Fact]
    public void Agosto_DeltaDistinto_DimensionaBloquesAlaFuenteReal()
    {
        var bloques = LeerBloques(AgostoR1, new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 }).ToArray();
        var salida = SalidaTemporal("agosto");
        try
        {
            new OpenXmlPlantillaWriter().EscribirEspejoR1(Insumos.GoldenQ2, salida, bloques);

            Assert.True(File.Exists(salida));
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var (nameRow, totalRow, filas) = MedirBloque(salida, NombresAse[aseId - 1]);
                Assert.True(nameRow > 0 && totalRow > nameRow, $"ASE {aseId}: anclas del bloque no encontradas.");
                Assert.Equal(ConteoAgosto[aseId - 1], filas);
                Assert.Equal(bloques[aseId - 1].Filas.Count, filas);
            }
        }
        finally
        {
            Borrar(salida);
        }
    }

    /// <summary>
    /// Reanclaje: la mutación (Δ≠0) inserta/borra filas de datos y reancla referencias, pero NUNCA
    /// agrega ni quita celdas-fórmula del bloque (las visibles viven debajo del bloque y solo se
    /// desplazan). W-2 (auditoría PR3): conteo exacto preservado, no una cota arbitraria.
    /// </summary>
    [Fact]
    public void Agosto_FormulasPreservadasTrasMutacion()
    {
        var bloques = LeerBloques(AgostoR1, new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 }).ToArray();
        var antes = ContarFormulas(Insumos.GoldenQ2, "Reporte Componentes R1");
        var salida = SalidaTemporal("agosto-formulas");
        try
        {
            new OpenXmlPlantillaWriter().EscribirEspejoR1(Insumos.GoldenQ2, salida, bloques);
            var despues = ContarFormulas(salida, "Reporte Componentes R1");
            Assert.Equal(antes, despues); // el dimensionado no crea ni destruye fórmulas
        }
        finally
        {
            Borrar(salida);
        }
    }

    /// <summary>
    /// Plan 21 (W-2, auditoría PR3, R-E-3): evidencia ESTRUCTURAL del reanclaje contra la
    /// expectativa derivada de T0a/T0c. La plantilla Q2 (45/75/43/73/52) y la fuente de agosto
    /// (42/66/37/79/60) fijan Δ = -3/-9/-6/+6/+8 por bloque. Se aserta (a) el Δ real por bloque,
    /// (b) que cada bloque con Δ≠0 reancló referencias (contador > 0) y (c) que ninguna referencia
    /// quedó en fila equivocada: CONSOLIDADO D9 (golden R1!F53) y D13 (golden R1!F558) deben
    /// apuntar a F50 y F554, el resultado del shift acumulado 5→1 con los totalRows Q2
    /// (50/162/283/450/554).
    /// </summary>
    [Fact]
    public void Agosto_Reanclaje_DeltasYReferenciasPorBloque_ContraExpectativaT0a()
    {
        var bloques = LeerBloques(AgostoR1, new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 }).ToArray();
        var salida = SalidaTemporal("agosto-reanclaje");
        try
        {
            var resultado = new OpenXmlPlantillaWriter().EscribirEspejoR1ConResultado(Insumos.GoldenQ2, salida, bloques);

            var deltasEsperados = new Dictionary<int, int> { [1] = -3, [2] = -9, [3] = -6, [4] = +6, [5] = +8 };
            foreach (var (aseId, delta) in deltasEsperados)
            {
                Assert.Equal(delta, resultado.DeltasPorBloque[aseId]);
                Assert.True(
                    resultado.ReferenciasReancladasPorBloque[aseId] > 0,
                    $"ASE {aseId} (Δ={delta}): el reanclaje debe mover referencias; contador={resultado.ReferenciasReancladasPorBloque[aseId]}.");
            }

            // Invariante del contador: total = suma por bloque.
            Assert.Equal(resultado.ReferenciasReancladasPorBloque.Values.Sum(), resultado.TotalReferenciasReancladas);

            _output.WriteLine("Reanclaje por bloque: " + string.Join(
                ", ",
                resultado.ReferenciasReancladasPorBloque.OrderBy(kv => kv.Key)
                    .Select(kv => $"ASE{kv.Key}(Δ={resultado.DeltasPorBloque[kv.Key]:+0;-0;0},refs={kv.Value})"))
                + $" total={resultado.TotalReferenciasReancladas}");

            // No-wrong-row (spot-check derivado de T0a): shift neto -3 para el visible ASE1 y -4
            // para el visible ASE5 (Σ deltas acumulados afectando cada fila).
            AssertReferenciaR1(salida, "D9", "F50");
            AssertReferenciaR1(salida, "D13", "F554");
        }
        finally
        {
            Borrar(salida);
        }
    }

    /// <summary>
    /// W-2: lee la fórmula de la celda de CONSOLIDADO y verifica que su referencia a la hoja R1
    /// apunte a la fila esperada (el token de fila reanclado por el espejo).
    /// </summary>
    private static void AssertReferenciaR1(string ruta, string celda, string filaEsperada)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, WorkbookLeafCellMap.HojaConsolidado);
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        Assert.NotNull(cell);
        var formula = cell!.CellFormula?.Text ?? string.Empty;
        Assert.Contains($"'Reporte Componentes R1'!{filaEsperada}", formula, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Plan 21 (T3, R-E-2/R-E-3/R-E-4): con la fuente de julio y el bloque ya dimensionado (Δ=0),
    /// el espejo debe reproducir EXACTAMENTE (dif ±0.5) el bloque R1 del golden: etiquetas A–E,
    /// valores por columna y fila Total final. El origen es el propio golden, de modo que cualquier
    /// divergencia de escritura queda aislada en las celdas que el espejo reescribe.
    /// </summary>
    [Fact]
    public void Q2_Julio_DeltaCero_ValoresEspejoVsGolden_Tolerancia()
    {
        var bloques = LeerBloques(Insumos.R1Q2, Insumos.PeriodoQ2()).ToArray();
        var salida = SalidaTemporal("q2-valores");
        try
        {
            new OpenXmlPlantillaWriter().EscribirEspejoR1(Insumos.GoldenQ2, salida, bloques);
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                CompararBloqueContraGolden(salida, Insumos.GoldenQ2, NombresAse[aseId - 1], aseId);
            }
        }
        finally
        {
            Borrar(salida);
        }
    }

    /// <summary>
    /// Plan 21 (T3, R-E-2/R-E-3/R-E-4): idéntico contraste para la quincena 1 (golden canónico Q1).
    /// </summary>
    [Fact]
    public void Q1_Julio_DeltaCero_ValoresEspejoVsGolden_Tolerancia()
    {
        var bloques = LeerBloques(Insumos.R1, Insumos.Periodo()).ToArray();
        var salida = SalidaTemporal("q1-valores");
        try
        {
            new OpenXmlPlantillaWriter().EscribirEspejoR1(Insumos.Plantilla, salida, bloques);
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                CompararBloqueContraGolden(salida, Insumos.Plantilla, NombresAse[aseId - 1], aseId);
            }
        }
        finally
        {
            Borrar(salida);
        }
    }

    private static void CompararBloqueContraGolden(string producido, string goldenRuta, string nombreAse, int aseId)
    {
        using var wbProducido = SpreadsheetDocument.Open(producido, false);
        using var wbGolden = SpreadsheetDocument.Open(goldenRuta, false);
        var pp = wbProducido.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null (producido)");
        var gp = wbGolden.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null (golden)");
        var wsP = Hoja(pp, "Reporte Componentes R1");
        var wsG = Hoja(gp, "Reporte Componentes R1");

        var (nameRow, totalRow) = RangoBloque(gp, wsG, nombreAse);
        Assert.True(nameRow > 0 && totalRow > nameRow, $"ASE {aseId}: no se localizó el bloque en el golden.");

        for (var r = nameRow; r <= totalRow; r++)
        {
            var filaP = FilaPorIndice(wsP, r);
            var filaG = FilaPorIndice(wsG, r);
            var referencias = ReferenciasDeFila(filaP, filaG, r);
            foreach (var referencia in referencias)
            {
                var vp = ValorCelda(pp, filaP, referencia);
                var vg = ValorCelda(gp, filaG, referencia);
                CompararValores(vp, vg, aseId, referencia);
            }
        }
    }

    private static (int NameRow, int TotalRow) RangoBloque(WorkbookPart workbookPart, Worksheet hoja, string nombreAse)
    {
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

    private static Row? FilaPorIndice(Worksheet hoja, int indice) =>
        hoja.Descendants<Row>().FirstOrDefault(r => (int)(r.RowIndex?.Value ?? 0) == indice);

    private static IEnumerable<string> ReferenciasDeFila(Row? filaP, Row? filaG, int indice)
    {
        var sufijo = indice.ToString(CultureInfo.InvariantCulture);
        var referencias = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var celda in (filaP?.Elements<Cell>() ?? Enumerable.Empty<Cell>()).Concat(filaG?.Elements<Cell>() ?? Enumerable.Empty<Cell>()))
        {
            var referencia = celda.CellReference?.Value;
            if (!string.IsNullOrWhiteSpace(referencia) && referencia.EndsWith(sufijo, StringComparison.Ordinal))
            {
                referencias.Add(referencia);
            }
        }

        return referencias;
    }

    private static (decimal? Numero, string? Texto) ValorCelda(WorkbookPart workbookPart, Row? fila, string referencia)
    {
        var celda = fila?.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        if (celda is null)
        {
            return (null, null);
        }

        if (celda.InlineString is not null)
        {
            return (null, celda.InlineString.InnerText ?? string.Empty);
        }

        if (celda.DataType?.Value == CellValues.SharedString && celda.CellValue is not null
            && int.TryParse(celda.CellValue.Text, out var indice))
        {
            return (null, workbookPart.SharedStringTablePart?.SharedStringTable?.ElementAt(indice).InnerText ?? string.Empty);
        }

        if (celda.CellValue is null || string.IsNullOrWhiteSpace(celda.CellValue.InnerText))
        {
            return (null, null);
        }

        return decimal.TryParse(celda.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var numero)
            ? (numero, null)
            : (null, celda.CellValue.InnerText);
    }

    private static void CompararValores((decimal? Numero, string? Texto) espejo, (decimal? Numero, string? Texto) golden, int aseId, string referencia)
    {
        var pEsNumero = espejo.Numero is not null;
        var gEsNumero = golden.Numero is not null;
        var pVacio = !pEsNumero && string.IsNullOrEmpty(espejo.Texto);
        var gVacio = !gEsNumero && string.IsNullOrEmpty(golden.Texto);

        if (pVacio && gVacio)
        {
            return;
        }

        if (pEsNumero && gEsNumero)
        {
            Assert.True(
                Math.Abs(espejo.Numero!.Value - golden.Numero!.Value) <= Insumos.Tolerancia,
                $"ASE {aseId} Reporte Componentes R1!{referencia}: espejo={espejo.Numero} golden={golden.Numero} (dif > {Insumos.Tolerancia}).");
            return;
        }

        if (!pEsNumero && !gEsNumero && !pVacio && !gVacio)
        {
            Assert.True(
                string.Equals(espejo.Texto, golden.Texto, StringComparison.Ordinal),
                $"ASE {aseId} Reporte Componentes R1!{referencia}: espejo='{espejo.Texto}' golden='{golden.Texto}'.");
            return;
        }

        Assert.Fail(
            $"ASE {aseId} Reporte Componentes R1!{referencia}: tipo distinto espejo=(num={espejo.Numero},texto='{espejo.Texto}') golden=(num={golden.Numero},texto='{golden.Texto}').");
    }

    private static IEnumerable<BloqueEspejoAseInputs> LeerBloques(Func<int, string> rutaPorAse, Periodo periodo)
    {
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            yield return reader.LeerEspejoR1(Insumos.Ase(aseId), rutaPorAse(aseId));
        }
    }

    private static (int NameRow, int TotalRow, int Filas) MedirBloque(string ruta, string nombreAse)
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

        var conteo = 0;
        foreach (var fila in filas)
        {
            var idx = (int)(fila.RowIndex?.Value ?? 0);
            if (idx <= nameRow || idx > totalRow)
            {
                continue;
            }

            if (new[] { "A", "B", "C", "D", "E" }.Any(c => !string.IsNullOrWhiteSpace(Texto(workbookPart, fila, c))))
            {
                conteo++;
            }
        }

        return (nameRow, totalRow, conteo);
    }

    private static int ContarFormulas(string ruta, string hojaNombre)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        return Hoja(workbookPart, hojaNombre).Descendants<Cell>().Count(c => c.CellFormula is not null);
    }

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet!;
    }

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

    private static string AgostoR1(int aseId)
    {
        var carpeta = Path.Combine(Insumos.Raiz(), "Docs", "Prueba2", "Insumos");
        var dir = Directory.EnumerateDirectories(carpeta)
            .First(d => Path.GetFileName(d).StartsWith($"{aseId}-", StringComparison.OrdinalIgnoreCase));
        return Directory.EnumerateFiles(dir, "Recaudoporcomponente*.xlsx", SearchOption.TopDirectoryOnly).First();
    }

    private static string SalidaTemporal(string etiqueta) =>
        Path.Combine(Path.GetTempPath(), $"remuneracion-espejo21-{etiqueta}-{Guid.NewGuid():N}.xlsx");

    private static void Borrar(string ruta)
    {
        try
        {
            if (File.Exists(ruta))
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
