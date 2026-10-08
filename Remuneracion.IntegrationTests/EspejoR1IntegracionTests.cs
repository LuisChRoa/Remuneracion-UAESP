using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 21 — PR 3 (T4): integración del espejo estructural R1.
///
/// El caso de regresión 2026082 (agosto) no tiene golden ni plantilla propia y
/// <c>Docs/Prueba2/Insumos</c> no trae <c>Conciliaciones/</c> (requeridas por el flujo completo de
/// período), por lo que el smoke es ESTRUCTURAL sobre las capas del espejo: la lectura
/// <c>LeerEspejoR1</c> + la escritura <c>EscribirEspejoR1</c> resuelven el defecto que rompía
/// <c>LeerLEspecialesMenores</c> (slot <c>TotalD_E</c> ocurrencia 2 ausente en ASE1) SIN
/// <c>ERR-VALIDACION</c>. Toda la evidencia sale de insumos reales; no se inventa salida esperada.
/// </summary>
public sealed class EspejoR1IntegracionTests
{
    private static readonly int[] ConteoAgosto = [42, 66, 37, 79, 60];

    [Fact]
    public void Agosto_ElDefectoDelMapaRolOcurrenciaSeResuelveConElEspejo()
    {
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var ase1 = Insumos.Ase(1);
        var r1Ase1 = AgostoR1(1);

        // Causa raíz documentada en T0 §1: el mapa congelado de julio exigía TotalD_E ocurrencia 2
        // (L20) que la fuente de agosto no trae → ERR-VALIDACION (CalculoInvalidoException). Ese
        // path legado se retiró en T5 (absorción 15/15, EspejoR1AbsorcionTests); ya no existe.
        // El lector espejo NO exige cardinalidades: la secuencia observada ES la especificación.
        var bloque = reader.LeerEspejoR1(ase1, r1Ase1);
        Assert.Equal(42, bloque.TotalFilas);
        Assert.True(bloque.TieneInvariantesDeCierre, "Agosto ASE1 debe conservar las 3 invariantes duras T0e.");
    }

    [Fact]
    public void Agosto_LecturaYEscrituraEspejoDeLos5Ase_SinErrValidacion()
    {
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var bloques = new List<BloqueEspejoAseInputs>();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var bloque = reader.LeerEspejoR1(Insumos.Ase(aseId), AgostoR1(aseId));
            Assert.Equal(ConteoAgosto[aseId - 1], bloque.TotalFilas);
            Assert.True(bloque.TieneInvariantesDeCierre, $"Agosto ASE {aseId}: faltan invariantes de cierre T0e.");
            bloques.Add(bloque);
        }

        // Dimensionado + reanclaje + escritura por encabezado sobre la plantilla Q2 real de julio:
        // no debe lanzar ERR-VALIDACION ni ERR-PLANTILLA espurio.
        var salida = Path.Combine(Path.GetTempPath(), $"remuneracion-espejo21-t4-{Guid.NewGuid():N}.xlsx");
        try
        {
            new OpenXmlPlantillaWriter().EscribirEspejoR1(Insumos.GoldenQ2, salida, bloques);
            Assert.True(File.Exists(salida));
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

    /// <summary>
    /// CRITICAL #1 (auditoría PR3): el flujo COMPLETO 5-ASE con espejo activo y Δ≠0 no debe
    /// recomputar el gate de desplazamiento DESPUÉS de que <c>AjustarEnWorkbook</c> ya mutó el
    /// workbook (recomputado devuelve false y dispara la validación por mapas absolutos de julio
    /// → <c>ERR-PLANTILLA</c> falso garantizado con Δ≠0). Reproducción con insumos reales:
    /// leafs Q2 reales + espejo de AGOSTO (42/66/37/79/60 vs plantilla Q2 45/75/43/73/52) → los
    /// 5 bloques tienen Δ≠0. Antes del fix este test falla con <c>ERR-PLANTILLA</c> espurio;
    /// después, la salida se certifica.
    /// </summary>
    [Fact]
    public void GenerarWorkbook_EspejoAgosto_DeltaDistinto_NoEmiteErrPlantillaEspurio()
    {
        var (datos, _, leafs) = AjustesSfTTests.LeerDatosQ2ConAjustes();
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        foreach (var leaf in leafs)
        {
            leaf.EspejoR1 = reader.LeerEspejoR1(leaf.Ase, AgostoR1(leaf.Ase.Id));
            leaf.DetRetriQ2 = new DetRetriQ2Inputs
            {
                Ase = leaf.Ase,
                TotalD104 = leaf.R1.TotalOportunoEsperadoPorAse
                    + leaf.R2.TotalOportunoEsperado
                    + leaf.R1.ExtemporaneoEsperadoPorAse
                    + leaf.R4.TotalReversionEsperada
                    + (leaf.AjustesSfT?.TotalAjustes ?? 0m)
            };
        }

        var periodo = Insumos.PeriodoQ2();
        var resultado = new CalculoRemuneracion().CalcularConsolidado(periodo, datos);

        var salida = Path.Combine(Path.GetTempPath(), $"remuneracion-espejo21-flujo-{Guid.NewGuid():N}.xlsx");
        try
        {
            new OpenXmlPlantillaWriter().GenerarWorkbook(Insumos.PlantillaQ2, salida, resultado, leafs);
            Assert.True(File.Exists(salida), "El flujo 5-ASE con espejo Δ≠0 debe certificar la salida sin ERR-PLANTILLA espurio.");

            // T6 (R-R-2): invariantes del flujo 5-ASE con Δ≠0. Los deltas de agosto son
            // {ASE1 −3, ASE2 −9, ASE3 −6, ASE4 +6, ASE5 +8} (ConteoAgosto vs plantilla Q2).
            var conteoPlantillaQ2 = new[] { 45, 75, 43, 73, 52 };
            var deltasEsperados = new[] { -3, -9, -6, +6, +8 };
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var filas = ContarFilasBloqueEnSalida(salida, NombresAse[aseId - 1]);
                Assert.Equal(ConteoAgosto[aseId - 1], filas);
                Assert.Equal(deltasEsperados[aseId - 1], ConteoAgosto[aseId - 1] - conteoPlantillaQ2[aseId - 1]);
            }

            // Fórmulas preservadas: el dimensionado 5-ASE no crea ni destruye fórmulas del bloque.
            // Plan 31 (T2): el EXTEMP visible de ASE3-agosto (sin filas Aplicacion) pasa a literal 0
            // por la recomposición por firma → una fórmula menos (antes/después, resto idéntico).
            Assert.Equal(
                ContarFormulas(Insumos.PlantillaQ2, "Reporte Componentes R1") - 1,
                ContarFormulas(salida, "Reporte Componentes R1"));
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

    private static readonly string[] NombresAse = ["PROMOAMBIENTAL", "LIME", "CIUDAD LIMPIA", "BOGOTA LIMPIA", "AREA LIMPIA"];

    private static int ContarFilasBloqueEnSalida(string ruta, string nombreAse)
    {
        using var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .First(s => string.Equals(s.Name?.Value, "Reporte Componentes R1", StringComparison.OrdinalIgnoreCase));
        var hoja = ((DocumentFormat.OpenXml.Packaging.WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet!;
        var filas = hoja.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().OrderBy(r => r.RowIndex?.Value ?? 0).ToList();

        var nameRow = 0;
        var totalRow = 0;
        foreach (var fila in filas)
        {
            if (nameRow == 0 && Normalizar(Texto(workbookPart, fila, "B")) == Normalizar(nombreAse))
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

        Assert.True(nameRow > 0 && totalRow > nameRow, $"ASE {nombreAse}: no se localizó el bloque en la salida.");
        return filas.Count(fila =>
        {
            var idx = (int)(fila.RowIndex?.Value ?? 0);
            return idx > nameRow && idx <= totalRow
                && new[] { "A", "B", "C", "D", "E" }.Any(c => !string.IsNullOrWhiteSpace(Texto(workbookPart, fila, c)));
        });
    }

    private static int ContarFormulas(string ruta, string hojaNombre)
    {
        using var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .First(s => string.Equals(s.Name?.Value, hojaNombre, StringComparison.OrdinalIgnoreCase));
        var hoja = ((DocumentFormat.OpenXml.Packaging.WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet!;
        return hoja.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Count(c => c.CellFormula is not null);
    }

    private static bool EsTotalFinal(DocumentFormat.OpenXml.Packaging.WorkbookPart workbookPart, DocumentFormat.OpenXml.Spreadsheet.Row fila) =>
        string.Equals(Texto(workbookPart, fila, "A"), "Total", StringComparison.OrdinalIgnoreCase)
        && string.IsNullOrWhiteSpace(Texto(workbookPart, fila, "B"));

    private static string Texto(DocumentFormat.OpenXml.Packaging.WorkbookPart workbookPart, DocumentFormat.OpenXml.Spreadsheet.Row fila, string columna)
    {
        var indice = ((uint?)fila.RowIndex?.Value ?? 0).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var celda = fila.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, columna + indice, StringComparison.OrdinalIgnoreCase));
        if (celda is null)
        {
            return string.Empty;
        }

        if (celda.InlineString is not null)
        {
            return celda.InlineString.InnerText ?? string.Empty;
        }

        if (celda.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString && celda.CellValue is not null
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
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta de agosto: {carpeta}");
        var dir = Directory.EnumerateDirectories(carpeta)
            .First(d => Path.GetFileName(d).StartsWith($"{aseId}-", StringComparison.OrdinalIgnoreCase));
        return Directory.EnumerateFiles(dir, "Recaudoporcomponente*.xlsx", SearchOption.TopDirectoryOnly).First();
    }
}
