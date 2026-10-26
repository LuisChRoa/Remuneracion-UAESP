using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 36 (T2, R-B-1/R-B-3/R-B-4/R-G-1 — D-A/D-B/D-E): integración del espejo R1 por su API
/// pública (<see cref="OpenXmlPlantillaWriter.EscribirEspejoR1(string, string, System.Collections.Generic.IReadOnlyList{Remuneracion.Core.Models.BloqueEspejoAseInputs}, Remuneracion.Core.Models.Periodo)"/>)
/// sobre la plantilla estandarizada única.
///
/// Tras el retiro del motor geométrico (Planes 21/31/32) la operación es <b>values-only por slot</b>:
/// se pegan los 5 bloques R1 de una fuente REAL (agosto-2026082) y se audita el invariante del plan:
///   - <c>&lt;f&gt;</c> de <c>Reporte Componentes R1</c> byte-idéntica al template (texto + t/ref/si);
///   - todo <c>&lt;v&gt;</c> cambiado vive en un slot del contrato (columna F/L, rango del ASE);
///   - dimension/reserva hidden del template intactas (D-E: NO-TOUCH).
///
/// Rutas EXPLÍCITAS (nunca <c>Insumos.Raiz()</c> ni la raíz <c>Docs/Insumos</c>), regen a temp
/// FRESCA, lectura con OpenXML en modo lectura (sin Excel/COM). PROHIBIDO usar cachés <c>&lt;v&gt;</c>
/// como oráculo de fórmulas: aquí el oráculo es el propio template.
/// </summary>
public sealed class EspejoR1IntegracionTests
{
    private const string HojaR1 = "Reporte Componentes R1";

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental",
        "2-Lime",
        "3-Ciudad Limpia",
        "4-Bogota Limpia",
        "5-Área Limpia",
    ];

    [Fact]
    public void Agosto_CincoAse_ValuesOnlySobrePlantillaUnica()
    {
        var lector = new ExcelDataReaderWorkbookLeafInputReader();
        var bloques = new List<BloqueEspejoAseInputs>(5);
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            bloques.Add(lector.LeerEspejoR1(AseFactory.DesdeId(aseId), FuenteR1Agosto(aseId)));
        }

        using var salida = new CopiaTemporal();
        var formulasAntes = FotoFormulas(salida.Ruta, HojaR1);
        var dimensionAntes = Dimension(salida.Ruta, HojaR1);
        var hiddenAntes = FilasHidden(salida.Ruta, HojaR1);

        // API pública: sobre una COPIA de la plantilla única. Período = base interna del oficial
        // (2026092): la Phase 0 es identidad y el oráculo de R1 sigue siendo el propio template.
        new OpenXmlPlantillaWriter().EscribirEspejoR1(Plantilla(), salida.Ruta, bloques, Periodo.Parse("2026092"));

        Assert.True(File.Exists(salida.Ruta));

        // R-G-1: <f> byte-idéntica al template en todo R1 (texto + t + ref + si).
        Assert.Equal(formulasAntes, FotoFormulas(salida.Ruta, HojaR1));

        // D-E: dimension y reserva hidden del template intactas (ni <v> ni <f> en la reserva).
        Assert.Equal(dimensionAntes, Dimension(salida.Ruta, HojaR1));
        Assert.Equal(hiddenAntes, FilasHidden(salida.Ruta, HojaR1));

        // Plan 44 (F1, §2.3): todo <v> cambiado vive en un slot del contrato multicolumna
        // (columnas F..T + Z, filas del ASE). `Exonera-*` (U/V/W/X) e `Y` quedan intactas.
        var cambios = CambiosDeValor(salida.Ruta, HojaR1);
        Assert.NotEmpty(cambios);
        var columnasContrato = new[] { 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'Z' };
        var bloquesDelimitados = BloquesDelimitados(salida.Ruta);
        Assert.All(cambios, cambio =>
        {
            var bloque = BloqueDelimitado(bloquesDelimitados, cambio.Fila);
            Assert.True(
                columnasContrato.Contains(cambio.Celda[0])
                && cambio.Fila >= bloque.Cabecera
                && cambio.Fila <= bloque.GranTotal,
                $"El escritor tocó {HojaR1}!{cambio.Celda} fuera del bloque delimitado del ASE {bloque.AseId} [{bloque.Cabecera},{bloque.GranTotal}] (multicolumna F..T/Z).");
        });
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Bloques de R1 DELIMITADOS por la propia hoja (Plan 59 H1.2): cabecera por rótulo del ASE +
    /// gran total declarado. Sin ventanas de filas congeladas.
    /// </summary>
    private static IReadOnlyList<DelimitadorBloquesHoja.BloqueHoja> BloquesDelimitados(string ruta)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        return DelimitadorBloquesHoja.Delimitar(parte, DelimitadorBloquesHoja.R1);
    }

    private static DelimitadorBloquesHoja.BloqueHoja BloqueDelimitado(
        IReadOnlyList<DelimitadorBloquesHoja.BloqueHoja> bloques,
        int fila)
    {
        foreach (var bloque in bloques)
        {
            if (fila >= bloque.Cabecera && fila <= bloque.GranTotal)
            {
                return bloque;
            }
        }

        throw new InvalidOperationException($"La fila {fila} no pertenece a ningún bloque delimitado de la hoja {HojaR1}.");
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

    private static string Plantilla() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string FuenteR1Agosto(int aseId)
    {
        var carpeta = Path.Combine(Raiz(), "Docs", "Prueba Agosto-2", "Insumos", CarpetasAse[aseId - 1]);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta explícita: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("Recaudoporcomponente", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Recaudoporcomponente del ASE{aseId} en {carpeta}");
        return archivo!;
    }

    private sealed class CopiaTemporal : IDisposable
    {
        private readonly string _dir;

        public CopiaTemporal()
        {
            _dir = Path.Combine(Path.GetTempPath(), "p36-t2-integracion-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
            Ruta = Path.Combine(_dir, "salida.xlsx");
            File.Copy(Plantilla(), Ruta, overwrite: false);
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

    private static Worksheet Hoja(WorkbookPart parte, string nombre)
    {
        var sheet = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!;
    }

    private static SortedDictionary<string, string> FotoFormulas(string ruta, string hoja)
    {
        var foto = new SortedDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        foreach (var cell in Hoja(parte, hoja).Descendants<Cell>())
        {
            var f = cell.CellFormula;
            if (f is null)
            {
                continue;
            }

            foto[cell.CellReference?.Value ?? "?"] =
                $"{f.Text}|t={f.FormulaType?.Value.ToString() ?? ""}|ref={f.Reference?.Value ?? ""}|si={f.SharedIndex?.Value.ToString() ?? ""}";
        }

        return foto;
    }

    private static string Dimension(string ruta, string hoja)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        return Hoja(parte, hoja).SheetDimension?.Reference?.Value ?? string.Empty;
    }

    private static SortedSet<int> FilasHidden(string ruta, string hoja)
    {
        var hidden = new SortedSet<int>();
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(parte, hoja);
        var cantidadFilas = 0;
        foreach (var row in worksheet.Descendants<Row>())
        {
            cantidadFilas++;
            if (row.Hidden is not null && row.Hidden.Value && row.RowIndex is not null)
            {
                hidden.Add((int)row.RowIndex.Value);
            }
        }

        Assert.Equal(711, cantidadFilas); // el retiro del motor NO dimensiona: siguen las 711 filas del template.
        return hidden;
    }

    private sealed record CambioValor(string Celda, int Fila);

    private static List<CambioValor> CambiosDeValor(string ruta, string hoja)
    {
        var temp = Path.Combine(Path.GetTempPath(), "p36-t2-base-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            File.Copy(Plantilla(), temp, overwrite: false);
            var baseVals = ValoresHoja(temp, hoja);
            var actual = ValoresHoja(ruta, hoja);
            var cambios = new List<CambioValor>();
            foreach (var (celda, valor) in actual)
            {
                if (!baseVals.TryGetValue(celda, out var previo) || previo != valor)
                {
                    cambios.Add(new CambioValor(celda, int.Parse(new string(celda.Where(char.IsAsciiDigit).ToArray()), CultureInfo.InvariantCulture)));
                }
            }

            return cambios;
        }
        finally
        {
            try
            {
                File.Delete(temp);
            }
            catch
            {
                // best-effort
            }
        }
    }

    private static Dictionary<string, string> ValoresHoja(string ruta, string hoja)
    {
        var valores = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        foreach (var cell in Hoja(parte, hoja).Descendants<Cell>())
        {
            if (cell.CellValue is null)
            {
                continue;
            }

            valores[cell.CellReference?.Value ?? "?"] = cell.CellValue.InnerText;
        }

        return valores;
    }
}
