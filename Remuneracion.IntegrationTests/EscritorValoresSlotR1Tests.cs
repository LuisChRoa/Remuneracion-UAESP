using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 36 (T1, R-B-1/R-B-2/R-B-4/R-T-2 — D-A/D-B/D-F): TDD del escritor values-only
/// <see cref="EscritorValoresSlotR1"/> (el writer aún no existe al escribir este test: primero ROJO).
///
/// Reglas del plan para lo nuevo: rutas EXPLÍCITAS (nunca <c>Insumos.Raiz</c> ni raíz
/// <c>Docs/Insumos</c>), regens SIEMPRE a temp FRESCA, tolerancia ±0.5 en comparaciones de
/// dominio, NUNCA <c>&lt;v&gt;</c> como oráculo de fórmulas (aquí <c>&lt;f&gt;</c> debe quedar
/// idéntica al template — se aserta). Los 511 rojos pre-existentes de fixture-root son
/// dependencia declarada: no se tocan, no se reconcilian aquí.
/// </summary>
public sealed class EscritorValoresSlotR1Tests
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

    // S2 (R-B-1): por cada ASE, la fuente real julio-2 se pega values-only en slots del contrato.
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void S2_FuenteRealJulio_PegaSoloValoresEnSlots(int aseId)
    {
        var rutaR1 = FuenteR1Julio(aseId);
        var bloque = new ExcelDataReaderWorkbookLeafInputReader().LeerEspejoR1(AseFactory.DesdeId(aseId), rutaR1);

        using var salida = new CopiaTemporal();
        var antes = FotoFormulas(salida.Ruta, HojaR1);

        EscritorValoresSlotR1.EscrituraSlotR1Resultado resultado;
        using (var documento = SpreadsheetDocument.Open(salida.Ruta, true))
        {
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            resultado = EscritorValoresSlotR1.Escribir(parte, bloque);
            parte.Workbook?.Save();
        }

        // Toda fila fuente se colocó (la fuente observada es subsecuencia ordenada del contrato).
        Assert.Equal(bloque.TotalFilas, resultado.FilasColocadas);

        // <f> byte-idéntica al template en todo R1 (texto + t + ref + si).
        var despues = FotoFormulas(salida.Ruta, HojaR1);
        Assert.Equal(antes, despues);

        // Todo <v> cambiado vive en un slot del contrato (columna F/L, filas del rango del ASE).
        var (inicio, fin) = RangoAse(aseId);
        Assert.All(CambiosDeValor(salida.Ruta, HojaR1), cambio =>
        {
            Assert.True(
                (cambio.Celda.StartsWith('F') || cambio.Celda.StartsWith('L')) && cambio.Fila >= inicio && cambio.Fila <= fin,
                $"El writer tocó {HojaR1}!{cambio.Celda} fuera de los slots del ASE {aseId} [{inicio},{fin}].");
        });

        // Reserva hidden intacta: sin celdas nuevas y hidden=1 preservado.
        AssertHiddenIntacta(salida.Ruta, HojaR1);

        // Spot-checks de dominio: primera fila y AFaseo del ASE llegan a su slot.
        var primera = bloque.Filas[0];
        Assert.Equal(
            primera.Valor("Total") ?? 0m,
            TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, $"F{inicio}"));
        // Spot-check AFaseo: si la fuente lo trae se coloca en F0; si no, el slot queda en
        // cero del template (ASE5-julio2 no trae AFaseo en ninguna variante: caso real observado).
        var afaseo = bloque.Filas.FirstOrDefault(f =>
            EscritorValoresSlotR1.EsFirmaAFaseo(f.A, f.B, f.D, f.E));
        Assert.Equal(
            afaseo?.Valor("Total") ?? 0m,
            TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, SlotAFaseo(aseId)));
    }

    // S5 (R-T-2): AFaseo dual-firma — ambas variantes resuelven al mismo slot F0 del ASE.
    [Fact]
    public void S5_AFaseoDualFirma_AmbasVariantesAlMismoSlot()
    {
        const string slot = "F17";
        foreach (var variante in new[] { VarianteA(), VarianteD() })
        {
            using var salida = new CopiaTemporal();
            var bloque = new BloqueEspejoAseInputs
            {
                Ase = AseFactory.DesdeId(1),
                TieneColumnaEspeciales = false,
                Filas = [variante],
            };

            using (var documento = SpreadsheetDocument.Open(salida.Ruta, true))
            {
                var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
                var resultado = EscritorValoresSlotR1.Escribir(parte, bloque);
                parte.Workbook?.Save();
                Assert.Equal(1, resultado.FilasColocadas);
            }

            Assert.Equal(123.45m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, slot));
            Assert.True(!TestHelpers.CeldaEsFormula(salida.Ruta, HojaR1, slot), $"El slot {slot} es fórmula tras el pegado.");
        }

        FilaEspejoR1 VarianteA() => new() { A = "AFaseo", B = "Total", C = "", D = "", E = "", ValoresPorColumna = SoloTotal() };
        FilaEspejoR1 VarianteD() => new() { A = "", B = "", C = "", D = "AFASEO", E = "Total", ValoresPorColumna = SoloTotal() };
        Dictionary<string, decimal?> SoloTotal() =>
            new(StringComparer.OrdinalIgnoreCase) { ["Total"] = 123.45m };
    }

    // S3 (R-B-2): fuente con empresa sin slot → ERR-SIN-SLOT con ASE + reporte + firma observada.
    [Fact]
    public void S3_EmpresaSinSlot_FailFastSinRemap()
    {
        using var salida = new CopiaTemporal();
        var bloque = new BloqueEspejoAseInputs
        {
            Ase = AseFactory.DesdeId(1),
            TieneColumnaEspeciales = true,
            Filas = [new FilaEspejoR1
            {
                C = "EMPRESA FANTASMA", D = "Total",
                ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["Total"] = 7m,
                    ["SERVICIO ESPECIALES"] = 0m,
                },
            }],
        };

        using var documento = SpreadsheetDocument.Open(salida.Ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var ex = Assert.Throws<CalculoInvalidoException>(() => EscritorValoresSlotR1.Escribir(parte, bloque));
        Assert.Equal(CodigoError.SinSlot, ex.Codigo);
        Assert.Contains("ERR-SIN-SLOT", ex.Message, StringComparison.Ordinal);
        Assert.Contains("ASE 1", ex.Message, StringComparison.Ordinal);
        Assert.Contains("Recaudoporcomponente", ex.Message, StringComparison.Ordinal);
        Assert.Contains("EMPRESA FANTASMA", ex.Message, StringComparison.Ordinal);
    }

    // S4 (D-A): slot cuya celda destino trae <f> → ERR-PLANTILLA, cero escritura parcial.
    [Fact]
    public void S4_DestinoConFormula_ErrPlantillaSinEscrituraParcial()
    {
        using var salida = new CopiaTemporal();
        CorromperConFormula(salida.Ruta, "F8");
        var antes = FotoFormulas(salida.Ruta, HojaR1);

        var bloque = new BloqueEspejoAseInputs
        {
            Ase = AseFactory.DesdeId(1),
            TieneColumnaEspeciales = false,
            Filas = [new FilaEspejoR1
            {
                C = "ENEL", D = "Total",
                ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase) { ["Total"] = 9m },
            }],
        };

        using (var documento = SpreadsheetDocument.Open(salida.Ruta, true))
        {
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var ex = Assert.Throws<CalculoInvalidoException>(() => EscritorValoresSlotR1.Escribir(parte, bloque));
            Assert.Equal(CodigoError.Plantilla, ex.Codigo);
            Assert.Contains("F8", ex.Message, StringComparison.Ordinal);
        }

        Assert.Equal(antes, FotoFormulas(salida.Ruta, HojaR1));
    }

    // Sinonimia como DATOS (Plan 34 intacto): fuente NUEVO ESQUEMA encaja en slot RECIPROCIDAD.
    [Fact]
    public void Sinonimia_NuevoEsquemaEncajaEnSlotReciprocidad()
    {
        var fuente = new FilaEspejoR1 { C = "NUEVO ESQUEMA", D = "Total" };
        var slot = new EscritorValoresSlotR1.SlotR1(135, "F135", "L135", "", "", "RECIPROCIDAD", "Total", "", null);
        Assert.True(EscritorValoresSlotR1.EncajaEnSlot(fuente, null, slot));
    }

    // Catálogo abierto: empresa no tabulada solo se iguala a sí misma.
    [Fact]
    public void CatalogoAbierto_EmpresaNoTabuladaSoloIgualASiMisma()
    {
        var fuente = new FilaEspejoR1 { C = "ENEL", D = "Total" };
        var otro = new EscritorValoresSlotR1.SlotR1(0, "F0", "L0", "", "", "OCCIDENTE", "Total", "", null);
        var mismo = new EscritorValoresSlotR1.SlotR1(0, "F0", "L0", "", "", "ENEL", "Total", "", null);
        Assert.False(EscritorValoresSlotR1.EncajaEnSlot(fuente, null, otro));
        Assert.True(EscritorValoresSlotR1.EncajaEnSlot(fuente, null, mismo));
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────────────

    private static (int Inicio, int Fin) RangoAse(int aseId) => aseId switch
    {
        1 => (6, 72),
        2 => (110, 210),
        3 => (289, 355),
        4 => (450, 550),
        5 => (603, 676),
        _ => throw new ArgumentOutOfRangeException(nameof(aseId)),
    };

    private static string SlotAFaseo(int aseId) => aseId switch
    {
        1 => "F17",
        2 => "F121",
        3 => "F300",
        4 => "F461",
        5 => "F614",
        _ => throw new ArgumentOutOfRangeException(nameof(aseId)),
    };

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

    private static string FuenteR1Julio(int aseId)
    {
        var carpeta = Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Insumos", CarpetasAse[aseId - 1]);
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
            _dir = Path.Combine(Path.GetTempPath(), "p36-t1-" + Guid.NewGuid().ToString("N"));
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

    private static void CorromperConFormula(string ruta, string celda)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, HojaR1, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)parte.GetPartById(hoja.Id!)).Worksheet!;
        var cell = ws.Descendants<Cell>().First(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        cell.CellFormula = new CellFormula("1+1");
        ws.Save();
        parte.Workbook.Save();
    }

    private static SortedDictionary<string, string> FotoFormulas(string ruta, string hoja)
    {
        var foto = new SortedDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!;
        foreach (var cell in ws.Descendants<Cell>())
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

    private sealed record CambioValor(string Celda, int Fila);

    private static List<CambioValor> CambiosDeValor(string ruta, string hoja)
    {
        // Compara contra una copia fresca del template: todo <v> distinto debe vivir en slots.
        var temp = Path.Combine(Path.GetTempPath(), "p36-base-" + Guid.NewGuid().ToString("N") + ".xlsx");
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
        var sheet = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!;
        foreach (var cell in ws.Descendants<Cell>())
        {
            if (cell.CellValue is null)
            {
                continue;
            }

            valores[cell.CellReference?.Value ?? "?"] = cell.CellValue.InnerText;
        }

        return valores;
    }

    private static void AssertHiddenIntacta(string ruta, string hoja)
    {
        var hidden = new HashSet<int>();
        for (var i = 213; i <= 251; i++)
        {
            hidden.Add(i);
        }

        for (var i = 356; i <= 412; i++)
        {
            hidden.Add(i);
        }

        for (var i = 552; i <= 566; i++)
        {
            hidden.Add(i);
        }

        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!;
        foreach (var row in ws.Descendants<Row>())
        {
            if (row.RowIndex is null || !hidden.Contains((int)row.RowIndex.Value))
            {
                continue;
            }

            Assert.True(row.Hidden is null || row.Hidden.Value, $"La fila {row.RowIndex.Value} perdió hidden=1.");
            foreach (var cell in row.Elements<Cell>())
            {
                Assert.True(cell.CellFormula is null, $"La fila hidden {row.RowIndex.Value} trae <f> en {cell.CellReference?.Value}.");
                var v = cell.CellValue?.InnerText;
                Assert.True(string.IsNullOrEmpty(v) || v == "0", $"La fila hidden {row.RowIndex.Value} trae <v>='{v}' en {cell.CellReference?.Value}.");
            }
        }
    }
}
