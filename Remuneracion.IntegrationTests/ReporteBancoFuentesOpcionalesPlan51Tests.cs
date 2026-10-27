using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 51 (H1, S2.1-S2.5): el reporte por banco resuelve sus columnas fuente POR HEADER
/// reconocido (A1) y la fuente esperada que no viene se OMITE con huella (A2), mientras el writer
/// y la plantilla quedan intactos (A4).
///
/// Reglas cubiertas:
/// <list type="bullet">
/// <item>repro 2026101: el Resumen {ENEL, OCCIDENTE, Total} de ASE2 ya NO lanza el throw
/// posicional; <c>NUEVO ESQUEMA</c> queda omitido con huella <c>BANCO:NUEVO ESQUEMA</c>;</item>
/// <item><c>Total</c> nunca es fuente ni evidencia;</item>
/// <item>etiqueta desconocida SIN ausencia = sólo evidencia; cero fuentes reconocidas → fail-fast;
/// fuente reconocida sin celdas destino (desync mapa↔plantilla) → fail-fast;</item>
/// <item>sinonimia citada <c>RECIPROCIDAD ↔ NUEVO ESQUEMA</c> resuelta contra la clave de la
/// plantilla;</item>
/// <item>writer: bloque con ausente escribe las presentes + C59 y deja las celdas de la ausente
/// como estaban en la plantilla (0).</item>
/// </list>
/// </summary>
public sealed class ReporteBancoFuentesOpcionalesPlan51Tests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    // ── A1/A2 — repro del rojo 2026101 ────────────────────────────────────────────────────────

    [Fact]
    public void LeerReporteBanco_Octubre1_Ase2_NoLanzaThrowPosicional_OmiteNuevoEsquemaConHuella()
    {
        // La fuente real de octubre-1 ASE2 trae {ENEL, OCCIDENTE, Total}: antes del Plan 51 el
        // mapa T0-0.3 exigía 'NUEVO ESQUEMA' en la columna 2 y abortaba. Ahora se resuelve por
        // header: la fuente ausente se omite con huella.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(
            Insumos.Ase(2),
            new Periodo { CodigoAAAAMM = "202610", NumeroQuincena = 1 },
            RutaReporteBanco("Prueba Octubre-1", 2));

        var bloque = resultado.Ases.Single();
        Assert.Equal(new[] { "ENEL", "OCCIDENTE" }, bloque.Empresas.Select(e => e.Empresa).ToArray());
        Assert.Equal(new[] { "BANCO:NUEVO ESQUEMA" }, bloque.FuentesAusentesBanco.ToArray());
        Assert.Empty(bloque.EtiquetasBancoNoReconocidas);

        var enel = bloque.Empresas.Single(e => e.Empresa == "ENEL");
        var occ = bloque.Empresas.Single(e => e.Empresa == "OCCIDENTE");
        Assert.InRange(enel.TotalFuente - 10309260573m, -Tolerancia, Tolerancia);
        Assert.InRange(occ.TotalFuente - 971950730m, -Tolerancia, Tolerancia);

        // Gate D5-i intacto por empresa PRESENTE (el reader ya lo valida antes de devolver).
        Assert.InRange(enel.Total - enel.TotalFuente, -Tolerancia, Tolerancia);
        Assert.InRange(occ.Total - occ.TotalFuente, -Tolerancia, Tolerancia);

        // Conceptos 1/2/3 leídos por columna re-resuelta y el 7 ausente en Q1 = 0 demostrable.
        Assert.InRange(enel.AplicadosFacturacion - 10286935425.87m, -Tolerancia, Tolerancia);
        Assert.InRange(enel.SaldosFavorGenerados - 22325147.13m, -Tolerancia, Tolerancia);
        Assert.Equal(0m, enel.FinanciacionesNuevas);
        Assert.Equal(0m, enel.RecibosServEspeciales);

        Assert.Equal(1, resultado.Quincena);
        Assert.Null(resultado.Consolidado);
    }

    [Fact]
    public void LeerReporteBanco_Octubre1_Ase4_TambienOmiteNuevoEsquemaConHuella()
    {
        // Dato medido en H0 (§5.1 del T0): ASE4 de octubre-1 también trae {ENEL, OCCIDENTE}.
        // Misma familia de ausencia legítima: omisión + huella, sin ramas.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(
            Insumos.Ase(4),
            new Periodo { CodigoAAAAMM = "202610", NumeroQuincena = 1 },
            RutaReporteBanco("Prueba Octubre-1", 4));

        var bloque = resultado.Ases.Single();
        Assert.Equal(new[] { "ENEL", "OCCIDENTE"}, bloque.Empresas.Select(e => e.Empresa).ToArray());
        Assert.Equal(new[] { "BANCO:NUEVO ESQUEMA"}, bloque.FuentesAusentesBanco.ToArray());
    }

    [Fact]
    public void LeerReporteBanco_Julio1_Ase2_TresFuentes_SinAusenciasNiEvidencia()
    {
        // Paridad con el comportamiento anterior (T0-0.3): con las 3 fuentes presentes no hay
        // huella ni evidencia y los valores son los de siempre.
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(Insumos.Ase(2), Insumos.Periodo(), Insumos.ReporteBanco(2));

        var bloque = resultado.Ases.Single();
        Assert.Equal(3, bloque.Empresas.Count);
        Assert.Empty(bloque.FuentesAusentesBanco);
        Assert.Empty(bloque.EtiquetasBancoNoReconocidas);
        Assert.InRange(bloque.Empresas.Single(e => e.Empresa == "NUEVO ESQUEMA").Total - 29345550m, -Tolerancia, Tolerancia);
    }

    // ── A1 — reconocimiento por header (sintéticos) ───────────────────────────────────────────

    [Fact]
    public void LeerReporteBanco_Total_NuncaEsFuenteNiEvidencia()
    {
        // "TOTAL" como columna intermedia: se excluye por normalizado ANTES del vocabulario.
        var ruta = CrearReporteSintetico(["ENEL", "TOTAL", "OCCIDENTE"]);
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(Insumos.Ase(1), Insumos.Periodo(), ruta);

        var bloque = resultado.Ases.Single();
        Assert.Equal(new[] { "ENEL", "OCCIDENTE"}, bloque.Empresas.Select(e => e.Empresa).ToArray());
        Assert.DoesNotContain(bloque.Empresas, e => e.Empresa.Equals("TOTAL", StringComparison.OrdinalIgnoreCase));
        Assert.Empty(bloque.EtiquetasBancoNoReconocidas);
    }

    [Fact]
    public void LeerReporteBanco_DesconocidaSinAusencia_EsSoloEvidencia()
    {
        // A3 regla 3: etiqueta real sin destino no convierte en rojo lo sano (gemelo 'EAB' Plan 50).
        var ruta = CrearReporteSintetico(["ENEL", "OCCIDENTE", "RARO"]);
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(Insumos.Ase(1), Insumos.Periodo(), ruta);

        var bloque = resultado.Ases.Single();
        Assert.Equal(new[] { "ENEL", "OCCIDENTE"}, bloque.Empresas.Select(e => e.Empresa).ToArray());
        Assert.Empty(bloque.FuentesAusentesBanco);
        Assert.Equal(new[] { "BANCO:RARO"}, bloque.EtiquetasBancoNoReconocidas.ToArray());
    }

    [Fact]
    public void LeerReporteBanco_CeroFuentesReconocidas_FailFastEstructuraRota()
    {
        // A3 regla 5 / R-VACIO: un resumen sin ninguna fuente del catálogo es estructura rota.
        var ruta = CrearReporteSintetico(["RARO"]);

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            new ExcelDataReaderWorkbookLeafInputReader().LeerReporteBanco(Insumos.Ase(1), Insumos.Periodo(), ruta));

        Assert.Contains("ASE 1", ex.Message, StringComparison.Ordinal);
        Assert.Contains("BANCO", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void LeerReporteBanco_FuenteReconocidaSinCeldasDestino_FailFastDesync()
    {
        // R-MAPA-DESYNC: ENERBIT es fuente del vocabulario pero la plantilla no le declara celdas
        // en ASE1 → mapa↔plantilla desincronizados, nunca skip silencioso.
        var ruta = CrearReporteSintetico(["ENEL", "OCCIDENTE", "ENERBIT"]);

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            new ExcelDataReaderWorkbookLeafInputReader().LeerReporteBanco(Insumos.Ase(1), Insumos.Periodo(), ruta));

        Assert.Contains("ASE 1", ex.Message, StringComparison.Ordinal);
        Assert.Contains("BANCO", ex.Message, StringComparison.Ordinal);
        Assert.Contains("ENERBIT", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void LeerReporteBanco_SinonimoReciprocidad_ResuelveLaClaveDeLaPlantilla()
    {
        // La fuente puede traer el rótulo legado 'RECIPROCIDAD'; la clave de la plantilla en ASE2
        // es 'NUEVO ESQUEMA' (sinonimia citada, Plan 34). No se inventa una cuarta empresa.
        var ruta = CrearReporteSintetico(["ENEL", "RECIPROCIDAD", "OCCIDENTE"]);
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(Insumos.Ase(2), Insumos.Periodo(), ruta);

        var bloque = resultado.Ases.Single();
        Assert.Equal(new[] { "ENEL", "NUEVO ESQUEMA", "OCCIDENTE"}, bloque.Empresas.Select(e => e.Empresa).ToArray());
        Assert.Empty(bloque.FuentesAusentesBanco);
        Assert.Empty(bloque.EtiquetasBancoNoReconocidas);
    }

    [Fact]
    public void LeerReporteBanco_HeaderDesplazado_ResuelvePorEtiquetaYNoPorPosicion()
    {
        // El mismo par de fuentes en orden físico invertido: el mapeo por header asigna cada
        // columna a su empresa (un mapa posicional leería los valores cruzados).
        var ruta = CrearReporteSintetico(["OCCIDENTE", "ENEL"]);
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var resultado = reader.LeerReporteBanco(Insumos.Ase(1), Insumos.Periodo(), ruta);

        var bloque = resultado.Ases.Single();
        var enel = bloque.Empresas.Single(e => e.Empresa == "ENEL");
        var occ = bloque.Empresas.Single(e => e.Empresa == "OCCIDENTE");

        // En el fixture la columna 0 de datos vale 100 y la 1 vale 110 (concepto 1).
        Assert.InRange(enel.AplicadosFacturacion - 110m, -Tolerancia, Tolerancia);
        Assert.InRange(occ.AplicadosFacturacion - 100m, -Tolerancia, Tolerancia);
    }

    // ── A4 — writer intacto: presentes + C59, celdas de la ausente como en la plantilla ───────

    [Fact]
    public void Writer_BloqueConFuenteAusente_EscribePresentesYC59_DejaLaAusenteEnCero()
    {
        var (resultado, leaf) = CrearCasoValidoAse1();
        leaf.ReporteBanco = new ExcelDataReaderWorkbookLeafInputReader().LeerReporteBanco(
            Insumos.Ase(1), Insumos.Periodo(), CrearReporteSintetico(["ENEL"]));

        var bloque = leaf.ReporteBanco.Ases.Single();
        Assert.Equal(new[] { "ENEL"}, bloque.Empresas.Select(e => e.Empresa).ToArray());
        Assert.Equal(new[] { "BANCO:OCCIDENTE"}, bloque.FuentesAusentesBanco.ToArray());

        var dir = Path.Combine(Path.GetTempPath(), "p51-writer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var salida = Path.Combine(dir, "salida.xlsx");
        var plantilla = Insumos.PlantillaHistoricaV2();

        // Valor de la celda destino de la fuente ausente ANTES de escribir (0 de plantilla).
        var occidenteAntes = LeerCelda(plantilla, "REPORTE RECAUDO x BANCO", "H13");

        new OpenXmlPlantillaWriter().GenerarWorkbook(plantilla, salida, resultado, new[] { leaf });

        using var doc = SpreadsheetDocument.Open(salida, false);
        var hoja = HojaBanco(doc);
        Assert.Equal(100m, LeerCeldaNumerica(hoja, "C13")); // ENEL concepto 1
        Assert.Equal(10m, LeerCeldaNumerica(hoja, "C14"));  // ENEL concepto 2
        Assert.Equal(5m, LeerCeldaNumerica(hoja, "C15"));   // ENEL concepto 3
        Assert.Equal(0m, LeerCeldaNumerica(hoja, "C16"));   // ENEL concepto 7 (fila ausente = 0)
        Assert.Equal(1m, LeerCeldaNumerica(hoja, "C59"));   // C59 = quincena (dominio)

        // La fuente ausente no se escribe: su celda queda como en la plantilla (0).
        Assert.Equal(0m, occidenteAntes);
        Assert.Equal(occidenteAntes, LeerCeldaNumerica(hoja, "H13"));
    }

    // ── Fixtures ──────────────────────────────────────────────────────────────────────────────

    private static string RutaReporteBanco(string periodo, int ase)
    {
        var insumos = Path.Combine(Insumos.Raiz(), "Docs", periodo, "Insumos");
        var carpeta = Directory.EnumerateDirectories(insumos)
            .First(d => Path.GetFileName(d).StartsWith($"{ase}-", StringComparison.Ordinal));
        return Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .First(f => Path.GetFileNameWithoutExtension(f).StartsWith("ReportePagosxBanco", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Fixture sintético (OpenXML) del Resumen: una columna por fuente (en el orden recibido) +
    /// columna <c>Total</c>, conceptos 1/2/3 y fila <c>Total</c>. Valores por columna de datos
    /// <c>i</c>: concepto 1 = <c>100 + 10·i</c>, concepto 2 = <c>10 + i</c>, concepto 3 = <c>5 + i</c>.
    /// </summary>
    private static string CrearReporteSintetico(IReadOnlyList<string> fuentes)
    {
        var ruta = Path.Combine(Path.GetTempPath(), "p51-banco-" + Guid.NewGuid().ToString("N") + ".xlsx");
        using var doc = SpreadsheetDocument.Create(ruta, SpreadsheetDocumentType.Workbook);
        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheetPart = workbookPart.AddNewPart<WorksheetPart>();

        var filas = new List<Row>
        {
            new(new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("Resumen Recaudo Aplicado Por Servicio") }),
            FilaHeader(fuentes),
            FilaConcepto("1-APLICADOS A FACTURACION", fuentes, i => 100 + (10 * i)),
            FilaConcepto("2-SALDOS A FAVOR GENERADOS", fuentes, i => 10 + i),
            FilaConcepto("3-APLICADOS A FINANCIACIONES NUEVAS", fuentes, i => 5 + i),
            FilaTotal(fuentes),
        };

        sheetPart.Worksheet = new Worksheet(new SheetData(filas));
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(sheetPart), SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        return ruta;
    }

    private static Row FilaHeader(IReadOnlyList<string> fuentes)
    {
        var row = new Row();
        for (var i = 0; i < fuentes.Count; i++)
        {
            row.AppendChild(new Cell
            {
                CellReference = Columna(i + 1) + "2",
                DataType = CellValues.String,
                CellValue = new CellValue(fuentes[i]),
            });
        }

        row.AppendChild(new Cell
        {
            CellReference = Columna(fuentes.Count + 1) + "2",
            DataType = CellValues.String,
            CellValue = new CellValue("Total"),
        });
        return row;
    }

    private static Row FilaConcepto(string etiqueta, IReadOnlyList<string> fuentes, Func<int, int> valor)
    {
        var numeroFila = etiqueta.StartsWith('1') ? 3 : etiqueta.StartsWith('2') ? 4 : 5;
        var row = new Row(new Cell
        {
            CellReference = "A" + numeroFila.ToString(CultureInfo.InvariantCulture),
            DataType = CellValues.String,
            CellValue = new CellValue(etiqueta),
        });
        for (var i = 0; i < fuentes.Count; i++)
        {
            row.AppendChild(new Cell
            {
                CellReference = Columna(i + 1) + numeroFila.ToString(CultureInfo.InvariantCulture),
                CellValue = new CellValue(valor(i).ToString(CultureInfo.InvariantCulture)),
            });
        }

        return row;
    }

    private static Row FilaTotal(IReadOnlyList<string> fuentes)
    {
        var row = new Row(new Cell
        {
            CellReference = "A6",
            DataType = CellValues.String,
            CellValue = new CellValue("Total"),
        });
        for (var i = 0; i < fuentes.Count; i++)
        {
            var total = (100 + (10 * i)) + (10 + i) + (5 + i);
            row.AppendChild(new Cell
            {
                CellReference = Columna(i + 1) + "6",
                CellValue = new CellValue(total.ToString(CultureInfo.InvariantCulture)),
            });
        }

        return row;
    }

    private static string Columna(int indice)
    {
        var letra = (char)('A' + indice);
        return letra.ToString();
    }

    private static (ResultadoRemuneracion Resultado, WorkbookLeafInputs Leaf) CrearCasoValidoAse1()
    {
        var reader = new ExcelDataReaderRecaudoReader();
        var r1 = reader.LeerR1(Insumos.R1(1));
        var r2 = reader.LeerR2(Insumos.R2(1));
        var r4 = reader.LeerR4(Insumos.R4(1));
        var ase = Insumos.Ase(1);
        var periodo = Insumos.Periodo();
        var resultado = new CalculoRemuneracion().CalcularConsolidado(periodo, [(ase, r1, r2, r4)]);
        var leaf = new ExcelDataReaderWorkbookLeafInputReader().LeerLeafInputs(ase, periodo, Insumos.R1(1), Insumos.R2(1), Insumos.R4(1));
        return (resultado, leaf);
    }

    private static Worksheet HojaBanco(SpreadsheetDocument doc)
    {
        var workbookPart = doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var libro = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook null");
        var sheet = libro.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, "REPORTE RECAUDO x BANCO", StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException("Worksheet null");
    }

    private static decimal LeerCelda(string ruta, string hoja, string celda)
    {
        using var doc = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var libro = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook null");
        var sheet = libro.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException("Worksheet null");
        return LeerCeldaNumerica(worksheet, celda);
    }

    private static decimal LeerCeldaNumerica(Worksheet worksheet, string celda)
    {
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        var valor = cell?.CellValue?.Text;
        return decimal.TryParse(valor, NumberStyles.Any, CultureInfo.InvariantCulture, out var numero) ? numero : 0m;
    }
}
