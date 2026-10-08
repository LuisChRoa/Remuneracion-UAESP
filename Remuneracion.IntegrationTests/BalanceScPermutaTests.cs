using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 29 (T1 / WU-1, Unidad B): verifica END-TO-END la permuta de destino de la hoja
/// <c>BCE SC POR FACT.</c> — <c>D = SUBSIDIO (← E-fuente)</c> y <c>E = CONTRIBUCION (← F-fuente)</c> —
/// contra el header de la plantilla (veredicto corregido T0-V4; el veredicto del Plan 10
/// "D=CONTRIBUCION" quedó refutado). Solo insumos REALES:
///   - julio-Q2: <c>Docs/Prueba Julio-2/Insumos</c> (5 ASE + Conciliaciones + R10) + salida actual
///     de la app <c>Docs/Prueba Julio-2/Resultado/Resultado1/Remuneración 202607-2 Total.xlsx</c>;
///   - agosto-Q2: <c>Docs/Prueba Agosto-2/Insumos</c> + salida actual
///     <c>Docs/Prueba Agosto-2/Resultado/Resultado2/Remuneración 202608-2 Total.xlsx</c>.
/// El "TOTAL GENERAL" de la fuente se localiza por LABEL (patrón Plan 25: roles por firma, no por
/// dirección fija). <c>F = D + E</c> queda invariante vs la salida actual (la permuta es
/// conmutativa, así que el dinero aguas abajo no cambia). Julio-ASE5 sigue fuente-fiel (R39): la
/// divergencia del manual se documenta como constante (R-B-2/S2), nunca como defecto.
/// </summary>
public sealed class BalanceScPermutaTests
{
    private const decimal Tolerancia = 0.5m;
    private const string HojaBce = WorkbookLeafCellMapBalanceSc.HojaBce;

    [Fact]
    public void JulioQ2_PermutaBce_SubsidioAD_ContribucionAE_VsFuente_E2E() =>
        VerificarPermutaPorPeriodo(
            carpetaInsumos: Insumos.CarpetaInsumosJulioQ2,
            salidaActualApp: Insumos.ResultadoJulioQ2(1),
            rutaPlantilla: Insumos.PlantillaQ2,
            periodo: new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 });

    [Fact]
    public void AgostoQ2_PermutaBce_SubsidioAD_ContribucionAE_VsFuente_E2E() =>
        VerificarPermutaPorPeriodo(
            carpetaInsumos: Insumos.CarpetaInsumosAgosto,
            salidaActualApp: Insumos.ResultadoAgostoQ2(2),
            rutaPlantilla: Insumos.PlantillaAgosto2026082,
            periodo: new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 });

    /// <summary>
    /// Julio-ASE5: la app es fuente-fiel (D7 ← E39 = Subsidio, E7 ← F39 = Contribución). El manual
    /// del administrativo trae OTROS valores (D7=-1138650714.67, E7=+1055691691.01) que no
    /// provienen de la fuente (T0a: 0 apariciones en 88 archivos). Se fijan como constante
    /// documentada, no como defecto (R-B-2 / S2); la rama de escritura de la app queda cubierta
    /// por <see cref="JulioQ2_PermutaBce_SubsidioAD_ContribucionAE_VsFuente_E2E"/>.
    /// </summary>
    [Fact]
    public void JulioAse5Bce_ManualDivergeDeLaFuente_ConstanteDocumentada()
    {
        var rutaFuente = LocalizarBalance(Insumos.CarpetaInsumosJulioQ2, aseId: 5);
        var (subsidioFuente, contribucionFuente, _) = LeerTotalGeneralPorLabel(rutaFuente);

        // La fuente R39 respalda a la app: Subsidio negativo en E, Contribución positiva en F.
        Assert.InRange(subsidioFuente - (-3615845886.78m), -Tolerancia, Tolerancia);
        Assert.InRange(contribucionFuente - 2322461234.81m, -Tolerancia, Tolerancia);

        var rutaManual = Insumos.ManualJulioQ2;
        Assert.True(File.Exists(rutaManual), $"Falta el manual de julio: {rutaManual}");
        var manualD7 = LeerCeldaNumerica(rutaManual, HojaBce, "D7");
        var manualE7 = LeerCeldaNumerica(rutaManual, HojaBce, "E7");

        // Constantes documentadas del manual (divergencia, no defecto de la app).
        Assert.InRange(manualD7 - (-1138650714.67m), -Tolerancia, Tolerancia);
        Assert.InRange(manualE7 - 1055691691.01m, -Tolerancia, Tolerancia);

        // La divergencia es real: el manual no coincide con la fuente (±0.5) que respalda a la app.
        Assert.False(Math.Abs(manualD7 - subsidioFuente) <= Tolerancia, "el manual D7 debería divergir de R39");
        Assert.False(Math.Abs(manualE7 - contribucionFuente) <= Tolerancia, "el manual E7 debería divergir de R39");
    }

    private static void VerificarPermutaPorPeriodo(string carpetaInsumos, string salidaActualApp, string rutaPlantilla, Periodo periodo)
    {
        Assert.True(Directory.Exists(carpetaInsumos), $"Falta la carpeta de insumos: {carpetaInsumos}");
        Assert.True(File.Exists(salidaActualApp), $"Falta la salida actual de la app: {salidaActualApp}");
        Assert.True(File.Exists(rutaPlantilla), $"Falta la plantilla del período {periodo.CodigoCompleto}: {rutaPlantilla}");

        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-bce-permuta-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, "salida.xlsx");
        try
        {
            var procesador = new ProcesadorPeriodo(
                new ExcelDataReaderRecaudoReader(),
                new ExcelDataReaderWorkbookLeafInputReader(),
                new CalculoRemuneracion(),
                new ValidadorBasico(),
                new OpenXmlPlantillaWriter(),
                new ArchivoFuenteLocator(),
                new ExcelDataReaderDetRetriR10Reader());

            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = periodo,
                CarpetaPeriodo = carpetaInsumos,
                RutaPlantilla = rutaPlantilla,
                RutaSalida = salida
            });

            Assert.True(File.Exists(salida), "El flujo 5-ASE debe certificar la salida.");

            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var rutaFuente = LocalizarBalance(carpetaInsumos, aseId);
                var (subsidioFuente, contribucionFuente, totalFuente) = LeerTotalGeneralPorLabel(rutaFuente);

                var fila = aseId + 2;
                var escritoD = LeerCeldaNumerica(salida, HojaBce, $"D{fila}");
                var escritoE = LeerCeldaNumerica(salida, HojaBce, $"E{fila}");

                // D = SUBSIDIO (E-fuente); E = CONTRIBUCION (F-fuente) — el header manda (T0-V4).
                Assert.InRange(escritoD - subsidioFuente, -Tolerancia, Tolerancia);
                Assert.InRange(escritoE - contribucionFuente, -Tolerancia, Tolerancia);

                // F = D + E invariante: mismo total que la salida actual de la app y que la fuente G.
                var appD = LeerCeldaNumerica(salidaActualApp, HojaBce, $"D{fila}");
                var appE = LeerCeldaNumerica(salidaActualApp, HojaBce, $"E{fila}");
                Assert.InRange((escritoD + escritoE) - (appD + appE), -Tolerancia, Tolerancia);
                Assert.InRange((escritoD + escritoE) - totalFuente, -Tolerancia, Tolerancia);
            }
        }
        finally
        {
            Borrar(salidaDir);
        }
    }

    /// <summary>
    /// Localiza el <c>R4-BalanceSubsidioyContribuciones*_*.xlsx</c> del ASE dentro de la carpeta
    /// del período (mismo criterio del locator: prefijo en el nombre; ASE5 usa
    /// <c>-Optimizado</c> como fallback y matchea el prefijo común normalizado).
    /// </summary>
    private static string LocalizarBalance(string carpetaInsumos, int aseId)
    {
        var carpetaAse = Directory.EnumerateDirectories(carpetaInsumos)
            .FirstOrDefault(d => Path.GetFileName(d).StartsWith($"{aseId}-", StringComparison.OrdinalIgnoreCase))
            ?? throw new DirectoryNotFoundException($"No se encontró la carpeta del ASE {aseId} en {carpetaInsumos}.");

        var archivo = Directory.EnumerateFiles(carpetaAse, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Normalizar(Path.GetFileNameWithoutExtension(f))
                .StartsWith("r4balancesubsidioycontribuciones", StringComparison.Ordinal))
            ?? throw new FileNotFoundException($"No se encontró el balance BCE del ASE {aseId} en {carpetaAse}.");

        return archivo;
    }

    /// <summary>
    /// Lee la fila "TOTAL GENERAL" del <c>Sheet1</c> de la fuente buscándola por LABEL normalizado
    /// (no por fila fija), y devuelve (Subsidio = E, Contribución = F, Total = G) de esa fila.
    /// </summary>
    private static (decimal Subsidio, decimal Contribucion, decimal Total) LeerTotalGeneralPorLabel(string ruta)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart
            ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, "Sheet1", StringComparison.OrdinalIgnoreCase))
            ?? workbookPart.Workbook.Descendants<Sheet>().First();
        var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{sheet.Name?.Value}' no tiene Worksheet.");

        Row? filaTotalGeneral = null;
        foreach (var row in worksheet.Descendants<Row>().OrderByDescending(r => r.RowIndex?.Value ?? 0))
        {
            if (row.Elements<Cell>().Any(c => string.Equals(
                    Normalizar(Texto(workbookPart, c)), WorkbookLeafCellMapBalanceSc.EtiquetaTotalGeneral,
                    StringComparison.Ordinal)))
            {
                filaTotalGeneral = row;
                break;
            }
        }

        if (filaTotalGeneral is null)
        {
            throw new InvalidOperationException($"No se encontró 'TOTAL GENERAL' en la fuente {Path.GetFileName(ruta)}.");
        }

        return (
            Numero(Texto(workbookPart, Celda(filaTotalGeneral, "E"))),
            Numero(Texto(workbookPart, Celda(filaTotalGeneral, "F"))),
            Numero(Texto(workbookPart, Celda(filaTotalGeneral, "G"))));
    }

    private static Cell? Celda(Row fila, string columna)
    {
        var indice = ((uint?)fila.RowIndex?.Value ?? 0).ToString(CultureInfo.InvariantCulture);
        return fila.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, columna + indice, StringComparison.OrdinalIgnoreCase));
    }

    private static string Texto(WorkbookPart workbookPart, Cell? celda)
    {
        if (celda?.CellValue is null)
        {
            return string.Empty;
        }

        if (celda.DataType?.Value == CellValues.SharedString
            && int.TryParse(celda.CellValue.Text, out var i))
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable?.ElementAt(i).InnerText ?? string.Empty;
        }

        return celda.CellValue.Text;
    }

    private static decimal Numero(string texto) =>
        decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor) ? valor : 0m;

    private static decimal LeerCeldaNumerica(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var celdaXml = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        return Numero(Texto(workbookPart, celdaXml));
    }

    /// <summary>Normaliza a minúsculas conservando solo letras/dígitos (sin espacios ni signos).</summary>
    private static string Normalizar(string texto)
    {
        var sb = new System.Text.StringBuilder(texto.Length);
        foreach (var ch in texto)
        {
            if (char.IsLetterOrDigit(ch))
            {
                sb.Append(char.ToLowerInvariant(ch));
            }
        }

        return sb.ToString();
    }

    private static void Borrar(string ruta)
    {
        try
        {
            if (Directory.Exists(ruta))
            {
                Directory.Delete(ruta, recursive: true);
            }
        }
        catch
        {
            // best-effort
        }
    }
}
