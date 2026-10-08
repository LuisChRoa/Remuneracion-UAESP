using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 29 (T5, Unidad P — R-P-3/T0f): verifica END-TO-END que el total de control de recaudo se
/// puebla como LITERAL en <c>'Valida - Control Recaudo'!F10</c> y que la cadena
/// <c>VALIDACION_TOTAL!C15 = (C14 = C9)</c> queda satisfecha.
///
/// Flujo 5-ASE Q2 real (<see cref="ProcesadorPeriodo"/>) contra la plantilla canónica en ceros del
/// período (julio: <c>Docs/Prueba Julio-2/Plantilla_Remuneracion.xlsx</c>; agosto: base
/// <c>Docs/Prueba Agosto-2/Plantilla_Remuneracion.xlsx</c>, Plan 30/T3), con los insumos REALES de
/// julio (<c>Docs/Prueba Julio-2/Insumos</c>) y agosto (<c>Docs/Prueba Agosto-2/Insumos</c>).
///
/// Oráculos independientes:
///   - FUENTE: Σ <c>RecaudoEmpresaInputs.Total</c> leída por el reader desde las conciliaciones
///     (el mismo valor que alimenta <c>VALIDACION_TOTAL!C9 = SUM(C3:C8)</c>).
///   - MANUAL del administrativo: F10 (julio 72497949967 / agosto 77064563298).
///
/// Sin emojis. Tolerancia ±0.5.
/// </summary>
public sealed class ValidaControlRecaudoF10Tests
{
    private const decimal Tol = 0.5m;

    private const string HojaControl = WorkbookLeafCellMapValidaciones.HojaValidaControlRecaudo;
    private const string CeldaControl = WorkbookLeafCellMapValidaciones.CeldaControlRecaudoTotal;
    private const string HojaValidacionTotal = WorkbookLeafCellMapValidaciones.HojaValidacionTotal;

    private static readonly string Plantilla = Insumos.PlantillaQ2;

    private static readonly string CarpetaJulio = Insumos.CarpetaInsumosJulioQ2;
    private static readonly string CarpetaAgosto = Insumos.CarpetaInsumosAgosto;

    private static readonly string ManualJulio = Insumos.ManualJulioQ2;

    private static readonly string ManualAgosto = Insumos.ManualAgosto2026082;

    [Fact]
    public void JulioQ2_F10_IgualaFuenteYManual()
    {
        var periodo = new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 };
        using var salida = new SalidaTemporal("remuneracion-f10-julio-");
        EjecutarFlujo(CarpetaJulio, periodo, salida.Ruta);

        var fuente = TotalControlFuente(periodo, CarpetaJulio);
        var f10 = LeerCeldaNumerica(salida.Ruta, HojaControl, CeldaControl);

        Assert.InRange(f10 - fuente, -Tol, Tol);
        Assert.InRange(f10 - 72497949967m, -Tol, Tol);

        Assert.True(File.Exists(ManualJulio), $"Falta el manual de julio: {ManualJulio}");
        Assert.InRange(LeerCeldaNumerica(ManualJulio, HojaControl, CeldaControl) - f10, -Tol, Tol);

        AssertCadenaValidacionTotalSatisfaceC15(salida.Ruta);
    }

    [Fact]
    public void AgostoQ2_F10_IgualaFuenteYManual()
    {
        var periodo = new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 };
        using var salida = new SalidaTemporal("remuneracion-f10-agosto-");
        EjecutarFlujo(CarpetaAgosto, periodo, salida.Ruta, Insumos.PlantillaAgosto2026082);

        var fuente = TotalControlFuente(periodo, CarpetaAgosto);
        var f10 = LeerCeldaNumerica(salida.Ruta, HojaControl, CeldaControl);

        Assert.InRange(f10 - fuente, -Tol, Tol);

        Assert.True(File.Exists(ManualAgosto), $"Falta el manual de agosto: {ManualAgosto}");
        Assert.InRange(LeerCeldaNumerica(ManualAgosto, HojaControl, CeldaControl) - f10, -Tol, Tol);

        AssertCadenaValidacionTotalSatisfaceC15(salida.Ruta);
    }

    /// <summary>
    /// Guard anti-fórmula intacto (R-P-3): si la celda destino F10 fuera fórmula en la plantilla,
    /// la escritura lanza <c>ERR-PLANTILLA</c> y NO deja archivo parcial (nunca sobrescritura).
    /// </summary>
    [Fact]
    public void F10_DestinoFormula_LanzaErrPlantilla_SinSobrescribir()
    {
        using var temp = new DirectorioTemporal("remuneracion-f10-formula-");
        var plantillaMutada = Path.Combine(temp.Ruta, "plantilla-mutada.xlsx");
        File.Copy(Plantilla, plantillaMutada);

        ConvertirEnFormula(plantillaMutada, HojaControl, CeldaControl, "1+1");

        var salida = Path.Combine(temp.Ruta, "salida.xlsx");
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, salida, plantillaMutada));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains("fórmula", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(salida), "No debe quedar archivo parcial ante F10 fórmula.");
    }

    // ── Oráculos y asertos ─────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Σ de los totales de las hojas <c>Recaudo *</c> leída por el reader desde las conciliaciones
    /// reales: el valor que alimenta <c>VALIDACION_TOTAL!C9</c>. Camino independiente del writer.
    /// </summary>
    private static decimal TotalControlFuente(Periodo periodo, string carpetaPeriodo)
    {
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var localizador = new ArchivoFuenteLocator();
        var recaudos = reader.LeerRecaudosEmpresa(
            periodo,
            empresa => localizador.BuscarConciliacion(carpetaPeriodo, empresa.PrefijoConciliacion));
        return recaudos.Sum(r => r.Total);
    }

    /// <summary>
    /// R-P-3: la cadena <c>C14 = 'Valida - Control Recaudo'!F10</c> y <c>C15 = (C14 = C9)</c> sigue
    /// intacta (fórmulas no sobrescritas) y C15 queda en True. El assert de F10 = Σ fuente (en los
    /// tests que llaman) es lo que garantiza que la igualdad se sostenga post-recálculo, no el caché.
    /// </summary>
    private static void AssertCadenaValidacionTotalSatisfaceC15(string ruta)
    {
        var formulaC15 = LeerFormula(ruta, HojaValidacionTotal, "C15");
        Assert.NotNull(formulaC15);
        Assert.Contains("C9", NormalizarFormula(formulaC15!), StringComparison.OrdinalIgnoreCase);

        var formulaC14 = LeerFormula(ruta, HojaValidacionTotal, "C14");
        Assert.NotNull(formulaC14);
        Assert.Contains("F10", NormalizarFormula(formulaC14!), StringComparison.OrdinalIgnoreCase);

        Assert.True(
            LeerBooleanoCelda(ruta, HojaValidacionTotal, "C15"),
            "VALIDACION_TOTAL!C15 debe quedar en True post-escritura (R-P-3): C14 (=F10) = C9 con F10 poblado.");
    }

    // ── Helpers ────────────────────────────────────────────────────────────────────────────────

    private static void EjecutarFlujo(string carpetaInsumos, Periodo periodo, string rutaSalida, string? rutaPlantilla = null)
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
            RutaPlantilla = rutaPlantilla ?? Plantilla,
            RutaSalida = rutaSalida
        });
    }

    private static decimal LeerCeldaNumerica(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, hoja);
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        if (objCelda.CellValue is null || string.IsNullOrWhiteSpace(objCelda.CellValue.InnerText))
        {
            return 0m;
        }

        return decimal.TryParse(objCelda.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor)
            ? valor
            : 0m;
    }

    private static string? LeerFormula(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, hoja);
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");
        return objCelda.CellFormula?.Text;
    }

    private static bool LeerBooleanoCelda(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, hoja);
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        var texto = objCelda.CellValue?.InnerText?.Trim() ?? string.Empty;
        return texto.Equals("1", StringComparison.OrdinalIgnoreCase)
            || texto.Equals("true", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizarFormula(string formula) =>
        new(formula.Where(c => !char.IsWhiteSpace(c) && c != '\'').ToArray());

    private static void ConvertirEnFormula(string ruta, string hoja, string celda, string formula)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var worksheet = Hoja(documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null"), hoja);
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        objCelda.CellValue?.Remove();
        objCelda.DataType = null;
        objCelda.CellFormula = new CellFormula(formula);
        worksheet.Save();
    }

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet.");
    }

    private sealed class SalidaTemporal : IDisposable
    {
        private readonly DirectorioTemporal _dir;

        public SalidaTemporal(string prefijo)
        {
            _dir = new DirectorioTemporal(prefijo);
            Ruta = Path.Combine(_dir.Ruta, "salida.xlsx");
        }

        public string Ruta { get; }

        public void Dispose() => _dir.Dispose();
    }

    private sealed class DirectorioTemporal : IDisposable
    {
        public DirectorioTemporal(string prefijo)
        {
            Ruta = Path.Combine(Path.GetTempPath(), prefijo + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Ruta);
        }

        public string Ruta { get; }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(Ruta))
                {
                    Directory.Delete(Ruta, recursive: true);
                }
            }
            catch
            {
                // best-effort
            }
        }
    }
}
