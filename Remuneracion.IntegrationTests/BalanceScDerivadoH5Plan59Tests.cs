using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (H5, §0.1/§0.2): el gate del balance de subsidio/contribución quedó DERIVADO
/// (<see cref="VerificadorBalanceScDerivado"/>), pero sigue VIVO: si una celda protegida de
/// <c>BCE SC POR FACT.</c> deja de ser fórmula, la corrida aborta con <c>ERR-PLANTILLA</c> nombrando
/// la identidad (banda · rol · ASE).
///
/// Estas pruebas son la contraparte conductual del candado estático: no basta con que el código no
/// tenga literales de período; hay que demostrar que la verificación que sustituyó a la foto congelada
/// sigue detectando la rotura. Se injerta sobre una COPIA temporal de la base del período (la plantilla
/// oficial nunca se abre en escritura).
/// </summary>
public sealed class BalanceScDerivadoH5Plan59Tests
{
    private static readonly Periodo PeriodoOctubre = Periodo.Parse("2026101");

    [Fact]
    public void H5_BaseDelPeriodoSinTocar_PasaLaVerificacionDerivada()
    {
        using var baseCopiada = new CopiaTemporal(Insumos.BaseDelPeriodoV4("2026101"));
        using var documento = SpreadsheetDocument.Open(baseCopiada.Ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");

        // No lanza: la plantilla vigente declara todos los roles que el verificador exige.
        VerificadorBalanceScDerivado.Validar(parte, PeriodoOctubre);
    }

    [Fact]
    public void H5_Injerto_VinculoConElDetalleDelPeriodoConvertidoEnValor_Enrojece()
    {
        using var baseCopiada = new CopiaTemporal(Insumos.BaseDelPeriodoV4("2026101"));

        // El rol 'SISTEMA' es el vínculo con el balance del ASE en la hoja de detalle REAL del período.
        var celda = Mutar(baseCopiada.Ruta, texto => texto.Contains("DetRetri", StringComparison.OrdinalIgnoreCase));

        var ex = Assert.Throws<CalculoInvalidoException>(() => Verificar(baseCopiada.Ruta));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains("ERR-PLANTILLA", ex.Message, StringComparison.Ordinal);
        Assert.Contains("SISTEMA", ex.Message, StringComparison.Ordinal);
        Assert.Contains("DetRetri", ex.Message, StringComparison.Ordinal);
        Assert.Contains(celda, ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void H5_Injerto_SumaDeTramoConvertidaEnValor_Enrojece()
    {
        using var baseCopiada = new CopiaTemporal(Insumos.BaseDelPeriodoV4("2026101"));

        // La suma sobre el tramo de ASE es un ROL de la banda de totales: si deja de ser fórmula, el
        // gate debe decirlo (antes lo decía la foto congelada de 78 celdas).
        var celda = Mutar(baseCopiada.Ruta, texto => texto.StartsWith("SUM(", StringComparison.OrdinalIgnoreCase));

        var ex = Assert.Throws<CalculoInvalidoException>(() => Verificar(baseCopiada.Ruta));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains("ERR-PLANTILLA", ex.Message, StringComparison.Ordinal);
        Assert.Contains("sumas", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("TOTAL BSC", ex.Message, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(celda));
    }

    [Fact]
    public void H5_Injerto_TotalBscDescuadrado_Enrojece()
    {
        using var baseCopiada = new CopiaTemporal(Insumos.BaseDelPeriodoV4("2026101"));

        // 'TOTAL BSC' = SUBSIDIO + CONTRIBUCION por ASE: se le cambia un término y la forma deja de
        // ser la que la geometría exige.
        var celda = Mutar(
            baseCopiada.Ruta,
            texto => texto.Replace(" ", string.Empty, StringComparison.Ordinal).Length > 0
                && !texto.Contains('!')
                && texto.Split('+').Length == 2,
            reemplazo: "1+1+1");

        var ex = Assert.Throws<CalculoInvalidoException>(() => Verificar(baseCopiada.Ruta));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains("ERR-PLANTILLA", ex.Message, StringComparison.Ordinal);
        Assert.Contains("TOTAL BSC", ex.Message, StringComparison.Ordinal);
        Assert.Contains(celda, ex.Message, StringComparison.Ordinal);
    }

    // ── helpers ──────────────────────────────────────────────────────────────────────────────────

    private static void Verificar(string ruta)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        VerificadorBalanceScDerivado.Validar(parte, PeriodoOctubre);
    }

    /// <summary>
    /// Convierte en VALOR la primera fórmula de <c>BCE SC POR FACT.</c> cuyo texto cumple el predicado
    /// (reemplazándola por <paramref name="reemplazo"/> cuando se indica). Devuelve la referencia de la
    /// celda mutada para poder exigirla en el mensaje del gate.
    /// </summary>
    private static string Mutar(string ruta, Func<string, bool> predicado, string? reemplazo = null)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = parte.Workbook?.Sheets?.Elements<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, WorkbookLeafCellMapBalanceSc.HojaBce, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException("No se encontró la hoja de balance.");
        var worksheet = ((WorksheetPart)parte.GetPartById(hoja.Id!)).Worksheet
            ?? throw new InvalidOperationException("La hoja de balance no tiene Worksheet válido.");

        foreach (var celda in worksheet.Descendants<Cell>())
        {
            var formula = celda.CellFormula;
            if (formula is null || string.IsNullOrWhiteSpace(formula.Text) || !predicado(formula.Text))
            {
                continue;
            }

            var referencia = celda.CellReference?.Value ?? string.Empty;
            celda.RemoveAllChildren<CellFormula>();

            if (reemplazo is null)
            {
                celda.CellValue = new CellValue("0");
            }
            else
            {
                celda.CellFormula = new CellFormula(reemplazo);
            }

            worksheet.Save();
            return referencia;
        }

        throw new InvalidOperationException("No se encontró ninguna fórmula de BCE que cumpla el predicado del injerto.");
    }

    private sealed class CopiaTemporal : IDisposable
    {
        private readonly string _directorio;

        internal CopiaTemporal(string origen)
        {
            _directorio = Path.Combine(
                Path.GetTempPath(),
                "p59-h5-bce-" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture));
            Directory.CreateDirectory(_directorio);
            Ruta = Path.Combine(_directorio, "base.xlsx");
            File.Copy(origen, Ruta, overwrite: true);
        }

        internal string Ruta { get; }

        public void Dispose()
        {
            try
            {
                Directory.Delete(_directorio, recursive: true);
            }
            catch (IOException)
            {
                // limpieza best-effort del temp
            }
        }
    }
}
