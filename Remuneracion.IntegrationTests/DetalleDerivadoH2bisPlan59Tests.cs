using DocumentFormat.OpenXml.Packaging;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (H2-bis): las protecciones de <c>DetRetri</c>/<c>DetValiRetri</c> se verifican por ROL +
/// FORMA sobre la gramática que la plantilla v3 declara (tres familias de filas por ASE + rótulos de
/// columna), no por la lista congelada de celdas y operandos de la v2.
///
/// Los dos casos que importan:
/// <list type="bullet">
/// <item><b>verde</b>: la base preparada del período 2026072 con la v3 pasa la verificación (antes
/// fallaba exigiendo <c>['F75','E47','D73','F77']</c> en <c>D32</c>);</item>
/// <item><b>rojo</b>: si una celda protegida deja de ser la fórmula de su rol, la corrida aborta
/// nombrando la identidad (hoja · familia · ASE · rótulo), no una coordenada.</item>
/// </list>
/// </summary>
public sealed class DetalleDerivadoH2bisPlan59Tests
{
    private static readonly string[] HojasBase = ["DetRetri2026072", "DetValiRetri2026072"];

    [Fact]
    public void Bases2026072_ConLaV3_PasanLaVerificacionPorRolYForma()
    {
        var ruta = BaseDelPeriodo();
        using var documento = SpreadsheetDocument.Open(ruta, isEditable: false);
        var parte = documento.WorkbookPart!;

        DetallePeriodoDerivado.Validar(parte, Periodo.Parse("2026072"));
    }

    [Fact]
    public void Base2026072_SiUnaCeldaProtegidaPierdeSuFormula_FallaNombrandoLaIdentidad()
    {
        var ruta = CopiaDeLaBase();
        try
        {
            using (var documento = SpreadsheetDocument.Open(ruta, isEditable: true))
            {
                var parte = documento.WorkbookPart!;
                var hoja = HojaFisica(parte, "DetRetri2026072");
                var celda = ResolvedorDestinoValor.BuscarCelda(hoja, "D32")
                    ?? throw new InvalidOperationException("La base preparada no declara DetRetri2026072!D32.");

                // Injerto: la celda de verificación del ASE 1 deja de ser la comparación de composición.
                celda.CellFormula = null;
                celda.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("0");
                hoja.Save();
            }

            using var reabierto = SpreadsheetDocument.Open(ruta, isEditable: false);
            var excepcion = Assert.Throws<CalculoInvalidoException>(
                () => DetallePeriodoDerivado.Validar(reabierto.WorkbookPart!, Periodo.Parse("2026072")));

            Assert.Contains("DetRetri2026072", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("familia", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("ASE 1", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("RECAUDO TOTAL", excepcion.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    private static string BaseDelPeriodo()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"p59-detalle-{Environment.ProcessId}");
        Directory.CreateDirectory(dir);
        var salida = Path.Combine(dir, "Plantilla_Remuneracion_2026072.xlsx");
        if (!File.Exists(salida))
        {
            PreparadorBasePipeline.Ejecutar(PlantillaV3(), Periodo.Parse("2026072"), salida);
        }

        return salida;
    }

    private static string CopiaDeLaBase()
    {
        var origen = BaseDelPeriodo();
        var copia = Path.Combine(Path.GetDirectoryName(origen)!, $"copia-{Guid.NewGuid():N}.xlsx");
        File.Copy(origen, copia, overwrite: true);
        return copia;
    }

    private static string PlantillaV3() =>
        Path.Combine(Insumos.Raiz(), "Docs", "Plantilla estandarizada", "v3", "Formato Remuneracion 2026xx-x Total v3.xlsx");

    private static DocumentFormat.OpenXml.Spreadsheet.Worksheet HojaFisica(WorkbookPart parte, string nombre)
    {
        var hoja = parte.Workbook!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)parte.GetPartById(hoja.Id!)).Worksheet!;
    }

    private static IReadOnlyList<string> Hojas() => HojasBase;
}
