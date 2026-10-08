using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 30 (T2, R-C-3/S3): fail-fast honesto cuando la plantilla NO trae la hoja Det del período.
///
/// Cubre el escenario S3 («base equivocada»): se parte de la plantilla en ceros
/// (<c>Docs/Prueba Julio-2/Plantilla_Remuneracion.xlsx</c>, byte-idéntica al canónico Q2), se renombra la hoja
/// <c>DetRetri2026072</c> en <c>xl/workbook.xml</c> con ZIP+XML puro (BCL, sin Excel/COM ni OpenXML
/// de escritura) y se llama DIRECTAMENTE al writer Q2 (<c>GenerarWorkbook</c>).
/// La resolución de la hoja Det del período (validación protegida del writer y, en el flujo real
/// con espejo, la escritura DetRetri-Q2) comparten el mismo fail-fast honesto de
/// <c>ObtenerHoja</c>.
///
/// Se verifica que el fallo es <c>ERR-PLANTILLA</c> (<see cref="CalculoInvalidoException"/> con
/// <see cref="CodigoError.Plantilla"/>) y que el mensaje NOMBRA la hoja esperada del período
/// (<c>DetRetri2026072</c>), el código de período (<c>2026072</c>), la operación
/// (<c>GenerarWorkbook</c>) y la marca <c>hoja Det</c> — sin rastro de una hoja de otro período
/// (nunca fallback silencioso, nunca escritura parcial).
/// </summary>
public sealed class FailFastHojaDetRetriPeriodoTests
{
    private const string HojaDetRetriVieja = "DetRetri2026072";
    private const string HojaDetRetriRenombrada = "DetRetriRENOMBRADA";
    private const string HojaDetRetriOtroPeriodo = "DetRetri2026082";

    [Fact]
    public void PlantillaSinHojaDetRetriDelPeriodo_FallaErrPlantillaNombrandoHojaPeriodoYOperacion()
    {
        // Mismo armado que DetRetriCapacidadTests: leafs Q2 reales con su DetRetri calculado.
        var (datos, _, leafs) = AjustesSfTTests.LeerDatosQ2ConAjustes();
        foreach (var leaf in leafs)
        {
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

        var resultado = new CalculoRemuneracion().CalcularConsolidado(Insumos.PeriodoQ2(), datos);

        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-s3-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        try
        {
            var plantillaSinDet = Path.Combine(salidaDir, "plantilla-sin-DetRetri-2026072.xlsx");
            File.Copy(Insumos.PlantillaQ2, plantillaSinDet, overwrite: true);
            RenombrarHojaEnWorkbookXml(plantillaSinDet, HojaDetRetriVieja, HojaDetRetriRenombrada);

            var salida = Path.Combine(salidaDir, "salida.xlsx");
            var ex = Assert.Throws<CalculoInvalidoException>(() =>
                new OpenXmlPlantillaWriter().GenerarWorkbook(plantillaSinDet, salida, resultado, leafs));

            // ERR-PLANTILLA: la ausencia de la hoja del período es un problema de plantilla.
            Assert.Equal(CodigoError.Plantilla, ex.Codigo);
            // El mensaje es honesto: hoja esperada del período + código de período + operación.
            Assert.Contains(HojaDetRetriVieja, ex.Message, StringComparison.Ordinal);
            Assert.Contains("2026072", ex.Message, StringComparison.Ordinal);
            Assert.Contains("GenerarWorkbook", ex.Message, StringComparison.Ordinal);
            Assert.Contains("hoja Det", ex.Message, StringComparison.Ordinal);

            // Nombra la hoja del período (no la de otro) y NO cae en fallback a otra hoja.
            Assert.DoesNotContain(HojaDetRetriOtroPeriodo, ex.Message, StringComparison.Ordinal);

            Assert.False(File.Exists(salida), "No debe quedar salida parcial ante fallo (patrón V5).");
        }
        finally
        {
            Borrar(salidaDir);
        }
    }

    /// <summary>
    /// Renombra la <c>name=</c> de una hoja dentro de <c>xl/workbook.xml</c> manipulando el ZIP y
    /// el XML con BCL (<see cref="ZipArchive"/> + <see cref="XDocument"/>), sin Excel/COM y sin el
    /// motor OpenXML. El <c>sheetId</c>/<c>r:id</c> se preservan: solo cambia el nombre visible.
    /// </summary>
    private static void RenombrarHojaEnWorkbookXml(string ruta, string nombreViejo, string nombreNuevo)
    {
        using var zip = ZipFile.Open(ruta, ZipArchiveMode.Update);
        var entrada = zip.GetEntry("xl/workbook.xml")
            ?? throw new InvalidOperationException("La plantilla no trae xl/workbook.xml.");

        string xml;
        using (var lector = new StreamReader(entrada.Open(), Encoding.UTF8))
        {
            xml = lector.ReadToEnd();
        }

        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var doc = XDocument.Parse(xml);
        var hoja = doc.Descendants(ns + "sheet")
            .FirstOrDefault(s => (string?)s.Attribute("name") == nombreViejo)
            ?? throw new InvalidOperationException($"La plantilla no trae la hoja '{nombreViejo}'.");
        hoja.SetAttributeValue("name", nombreNuevo);

        // Recrear la entrada (borrar + crear) garantiza que la longitud quede exacta: escribir
        // sobre la entrada existente en modo Update deja bytes sobrantes del XML original.
        entrada.Delete();
        var nueva = zip.CreateEntry("xl/workbook.xml");
        using var escritor = new StreamWriter(nueva.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        escritor.Write(doc.ToString(SaveOptions.DisableFormatting));
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
            // best-effort (copia temporal)
        }
    }
}
