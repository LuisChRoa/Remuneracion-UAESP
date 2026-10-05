using System.Text;
using Remuneracion.Core.Services;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 27 (T3, R-F-4/D-C): clasificador de firma puro (bytes → veredicto, sin archivos) y
/// lectura acotada de archivos reales/temporales. Regla mínima: PK (xlsx) y OLE (xls legacy) son
/// válidos; el texto con marcadores web es "copia de página web"; el resto es formato desconocido.
/// </summary>
public sealed class InspectorFirmaExcelTests
{
    [Fact]
    public void Clasificar_FirmaXlsxPk_EsValida()
    {
        Assert.Equal(FirmaExcel.Valida, InspectorFirmaExcel.Clasificar([0x50, 0x4B, 0x03, 0x04, 0x14, 0x00]));
    }

    [Fact]
    public void Clasificar_FirmaXlsOle_EsValida()
    {
        // D-D: el CreateReader genérico del proyecto abre xls legacy; se acepta como válido.
        Assert.Equal(FirmaExcel.Valida, InspectorFirmaExcel.Clasificar([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1]));
    }

    [Theory]
    [InlineData("<!DOCTYPE html><html><head></head></html>")]
    [InlineData("<html><body>hola</body></html>")]
    [InlineData("From: <Saved by Blink>\r\nMIME-Version: 1.0\r\nContent-Type: multipart/related;")]
    [InlineData("MIME-Version: 1.0")]
    [InlineData("Saved by Blink")]
    public void Clasificar_TextoWeb_EsCopiaDeWeb(string contenido)
    {
        Assert.Equal(FirmaExcel.NoExcelPareceCopiaWeb, InspectorFirmaExcel.Clasificar(Encoding.UTF8.GetBytes(contenido)));
    }

    [Fact]
    public void Clasificar_TextoWebConBom_EsCopiaDeWeb()
    {
        var bytes = new byte[] { 0xEF, 0xBB, 0xBF }.Concat(Encoding.UTF8.GetBytes("<html></html>")).ToArray();
        Assert.Equal(FirmaExcel.NoExcelPareceCopiaWeb, InspectorFirmaExcel.Clasificar(bytes));
    }

    [Fact]
    public void Clasificar_OtraFirma_EsFormatoDesconocido()
    {
        Assert.Equal(FirmaExcel.NoExcelFormatoDesconocido, InspectorFirmaExcel.Clasificar(Encoding.ASCII.GetBytes("NO_ES_UN_EXCEL")));
    }

    [Fact]
    public void Clasificar_Vacio_EsFormatoDesconocido()
    {
        Assert.Equal(FirmaExcel.NoExcelFormatoDesconocido, InspectorFirmaExcel.Clasificar([]));
    }

    [Fact]
    public void Inspeccionar_ArchivoXlsxReal_EsValida()
    {
        // Insumo real de disco: la conciliación de Reciprocidad de agosto (multi-hoja).
        Assert.Equal(FirmaExcel.Valida, InspectorFirmaExcel.Inspeccionar(Insumos.Conciliacion(Insumos.CarpetaInsumosAgosto, "Conjunta Recip")).Firma);
    }

    [Fact]
    public void Inspeccionar_MhtmlTemporal_EsCopiaDeWeb()
    {
        var ruta = Path.Combine(Path.GetTempPath(), "remuneracion-mhtml-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            File.WriteAllText(ruta, MhtmlBlink());
            var resultado = InspectorFirmaExcel.Inspeccionar(ruta);
            Assert.Equal(FirmaExcel.NoExcelPareceCopiaWeb, resultado.Firma);
            // Los primeros bytes se devuelven para el log técnico (nunca al mensaje administrativo).
            Assert.NotEmpty(resultado.PrimerosBytes);
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    [Fact]
    public void Inspeccionar_ArchivoInexistente_EsIlegible()
    {
        var ruta = Path.Combine(Path.GetTempPath(), "remuneracion-no-existe-" + Guid.NewGuid().ToString("N") + ".xlsx");
        Assert.Equal(FirmaExcel.Ilegible, InspectorFirmaExcel.Inspeccionar(ruta).Firma);
    }

    /// <summary>MHTML real (patrón "Saved by Blink") reutilizado por los tests de preflight.</summary>
    internal static string MhtmlBlink() =>
        "From: <Saved by Blink>\r\n" +
        "Snapshot-Content-Location: https://www.example.com/reporte\r\n" +
        "Subject: Descarga de reporte\r\n" +
        "Date: Mon, 7 Sep 2026 10:00:00 -0500\r\n" +
        "MIME-Version: 1.0\r\n" +
        "Content-Type: multipart/related; type=\"text/html\"; boundary=\"----MultipartBoundary--abc\"\r\n" +
        "\r\n" +
        "------MultipartBoundary--abc\r\n" +
        "Content-Type: text/html\r\n" +
        "Content-Transfer-Encoding: quoted-printable\r\n" +
        "\r\n" +
        "<!DOCTYPE html><html><head><meta charset=\"utf-8\"></head><body>reporte</body></html>\r\n";
}
