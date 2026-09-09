using System.IO;
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
/// HU-15 (W-2.3, D5): códigos explícitos del writer. Cada fallo de plantilla/escritura porta el
/// código correcto (ERR-PLANTILLA → salida 2 / ERR-ESCRITURA → salida 3 con InnerException), la
/// atomicidad (borrar parcial) se conserva y los gates de coherencia siguen ERR-VALIDACION (1).
/// Solo usa copias en temp; <c>Docs/Insumos/</c> jamás es destino.
/// </summary>
public sealed class EscrituraCodigosTests
{
    [Fact]
    public void PlantillaAusente_CodigoEsPlantilla()
    {
        var writer = new OpenXmlPlantillaWriter();
        var salida = Path.Combine(Path.GetTempPath(), "remuneracion-w2-3-missing-" + Guid.NewGuid().ToString("N"), "salida.xlsx");

        var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
            writer.GenerarWorkbook(Path.Combine(Path.GetTempPath(), "no-existe.xlsx"), salida, new ResultadoRemuneracion(), new WorkbookLeafInputs()));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Equal(CodigosSalida.FuenteOPlantilla, CatalogoErrores.CodigoSalidaPara(ex.Codigo));
    }

    [Fact]
    public void SalidaIgualPlantilla_CodigoEsPlantilla_YPlantillaIntacta()
    {
        var dir = Path.Combine(Path.GetTempPath(), "remuneracion-w2-3-inplace-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var plantilla = Path.Combine(dir, "plantilla.xlsx");
        File.Copy(Insumos.Plantilla, plantilla, overwrite: true);
        var hashAntes = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(plantilla)));

        var writer = new OpenXmlPlantillaWriter();
        // ValidarNoInPlace corre ANTES de la coherencia: un resultado/leaf vacío basta para llegar.
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            writer.GenerarWorkbook(plantilla, plantilla, new ResultadoRemuneracion(), new WorkbookLeafInputs()));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Equal(CodigosSalida.FuenteOPlantilla, CatalogoErrores.CodigoSalidaPara(ex.Codigo));
        var hashDespues = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(plantilla)));
        Assert.Equal(hashAntes, hashDespues); // la plantilla no se mutó
    }

    [Fact]
    public void EstructuraCorrupta_ValorFijoEnFormulaProtegida_CodigoEsPlantilla_SinParcial()
    {
        var (resultado, leaf) = CrearCasoValidoAse1();
        var dir = Path.Combine(Path.GetTempPath(), "remuneracion-w2-3-estruct-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var plantillaRota = Path.Combine(dir, "plantilla-rota.xlsx");
        File.Copy(Insumos.Plantilla, plantillaRota, overwrite: true);
        FijarValor(plantillaRota, "CONSOLIDADO_TOTAL RECAUDO", "D9", "123");

        var writer = new OpenXmlPlantillaWriter();
        var salida = Path.Combine(dir, "salida.xlsx");

        // La coherencia del leaf pasa (inputs reales); la estructura rota de la copia se detecta
        // en ValidarFormulasProtegidas (D9 con valor fijo) → ERR-PLANTILLA y parcial borrado.
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            writer.GenerarWorkbook(plantillaRota, salida, resultado, leaf));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.False(File.Exists(salida), "No debe quedar archivo parcial ante estructura rota.");
    }

    [Fact]
    public void IoBloqueadoEnCopia_CodigoEsEscritura_ConInnerException()
    {
        var (resultado, leaf) = CrearCasoValidoAse1();

        var dir = Path.Combine(Path.GetTempPath(), "remuneracion-w2-3-io-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var salida = Path.Combine(dir, "salida.xlsx");
        // Bloqueo exclusivo del destino: File.Copy lanza IOException (sharing violation).
        using (var bloqueo = new FileStream(salida, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
        {
            var writer = new OpenXmlPlantillaWriter();
            var ex = Assert.Throws<CalculoInvalidoException>(() =>
                writer.GenerarWorkbook(Insumos.Plantilla, salida, resultado, leaf));

            Assert.Equal(CodigoError.Escritura, ex.Codigo);
            Assert.Equal(CodigosSalida.Escritura, CatalogoErrores.CodigoSalidaPara(ex.Codigo));
            Assert.IsType<IOException>(ex.InnerException);
        }

        // El borrado best-effort no pudo eliminar el archivo bloqueado; al liberar el handle,
        // el contenido sigue siendo el placeholder del test (la copia nunca ocurrió).
        Assert.True(File.Exists(salida));
    }

    /// <summary>
    /// Caso single-ASE Q1 coherente (ASE 1) para llegar a las fases de copia/escritura del writer.
    /// </summary>
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

    /// <summary>
    /// Convierte una celda con fórmula en valor fijo (simula plantilla "recalculada a mano" o
    /// corrupta) y guarda el workbook. Read-only del golden; la copia vive en temp.
    /// </summary>
    private static void FijarValor(string ruta, string hoja, string celda, string valor)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var workbookXml = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook null");
        var sheet = workbookXml.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Falta hoja {hoja}");
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException($"Falta Worksheet de {hoja}");
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Falta celda {hoja}!{celda}");
        cell.CellFormula = null;
        cell.DataType = CellValues.String;
        cell.CellValue = new CellValue(valor);
        workbook.Save();
    }
}