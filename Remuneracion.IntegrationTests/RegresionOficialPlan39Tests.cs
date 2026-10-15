using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 39 (T5, S8 — R-G-1/R-G-2): regresión end-to-end DESDE LA PLANTILLA OFICIAL.
///
/// La plantilla oficial (<c>Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx</c>)
/// es el INPUT DIRECTO de la App para AMBOS períodos. Phase 0 prepara la base del período sobre la
/// copia de trabajo (3 hojas + 11 tokens + N3 + columna L por quincena — R-INTERV-Q) y el
/// procesamiento es values-only. La regresión corre el flujo 5-ASE para Q1 y Q2 a rutas temp FRESCAS
/// y aserta:
///   - healthy-open (sin <c>xl/calcChain.xml</c>, <c>fullCalcOnLoad="1"</c>, <c>CalculationChainPart</c>
///     ausente — clase de defecto "diálogo de reparación" del Plan 28) + invariante shared R1 (Plan 33);
///   - <c>&lt;f&gt;</c> de <c>CONSOLIDADO_TOTAL RECAUDO!L9:L13</c> == la variante de la quincena
///     (Q1 → N, Q2 → M) + <c>N3 == período</c>;
///   - el oficial queda intacto (SHA256 idéntico);
///   - DetRetri-D vs R10 ±0.5 en los 5 ASE (ambos períodos; Q2 además escrito en la salida).
///
/// Sin Excel/COM; verificación estructural por parser (ZIP + XML). Rutas EXPLÍCITAS (el oficial +
/// insumos reales por ruta explícita + temp FRESCA): prohibido <c>Docs/Insumos</c> raíz.
/// </summary>
public sealed class RegresionOficialPlan39Tests
{
    private const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";
    private const string HojaR1 = "Reporte Componentes R1";
    private const string EntradaCalcChain = "xl/calcChain.xml";
    private const decimal Tolerancia = Insumos.Tolerancia;

    [Theory]
    [InlineData("2026071", "Docs/Prueba Julio-1/Insumos", "N")]
    [InlineData("2026072", "Docs/Prueba Julio-2/Insumos", "M")]
    public void T5_RegresionOficial_HealthyOpen_VarianteL_DetRetriVsR10(string codigo, string carpetaRel, string columna)
    {
        var oficial = Insumos.PlantillaOficial();
        Assert.True(File.Exists(oficial), $"Falta la plantilla oficial: {oficial}");
        var hashAntes = Sha256(oficial);

        var periodo = Periodo.Parse(codigo);
        var salidaDir = Path.Combine(Path.GetTempPath(), "p39-t5-regresion-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, periodo.NombreArchivo);
        try
        {
            var resultado = CrearProcesador().Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = periodo,
                CarpetaPeriodo = Path.Combine(Insumos.Raiz(), carpetaRel),
                RutaPlantilla = oficial,
                RutaSalida = salida
            });

            Assert.True(File.Exists(salida), $"El flujo 5-ASE {codigo} desde el OFICIAL debe certificar la salida.");
            Assert.Equal(5, resultado.Leafs.Count);

            // (1) Healthy-open: sin calcChain + fullCalcOnLoad + CalculationChainPart ausente.
            var estructura = InspeccionarEstructura(salida);
            Assert.False(estructura.ContieneCalcChain, "La salida no debe contener xl/calcChain.xml.");
            Assert.True(estructura.FullCalcOnLoad, "xl/workbook.xml debe traer fullCalcOnLoad=\"1\".");
            Assert.True(estructura.CalculationChainPartAusente, "El WorkbookPart no debe exponer CalculationChainPart.");

            // (2) R-INTERV-Q: L9:L13 == variante de la quincena + N3 == período.
            using (var documento = SpreadsheetDocument.Open(salida, false))
            {
                var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
                var consolidado = Hoja(parte, HojaConsolidado);
                for (var i = 0; i < 5; i++)
                {
                    Assert.Equal($"INTERVENTORIA!{columna}{26 + i}", FormulaDe(consolidado, $"L{9 + i}"));
                }

                Assert.Equal(codigo, ValorDe(consolidado, "N3"));
            }

            // (3) El OFICIAL no se modificó (Phase 0 escribe en la copia de trabajo).
            Assert.Equal(hashAntes, Sha256(oficial));

            // (4) DetRetri-D calculado vs R10 ±0.5 en los 5 ASE (ambos períodos).
            var r10 = new ExcelDataReaderDetRetriR10Reader().LeerDetRetri(periodo, RutaR10(codigo));
            foreach (var leaf in resultado.Leafs)
            {
                var calculado = leaf.DetRetriQ2
                    ?? throw new InvalidOperationException($"ASE {leaf.Ase.Id}: el leaf debe traer DetRetriQ2.");
                Assert.InRange(calculado.Detalle - r10.DetRetriPorAse[leaf.Ase.Id], -Tolerancia, Tolerancia);
            }

            // (5) En Q2 el DetRetri-D se escribe en la salida; se contrasta contra el R10.
            if (periodo.NumeroQuincena == 2)
            {
                var hojaDetRetri = NombresHojaPeriodo.DetRetri(periodo);
                for (var aseId = 1; aseId <= 5; aseId++)
                {
                    var escrito = TestHelpers.LeerCeldaNumerica(salida, hojaDetRetri, $"D{8 + aseId}");
                    Assert.InRange(escrito - r10.DetRetriPorAse[aseId], -Tolerancia, Tolerancia);
                }
            }

            // (6) Red ciega shared R1 (clase "diálogo de reparación" Plan 33): sin seguidoras huérfanas.
            SharedFormulaR1Gate.ExigirInvariante(salida, HojaR1);
        }
        finally
        {
            BorrarArbol(salidaDir);
        }
    }

    private static ProcesadorPeriodo CrearProcesador() =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader());

    private static string RutaR10(string codigo) => codigo == "2026071" ? Insumos.R10 : Insumos.R10Q2;

    private sealed record EstructuraSalida(bool ContieneCalcChain, bool FullCalcOnLoad, bool CalculationChainPartAusente);

    private static EstructuraSalida InspeccionarEstructura(string ruta)
    {
        bool contiene;
        bool full;
        using (var zip = System.IO.Compression.ZipFile.OpenRead(ruta))
        {
            contiene = zip.Entries.Any(e => string.Equals(e.FullName, EntradaCalcChain, StringComparison.OrdinalIgnoreCase));
            var wb = zip.GetEntry("xl/workbook.xml")
                ?? throw new InvalidOperationException("La salida no tiene xl/workbook.xml.");
            using var lector = new StreamReader(wb.Open());
            var xml = lector.ReadToEnd();
            full = xml.Contains("fullCalcOnLoad=\"1\"", StringComparison.Ordinal);
        }

        bool partAusente;
        using (var documento = SpreadsheetDocument.Open(ruta, false))
        {
            var wp = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            partAusente = wp.CalculationChainPart is null;
        }

        return new EstructuraSalida(contiene, full, partAusente);
    }

    private static string Sha256(string ruta) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(ruta)));

    private static Worksheet Hoja(WorkbookPart parte, string nombre)
    {
        var sheet = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)parte.GetPartById(sheet.Id!)).Worksheet!;
    }

    private static Cell? Celda(Worksheet hoja, string referencia) =>
        hoja.Descendants<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));

    private static string? FormulaDe(Worksheet hoja, string referencia) => Celda(hoja, referencia)?.CellFormula?.Text;

    private static string ValorDe(Worksheet hoja, string referencia) => Celda(hoja, referencia)?.CellValue?.InnerText?.Trim() ?? string.Empty;

    private static void BorrarArbol(string ruta)
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
