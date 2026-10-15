using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 40 (T4, R-G-2/R-R-1 — enmienda E-1): regresión end-to-end DESDE LA PLANTILLA OFICIAL con
/// el gate W-1 EVALUADOR activo (<see cref="ValidacionOracleReader"/> inyectado), Q1 y Q2, 5 ASE,
/// a rutas temp FRESCAS.
///
/// <para>
/// <b>Veredicto (E1-2/E1-3):</b> tras revertir <c>SeGatea=false</c>, la corrida evalúa honestamente
/// los 5 sub-bloques. El recaudo cierra (<c>C15</c>/<c>D25</c> TRUE — el bloqueante Q1 de <c>C3:C8</c>
/// se resolvió en T3 con la columna D); el <b>espejo de remuneración NO cierra</b>
/// (<c>O25</c>/<c>D34</c>/<c>F34</c> FALSE) — <b>DEFECTO REAL E1-2</b>: el archivo entregado mostraría
/// validaciones en FALSE tras el recálculo de Excel. T4 se DETIENE aquí y lo declara con los números
/// exactos (no se apaga, no se sella un TRUE cosmético). El frente propio (espejo R1 / remuneración)
/// queda fuera del Plan 40 (E1-2).
/// </para>
///
/// <para>
/// El writer SÍ produce la salida (el gate corre DESPUÉS, en el orquestador): el sellado de cachés
/// deja el archivo autoconsistente — sellado == evaluación (E1-2/D-E), nunca un TRUE inventado.
/// </para>
/// </summary>
public sealed class RegresionEvaluaSellaPlan40Tests
{
    private const string Hoja = "VALIDACION_TOTAL";

    [Theory]
    [InlineData("2026071", "Docs/Prueba Julio-1/Insumos")]
    [InlineData("2026072", "Docs/Prueba Julio-2/Insumos")]
    public void T4_EndToEnd_Oficial_EvaluaSella_DefectoEspejoDeclarado(string codigo, string carpetaRel)
    {
        var oficial = Insumos.PlantillaOficial();
        Assert.True(File.Exists(oficial), $"Falta la plantilla oficial: {oficial}");
        var hashAntes = Sha256(oficial);

        var periodo = Periodo.Parse(codigo);
        var salidaDir = Path.Combine(Path.GetTempPath(), "p40-t4-e2e-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, periodo.NombreArchivo);
        try
        {
            var procesador = new ProcesadorPeriodo(
                new ExcelDataReaderRecaudoReader(),
                new ExcelDataReaderWorkbookLeafInputReader(),
                new CalculoRemuneracion(),
                new ValidadorBasico(),
                new OpenXmlPlantillaWriter(),
                new ArchivoFuenteLocator(),
                new ExcelDataReaderDetRetriR10Reader(),
                new ValidacionOracleReader());

            // E1-3: T4 NO declara éxito — el gate falla honestamente por el defecto E1-2.
            var ex = Assert.Throws<CalculoInvalidoException>(() => procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = periodo,
                CarpetaPeriodo = Path.Combine(Insumos.Raiz(), carpetaRel),
                RutaPlantilla = oficial,
                RutaSalida = salida
            }));

            Assert.Contains("[ERR-VALIDACION]", ex.Message, StringComparison.Ordinal);

            // El recaudo (C15/D25) NO aparece en el error: cierra en ambas quincenas (E1-4 resuelto).
            Assert.DoesNotContain("VALIDACION_TOTAL · C15", ex.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("VALIDACION_TOTAL · D25", ex.Message, StringComparison.Ordinal);

            // El espejo de remuneración SÍ aparece (E1-2): O25/D34/F34 en los 5 ASE.
            foreach (var celda in new[] { "O25", "D34", "F34" })
            {
                Assert.Contains($"VALIDACION_TOTAL · {celda}", ex.Message, StringComparison.Ordinal);
            }

            // 5 ASE × 3 sub-bloques = 15 errores (el defecto se declara por ASE y por celda).
            Assert.Equal(15, ex.Message.Split("ASE ").Length - 1);

            // El OFICIAL no se modificó (Phase 0 escribe en la copia de trabajo).
            Assert.Equal(hashAntes, Sha256(oficial));

            // Plan 43 (O-4): la corrida no cerró bien (Q2: declarada pendiente de reconocimiento;
            // Q1: ERR-SIN-SLOT antes de escribir) ⇒ el archivo NO lleva el nombre definitivo. En Q2
            // el writer corrió y dejó el archivo MARCADO "sin validar" — autoconsistente: sellado ==
            // evaluación (C15/D25 = TRUE; O25/D34/F34 = FALSE); en Q1 la aserción previa ya falla por
            // ERR-SIN-SLOT (antes de escribir) y no hay archivo alguno.
            var rutaMarcada = NombresSalida.RutaSinValidar(salida);
            Assert.True(File.Exists(rutaMarcada), "El writer debió producir la salida marcada 'sin validar' antes del gate.");
            Assert.False(File.Exists(salida), "Una corrida no cerrada no produce el nombre definitivo.");
            Assert.Equal("1", Valor(rutaMarcada, Hoja, "C15"));
            Assert.Equal("1", Valor(rutaMarcada, Hoja, "D25"));
            foreach (var celda in new[] { "O25", "D34", "F34" })
            {
                Assert.Equal("0", Valor(rutaMarcada, Hoja, celda));
            }

            // La evaluación independiente del gate coincide con el sellado (sellado == evaluación).
            var snapshots = new ValidacionOracleReader().LeerSnapshots(rutaMarcada, periodo);
            foreach (var snapshot in snapshots)
            {
                Assert.True(snapshot.SubBloquesValidacionTotal["C15"]);
                Assert.True(snapshot.SubBloquesValidacionTotal["D25"]);
                Assert.False(snapshot.SubBloquesValidacionTotal["O25"]);
                Assert.False(snapshot.SubBloquesValidacionTotal["D34"]);
                Assert.False(snapshot.SubBloquesValidacionTotal["F34"]);
            }
        }
        finally
        {
            try
            {
                Directory.Delete(salidaDir, recursive: true);
            }
            catch
            {
                // best-effort
            }
        }
    }

    private static string Sha256(string ruta) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(ruta)));

    private static string? Valor(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        return ws.Descendants<Cell>()
            .First(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            .CellValue?.InnerText;
    }
}
