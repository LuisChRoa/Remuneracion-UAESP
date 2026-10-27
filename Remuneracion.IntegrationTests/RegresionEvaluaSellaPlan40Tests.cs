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
/// Plan 40 (T4, R-G-2/R-R-1 — enmienda E-1): regresión end-to-end DESDE LA PLANTILLA VIGENTE con
/// el gate W-1 EVALUADOR activo (<see cref="ValidacionOracleReader"/> inyectado), Q1 y Q2, 5 ASE,
/// a rutas temp FRESCAS.
///
/// <para>
/// <b>Qué cubre hoy (Plan 59 §15, 2026-10-27):</b> el "defecto espejo declarado" (E1-2) que este
/// fixture caracterizaba ya NO existe en la plantilla vigente: la v4 corrige el typo de fórmula
/// <c>Reporte Componentes R1!F328</c> (<c>−M213</c> → <c>−L213</c>) y los sub-bloques del espejo de
/// remuneración evalúan TRUE. Por eso el caso Q1 exige ahora el comportamiento CORRECTO (la corrida
/// cierra con el nombre definitivo y los 5 sub-bloques sellados en <c>1</c>), no el rojo histórico.
/// El caso Q2 conserva su objeto original: el gate se detiene HONESTAMENTE ante la divergencia
/// declarada SINTÉTICA (<c>SINTETICA-C15</c>, Δ = 1) y el sellado se prueba en ambos sentidos.
/// </para>
///
/// <para>
/// <b>Evidencia del cambio de expectativa (no relajación):</b> la corrida real <c>2026072</c> con la
/// v4 da <c>RESULTADO OK</c> (exit 0) y <c>2026071</c> cierra contra el R10 oficial
/// (<c>R10_Remuneracion_2026071.xlsx</c>), con <c>VALIDACION_TOTAL O9 = 0 / P9 = TRUE</c> en los 5 ASE
/// (§14.3 del Plan 59). La plantilla se re-verifica byte-intacta al final de cada caso.
/// </para>
/// </summary>
public sealed class RegresionEvaluaSellaPlan40Tests
{
    private const string Hoja = "VALIDACION_TOTAL";

    [Theory]
    [InlineData("2026071", "Docs/Prueba Julio-1/Insumos")]
    [InlineData("2026072", "Docs/Prueba Julio-2/Insumos")]
    public void T4_EndToEnd_PlantillaVigente_EvaluaSella(string codigo, string carpetaRel)
    {
        var oficial = Insumos.PlantillaOficialV4();
        Assert.True(File.Exists(oficial), $"Falta la plantilla vigente: {oficial}");
        var hashAntes = Sha256(oficial);

        var periodo = Periodo.Parse(codigo);
        var salidaDir = Path.Combine(Path.GetTempPath(), "p40-t4-e2e-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, periodo.NombreArchivo);
        try
        {
            if (periodo.NumeroQuincena == 2)
            {
                // ── Q2 (2026072) — Plan 46 (H2, decisión del Ingeniero 2026-10-17) ──────────────
                // La divergencia REAL del espejo quedó RESUELTA (O25/D34/F34 TRUE) y su entrada se
                // retiró del catálogo. El gate sigue deteniéndose HONESTAMENTE con la divergencia
                // declarada SINTÉTICA (SINTETICA-C15, Δ = 1) sin reconocimiento: el sellado se prueba
                // en ambos sentidos (FALSE→0, TRUE→1).
                var procesadorSintetico = FixtureDivergenciaSintetica.CrearProcesador();
                var exSintetico = Assert.Throws<CalculoInvalidoException>(() => procesadorSintetico.Ejecutar(new SolicitudProcesoPeriodo
                {
                    Periodo = periodo,
                    CarpetaPeriodo = Path.Combine(Insumos.Raiz(), carpetaRel),
                    RutaPlantilla = oficial,
                    RutaSalida = salida
                }));

                Assert.Contains("[ERR-VALIDACION]", exSintetico.Message, StringComparison.Ordinal);
                Assert.Contains($"VALIDACION_TOTAL · {FixtureDivergenciaSintetica.Celda}", exSintetico.Message, StringComparison.Ordinal);
                // Los 3 sub-bloques del espejo de remuneración YA NO divergen (Plan 46).
                foreach (var celda in new[] { "O25", "D34", "F34" })
                {
                    Assert.DoesNotContain($"VALIDACION_TOTAL · {celda}", exSintetico.Message, StringComparison.Ordinal);
                }

                Assert.Equal(5, exSintetico.Message.Split("ASE ").Length - 1); // 5 ASE × 1 celda
                Assert.Equal(hashAntes, Sha256(oficial));

                var rutaMarcadaQ2 = NombresSalida.RutaSinValidar(salida);
                Assert.True(File.Exists(rutaMarcadaQ2), "El writer debió producir la salida marcada 'sin validar' antes del gate.");
                Assert.False(File.Exists(salida), "Una corrida no cerrada no produce el nombre definitivo.");
                Assert.Equal("0", Valor(rutaMarcadaQ2, Hoja, "C15"));
                foreach (var celda in new[] { "D25", "O25", "D34", "F34" })
                {
                    Assert.Equal("1", Valor(rutaMarcadaQ2, Hoja, celda));
                }

                // La evaluación independiente del gate coincide con el sellado (sellado == evaluación).
                var snapshotsQ2 = new ValidacionOracleReader().LeerSnapshots(rutaMarcadaQ2, periodo);
                foreach (var snapshot in snapshotsQ2)
                {
                    Assert.False(snapshot.SubBloquesValidacionTotal["C15"]);
                    Assert.True(snapshot.SubBloquesValidacionTotal["D25"]);
                    Assert.True(snapshot.SubBloquesValidacionTotal["O25"]);
                    Assert.True(snapshot.SubBloquesValidacionTotal["D34"]);
                    Assert.True(snapshot.SubBloquesValidacionTotal["F34"]);
                }

                return;
            }

            // ── Q1 (2026071) — Plan 59 (§15): la plantilla vigente CIERRA ─────────────────────────
            // El bloqueo histórico (ERR-SIN-SLOT por las 5 fronteras fijas de la v2 + el typo de
            // `R1!F328` que dejaba el espejo en FALSE) era de PLANTILLA y quedó cerrado por la v4:
            // la corrida cierra y el entregable lleva el nombre definitivo (no el marcado
            // "sin validar"). No se relaja nada: se exige el resultado verdadero —los 5 sub-bloques
            // evalúan y sellan TRUE— porque la verdad cambió (oráculo: R10 de julio-1).
            var procesador = new ProcesadorPeriodo(
                new ExcelDataReaderRecaudoReader(),
                new ExcelDataReaderWorkbookLeafInputReader(),
                new CalculoRemuneracion(),
                new ValidadorBasico(),
                new OpenXmlPlantillaWriter(),
                new ArchivoFuenteLocator(),
                new ExcelDataReaderDetRetriR10Reader(),
                new ValidacionOracleReader());

            var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = periodo,
                CarpetaPeriodo = Path.Combine(Insumos.Raiz(), carpetaRel),
                RutaPlantilla = oficial,
                RutaSalida = salida
            });

            Assert.NotEmpty(resultado.Validaciones);

            // El OFICIAL no se modificó (Phase 0 escribe en la copia de trabajo).
            Assert.Equal(hashAntes, Sha256(oficial));

            // Plan 43 (O-4): una corrida que cierra bien se promueve al nombre DEFINITIVO.
            Assert.True(File.Exists(salida), "Con la plantilla vigente la corrida cierra y produce el nombre definitivo.");

            // Sellado == evaluación: los 5 sub-bloques quedan en 1.
            foreach (var celda in new[] { "C15", "D25", "O25", "D34", "F34" })
            {
                Assert.Equal("1", Valor(salida, Hoja, celda));
            }

            var snapshotsQ1 = new ValidacionOracleReader().LeerSnapshots(salida, periodo);
            Assert.Equal(5, snapshotsQ1.Count);
            foreach (var snapshot in snapshotsQ1)
            {
                foreach (var celda in new[] { "C15", "D25", "O25", "D34", "F34" })
                {
                    Assert.True(snapshot.SubBloquesValidacionTotal[celda], $"VALIDACION_TOTAL · {celda} debe cerrar con la plantilla vigente.");
                }
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
