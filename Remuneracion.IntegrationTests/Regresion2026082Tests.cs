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
/// Plan 25 (WU-2 = T3, R-R-2/D-E): regresión END-TO-END del período 2026082 (agosto Q2) modo
/// 5 ASE vía <see cref="ProcesadorPeriodo"/>, con los insumos REALES de <c>Docs/Prueba Agosto-2/Insumos</c>.
///
/// Contexto y límites (decisión del Ingeniero, memoria <c>verificacion/2026082-sin-oraculo</c>):
/// UAESP no entregó plantilla ni salida de agosto; el R10 SÍ existe
/// (<c>R10_Remuneracion_2026082.xlsx</c>, hoja <c>DetRetri2026082</c>, D9:D13) y es el ÚNICO
/// oráculo de agosto (D-E). Por eso la verificación es:
///   - DetRetri calculado de los 5 ASE == R10 D9:D13 ±0.5 post-redondeo (ASE3 = 16.369.059.896);
///   - invariantes duras de cierre T0e por bloque espejo;
///   - Δ por bloque = {ASE1 −3, ASE2 −9, ASE3 −6, ASE4 +6, ASE5 +8} y fórmulas preservadas.
///   - sin asserts de negocio contra salida de agosto fuera del R10 (prohibido inventar golden).
///
/// La plantilla destino es la base canónica de agosto
/// (<c>Docs/Prueba Agosto-2/Plantilla_Remuneracion.xlsx</c>,
/// hojas <c>DetRetri2026082</c>/<c>DetValiRetri2026082</c>): el espejo R1 la redimensiona a la
/// forma de agosto. <c>Docs/Prueba Agosto-2/Insumos</c> NO trae <c>Conciliaciones/</c> (limitación
/// declarada); se aporta una copia temporal con las conciliaciones REALES del canónico Q2 de julio
/// (mismo layout RESUMEN MES, G2-D2).
///
/// Los blockers 2.5 (SALDOS POR NOTA, Plan 23) y R1-Q2 (Plan 25, WU-1) quedaron RESUELTOS: el
/// flujo 5-ASE de agosto corre END-TO-END sin <c>ERR-VALIDACION</c>. Regresión dura: ninguno de
/// esos fail-fast debe reaparecer.
/// </summary>
public sealed class Regresion2026082Tests
{
    private static readonly int[] ConteoPlantillaQ2 = [45, 75, 43, 73, 52];
    private static readonly int[] ConteoAgosto = [42, 66, 37, 79, 60];
    private static readonly int[] DeltasEsperados = [-3, -9, -6, +6, +8];
    private static readonly string[] NombresAse = ["PROMOAMBIENTAL", "LIME", "CIUDAD LIMPIA", "BOGOTA LIMPIA", "AREA LIMPIA"];

    /// <summary>
    /// Oráculo R10 de agosto (único de agosto): DetRetri D9:D13 por ASE (T0b §3.4). ASE3 = el
    /// blocker histórico; el resto cierra como el resto de períodos.
    /// </summary>
    private static readonly decimal[] DetRetriR10Agosto =
    [
        18378829331m, 21599709648m, 16369059896m, 8372092112m, 12137660178m
    ];

    [Fact]
    public void Ejecutar_Periodo2026082_Modo5Ase_EndToEnd_DetRetri5De5VsR10()
    {
        var carpetaPeriodo = PrepararPeriodoAgosto();
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-2026082-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, "Remuneracion2026082.xlsx");
        try
        {
            var procesador = CrearProcesador();
            Exception? fallo = null;
            ResultadoProcesoPeriodo? resultado = null;
            try
            {
                resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
                {
                    Periodo = new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 },
                    CarpetaPeriodo = carpetaPeriodo,
                    RutaPlantilla = Insumos.PlantillaAgosto2026082,
                    RutaSalida = salida
                });
            }
            catch (Exception ex)
            {
                fallo = ex;
            }

            // Regresión dura: los blockers históricos 2.5 (SALDOS POR NOTA) y R1-Q2 quedaron
            // resueltos. Correr sin excepción es la garantía de que no reaparecen; si algo
            // fallara, el mensaje del catch de abajo nombra el blocker que volvió.
            if (fallo is not null)
            {
                Assert.Fail(
                    $"El flujo 5-ASE de agosto debe correr END-TO-END (blockers 2.5 y R1-Q2 resueltos). Falló: {fallo.GetType().Name}: {fallo.Message}");
            }

            Assert.NotNull(resultado);
            Assert.Equal(5, resultado!.Resultado.Consolidados.Count);
            Assert.Equal(5, resultado.Leafs.Count);
            Assert.True(File.Exists(salida), "El flujo 5-ASE de agosto debe certificar la salida.");

            // D-E: DetRetri calculado (ROUND(D104:D108,0)) de los 5 ASE == R10 D9:D13 ±0.5
            // post-redondeo. Es el ÚNICO oráculo de agosto; ASE3 = 16.369.059.896.
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var leaf = resultado.Leafs.Single(l => l.Ase.Id == aseId);
                var detretri = leaf.DetRetriQ2 ?? throw new InvalidOperationException($"ASE {aseId}: falta DetRetriQ2.");
                Assert.InRange(detretri.Detalle - DetRetriR10Agosto[aseId - 1], -Insumos.Tolerancia, Insumos.Tolerancia);
            }

            // Invariantes de cierre T0e por ASE (la fuente define la forma; no se simulan filas).
            foreach (var leaf in resultado.Leafs)
            {
                Assert.NotNull(leaf.EspejoR1);
                Assert.True(leaf.EspejoR1!.TieneInvariantesDeCierre,
                    $"ASE {leaf.Ase.Id}: el espejo debe traer las 3 invariantes de cierre T0e.");
            }

            // Δ por bloque = fuente agosto − bloque plantilla Q2 (45/75/43/73/52 − 42/66/37/79/60).
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var filasBloque = ContarFilasBloqueEnSalida(salida, NombresAse[aseId - 1]);
                Assert.Equal(ConteoAgosto[aseId - 1], filasBloque);
                Assert.Equal(DeltasEsperados[aseId - 1], ConteoAgosto[aseId - 1] - ConteoPlantillaQ2[aseId - 1]);
            }

            // Fórmulas preservadas: el dimensionado del espejo no crea ni destruye fórmulas en R1.
            // Excepción única del contrato T2 (Plan 31, §2.3 / escenario S5): el EXTEMP del
            // ASE3-agosto (0 filas Aplicacion) se escribe como literal 0 auditado, en vez de
            // conservar la fórmula del template (que apuntaba a filas de otro rol). Por eso el
            // conteo baja EXACTAMENTE en 1 (base 1155 -> salida 1154). No es regresión: es el
            // contrato funcionando.
            var formulasBase = ContarFormulas(Insumos.PlantillaAgosto2026082, "Reporte Componentes R1");
            var formulasSalida = ContarFormulas(salida, "Reporte Componentes R1");
            Assert.Equal(formulasBase - 1, formulasSalida);

            // El único visible que cambió: Reporte Componentes R1!F327 (EXTEMP ASE3-agosto) es
            // literal 0 sin <f> (S5); el resto de los visibles conserva su fórmula.
            Assert.True(
                EsLiteralSinFormula(salida, "Reporte Componentes R1", "F327", "0"),
                "Reporte Componentes R1!F327 (EXTEMP ASE3-agosto) debe quedar como literal 0 sin <f> (contrato T2 §2.3/S5).");

            // Invariante fuerte de "ninguna otra fórmula desapareció": el conteo de fórmulas POR
            // COLUMNA es invariante a la geometría del espejo, que desplaza filas (nunca columnas);
            // el set de referencias absolutas base-vs-salida NO es comparable porque cada bloque R1
            // cambia de filas al dimensionar. El único delta admisible es la columna F = base - 1
            // (F327). Cualquier otra columna que cambiara su conteo delataría una fórmula creada o
            // destruida fuera del contrato T2.
            var porColumnaBase = ContarFormulasPorColumna(Insumos.PlantillaAgosto2026082, "Reporte Componentes R1");
            var porColumnaSalida = ContarFormulasPorColumna(salida, "Reporte Componentes R1");
            foreach (var columna in porColumnaBase.Keys.Union(porColumnaSalida.Keys).OrderBy(c => c, StringComparer.Ordinal))
            {
                var esperado = porColumnaBase.TryGetValue(columna, out var enBase) ? enBase : 0;
                if (string.Equals(columna, "F", StringComparison.Ordinal))
                {
                    esperado -= 1;
                }

                var real = porColumnaSalida.TryGetValue(columna, out var enSalida) ? enSalida : 0;
                Assert.Equal(esperado, real);
            }
        }
        finally
        {
            BorrarDecimal(carpetaPeriodo);
            BorrarDecimal(salidaDir);
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

    /// <summary>
    /// Copia temporal del período 2026082: insumos REALES de <c>Docs/Prueba Agosto-2/Insumos</c>
    /// (5 ASE + R10) más las conciliaciones reales del canónico Q2 de julio (limitación declarada:
    /// la carpeta de agosto no trae <c>Conciliaciones/</c>).
    /// </summary>
    private static string PrepararPeriodoAgosto()
    {
        var destino = Path.Combine(Path.GetTempPath(), "remuneracion-2026082-insumos-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(destino);

        var origen = Insumos.CarpetaInsumosAgosto;
        foreach (var archivo in Directory.EnumerateFiles(origen, "*.xlsx", SearchOption.AllDirectories))
        {
            var relativo = Path.GetRelativePath(origen, archivo);
            var destinoArchivo = Path.Combine(destino, relativo);
            Directory.CreateDirectory(Path.GetDirectoryName(destinoArchivo)!);
            File.Copy(archivo, destinoArchivo);
        }

        var conciliacionesDestino = Path.Combine(destino, "Conciliaciones");
        Directory.CreateDirectory(conciliacionesDestino);
        foreach (var archivo in Directory.EnumerateFiles(Insumos.CarpetaConciliacionesQ2, "*.xlsx", SearchOption.TopDirectoryOnly))
        {
            File.Copy(archivo, Path.Combine(conciliacionesDestino, Path.GetFileName(archivo)));
        }

        return destino;
    }

    private static int ContarFilasBloqueEnSalida(string ruta, string nombreAse)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = Hoja(workbookPart, "Reporte Componentes R1");
        var filas = hoja.Descendants<Row>().OrderBy(r => r.RowIndex?.Value ?? 0).ToList();

        var nameRow = 0;
        var totalRow = 0;
        foreach (var fila in filas)
        {
            if (nameRow == 0 && Normalizar(Texto(workbookPart, fila, "B")) == Normalizar(nombreAse))
            {
                nameRow = (int)(fila.RowIndex?.Value ?? 0);
                continue;
            }

            if (nameRow > 0 && EsTotalFinal(workbookPart, fila))
            {
                totalRow = (int)(fila.RowIndex?.Value ?? 0);
                break;
            }
        }

        Assert.True(nameRow > 0 && totalRow > nameRow, $"ASE {nombreAse}: no se localizó el bloque en la salida.");

        var conteo = 0;
        foreach (var fila in filas)
        {
            var idx = (int)(fila.RowIndex?.Value ?? 0);
            if (idx <= nameRow || idx > totalRow)
            {
                continue;
            }

            if (new[] { "A", "B", "C", "D", "E" }.Any(c => !string.IsNullOrWhiteSpace(Texto(workbookPart, fila, c))))
            {
                conteo++;
            }
        }

        return conteo;
    }

    private static int ContarFormulas(string ruta, string hojaNombre)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        return Hoja(workbookPart, hojaNombre).Descendants<Cell>().Count(c => c.CellFormula is not null);
    }

    /// <summary>
    /// Plan 31 (T2, §2.3/S5): cuenta las fórmulas de una hoja agrupadas por columna. Es el
    /// invariante fuerte de "ninguna otra fórmula desapareció" tras el dimensionado del espejo:
    /// el espejo desplaza filas (no columnas), de modo que el conteo por columna es estable y el
    /// único delta legítimo es la columna F (F327 = literal 0 del EXTEMP ASE3-agosto).
    /// </summary>
    private static Dictionary<string, int> ContarFormulasPorColumna(string ruta, string hojaNombre)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var conteo = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var celda in Hoja(workbookPart, hojaNombre).Descendants<Cell>())
        {
            if (celda.CellFormula is null)
            {
                continue;
            }

            var columna = ColumnaDe(celda.CellReference?.Value);
            conteo[columna] = conteo.TryGetValue(columna, out var actual) ? actual + 1 : 1;
        }

        return conteo;
    }

    /// <summary>
    /// Plan 31 (T2, §2.3/S5): true si la celda indicada es un literal con el valor esperado y NO
    /// tiene <c>&lt;f&gt;</c> (lectura ZIP+XML vía OpenXML, sin Excel/COM). Es el aserto específico
    /// del único 0-Aplic de agosto (EXTEMP ASE3-agosto, <c>Reporte Componentes R1!F327</c>).
    /// </summary>
    private static bool EsLiteralSinFormula(string ruta, string hojaNombre, string referencia, string valorEsperado)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var celda = Hoja(workbookPart, hojaNombre).Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        return celda is not null
            && celda.CellFormula is null
            && string.Equals(celda.CellValue?.InnerText, valorEsperado, StringComparison.Ordinal);
    }

    private static string ColumnaDe(string? referencia)
    {
        var letras = new string((referencia ?? string.Empty).Where(char.IsLetter).ToArray());
        if (letras.Length == 0)
        {
            throw new InvalidOperationException($"Celda sin referencia de columna: '{referencia}'.");
        }

        return letras.ToUpperInvariant();
    }

    private static bool EsTotalFinal(WorkbookPart workbookPart, Row fila) =>
        string.Equals(Texto(workbookPart, fila, "A"), "Total", StringComparison.OrdinalIgnoreCase)
        && string.IsNullOrWhiteSpace(Texto(workbookPart, fila, "B"));

    private static string Texto(WorkbookPart workbookPart, Row fila, string columna)
    {
        var indice = ((uint?)fila.RowIndex?.Value ?? 0).ToString(CultureInfo.InvariantCulture);
        var celda = fila.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, columna + indice, StringComparison.OrdinalIgnoreCase));
        if (celda is null)
        {
            return string.Empty;
        }

        if (celda.InlineString is not null)
        {
            return celda.InlineString.InnerText ?? string.Empty;
        }

        if (celda.DataType?.Value == CellValues.SharedString && celda.CellValue is not null
            && int.TryParse(celda.CellValue.Text, out var i))
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable?.ElementAt(i).InnerText ?? string.Empty;
        }

        return celda.CellValue?.InnerText ?? string.Empty;
    }

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet!;
    }

    private static string Normalizar(string texto)
    {
        var d = texto.Normalize(System.Text.NormalizationForm.FormD);
        var sb = new System.Text.StringBuilder();
        foreach (var ch in d)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(ch) != System.Globalization.UnicodeCategory.NonSpacingMark)
            {
                sb.Append(char.ToUpperInvariant(ch));
            }
        }

        return sb.ToString();
    }

    private static void BorrarDecimal(string ruta)
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
