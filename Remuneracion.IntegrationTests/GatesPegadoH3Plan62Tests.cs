using System.Globalization;
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
/// Plan 62 (H3 — T3.1/T3.2/T3.3): reformulación de gates tras el pegado del bloque R10
/// (<c>DetRetri{AAAAMMQ}</c> / <c>DetValiRetri{AAAAMMQ}</c>).
///
/// <para><b>Doctrina D-61R10:</b> el gate TOTAL calculado-vs-R10 se SUSTITUYE por la verificación de
/// integridad del pegado (<c>OpenXmlPlantillaWriter.VerificarIntegridadPegadoR10</c>); el gate
/// PER-ASE se conserva porque compara el cálculo del DOMINIO contra el insumo R10, no lo pegado
/// contra sí mismo. El <c>-1</c> residual vive en <c>DetValiRetri!D21</c> (divergencia documentada
/// D6, excluida del gate de validaciones cruzadas) y NO se abre categoría nueva.</para>
///
/// <para><b>Fuentes:</b> <c>Docs/Prueba Septiembre 1/Insumos/R10_Remuneracion_2026091.xlsx</c> por
/// ruta explícita; la corrida real de septiembre-1 se comparte con
/// <see cref="EspecialesR4H2Plan61Tests.Generacion2026091"/> (su Lazy es thread-safe). El caso del
/// read-back trabaja sobre una COPIA temporal: los insumos y la salida compartida no se mutan.</para>
/// </summary>
public sealed class GatesPegadoH3Plan62Tests
{
    private const decimal Tolerancia = 0.5m;

    private const string CodigoPeriodo = "2026091";

    private const string DuenoAdministrativo = "Área administrativa (elabora el reporte oficial)";

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T3.2 — El gate de TOTAL calculado-vs-R10 ya no existe para la hoja (tautología imposible)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T3.2 (R-TAUTOLOGIA / D-GATE-TOTAL): con los cinco ASE cuadrando per-ASE y un
    /// <c>r10.Total</c> absurdo, <see cref="ValidadorBasico.ValidarDetRetriContraR10"/> ya NO emite
    /// el error de total. Antes de H3 esa capa reportaba el <c>-1</c> de composición (Σ del dominio
    /// vs total declarado) aunque cada ASE cuadrara; ahora ese verde/rojo no prueba nada porque la
    /// hoja es el R10 por construcción.
    /// </summary>
    [Fact]
    public void T3_2_ElGateDeTotalCalculadoVsR10_YaNoExiste()
    {
        var leafs = CrearLeafsSinteticos(1_000m, 2_000m, 3_000m, 4_000m, 5_000m);
        var r10 = R10Sintetico(leafs, total: 999_999_999m);

        var errores = new ValidadorBasico().ValidarDetRetriContraR10(leafs, r10);

        Assert.Empty(errores);
        Assert.DoesNotContain(errores, e => e.Contains("total calculado", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(errores, e => e.Contains("R10 total", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(errores, e => e.Contains('Σ'));
    }

    /// <summary>
    /// T3.2 (D-61R10): <c>ClasificadorNovedadesEntrega.DesdeR10</c> deja de emitir la ficha de
    /// TOTAL ("Total del período", <c>AseId = 0</c>) aunque el total declarado del R10 no cuadre:
    /// esa ficha era la que producía la novedad <c>-1</c> de septiembre-1.
    /// </summary>
    [Fact]
    public void T3_2_DesdeR10_NoEmiteNovedadDeTotal()
    {
        var leafs = CrearLeafsSinteticos(1_000m, 2_000m, 3_000m, 4_000m, 5_000m);
        var r10 = R10Sintetico(leafs, total: 999_999_999m);

        var novedades = ClasificadorNovedadesEntrega.DesdeR10(
            leafs, r10, Periodo.Parse(CodigoPeriodo), "R10_Remuneracion_2026091.xlsx");

        Assert.Empty(novedades);
        Assert.DoesNotContain(novedades, n => n.AseId == 0);
        Assert.DoesNotContain(novedades, n => n.ConceptoOCelda == "Total del período");
    }

    /// <summary>
    /// T3.2: el retiro del TOTAL no apaga el gate per-ASE (D-PERASE-CONSERVA). Con un ASE desviado
    /// en 1 peso, el validador sigue reportando UNA ficha que nombra el ASE, y
    /// <c>DesdeR10</c> la entrega con el dueño administrativo: el cálculo de la plantilla se
    /// contrasta contra el reporte oficial del operador, no contra la copia pegada.
    /// </summary>
    [Fact]
    public void T3_1_LaNovedadPerAseConservaElDuenoAdministrativo()
    {
        var correctos = CrearLeafsSinteticos(1_000m, 2_000m, 3_000m, 4_000m, 5_000m);
        var r10 = R10Sintetico(correctos, total: 15_000m);
        var conBrecha = correctos
            .Select(l => l.Ase.Id == 3 ? CrearLeaf(3, 3_001m) : l)
            .ToList();

        var errores = new ValidadorBasico().ValidarDetRetriContraR10(conBrecha, r10);
        var error = Assert.Single(errores);
        Assert.Contains("ASE 3", error, StringComparison.Ordinal);
        Assert.Contains("3001", error, StringComparison.Ordinal);
        Assert.Contains("3000", error, StringComparison.Ordinal);

        var novedad = Assert.Single(ClasificadorNovedadesEntrega.DesdeR10(
            conBrecha, r10, Periodo.Parse(CodigoPeriodo), "R10_Remuneracion_2026091.xlsx"));
        Assert.Equal(3, novedad.AseId);
        Assert.Equal(DuenoAdministrativo, novedad.Dueno);
        Assert.Contains("etiqueta del ASE", novedad.QueRevisar, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("DetRetri2026091", novedad.Hoja, StringComparison.Ordinal);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T3.1 — Per-ASE conservado contra el R10 real de septiembre-1
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T3.1 (D-PERASE-CONSERVA): con el R10 del período y los datos de septiembre-1, el gate
    /// per-ASE queda VERDE por los cinco operadores. El operando izquierdo es el cálculo del
    /// DOMINIO (<c>leaf.DetRetriQ2.Detalle</c>, compuesto en <c>ProcesadorPeriodo</c> desde
    /// R1/R2/R4 + AJUSTES-SF-T), no la hoja pegada: no hay tautología.
    /// </summary>
    [Fact]
    public void T3_1_Septiembre1_PerAseVerdeContraElR10()
    {
        var (resultado, leafs) = EspecialesR4H2Plan61Tests.Generacion2026091.Value;
        var periodo = resultado.Resultado.Periodo;
        var r10 = new ExcelDataReaderDetRetriR10Reader().LeerDetRetri(periodo, RutaR10());

        var errores = new ValidadorBasico().ValidarDetRetriContraR10(leafs, r10);

        Assert.Empty(errores);
        foreach (var leaf in leafs.OrderBy(l => l.Ase.Id))
        {
            var calculado = leaf.DetRetriQ2?.Detalle
                ?? throw new InvalidOperationException($"ASE {leaf.Ase.Id}: falta el DetRetri del dominio.");
            Assert.True(r10.DetRetriPorAse.ContainsKey(leaf.Ase.Id), $"El R10 no trae el ASE {leaf.Ase.Id}.");
            Assert.InRange(calculado - r10.DetRetriPorAse[leaf.Ase.Id], -Tolerancia, Tolerancia);
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T3.3 — Integridad del pegado: pegado corrupto ⇒ rojo con hoja + celda + esperado + escrito
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T3.3 (D-H3.2 / D61R10): el sustituto del gate de total es el read-back de integridad. Si el
    /// bloque del workbook quedó con un monto distinto al del R10 (pegado corrupto), la verificación
    /// falla con <c>ERR-ESCRITURA</c> nombrando HOJA + CELDA + esperado + escrito. Se perturba una
    /// COPIA temporal de la salida real para no tocar el fixture compartido.
    /// </summary>
    [Fact]
    public void T3_3_ReadBack_EnrojeceAntePegadoCorrupto()
    {
        var periodo = Periodo.Parse(CodigoPeriodo);
        var lector = new ExcelDataReaderDetRetriR10Reader();
        var bloqueDet = lector.LeerBloqueDetRetri(periodo, RutaR10());
        var bloqueVali = lector.LeerBloqueDetValiRetri(periodo, RutaR10());

        using var copia = CopiaTemporalDeLaSalida.Nueva();
        var hoja = bloqueDet.Hoja;
        var esperado = bloqueDet.Celda(9, "D")?.Numerico
            ?? throw new InvalidOperationException("El bloque del R10 no trae D9.");
        var escrito = esperado + 1m;
        EscribirNumero(copia.Ruta, hoja, "D9", escrito);

        using var documento = SpreadsheetDocument.Open(copia.Ruta, false);
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            OpenXmlPlantillaWriter.VerificarIntegridadPegadoR10(
                documento.WorkbookPart!, periodo, bloqueDet, bloqueVali));

        Assert.Equal(CodigoError.Escritura, ex.Codigo);
        Assert.Contains(hoja, ex.Message, StringComparison.Ordinal);
        Assert.Contains("D9", ex.Message, StringComparison.Ordinal);
        Assert.Contains(esperado.ToString(CultureInfo.InvariantCulture), ex.Message, StringComparison.Ordinal);
        Assert.Contains(escrito.ToString(CultureInfo.InvariantCulture), ex.Message, StringComparison.Ordinal);
        Assert.Contains("se esperaba", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T3.3 — E2E combinado: septiembre-1 verde, CERTIFICADA, sin novedad R10 "-1"
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// T3.3 (CA-3): la corrida completa de <c>ProcesadorPeriodo.Ejecutar</c> en septiembre-1 cierra
    /// CERTIFICADA, sin novedades y sin la ficha R10 <c>-1</c> (ni per-ASE ni de total). La hoja
    /// entregada trae el bloque del R10 por construcción (N1); la integridad del pegado es lo que
    /// se prueba, no una suma del mismo dato contra sí mismo.
    /// </summary>
    [Fact]
    public void T3_3_Septiembre1_E2E_CierraCertificadaSinNovedadR10()
    {
        var (resultado, _) = EspecialesR4H2Plan61Tests.Generacion2026091.Value;

        Assert.Equal(EstadoEntrega.Certificada, resultado.EstadoEntrega);
        Assert.Empty(resultado.Novedades);
        Assert.Equal(string.Empty, resultado.RutaReporteNovedades);
        Assert.True(File.Exists(resultado.RutaSalida), $"La corrida certificada debía dejar el Excel limpio: {resultado.RutaSalida}");
        Assert.DoesNotContain("(sin validar)", Path.GetFileName(resultado.RutaSalida), StringComparison.OrdinalIgnoreCase);

        Assert.DoesNotContain(resultado.Novedades, n => n.Delta == -1m);
        Assert.DoesNotContain(resultado.Novedades, n => n.ConceptoOCelda == "Total del período");

        // Discriminador del pegado: el asiento del ASE 3 de la hoja es el del R10 del período.
        var periodo = resultado.Resultado.Periodo;
        var bloqueDet = new ExcelDataReaderDetRetriR10Reader().LeerBloqueDetRetri(periodo, RutaR10());
        using var documento = SpreadsheetDocument.Open(resultado.RutaSalida, false);
        var celda = Celda(Hoja(documento.WorkbookPart!, bloqueDet.Hoja), "D11")
            ?? throw new InvalidOperationException("Falta DetRetri D11 en la salida.");
        Assert.InRange(
            NumeroDe(celda),
            (bloqueDet.Celda(11, "D")!.Numerico ?? 0m) - Tolerancia,
            (bloqueDet.Celda(11, "D")!.Numerico ?? 0m) + Tolerancia);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────────────────

    private static WorkbookLeafInputs CrearLeaf(int aseId, decimal detalle)
    {
        var ase = AseFactory.DesdeId(aseId);
        return new WorkbookLeafInputs
        {
            Ase = ase,
            Periodo = Periodo.Parse(CodigoPeriodo),
            DetRetriQ2 = new DetRetriQ2Inputs { Ase = ase, TotalD104 = detalle }
        };
    }

    private static List<WorkbookLeafInputs> CrearLeafsSinteticos(params decimal[] detalles) =>
        detalles.Select((detalle, indice) => CrearLeaf(indice + 1, detalle)).ToList();

    private static DetRetriInputs R10Sintetico(IReadOnlyList<WorkbookLeafInputs> leafs, decimal total) =>
        new()
        {
            CodigoRemuneracion = CodigoPeriodo,
            DetRetriPorAse = leafs.ToDictionary(l => l.Ase.Id, l => l.DetRetriQ2!.Detalle),
            Total = total
        };

    private static string RutaR10() => Path.Combine(
        Raiz(), "Docs", "Prueba Septiembre 1", "Insumos", "R10_Remuneracion_2026091.xlsx");

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }

    private sealed class CopiaTemporalDeLaSalida : IDisposable
    {
        private readonly string _carpeta;

        private CopiaTemporalDeLaSalida(string carpeta, string ruta)
        {
            _carpeta = carpeta;
            Ruta = ruta;
        }

        public string Ruta { get; }

        public static CopiaTemporalDeLaSalida Nueva()
        {
            var origen = EspecialesR4H2Plan61Tests.Generacion2026091.Value.Resultado.RutaSalida;
            if (!File.Exists(origen))
            {
                throw new InvalidOperationException($"La corrida de septiembre-1 no dejó el Excel: {origen}");
            }

            var carpeta = Path.Combine(Path.GetTempPath(), "remuneracion-p62-h3-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(carpeta);
            var ruta = Path.Combine(carpeta, Path.GetFileName(origen));
            File.Copy(origen, ruta, overwrite: false);
            return new CopiaTemporalDeLaSalida(carpeta, ruta);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_carpeta))
                {
                    Directory.Delete(_carpeta, recursive: true);
                }
            }
            catch (IOException)
            {
                // Limpieza best-effort (temp).
            }
        }
    }

    private static void EscribirNumero(string ruta, string hoja, string referencia, decimal valor)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var destino = Celda(Hoja(documento.WorkbookPart!, hoja), referencia)
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no trae '{referencia}'.");
        destino.CellFormula = null;
        destino.InlineString = null;
        destino.DataType = CellValues.Number;
        destino.CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture));
        documento.WorkbookPart!.Workbook!.Save();
    }

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"El workbook no trae la hoja '{nombre}'.");
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet.");
    }

    private static Cell? Celda(Worksheet worksheet, string referencia) =>
        worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));

    private static decimal NumeroDe(Cell celda) =>
        decimal.Parse(celda.CellValue?.InnerText ?? "0", NumberStyles.Any, CultureInfo.InvariantCulture);
}
