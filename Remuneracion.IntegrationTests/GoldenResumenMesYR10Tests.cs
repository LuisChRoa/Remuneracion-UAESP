using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 20 / T6 (goldens Capa A Q1+Q2 de G2/G3): cierra el golden que quedó pendiente y la
/// reorganización de insumos (T6 bloqueado y la plantilla control Q2 eliminada a propósito).
///
/// B1 (S4, G2): el RESUMEN MES de <c>{2026072}/Conciliaciones/</c> (VALOR 2°Q en F/G) es copia
/// 1:1 (dif 0) de las hojas <c>Recaudo *</c> del golden Q2 <c>Remuneracion 202607-2 Total.xlsx</c>.
/// B2 (S3, regresión Q1): idem con <c>{2026071}/Conciliaciones/</c> (VALOR 1°Q en D/E) contra el
/// golden Q1 <c>Remuneracion 202607-1 Total.xlsx</c>.
/// B3 (S6/R-G3-2): el reader R10 expone el DetRetri por ASE (D9:D13) + total (D14) de ambos
/// períodos (58.210.094.822 Q1 / 72.441.209.168 Q2; rangos 01-15/07 y 16-31/07).
/// B4 (S7): fail-fast sin R10 (nombra período + ruta) y divergencia &gt;±0.5 del DetRetri
/// calculado contra el R10 (nombra período + archivo + ambos valores), sobre fixtures temporales.
///
/// Todos los insumos referenciados EXISTEN en <c>Docs/Insumos</c> (nunca <c>Consolidado/</c> ni
/// plantillas eliminadas). Los casos negativos copian el período a temp: <c>Docs/Insumos</c> no se
/// toca.
/// </summary>
public sealed class GoldenResumenMesYR10Tests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    // ── B1 ────────────────────────────────────────────────────────────────────────────────
    [Fact]
    public void GoldenQ2_S4_ResumenMes_FG_EsCopia1a1DeLosRecaudos()
    {
        var lector = new ExcelDataReaderWorkbookLeafInputReader();
        var locator = new ArchivoFuenteLocator();

        var recaudos = lector.LeerRecaudosEmpresa(
            Insumos.PeriodoQ2(),
            empresa => locator.BuscarConciliacion(Insumos.CarpetaPeriodoQ2, empresa.PrefijoConciliacion));

        Assert.Equal(5, recaudos.Count);
        CompararBloquesContraGolden(recaudos, Insumos.GoldenQ2, letraValor: "F", letraRegistros: "G");
    }

    // ── B2 ────────────────────────────────────────────────────────────────────────────────
    [Fact]
    public void GoldenQ1_S3_ResumenMes_DE_EsCopia1a1DeLosRecaudos()
    {
        var lector = new ExcelDataReaderWorkbookLeafInputReader();
        var locator = new ArchivoFuenteLocator();

        var recaudos = lector.LeerRecaudosEmpresa(
            Insumos.Periodo(),
            empresa => locator.BuscarConciliacion(Insumos.CarpetaPeriodo, empresa.PrefijoConciliacion));

        Assert.Equal(5, recaudos.Count);
        CompararBloquesContraGolden(recaudos, Insumos.Plantilla, letraValor: "D", letraRegistros: "E");
    }

    /// <summary>
    /// Copia 1:1 (tolerancia 0) de cada celda mapeada del reader contra la hoja <c>Recaudo *</c>
    /// del golden. Ancla los totales OPORTUNO/EXTEMP/TOTAL por ASE y exige que la letra leída sea
    /// la de la quincena (F/G en Q2, D/E en Q1) — nunca la otra.
    /// </summary>
    private static void CompararBloquesContraGolden(
        IReadOnlyList<RecaudoEmpresaInputs> recaudos,
        string golden,
        string letraValor,
        string letraRegistros)
    {
        foreach (var recaudo in recaudos)
        {
            Assert.False(string.IsNullOrWhiteSpace(recaudo.HojaRecaudo));

            foreach (var (celda, valor) in recaudo.Celdas)
            {
                Assert.True(
                    celda.StartsWith(letraValor, StringComparison.OrdinalIgnoreCase)
                    || celda.StartsWith(letraRegistros, StringComparison.OrdinalIgnoreCase),
                    $"{recaudo.HojaRecaudo}!{celda}: la quincena leída debe ser {letraValor}/{letraRegistros}.");

                var esperado = TestHelpers.LeerCeldaNumerica(golden, recaudo.HojaRecaudo, celda);
                Assert.Equal(esperado, valor);
            }

            // Anclas de bloque: OPORTUNO (fila 9), EXTEMP (18) y TOTAL (27) contra el golden.
            Assert.Equal(TestHelpers.LeerCeldaNumerica(golden, recaudo.HojaRecaudo, $"{letraValor}9"), recaudo.TotalOportuno);
            Assert.Equal(TestHelpers.LeerCeldaNumerica(golden, recaudo.HojaRecaudo, $"{letraValor}18"), recaudo.TotalExtemporaneo);
            Assert.Equal(TestHelpers.LeerCeldaNumerica(golden, recaudo.HojaRecaudo, $"{letraValor}27"), recaudo.Total);
        }
    }

    // ── B3 ────────────────────────────────────────────────────────────────────────────────
    [Fact]
    public void GoldenR10_S6_DetRetriPorAseYTotal_Q1YQ2()
    {
        var reader = new ExcelDataReaderDetRetriR10Reader();

        var q1 = reader.LeerDetRetri(Insumos.Periodo(), Insumos.R10);
        Assert.Equal(5, q1.DetRetriPorAse.Count);
        Assert.InRange(q1.Total - 58210094822m, -Tolerancia, Tolerancia);
        foreach (var (aseId, esperado) in new Dictionary<int, decimal>
                 {
                     [1] = 16758167585m,
                     [2] = 12227292053m,
                     [3] = 10127195599m,
                     [4] = 10553776376m,
                     [5] = 8543663209m
                 })
        {
            Assert.InRange(q1.DetRetriPorAse[aseId] - esperado, -Tolerancia, Tolerancia);
        }

        Assert.Equal(new DateTime(2026, 7, 1), q1.FechaDesde);
        Assert.Equal(new DateTime(2026, 7, 15), q1.FechaHasta);

        var q2 = reader.LeerDetRetri(Insumos.PeriodoQ2(), Insumos.R10Q2);
        Assert.Equal(5, q2.DetRetriPorAse.Count);
        Assert.InRange(q2.Total - 72441209168m, -Tolerancia, Tolerancia);
        foreach (var (aseId, esperado) in new Dictionary<int, decimal>
                 {
                     [1] = 17450228673m,
                     [2] = 20516143970m,
                     [3] = 15221896467m,
                     [4] = 7179595396m,
                     [5] = 12073344662m
                 })
        {
            Assert.InRange(q2.DetRetriPorAse[aseId] - esperado, -Tolerancia, Tolerancia);
        }

        Assert.Equal(new DateTime(2026, 7, 16), q2.FechaDesde);
        Assert.Equal(new DateTime(2026, 7, 31), q2.FechaHasta);
    }

    // ── B4a ───────────────────────────────────────────────────────────────────────────────
    [Fact]
    public void GoldenR10_S7_FaltaR10EnElPeriodo_FallaNombrandoPeriodoYRuta()
    {
        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-r10-ausente-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodo, carpetaPeriodo);
        var rutaR10 = Path.Combine(carpetaPeriodo, "R10_Remuneracion_2026071.xlsx");
        Assert.True(File.Exists(rutaR10), $"El fixture debía copiar el R10: {rutaR10}");
        File.Delete(rutaR10);

        var salida = Path.Combine(Path.GetTempPath(), "remuneracion-r10-ausente-salida-" + Guid.NewGuid().ToString("N"), Insumos.Periodo().NombreArchivo);
        var procesador = CrearProcesadorConR10();

        var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.Periodo(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.Plantilla,
                RutaSalida = salida
            }));

        Assert.Contains("R10_Remuneracion_2026071", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2026071", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(carpetaPeriodo, ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(salida), "No debe existir salida certificada si falta el R10.");
    }

    // ── B4b ───────────────────────────────────────────────────────────────────────────────
    [Fact]
    public void GoldenR10_S7_DivergenciaMayorAMediaUnidad_FallaConPeriodoArchivoYAmbosValores()
    {
        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-r10-divergente-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodoQ2, carpetaPeriodo);
        var rutaR10 = Path.Combine(carpetaPeriodo, "R10_Remuneracion_2026072.xlsx");
        Assert.True(File.Exists(rutaR10), $"El fixture debía copiar el R10 Q2: {rutaR10}");

        // Calculado (bottom-up) ASE1 = ROUND(D104) = 17.450.228.673; el R10 se altera a ...673+1.000.
        const decimal calculadoAse1 = 17450228673m;
        const decimal r10Alterado = 17450229673m;
        ModificarCeldaNumerica(rutaR10, WorkbookLeafCellMapQ2.HojaDetRetri(Insumos.PeriodoQ2()), "D9", r10Alterado);

        var salida = Path.Combine(Path.GetTempPath(), "remuneracion-r10-divergente-salida-" + Guid.NewGuid().ToString("N"), Insumos.PeriodoQ2().NombreArchivo);
        var procesador = CrearProcesadorConR10();

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.PlantillaQ2,
                RutaSalida = salida
            }));

        Assert.Contains("2026072", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("R10_Remuneracion_2026072.xlsx", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(calculadoAse1.ToString(CultureInfo.InvariantCulture), ex.Message, StringComparison.Ordinal);
        Assert.Contains(r10Alterado.ToString(CultureInfo.InvariantCulture), ex.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(salida), "No debe existir salida certificada ante divergencia DetRetri-vs-R10.");
    }

    // ── B4c (Q1) ──────────────────────────────────────────────────────────────────────────
    [Fact]
    public void GoldenR10_S7_Q1_Divergencia_FallaConPeriodoArchivoYAmbosValores()
    {
        // HU-20/G3: la validación DetRetri calculado-vs-R10 corre en AMBAS quincenas. En Q1 el
        // DetRetri calculado es ROUND(D104:D108,0) (= 16.758.167.585 para ASE1) y también se
        // contrasta contra el R10 (hoja DetRetri2026071). Este test PRUEBA que Q1 no es un no-op.
        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-r10-divergente-q1-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodo, carpetaPeriodo);
        var rutaR10 = Path.Combine(carpetaPeriodo, "R10_Remuneracion_2026071.xlsx");
        Assert.True(File.Exists(rutaR10), $"El fixture debía copiar el R10 Q1: {rutaR10}");

        const decimal calculadoAse1 = 16758167585m;
        const decimal r10Alterado = 16758168585m;
        ModificarCeldaNumerica(rutaR10, WorkbookLeafCellMapQ2.HojaDetRetri(Insumos.Periodo()), "D9", r10Alterado);

        var salida = Path.Combine(Path.GetTempPath(), "remuneracion-r10-divergente-q1-salida-" + Guid.NewGuid().ToString("N"), Insumos.Periodo().NombreArchivo);
        var procesador = CrearProcesadorConR10();

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.Periodo(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.Plantilla,
                RutaSalida = salida
            }));

        Assert.Contains("2026071", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("R10_Remuneracion_2026071.xlsx", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(calculadoAse1.ToString(CultureInfo.InvariantCulture), ex.Message, StringComparison.Ordinal);
        Assert.Contains(r10Alterado.ToString(CultureInfo.InvariantCulture), ex.Message, StringComparison.Ordinal);
        Assert.False(File.Exists(salida), "No debe existir salida certificada ante divergencia DetRetri-vs-R10 en Q1.");
    }

    /// <summary>
    /// Procesador de período con el reader R10 inyectado (7.º parámetro, OBLIGATORIO desde
    /// HU-20/G3): los casos negativos necesitan que el oráculo de período se consulte
    /// (G3-D2/R-G3-3) en AMBAS quincenas.
    /// </summary>
    private static ProcesadorPeriodo CrearProcesadorConR10() =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            detRetriR10Reader: new ExcelDataReaderDetRetriR10Reader());

    private static void CopiarArbol(string origen, string destino)
    {
        Directory.CreateDirectory(destino);
        foreach (var archivo in Directory.EnumerateFiles(origen, "*.xlsx", SearchOption.AllDirectories))
        {
            var relativo = Path.GetRelativePath(origen, archivo);
            var destinoArchivo = Path.Combine(destino, relativo);
            Directory.CreateDirectory(Path.GetDirectoryName(destinoArchivo)!);
            File.Copy(archivo, destinoArchivo);
        }
    }

    private static void ModificarCeldaNumerica(string ruta, string hoja, string celda, decimal valor)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        objCelda.CellFormula?.Remove();
        objCelda.DataType = null;
        objCelda.CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture));
        worksheet.Save();
    }
}
