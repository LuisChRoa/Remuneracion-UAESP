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
/// Plan 25 (WU-1 = T1+T2): roles R1-Q2 resueltos por FIRMA de etiquetas sobre la secuencia del
/// período (fin del mapa congelado por conteo), corrección de rol D-A (<c>F37/F270 → Aplic1</c>),
/// opcionalidad de <c>Aplic</c> (ausente = 0) y agregado D-C (<c>EXTEMP = Σ Aplicacion</c>).
///
/// Evidencia: <c>plans/25 - roles-r1-q2-por-firma-sobre-secuencia-espejo.md</c> y
/// <c>plans/24 - T0 Evidencia.md</c>. Verificación SOLO con insumos reales de
/// <c>Docs/Insumos</c> (julio) y <c>Docs/Prueba2</c> (agosto); sin fixtures de valores.
/// </summary>
public sealed class R1Q2ResolucionPorFirmaTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    // S1 (Plan 25 §3.2): agosto-ASE3 sin filas `Aplicacion` → no lanza; EXTEMP = 0 explícito.
    [Fact]
    public void S1_AgostoAse3_SinAplicacion_ExtempCeroYSlotsEnCero()
    {
        var leaf = new ExcelDataReaderWorkbookLeafInputReader().LeerLeafInputs(
            Insumos.Ase(3), Insumos.PeriodoQ2(),
            Insumos.R1Agosto(3), Insumos.R2Agosto(3), Insumos.R4Agosto(3));

        Assert.Equal(0m, leaf.R1.ExtemporaneoEsperadoPorAse);
        Assert.Equal(0m, leaf.R1.CeldasPorAse["F250"]); // Aplic0
        Assert.Equal(0m, leaf.R1.CeldasPorAse["F270"]); // Aplic1 (D-A)
        Assert.Equal(0m, leaf.R1.CeldasPorAse["L250"]); // LAplic0

        // TOT_OPT del espejo T0c §4.3 (3 filas Mes) — la ausencia de Aplic no lo altera.
        Assert.InRange(leaf.R1.TotalOportunoEsperadoPorAse - 16345751509.86m, -Tolerancia, Tolerancia);
    }

    // S2 (Plan 25 §3.2, R-F-1): julio-ASE1/ASE3 con `Aplicacion` presente → valores reales
    // (F37 = Aplic1, F270 = Aplic1), NO `Subs0`. Aplic0 de ASE1 = 9.215,69 (R16 de la fuente).
    [Fact]
    public void S2_Q2Julio_AplicacionPresente_ValoresRealesIntactos()
    {
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        var ase1 = reader.LeerLeafInputs(Insumos.Ase(1), Insumos.PeriodoQ2(), Insumos.R1Q2(1), Insumos.R2Q2(1), Insumos.R4Q2(1));
        Assert.InRange(ase1.R1.CeldasPorAse["F37"] - 40364.31m, -Tolerancia, Tolerancia); // Aplic1 (D-A)
        Assert.InRange(ase1.R1.CeldasPorAse["F17"] - 9215.69m, -Tolerancia, Tolerancia);  // Aplic0
        Assert.Equal(0m, ase1.R1.CeldasPorAse["L17"]);

        var ase3 = reader.LeerLeafInputs(Insumos.Ase(3), Insumos.PeriodoQ2(), Insumos.R1Q2(3), Insumos.R2Q2(3), Insumos.R4Q2(3));
        Assert.InRange(ase3.R1.CeldasPorAse["F270"] - 627141.97m, -Tolerancia, Tolerancia); // Aplic1 (D-A)
    }

    // S3 (Plan 25 §3.2, R-F-3): rol core `Mes` ausente → fail-fast que nombra ASE + reporte + celda.
    // Estructural (no hay fuente real con menos filas Mes): parte de una plantilla mínima en memoria
    // derivada de la firma real; sin valores de negocio inventados.
    [Fact]
    public void S3_CoreMes_Ausente_FailFastNombraAseReporteCelda()
    {
        var ruta = CrearFuenteR1SinMes2();
        var reader = new ExcelDataReaderWorkbookLeafInputReader();

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            reader.LeerLeafInputs(Insumos.Ase(1), Insumos.PeriodoQ2(), ruta, Insumos.R2Q2(1), Insumos.R4Q2(1)));

        Assert.Contains("ASE 1", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("R1-Q2", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Mes2", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("F48", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // S5 (Plan 25 §3.2, D-D): el reader es compartido; la resolución por firma no depende del
    // espejo ni del orquestador. Lectura directa del reader (condición single-ASE, sin EspejoR1)
    // == leaf del ASE3 en el flujo 5-ASE (ProcesadorPeriodo). El path single-ASE no puede correr Q2
    // end-to-end (guard pre-existente de CalculoRemuneracion), por eso la paridad se asienta en el
    // reader compartido, que es donde vive el fix (D-D).
    [Fact]
    public void S5_ParidadReaderCompartido_DirectoVsFlujo5Ase_AgostoAse3()
    {
        var leafDirecto = new ExcelDataReaderWorkbookLeafInputReader().LeerLeafInputs(
            Insumos.Ase(3), Insumos.PeriodoQ2(),
            Insumos.R1Agosto(3), Insumos.R2Agosto(3), Insumos.R4Agosto(3));
        Assert.Equal(0m, leafDirecto.R1.ExtemporaneoEsperadoPorAse);

        var carpetaPeriodo = PrepararPeriodoAgosto();
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-wu1-paridad-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, "Paridad2026082.xlsx");
        try
        {
            var procesador = new ProcesadorPeriodo(
                new ExcelDataReaderRecaudoReader(),
                new ExcelDataReaderWorkbookLeafInputReader(),
                new CalculoRemuneracion(),
                new ValidadorBasico(),
                new OpenXmlPlantillaWriter(),
                new ArchivoFuenteLocator(),
                new ExcelDataReaderDetRetriR10Reader());

            var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 },
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.PlantillaAgosto2026082,
                RutaSalida = salida
            });

            var leaf5Ase = resultado.Leafs.Single(l => l.Ase.Id == 3);
            Assert.Equal(leafDirecto.R1.ExtemporaneoEsperadoPorAse, leaf5Ase.R1.ExtemporaneoEsperadoPorAse);
            Assert.Equal(0m, leaf5Ase.R1.ExtemporaneoEsperadoPorAse);
            Assert.InRange(leaf5Ase.R1.TotalOportunoEsperadoPorAse - 16345751509.86m, -Tolerancia, Tolerancia);
            Assert.InRange(leaf5Ase.DetRetriQ2!.Detalle - 16369059896m, -Tolerancia, Tolerancia);
        }
        finally
        {
            Borrar(carpetaPeriodo);
            Borrar(salidaDir);
        }
    }

    /// <summary>
    /// Fuente mínima R1 con SOLO 2 filas <c>Mes/Total</c> (ASE1-Q2 exige 3 → rol Mes2 ausente).
    /// Estructural: no aporta valores de negocio; solo ejercita la rama del fail-fast de rol core.
    /// </summary>
    private static string CrearFuenteR1SinMes2()
    {
        var ruta = Path.Combine(Path.GetTempPath(), "r1-sin-mes2-" + Guid.NewGuid().ToString("N") + ".xlsx");
        using (var doc = SpreadsheetDocument.Create(ruta, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var sheetPart = workbookPart.AddNewPart<WorksheetPart>();

            Row FilaMes(int fila) => new(
                new Cell { CellReference = $"B{fila}", DataType = CellValues.String, CellValue = new CellValue("Mes") },
                new Cell { CellReference = $"C{fila}", DataType = CellValues.String, CellValue = new CellValue("Total") },
                new Cell { CellReference = $"F{fila}", CellValue = new CellValue("100") });

            sheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(new Cell
                {
                    CellReference = "A1",
                    DataType = CellValues.String,
                    CellValue = new CellValue("Recaudo Desde: 16/07/2026 Hasta: 31/07/2026")
                }),
                new Row(
                    new Cell { CellReference = "F3", DataType = CellValues.String, CellValue = new CellValue("Total") },
                    new Cell { CellReference = "G3", DataType = CellValues.String, CellValue = new CellValue("Componente TDF") }),
                new Row(
                    new Cell { CellReference = "A4", DataType = CellValues.String, CellValue = new CellValue("Componente") },
                    new Cell { CellReference = "B4", DataType = CellValues.String, CellValue = new CellValue("Total") },
                    new Cell { CellReference = "F4", CellValue = new CellValue("100") }),
                FilaMes(11),
                FilaMes(31)));

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(sheetPart), SheetId = 1, Name = "Sheet1" });
            workbookPart.Workbook.Save();
        }

        return ruta;
    }

    /// <summary>
    /// Copia temporal del período 2026082: insumos REALES de <c>Docs/Prueba2/Insumos</c> (5 ASE +
    /// R10) más las conciliaciones reales del canónico Q2 de julio (limitación declarada: Prueba2
    /// no trae <c>Conciliaciones/</c>).
    /// </summary>
    private static string PrepararPeriodoAgosto()
    {
        var destino = Path.Combine(Path.GetTempPath(), "remuneracion-wu1-insumos-" + Guid.NewGuid().ToString("N"));
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
            // best-effort
        }
    }
}
