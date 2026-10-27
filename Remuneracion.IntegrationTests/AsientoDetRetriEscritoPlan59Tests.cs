using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (H4): el asiento <c>DetRetri{AAAAMMQ}</c> (tabla por ASE · columna de valor con encabezado
/// <c>'Recaudo Total'</c>) se escribe en AMBAS quincenas y lo ESCRITO se contrasta contra el esperado
/// del dominio.
///
/// Defecto que cubre (medido, corrida real 2026101 — quincena 1): el asiento se escribía solo en Q2,
/// así que en Q1 el entregable quedaba con el literal 0 de la plantilla y la corrida certificaba
/// <c>RESULTADO OK</c>: el gate de forma leía la fila de diferencias de <c>DetValiRetri</c> desde el
/// caché de la salida y «0 − 0» cerraba.
///
/// Este fixture prueba las dos mitades del arreglo:
/// <list type="number">
/// <item>Q1: el asiento escrito lleva el valor del dominio (no el 0 de la plantilla).</item>
/// <item>Inyección: si el asiento queda en cero, el gate de valores ENROJECE y el workbook no se
/// entrega (el falso verde ya no es posible).</item>
/// </list>
/// Los destinos se resuelven por ETIQUETA igual que el camino productivo; los números del oráculo
/// (2026101) se citan solo como evidencia en el nombre del test, no como coordenada de escritura.
/// </summary>
public sealed class AsientoDetRetriEscritoPlan59Tests
{
    [Fact]
    public void Q1_AsientoDetRetri_SeEscribeConElValorDelDominio_EnAmbasQuincenas()
    {
        var (resultado, leafs) = ConstruirPeriodoQ1ConDetRetri();
        var salidaDir = CrearDirectorioTemporal();
        try
        {
            var salida = Path.Combine(salidaDir, "salida.xlsx");
            new OpenXmlPlantillaWriter().GenerarWorkbook(Insumos.Plantilla, salida, resultado, leafs);

            // El asiento existe, no está en cero y coincide con el dominio (ROUND(D104:D108) por ASE).
            foreach (var leaf in leafs)
            {
                var escrito = LeerAsientoPorEtiqueta(salida, leaf);
                Assert.Equal(leaf.DetRetriQ2!.Detalle, escrito);
            }

            // Cierre: Σ de las filas por ASE del asiento (lo que el R10 oficial declara como total) y
            // cuadre consigo mismo.
            var cierre = LeerCierrePorEtiqueta(salida, leafs[0].Periodo);
            Assert.Equal(DetRetriRounder.Round(leafs.Sum(l => l.DetRetriQ2!.Detalle)), cierre);
            Assert.Equal(leafs.Sum(l => l.DetRetriQ2!.Detalle), cierre);

            // Evidencia del defecto cerrado: en Q1 el asiento ya no queda en el 0 de la plantilla.
            Assert.NotEqual(0m, LeerAsientoPorEtiqueta(salida, leafs[0]));

            // Y reproduce, fila por fila, el DetRetri del R10 oficial del período (oráculo G3).
            var r10 = new ExcelDataReaderDetRetriR10Reader().LeerDetRetri(leafs[0].Periodo, Insumos.R10);
            foreach (var leaf in leafs)
            {
                Assert.Equal(r10.DetRetriPorAse[leaf.Ase.Id], LeerAsientoPorEtiqueta(salida, leaf));
            }

            Assert.Equal(r10.Total, cierre);
        }
        finally
        {
            BorrarDirectorio(salidaDir);
        }
    }

    [Fact]
    public void Inyeccion_AsientoDetRetriForzadoACero_EnrojeceYNoCertifica()
    {
        var (resultado, leafs) = ConstruirPeriodoQ1ConDetRetri();
        var salidaDir = CrearDirectorioTemporal();
        try
        {
            var salida = Path.Combine(salidaDir, "salida.xlsx");
            new OpenXmlPlantillaWriter().GenerarWorkbook(Insumos.Plantilla, salida, resultado, leafs);

            // Control: tal como lo dejó el writer, el gate de valores pasa.
            EjecutarGateDeValores(salida, leafs);

            // Inyección: se fuerza el escenario histórico (asiento en 0, como el literal de la
            // plantilla cuando no se escribía en Q1) manteniendo intacto el esperado del dominio.
            ForzarAsientoACero(salida, leafs);

            var ex = Assert.Throws<CalculoInvalidoException>(() => EjecutarGateDeValores(salida, leafs));
            Assert.Equal(Remuneracion.Core.Errors.CodigoError.Escritura, ex.Codigo);
            Assert.Contains("DetRetri", ex.Message, StringComparison.Ordinal);
            Assert.Contains("ASE 1", ex.Message, StringComparison.Ordinal);
            Assert.Contains("Recaudo Total", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(leafs[0].DetRetriQ2!.Detalle.ToString(CultureInfo.InvariantCulture), ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            BorrarDirectorio(salidaDir);
        }
    }

    // ── Armado del período Q1 (Julio-1) con su DetRetri calculado ───────────────────────────────

    /// <summary>
    /// Réplica de la composición que puebla el proceso (<c>ProcesadorPeriodo</c>): DetRetriQ2 se
    /// calcula en AMBAS quincenas, así que el writer recibe el asiento también en Q1.
    /// </summary>
    private static (ResultadoRemuneracion Resultado, List<WorkbookLeafInputs> Leafs) ConstruirPeriodoQ1ConDetRetri()
    {
        var lector = new ExcelDataReaderRecaudoReader();
        var leafReader = new ExcelDataReaderWorkbookLeafInputReader();
        var periodo = Insumos.Periodo();

        var datos = Enumerable.Range(1, 5).Select(i =>
        {
            var r1 = lector.LeerR1(Insumos.R1(i));
            var r2 = lector.LeerR2(Insumos.R2(i));
            var r4 = lector.LeerR4(Insumos.R4(i));
            return (ase: Insumos.Ase(i), r1, r2, r4);
        }).ToList();

        var resultado = new CalculoRemuneracion().CalcularConsolidado(periodo, datos);

        var leafs = datos.Select(d =>
        {
            var leaf = leafReader.LeerLeafInputs(d.ase, periodo, Insumos.R1(d.ase.Id), Insumos.R2(d.ase.Id), Insumos.R4(d.ase.Id));
            // El balance de subsidios y contribuciones es un paso APARTE del lector de leafs en la
            // composición del período (ProcesadorPeriodo); el desglose trazable lo necesita en ambas
            // quincenas (DetRetri 'Balance de subsidio o contribuciones').
            leaf.BalanceSc = leafReader.LeerBalanceSc(d.ase, periodo, Insumos.Balance(d.ase.Id));
            leaf.DetRetriQ2 = new DetRetriQ2Inputs
            {
                Ase = leaf.Ase,
                TotalD104 = leaf.R1.TotalOportunoEsperadoPorAse
                    + leaf.R2.TotalOportunoEsperado
                    + leaf.R1.ExtemporaneoEsperadoPorAse
                    + leaf.R4.TotalReversionEsperada
                    + (leaf.AjustesSfT?.TotalAjustes ?? 0m)
            };
            return leaf;
        }).ToList();

        return (resultado, leafs);
    }

    // ── Resolución del asiento por ETIQUETA (mismo resolvedor del camino productivo) ────────────

    private static decimal LeerAsientoPorEtiqueta(string ruta, WorkbookLeafInputs leaf)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = HojaDetRetri(leaf.Periodo);
        var worksheet = Hoja(workbookPart, hoja);
        var destino = ResolvedorDestinoValor.ResolverTabla(
            workbookPart, worksheet,
            new ResolvedorDestinoValor.IdentidadTabla(hoja, null, leaf.Ase.Id, ResolvedorDestinoValor.EncabezadoRecaudoTotal));
        return Valor(destino.CeldaXml);
    }

    private static decimal LeerCierrePorEtiqueta(string ruta, Periodo periodo)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = HojaDetRetri(periodo);
        var worksheet = Hoja(workbookPart, hoja);
        var destino = ResolvedorDestinoValor.ResolverTabla(
            workbookPart, worksheet,
            new ResolvedorDestinoValor.IdentidadTabla(hoja, null, AseId: null, ResolvedorDestinoValor.EncabezadoRecaudoTotal));
        return Valor(destino.CeldaXml);
    }

    private static void EjecutarGateDeValores(string ruta, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        OpenXmlPlantillaWriter.VerificarAsientoDetRetri(workbookPart, leafs);
    }

    /// <summary>Fuerza el asiento a 0 (solo valores: la forma de las fórmulas no se toca).</summary>
    private static void ForzarAsientoACero(string ruta, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = HojaDetRetri(leafs[0].Periodo);
        var worksheet = Hoja(workbookPart, hoja);

        foreach (var leaf in leafs)
        {
            var destino = ResolvedorDestinoValor.ResolverTabla(
                workbookPart, worksheet,
                new ResolvedorDestinoValor.IdentidadTabla(hoja, null, leaf.Ase.Id, ResolvedorDestinoValor.EncabezadoRecaudoTotal));
            destino.CeldaXml.CellValue = new CellValue("0");
            destino.CeldaXml.DataType = CellValues.Number;
        }

        worksheet.Save();
    }

    private static string HojaDetRetri(Periodo periodo) => NombresHojaPeriodo.DetRetri(periodo);

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet.");
    }

    private static decimal Valor(Cell celda) =>
        decimal.TryParse(celda.CellValue?.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor)
            ? valor
            : 0m;

    private static string CrearDirectorioTemporal()
    {
        var dir = Path.Combine(Path.GetTempPath(), "p59-asiento-detretri-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void BorrarDirectorio(string ruta)
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
