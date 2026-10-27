using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 59 (§11 — FRENTE 1): el DESGLOSE trazable de las hojas de detalle del período se escribe en
/// AMBAS quincenas y lo escrito se contrasta contra el dominio y contra el R10 del período.
///
/// Defecto que cubre (medido en el entregable real de 2026101 — quincena 1): el writer escribía el
/// desglose sólo bajo <c>if (esQuincena2)</c>, así que en la primera quincena el entregable quedaba con
/// el literal 0 de la plantilla en <c>DetRetri</c> ('Balance de subsidio o contribuciones',
/// 'Recaudo base de remuneración', 'Costo de la Interventoría') y en <c>DetValiRetri</c> ('Costo de la
/// Interventoría', 'Validación retribución'), mientras la corrida certificaba <c>RESULTADO OK</c>.
/// El oráculo administrativo del mismo período trae esas columnas llenas.
///
/// Lo que se prueba:
/// <list type="number">
/// <item>Q1: lo escrito coincide con el dominio (balance = BCE Subsidio + Contribución redondeado;
/// costo = la mitad declarada para la PRIMERA quincena) y con el R10 oficial del período, fila por
/// fila y en la fila de cierre.</item>
/// <item>Q2: el costo es la mitad declarada para la SEGUNDA quincena (la única dependencia legítima de
/// quincena) — y el balance no cambia con la quincena.</item>
/// <item>Inyección: si el desglose queda en cero, el gate de valores ENROJECE con
/// <c>ERR-ESCRITURA</c> y el workbook no se entrega.</item>
/// </list>
///
/// Los destinos se resuelven por el RÓTULO de la columna y el id del ASE (mismo resolvedor del camino
/// productivo): el fixture no escribe ninguna coordenada.
/// </summary>
public sealed class DesgloseTrazableAmbasQuincenasPlan59Tests
{
    private const string RotuloBalance = "Balance de subsidio o contribuciones";
    private const string RotuloInterventoria = "Costo de la Interventoría";
    private const string RotuloValidacionRetribucion = "Validación retribución";

    [Fact]
    public void Q1_Desglose_CoincideConElDominioYConElR10DelPeriodo()
    {
        var periodo = Insumos.Periodo();
        var (leafs, resultado) = ConstruirPeriodo(periodo, Insumos.R1, Insumos.R2, Insumos.R4, i => BalanceDe(Insumos.CarpetasAse[i - 1], i));
        var salidaDir = NuevoDirectorio();
        try
        {
            var salida = Path.Combine(salidaDir, "salida.xlsx");
            new OpenXmlPlantillaWriter().GenerarWorkbook(Insumos.Plantilla, salida, resultado, leafs);

            var hojaDetRetri = NombresHojaPeriodo.DetRetri(periodo);
            var hojaDetValiRetri = NombresHojaPeriodo.DetValiRetri(periodo);
            var interventoriaPrimera = InterventoriaDeclarada.PorQuincena(1);

            foreach (var leaf in leafs)
            {
                var aseId = leaf.Ase.Id;

                // J: balance de subsidio o contribuciones = BCE SC POR FACT. (Subsidio + Contribución).
                var balance = Redondear(leaf.BalanceSc!.Ases.Single(a => a.Ase.Id == aseId).TotalBsc);
                var balanceEscrito = LeerPorRotulo(salida, hojaDetRetri, aseId, RotuloBalance);
                Assert.Equal(balance, balanceEscrito);
                Assert.NotEqual(0m, balanceEscrito);

                // L: costo de la interventoría de la PRIMERA quincena (la que el período declara).
                var costoEscrito = LeerPorRotulo(salida, hojaDetRetri, aseId, RotuloInterventoria);
                Assert.Equal(interventoriaPrimera[aseId], costoEscrito);

                // DetValiRetri: el mismo costo negado + bandera de validación en 0.
                Assert.Equal(-costoEscrito, LeerPorRotulo(salida, hojaDetValiRetri, aseId, RotuloInterventoria));
                Assert.Equal(0m, LeerPorRotulo(salida, hojaDetValiRetri, aseId, RotuloValidacionRetribucion));
            }

            // Oráculo: el R10 oficial del período declara el desglose completo. Se compara fila por
            // fila y la fila de cierre (Σ de las filas por ASE, sin re-redondeo).
            var r10 = Insumos.R10;
            foreach (var leaf in leafs)
            {
                var aseId = leaf.Ase.Id;
                Assert.Equal(
                    LeerPorRotulo(r10, hojaDetRetri, aseId, RotuloBalance),
                    LeerPorRotulo(salida, hojaDetRetri, aseId, RotuloBalance));
                Assert.Equal(
                    LeerPorRotulo(r10, hojaDetRetri, aseId, RotuloInterventoria),
                    LeerPorRotulo(salida, hojaDetRetri, aseId, RotuloInterventoria));
                Assert.Equal(
                    LeerPorRotulo(r10, hojaDetValiRetri, aseId, RotuloInterventoria),
                    LeerPorRotulo(salida, hojaDetValiRetri, aseId, RotuloInterventoria));
            }

            var cierreBalance = LeerPorRotulo(salida, hojaDetRetri, aseId: null, RotuloBalance);
            var cierreCosto = LeerPorRotulo(salida, hojaDetRetri, aseId: null, RotuloInterventoria);
            Assert.Equal(leafs.Sum(l => Redondear(l.BalanceSc!.Ases.Single(a => a.Ase.Id == l.Ase.Id).TotalBsc)), cierreBalance);
            Assert.Equal(leafs.Sum(l => interventoriaPrimera[l.Ase.Id]), cierreCosto);
            Assert.Equal(LeerPorRotulo(r10, hojaDetRetri, aseId: null, RotuloBalance), cierreBalance);
            Assert.Equal(LeerPorRotulo(r10, hojaDetRetri, aseId: null, RotuloInterventoria), cierreCosto);
            Assert.Equal(
                -cierreCosto,
                LeerPorRotulo(salida, hojaDetValiRetri, aseId: null, RotuloInterventoria));
        }
        finally
        {
            BorrarDirectorio(salidaDir);
        }
    }

    [Fact]
    public void Q2_Desglose_UsaLaSegundaQuincena_Y_ElBalanceNoCambiaConLaQuincena()
    {
        var q1 = Insumos.Periodo();
        var q2 = Insumos.PeriodoQ2();
        var (leafsQ1, _) = ConstruirPeriodo(q1, Insumos.R1, Insumos.R2, Insumos.R4, i => BalanceDe(Insumos.CarpetasAse[i - 1], i));
        var (leafsQ2, resultadoQ2) = ConstruirPeriodo(q2, Insumos.R1Q2, Insumos.R2Q2, Insumos.R4Q2, i => BalanceDe(Insumos.CarpetasAseQ2[i - 1], i));

        var salidaDir = NuevoDirectorio();
        try
        {
            var salidaQ2 = Path.Combine(salidaDir, "q2.xlsx");
            new OpenXmlPlantillaWriter().GenerarWorkbook(
                Insumos.PlantillaQ2, salidaQ2, resultadoQ2, leafsQ2);

            var hojaDetRetri = NombresHojaPeriodo.DetRetri(q2);
            var segunda = InterventoriaDeclarada.PorQuincena(2);
            var primera = InterventoriaDeclarada.PorQuincena(1);

            foreach (var leaf in leafsQ2)
            {
                var aseId = leaf.Ase.Id;
                var costo = LeerPorRotulo(salidaQ2, hojaDetRetri, aseId, RotuloInterventoria);
                Assert.Equal(segunda[aseId], costo);
                Assert.Equal(
                    Redondear(leaf.BalanceSc!.Ases.Single(a => a.Ase.Id == aseId).TotalBsc),
                    LeerPorRotulo(salidaQ2, hojaDetRetri, aseId, RotuloBalance));
            }

            // ASE1: las dos mitades declaradas difieren en 1 peso ⇒ la quincena se selecciona de
            // verdad (si alguien volviera a fijar una sola, este assert lo delata).
            Assert.NotEqual(primera[1], segunda[1]);
            Assert.Equal(
                segunda.Values.Sum(),
                LeerPorRotulo(salidaQ2, hojaDetRetri, aseId: null, RotuloInterventoria));

            // Y el balance no depende de la quincena en su REGLA (BCE del período ya escrito): se
            // verifica en las dos direcciones y no se compara entre períodos, porque el balance que la
            // fuente trae para la 1.ª y la 2.ª quincena del mismo mes son observaciones distintas.
            Assert.All(leafsQ2, l => Assert.NotEqual(
                0m,
                LeerPorRotulo(salidaQ2, hojaDetRetri, l.Ase.Id, RotuloBalance)));
        }
        finally
        {
            BorrarDirectorio(salidaDir);
        }
    }

    [Fact]
    public void Inyeccion_DesgloseForzadoACero_EnrojeceConErrEscritura()
    {
        var periodo = Insumos.Periodo();
        var (leafs, resultado) = ConstruirPeriodo(periodo, Insumos.R1, Insumos.R2, Insumos.R4, i => BalanceDe(Insumos.CarpetasAse[i - 1], i));
        var salidaDir = NuevoDirectorio();
        try
        {
            var salida = Path.Combine(salidaDir, "salida.xlsx");
            new OpenXmlPlantillaWriter().GenerarWorkbook(Insumos.Plantilla, salida, resultado, leafs);

            // Control: tal como lo dejó el writer, el gate de valores pasa.
            EjecutarGateDeValores(salida, leafs);

            // Inyección: se reproduce el escenario histórico (desglose en el 0 de la plantilla)
            // dejando intacto el esperado del dominio.
            ForzarACero(salida, NombresHojaPeriodo.DetRetri(periodo), leafs[0].Ase.Id, RotuloInterventoria);

            var ex = Assert.Throws<CalculoInvalidoException>(() => EjecutarGateDeValores(salida, leafs));
            Assert.Equal(CodigoError.Escritura, ex.Codigo);
            Assert.Contains("desglose trazable", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("ASE 1", ex.Message, StringComparison.Ordinal);
            Assert.Contains(RotuloInterventoria, ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            BorrarDirectorio(salidaDir);
        }
    }

    // ── Armado del período (misma composición que ProcesadorPeriodo) ────────────────────────────

    private static (List<WorkbookLeafInputs> Leafs, ResultadoRemuneracion Resultado) ConstruirPeriodo(
        Periodo periodo,
        Func<int, string> rutaR1,
        Func<int, string> rutaR2,
        Func<int, string> rutaR4,
        Func<int, string> rutaBalance)
    {
        var leafReader = new ExcelDataReaderWorkbookLeafInputReader();
        var lector = new ExcelDataReaderRecaudoReader();
        var leafs = new List<WorkbookLeafInputs>();
        var datos = new List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4, decimal ajustesSfT)>();

        for (var i = 1; i <= 5; i++)
        {
            var ase = Insumos.Ase(i);
            var r1 = rutaR1(i);
            var r2 = rutaR2(i);
            var r4 = rutaR4(i);

            var leaf = leafReader.LeerLeafInputs(ase, periodo, r1, r2, r4);
            // Misma composición que ProcesadorPeriodo: el balance de subsidios y contribuciones es un
            // paso aparte del lector de leafs (el desglose trazable lo necesita en ambas quincenas).
            leaf.BalanceSc = leafReader.LeerBalanceSc(ase, periodo, rutaBalance(i));
            leaf.DetRetriQ2 = new DetRetriQ2Inputs
            {
                Ase = leaf.Ase,
                TotalD104 = leaf.R1.TotalOportunoEsperadoPorAse
                    + leaf.R2.TotalOportunoEsperado
                    + leaf.R1.ExtemporaneoEsperadoPorAse
                    + leaf.R4.TotalReversionEsperada
                    + (leaf.AjustesSfT?.TotalAjustes ?? 0m)
            };
            leafs.Add(leaf);
            datos.Add((ase, lector.LeerR1(r1), lector.LeerR2(r2), lector.LeerR4(r4), 0m));
        }

        return (leafs, new CalculoRemuneracion().CalcularConsolidado(periodo, datos));
    }

    /// <summary>
    /// Ruta del <c>R4-BalanceSubsidioyContribuciones</c> del ASE en la carpeta dada (mismos prefijos
    /// del locator: base y variante <c>-Optimizado_</c>).
    /// </summary>
    private static string BalanceDe(string carpetaAse, int aseId)
    {
        var archivo = Directory.EnumerateFiles(carpetaAse, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f)
                .StartsWith("R4-BalanceSubsidioyContribuciones_", StringComparison.OrdinalIgnoreCase))
            ?? Directory.EnumerateFiles(carpetaAse, "*.xlsx", SearchOption.TopDirectoryOnly)
                .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f)
                    .StartsWith("R4-BalanceSubsidioyContribuciones-Optimizado_", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Balance ASE{aseId} en {carpetaAse}");
        return archivo!;
    }

    // ── Lectura por RÓTULO (mismo resolvedor del camino productivo) ─────────────────────────────

    private static decimal LeerPorRotulo(string ruta, string hoja, int? aseId, string rotulo)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, hoja);
        var destino = ResolvedorDestinoValor.ResolverTablaPorConcepto(
            workbookPart, worksheet,
            new ResolvedorDestinoValor.IdentidadTabla(hoja, Titulo: null, aseId, rotulo));
        return Valor(destino.CeldaXml, workbookPart);
    }

    /// <summary>
    /// Valor de una celda: numérico directo o TEXTO (el R10 del administrativo trae la tabla formateada
    /// como moneda, <c>"$2.032.363.679"</c>). El texto se normaliza (sin <c>$</c>, separador de miles
    /// <c>.</c>, decimal <c>,</c>) antes de compararlo — nunca se interpreta como índice de shared string.
    /// </summary>
    private static decimal Valor(Cell celda, WorkbookPart workbookPart)
    {
        string? texto;
        if (celda.DataType?.Value == CellValues.SharedString)
        {
            var tabla = workbookPart.SharedStringTablePart?.SharedStringTable
                ?? throw new InvalidOperationException("El workbook no trae tabla de cadenas compartidas.");
            texto = tabla.ChildElements[int.Parse(celda.CellValue!.InnerText, CultureInfo.InvariantCulture)].InnerText;
        }
        else if (celda.DataType?.Value == CellValues.InlineString)
        {
            texto = celda.InlineString?.Text?.Text;
        }
        else
        {
            texto = celda.CellValue?.InnerText;
        }

        if (string.IsNullOrWhiteSpace(texto))
        {
            return 0m;
        }

        if (decimal.TryParse(texto, NumberStyles.Number, CultureInfo.InvariantCulture, out var numerico))
        {
            return numerico;
        }

        var normalizado = texto.Replace("$", string.Empty, StringComparison.Ordinal)
            .Replace(".", string.Empty, StringComparison.Ordinal)
            .Replace(",", ".", StringComparison.Ordinal)
            .Trim();
        return decimal.TryParse(normalizado, NumberStyles.Number, CultureInfo.InvariantCulture, out var moneda)
            ? moneda
            : throw new InvalidOperationException($"No se pudo leer el valor '{texto}' como número.");
    }

    private static void EjecutarGateDeValores(string ruta, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        OpenXmlPlantillaWriter.VerificarAsientoDetRetri(workbookPart, leafs);
    }

    private static void ForzarACero(string ruta, string hoja, int aseId, string rotulo)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, hoja);
        var destino = ResolvedorDestinoValor.ResolverTablaPorConcepto(
            workbookPart, worksheet,
            new ResolvedorDestinoValor.IdentidadTabla(hoja, Titulo: null, aseId, rotulo));
        destino.CeldaXml.CellValue = new CellValue("0");
        destino.CeldaXml.DataType = CellValues.Number;
        worksheet.Save();
    }

    private static decimal Redondear(decimal valor) => DetRetriRounder.Round(valor);

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet.");
    }

    private static string NuevoDirectorio()
    {
        var dir = Path.Combine(Path.GetTempPath(), "p59-desglose-" + Guid.NewGuid().ToString("N"));
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
