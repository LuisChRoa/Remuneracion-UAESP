using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 29 (T4, Unidad D — desglose trazable): verifica END-TO-END que el writer escribe las
/// columnas de <c>DetRetri{AAAAMMQ}</c>/<c>DetValiRetri{AAAAMMQ}</c> cuyo origen workbook-interno
/// es de UNA arista (gate T0d, <c>plans/29-T0-Evidencia-B.md</c> §1.4) contra la plantilla canónica
/// en ceros del período (julio: <c>Docs/Prueba Julio-2/Plantilla_Remuneracion.xlsx</c>; agosto:
/// base <c>Docs/Prueba Agosto-2/Plantilla_Remuneracion.xlsx</c>, Plan 30/T3), con los insumos
/// REALES de julio-Q2 (<c>Docs/Prueba Julio-2/Insumos</c>) y agosto-Q2
/// (<c>Docs/Prueba Agosto-2/Insumos</c>).
///
/// Cobertura (R-D-1 + S5), reescrita por el Plan 62 (H2/H3 — D-59H4/D-61R10/D-J9):
///   - DetRetri J y L y DetValiRetri I/O viajan PEGADOS del R10 del período (N1). La derivación
///     histórica "J ← BCE F3" se conserva como diagnóstico y coincide en julio; el manual del
///     administrativo coincide con el R10 en los 5 ASE (la antigua divergencia ASE5 desaparece
///     porque la app ya no calcula J/L: copia el R10, fuente también del manual).
///   - DetValiRetri J (AJUSTE A LA DECENA) también entra pegado (D-J9); la exclusión SALE del
///     comparador (Plan 29) se retira en H5.
///   - DetRetri D 5/5 intacto (el desglose no mueve el dinero).
///   - R-D-1/guard: una celda-destino convertida en fórmula ⇒ ERR-PLANTILLA, sin sobrescritura.
///
/// Sin emojis. Tolerancia ±0.5.
/// </summary>
public sealed class DetRetriDesgloseTrazableTests
{
    private const decimal Tol = 0.5m;

    private static readonly Periodo PeriodoQ2 = Periodo.Parse("2026072");
    private static readonly Periodo PeriodoAgosto = Periodo.Parse("2026082");
    private static readonly string HojaDetRetri = WorkbookLeafCellMapDetRetri.HojaDetRetri(PeriodoQ2);
    private static readonly string HojaDetValiRetri = WorkbookLeafCellMapDetRetri.HojaDetValiRetri(PeriodoQ2);

    private static readonly string ManualJulio = Insumos.ManualJulioQ2;

    private static readonly string ManualAgosto = Insumos.ManualAgosto2026082;

    private static readonly string CarpetaJulio = Insumos.CarpetaInsumosJulioQ2;
    private static readonly string CarpetaAgosto = Insumos.CarpetaInsumosAgosto;

    // Hojas del período de agosto (sufijo 2026082): la base canónica 2026082 las trae y el writer
    // las resuelve por período (Plan 30/T1); el manual del administrativo usa el mismo sufijo.
    private static readonly string HojaDetRetriAgosto = WorkbookLeafCellMapDetRetri.HojaDetRetri(PeriodoAgosto);
    private static readonly string HojaDetValiRetriAgosto = WorkbookLeafCellMapDetRetri.HojaDetValiRetri(PeriodoAgosto);

    [Fact]
    public void JulioQ2_DesgloseTrazable_DetRetri_JyL_Y_DetVali_IyO()
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, Insumos.PlantillaQ2, salida.Ruta);
        Assert.True(File.Exists(ManualJulio), $"Falta el manual de julio: {ManualJulio}");

        var bloqueDet = new ExcelDataReaderDetRetriR10Reader().LeerBloqueDetRetri(PeriodoQ2, Insumos.R10Q2);
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var fila = WorkbookLeafCellMapDetRetri.ObtenerFila(aseId);

            // Plan 62 (H2/H3, D-59H4 + D-61R10): J/L de DetRetri y I de DetValiRetri viajan PEGADOS
            // del R10 del período. La derivación "J ← BCE F3" ya no es la fuente de la hoja; se
            // conserva como diagnóstico (en julio el R10 coincide con ROUND(BCE) en ambos canónicos).
            var bceFila = 2 + aseId; // D3..D7
            var esperadoJ = Redondear(LeerCeldaNumerica(salida.Ruta, "BCE SC POR FACT.", $"D{bceFila}")
                + LeerCeldaNumerica(salida.Ruta, "BCE SC POR FACT.", $"E{bceFila}"));
            var jEscrito = LeerCeldaNumerica(salida.Ruta, HojaDetRetri, $"J{fila}");
            Assert.True(
                Math.Abs(jEscrito - esperadoJ) <= Tol,
                $"julio ASE{aseId}: DetRetri J{fila} escrito={jEscrito} esperado ROUND(BCE.F{bceFila})={esperadoJ}.");

            // L = INTERVENTORIA F15 (2ª quincena) y DetValiRetri I = -L, O = 0: todos pegados del R10.
            var l = LeerCeldaNumerica(salida.Ruta, HojaDetRetri, $"L{fila}");
            Assert.InRange(l - LeerCeldaNumerica(ManualJulio, HojaDetRetri, $"L{fila}"), -Tol, Tol);
            Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaDetValiRetri, $"I{fila}") - (-l), -Tol, Tol);
            Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaDetValiRetri, $"O{fila}") - 0m, -Tol, Tol);

            // Reescritura formal Plan 62 (H3): el asiento de la hoja ES el R10 del período (N1). J
            // coincide con el bloque pegado en los 5 ASE —incluido ASE5— y con el manual del
            // administrativo; la antigua "divergencia del manual" (app fuente-fiel vs manual, T0a) ya
            // no aplica: la app no calcula J/L, los copia del R10, que es también la fuente del manual.
            Assert.InRange(jEscrito - (bloqueDet.Celda(fila, "J")?.Numerico ?? 0m), -Tol, Tol);
            Assert.InRange(l - (bloqueDet.Celda(fila, "L")?.Numerico ?? 0m), -Tol, Tol);
            Assert.InRange(
                LeerCeldaNumerica(ManualJulio, HojaDetRetri, $"J{fila}") - jEscrito,
                -Tol,
                Tol);
        }
    }

    [Fact]
    public void JulioQ2_DetValiRetri_J_TraeElDatoDelR10Pegado()
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, Insumos.PlantillaQ2, salida.Ruta);

        // Plan 62 (D-J9 + T0-09 + T2.2, reclassificado): "AJUSTE A LA DECENA" (DetValiRetri J) ya NO
        // se queda en el 0 de la plantilla — viaja PEGADO del R10 del periodo. La exclusion SALE del
        // comparador (Plan 29) se retira en H5/T5.1; aqui se fija que el dato existe y es el del R10.
        var bloqueVali = new ExcelDataReaderDetRetriR10Reader().LeerBloqueDetValiRetri(PeriodoQ2, Insumos.R10Q2);
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var fila = WorkbookLeafCellMapDetRetri.ObtenerFila(aseId);
            var esperado = bloqueVali.Celda(fila, "J")?.Numerico
                ?? throw new InvalidOperationException($"El bloque del R10 no trae J{fila}.");
            Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaDetValiRetri, $"J{fila}") - esperado, -Tol, Tol);
        }

        Assert.Contains(
            WorkbookLeafCellMapDetRetri.ColumnasExcluidas(PeriodoQ2),
            e => e.Hoja == HojaDetValiRetri && e.Columna == "J" && e.Motivo == WorkbookLeafCellMapDetRetri.MotivoSaleAjusteDecena);
    }

    [Fact]
    public void JulioQ2_DetRetri_D_5vs5_Intacto()
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, Insumos.PlantillaQ2, salida.Ruta);

        // El desglose trazable NO mueve el dinero: DetRetri D9:D13 sigue igual al manual (R-D-1).
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var celda = $"D{WorkbookLeafCellMapDetRetri.ObtenerFila(aseId)}";
            Assert.InRange(
                LeerCeldaNumerica(salida.Ruta, HojaDetRetri, celda) - LeerCeldaNumerica(ManualJulio, HojaDetRetri, celda),
                -Tol,
                Tol);
        }
    }

    [Fact]
    public void AgostoQ2_DesgloseTrazable_DetRetri_JyL_Y_DetVali_I()
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaAgosto, new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 }, Insumos.PlantillaAgosto2026082, salida.Ruta);
        Assert.True(File.Exists(ManualAgosto), $"Falta el manual de agosto: {ManualAgosto}");

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            var fila = WorkbookLeafCellMapDetRetri.ObtenerFila(aseId);

            var bceFila = 2 + aseId;
            var esperadoJ = Redondear(LeerCeldaNumerica(salida.Ruta, "BCE SC POR FACT.", $"D{bceFila}")
                + LeerCeldaNumerica(salida.Ruta, "BCE SC POR FACT.", $"E{bceFila}"));
            Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaDetRetriAgosto, $"J{fila}") - esperadoJ, -Tol, Tol);

            // Agosto: la app es fuente-fiel y el manual (DetRetri2026082) coincide en J/L.
            Assert.InRange(
                LeerCeldaNumerica(salida.Ruta, HojaDetRetriAgosto, $"J{fila}") - LeerCeldaNumerica(ManualAgosto, HojaDetRetriAgosto, $"J{fila}"),
                -Tol,
                Tol);
            var l = LeerCeldaNumerica(salida.Ruta, HojaDetRetriAgosto, $"L{fila}");
            Assert.InRange(l - LeerCeldaNumerica(ManualAgosto, HojaDetRetriAgosto, $"L{fila}"), -Tol, Tol);
            Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaDetValiRetriAgosto, $"I{fila}") - (-l), -Tol, Tol);
        }
    }

    [Fact]
    public void Mapa_ColumnasExcluidas_NoSolapanTrazables_Y_ListanMotivo()
    {
        // R-D-1: toda columna escritura es trazable; toda excluida tiene motivo no vacío.
        Assert.All(WorkbookLeafCellMapDetRetri.ColumnasExcluidas(PeriodoQ2), e => Assert.False(string.IsNullOrWhiteSpace(e.Motivo)));

        var trazables = WorkbookLeafCellMapDetRetri.ColumnasDetRetri
            .Select(c => (WorkbookLeafCellMapDetRetri.HojaDetRetri(PeriodoQ2), c.Columna))
            .Concat(WorkbookLeafCellMapDetRetri.ColumnasDetValiRetri.Select(c => (WorkbookLeafCellMapDetRetri.HojaDetValiRetri(PeriodoQ2), c.Columna)))
            .ToList();

        foreach (var (hoja, columna) in trazables)
        {
            Assert.DoesNotContain(
                WorkbookLeafCellMapDetRetri.ColumnasExcluidas(PeriodoQ2),
                e => e.Hoja == hoja && e.Columna == columna);
        }
    }

    [Fact]
    public void DesgloseTrazable_DestinoFormula_LanzaErrPlantilla_SinSobrescribir()
    {
        using var temp = new DirectorioTemporal("remuneracion-detretri-desglose-");
        var plantillaMutada = Path.Combine(temp.Ruta, "plantilla-mutada.xlsx");
        File.Copy(Insumos.PlantillaQ2, plantillaMutada);

        // J9 = destino-trazable DetRetri del ASE1: se convierte en fórmula ⇒ guard ⇒ ERR-PLANTILLA.
        ConvertirEnFormula(plantillaMutada, HojaDetRetri, "J9", "1+1");

        var salida = Path.Combine(temp.Ruta, "salida.xlsx");
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, plantillaMutada, salida));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains("fórmula", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────────────────

    private static decimal Redondear(decimal valor) => decimal.Round(valor, 0, MidpointRounding.AwayFromZero);

    private static void EjecutarFlujo(string carpetaInsumos, Periodo periodo, string rutaPlantilla, string rutaSalida)
    {
        var procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader());

        procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = periodo,
            CarpetaPeriodo = carpetaInsumos,
            RutaPlantilla = rutaPlantilla,
            RutaSalida = rutaSalida
        });
    }

    private static decimal LeerCeldaNumerica(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var worksheet = Hoja(workbookPart, hoja);
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        if (objCelda.CellValue is null || string.IsNullOrWhiteSpace(objCelda.CellValue.InnerText))
        {
            return 0m;
        }

        return decimal.TryParse(objCelda.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor)
            ? valor
            : 0m;
    }

    private static void ConvertirEnFormula(string ruta, string hoja, string celda, string formula)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var worksheet = Hoja(documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null"), hoja);
        var objCelda = worksheet.Descendants<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No existe {hoja}!{celda} en {Path.GetFileName(ruta)}.");

        objCelda.CellValue?.Remove();
        objCelda.DataType = null;
        objCelda.CellFormula = new CellFormula(formula);
        worksheet.Save();
    }

    private static Worksheet Hoja(WorkbookPart workbookPart, string nombre)
    {
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet.");
    }

    private static SalidaTemporal NuevaSalida() => new("remuneracion-detretri-" + Guid.NewGuid().ToString("N"));

    private sealed class SalidaTemporal : IDisposable
    {
        private readonly DirectorioTemporal _dir;

        public SalidaTemporal(string prefijo)
        {
            _dir = new DirectorioTemporal(prefijo);
            Ruta = Path.Combine(_dir.Ruta, "salida.xlsx");
        }

        public string Ruta { get; }

        public void Dispose() => _dir.Dispose();
    }

    private sealed class DirectorioTemporal : IDisposable
    {
        public DirectorioTemporal(string prefijo)
        {
            Ruta = Path.Combine(Path.GetTempPath(), prefijo + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Ruta);
        }

        public string Ruta { get; }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(Ruta))
                {
                    Directory.Delete(Ruta, recursive: true);
                }
            }
            catch
            {
                // best-effort
            }
        }
    }
}
