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
/// Plan 29 (T3, Unidad R — ESCRITURA): verifica END-TO-END que el desglose-detalle por componente
/// de <c>Rem. Anticipos R2</c> y <c>Reversion Pagos R4</c> se escribe contra la plantilla canónica
/// en ceros del período (julio: <c>Docs/Prueba Julio-2/Plantilla_Remuneracion.xlsx</c>;
/// agosto: base <c>Docs/Prueba Agosto-2/Plantilla_Remuneracion.xlsx</c>, Plan 30/T3), con los
/// insumos REALES de julio-Q2 (<c>Docs/Prueba Julio-2/Insumos</c>) y agosto-Q2
/// (<c>Docs/Prueba Agosto-2/Insumos</c>).
///
/// Cobertura (S3/S4 + R-R-1/R-R-2/R-R-3):
///   - Julio: bloque-destino completo (todas las filas y columnas del mapa) contra el MANUAL del
///     administrativo (misma malla canónica ⇒ oráculo independiente) y la primera fila
///     <c>Vlr Servicio</c> contra la FUENTE leída por el reader.
///   - Agosto: primera fila <c>Vlr Servicio</c> contra la FUENTE (la malla de agosto deriva en
///     LIME/BOGOTA/PROMO/CIUDAD — T0b; por eso no se compara bloque a bloque con el manual, cuya
///     malla de agosto está redimensionada) + la invariante <c>Especiales K</c> del ASE1 (144941.68).
///   - R-R-3: una celda-destino convertida en fórmula en una copia temporal ⇒ ERR-PLANTILLA,
///     sin sobrescritura.
///
/// Sin emojis. Tolerancia ±0.5.
/// </summary>
public sealed class DetalleR2R4EscrituraTests
{
    private const decimal Tol = 0.5m;

    private const string HojaR2 = WorkbookLeafCellMapDetalleR2R4.HojaR2;
    private const string HojaR4 = WorkbookLeafCellMapDetalleR2R4.HojaR4;

    private static readonly string ManualJulio = Insumos.ManualJulioQ2;

    private static readonly string ManualAgosto = Insumos.ManualAgosto2026082;

    private static readonly string CarpetaJulio = Insumos.CarpetaInsumosJulioQ2;
    private static readonly string CarpetaAgosto = Insumos.CarpetaInsumosAgosto;

    // ── S3/S4: celdas citadas de la evidencia T0b (julio ASE1) ────────────────────────────────

    [Fact]
    public void JulioAse1_CeldasCitadas_IgualanFuenteYManual()
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, Insumos.PlantillaQ2, salida.Ruta);

        // R2 E3 = 104754634.94, F3 = 6975298.69, K3 (Especiales) = 0 (julio sin columna).
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "E3") - 104754634.94m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "F3") - 6975298.69m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "K3") - 0m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(ManualJulio, HojaR2, "E3") - 104754634.94m, -Tol, Tol);

        // R4 D3 = -15794348.21 (primer Vlr Servicio), D9 = -762167.27 (OCCIDENTE).
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR4, "D3") - (-15794348.21m), -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR4, "D9") - (-762167.27m), -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(ManualJulio, HojaR4, "D3") - (-15794348.21m), -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(ManualJulio, HojaR4, "D9") - (-762167.27m), -Tol, Tol);
    }

    // ── R-R-1: julio, bloque-destino completo vs manual (misma malla canónica) ────────────────

    [Fact]
    public void JulioQ2_DetalleR2_BloqueCompleto_IgualaManual()
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, Insumos.PlantillaQ2, salida.Ruta);
        Assert.True(File.Exists(ManualJulio), $"Falta el manual de julio: {ManualJulio}");

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            AssertBloqueVsManual(salida.Ruta, ManualJulio, HojaR2, WorkbookLeafCellMapDetalleR2R4.BloquesR2[aseId], WorkbookLeafCellMapDetalleR2R4.ColumnasR2, $"R2 ASE{aseId} julio");
        }
    }

    [Fact]
    public void JulioQ2_DetalleR4_BloqueCompleto_IgualaManual()
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, Insumos.PlantillaQ2, salida.Ruta);
        Assert.True(File.Exists(ManualJulio), $"Falta el manual de julio: {ManualJulio}");

        for (var aseId = 1; aseId <= 5; aseId++)
        {
            AssertBloqueVsManual(salida.Ruta, ManualJulio, HojaR4, WorkbookLeafCellMapDetalleR2R4.BloquesR4[aseId], WorkbookLeafCellMapDetalleR2R4.ColumnasR4, $"R4 ASE{aseId} julio");
        }
    }

    // ── R-R-1/R-R-2: julio y agosto, primera fila Vlr Servicio vs fuente (reader) ─────────────

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void JulioQ2_DetalleR2_PrimeraVlrServicio_IgualaFuente(int aseId)
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, Insumos.PlantillaQ2, salida.Ruta);

        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var fuente = reader.LeerDetalleR2(Insumos.Ase(aseId), Insumos.R2JulioQ2(aseId));
        AssertFilaVlrServicioR2(salida.Ruta, aseId, fuente, $"julio ASE{aseId}");
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void JulioQ2_DetalleR4_PrimeraVlrServicio_IgualaFuente(int aseId)
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, Insumos.PlantillaQ2, salida.Ruta);

        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var fuente = reader.LeerDetalleR4(Insumos.Ase(aseId), Insumos.R4JulioQ2(aseId));
        AssertFilaVlrServicioR4(salida.Ruta, aseId, fuente, $"julio ASE{aseId}");
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void AgostoQ2_DetalleR2_PrimeraVlrServicio_IgualaFuente(int aseId)
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaAgosto, new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 }, Insumos.PlantillaAgosto2026082, salida.Ruta);

        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var fuente = reader.LeerDetalleR2(Insumos.Ase(aseId), Insumos.R2AgostoQ2(aseId));
        AssertFilaVlrServicioR2(salida.Ruta, aseId, fuente, $"agosto ASE{aseId}");
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void AgostoQ2_DetalleR4_PrimeraVlrServicio_IgualaFuente(int aseId)
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaAgosto, new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 }, Insumos.PlantillaAgosto2026082, salida.Ruta);

        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var fuente = reader.LeerDetalleR4(Insumos.Ase(aseId), Insumos.R4AgostoQ2(aseId));
        AssertFilaVlrServicioR4(salida.Ruta, aseId, fuente, $"agosto ASE{aseId}");
    }

    /// <summary>
    /// Agosto ASE1: la fuente trae <c>Especiales</c> en K y el destino la escribe (invariante
    /// de la columna ausente en julio). E3 = 68638686.24 / F3 = 3508770.06 / K3 = 144941.68.
    /// </summary>
    [Fact]
    public void AgostoAse1_EspecialesEnK_IgualaFuenteYManual()
    {
        using var salida = NuevaSalida();
        EjecutarFlujo(CarpetaAgosto, new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 }, Insumos.PlantillaAgosto2026082, salida.Ruta);

        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "E3") - 68638686.24m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "F3") - 3508770.06m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(salida.Ruta, HojaR2, "K3") - 144941.68m, -Tol, Tol);

        Assert.True(File.Exists(ManualAgosto), $"Falta el manual de agosto: {ManualAgosto}");
        Assert.InRange(LeerCeldaNumerica(ManualAgosto, HojaR2, "E3") - 68638686.24m, -Tol, Tol);
        Assert.InRange(LeerCeldaNumerica(ManualAgosto, HojaR2, "K3") - 144941.68m, -Tol, Tol);
    }

    // ── R-R-3: destino-detalle fórmula ⇒ ERR-PLANTILLA (copia temp, nunca sobrescritura) ──────

    [Fact]
    public void DetalleR2_DestinoFormula_LanzaErrPlantilla_SinSobrescribir()
    {
        using var temp = new DirectorioTemporal("remuneracion-detalle-formula-");
        var plantillaMutada = Path.Combine(temp.Ruta, "plantilla-mutada.xlsx");
        File.Copy(Insumos.PlantillaQ2, plantillaMutada);

        // E3 = primera celda de captura del detalle R2 del ASE1: se convierte en fórmula.
        ConvertirEnFormula(plantillaMutada, HojaR2, "E3", "1+1");

        var salida = Path.Combine(temp.Ruta, "salida.xlsx");
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            EjecutarFlujo(CarpetaJulio, new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 }, plantillaMutada, salida));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains("fórmula", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────────────────

    private static void AssertBloqueVsManual(
        string salida,
        string manual,
        string hoja,
        (int FilaInicio, int FilaFin) bloque,
        IReadOnlyList<(string? Encabezado, string Columna)> columnas,
        string contexto)
    {
        for (var fila = bloque.FilaInicio; fila <= bloque.FilaFin; fila++)
        {
            foreach (var (_, columna) in columnas)
            {
                var celda = columna + fila.ToString(CultureInfo.InvariantCulture);
                var real = LeerCeldaNumerica(salida, hoja, celda);
                var esperado = LeerCeldaNumerica(manual, hoja, celda);
                Assert.True(
                    Math.Abs(real - esperado) <= Tol,
                    $"{contexto}: {hoja}!{celda} escrito={real} manual={esperado} (Δ={real - esperado}).");
            }
        }
    }

    private static void AssertFilaVlrServicioR2(string salida, int aseId, DetalleR2AseInputs fuente, string contexto)
    {
        var fila = fuente.Filas.FirstOrDefault(f => f.Firma == "|||Vlr Servicio")
            ?? throw new InvalidOperationException($"{contexto}: la fuente no trae la fila 'Vlr Servicio'.");
        var inicio = WorkbookLeafCellMapDetalleR2R4.BloquesR2[aseId].FilaInicio;
        foreach (var (encabezado, columna) in WorkbookLeafCellMapDetalleR2R4.ColumnasR2)
        {
            var esperado = encabezado is null ? 0m : fila.Valor(encabezado) ?? 0m;
            var real = LeerCeldaNumerica(salida, HojaR2, columna + inicio.ToString(CultureInfo.InvariantCulture));
            Assert.True(
                Math.Abs(real - esperado) <= Tol,
                $"{contexto}: R2 {columna}{inicio} escrito={real} fuente={esperado} (Δ={real - esperado}).");
        }
    }

    private static void AssertFilaVlrServicioR4(string salida, int aseId, DetalleR4AseInputs fuente, string contexto)
    {
        var fila = fuente.Filas.FirstOrDefault(f => f.C == "Vlr Servicio")
            ?? throw new InvalidOperationException($"{contexto}: la fuente no trae la fila 'Vlr Servicio'.");
        var inicio = WorkbookLeafCellMapDetalleR2R4.BloquesR4[aseId].FilaInicio;
        foreach (var (encabezado, columna) in WorkbookLeafCellMapDetalleR2R4.ColumnasR4)
        {
            var esperado = encabezado is null ? 0m : fila.Valor(encabezado) ?? 0m;
            var real = LeerCeldaNumerica(salida, HojaR4, columna + inicio.ToString(CultureInfo.InvariantCulture));
            Assert.True(
                Math.Abs(real - esperado) <= Tol,
                $"{contexto}: R4 {columna}{inicio} escrito={real} fuente={esperado} (Δ={real - esperado}).");
        }
    }

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

    private static SalidaTemporal NuevaSalida() => new("remuneracion-detalle-" + Guid.NewGuid().ToString("N"));

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
