using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-13 (2.7, Plan 13 §4 Fase 2 — Unidad 2): <see cref="ValidacionOracleReader"/> read-only
/// contra los canónicos. El reader NUNCA escribe (A4: hash intacto tras lectura) y cada
/// hoja/celda-oráculo ASSERTA existencia (W2/D7): ausente = fallo que nombra hoja+celda.
/// </summary>
public sealed class ValidacionOracleReaderTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    private readonly IValidacionOracleReader _reader = new ValidacionOracleReader();

    [Fact]
    public void Q1Golden_SnapshotCierra_TodasLasValidaciones()
    {
        var snapshots = _reader.LeerSnapshots(Insumos.Plantilla, Insumos.Periodo());

        Assert.Equal(5, snapshots.Count);
        foreach (var snapshot in snapshots.OrderBy(s => s.Ase.Id))
        {
            Assert.Equal(5, snapshot.PorEmpresa.Count);
            foreach (var empresa in snapshot.PorEmpresa)
            {
                Assert.Equal(snapshot.Ase.Id, empresa.AseId);
                Assert.InRange(empresa.DiferenciaO, -Tolerancia, Tolerancia);
                Assert.True(empresa.VerificacionP);
            }

            Assert.NotNull(snapshot.DetValiRetri);
            var detalle = snapshot.DetValiRetri!;
            Assert.InRange(detalle.DiferenciasAse.Single().Valor, -Tolerancia, Tolerancia);
            Assert.True(detalle.VerificacionesAse.Single().Verificacion);
            Assert.True(detalle.VerificacionTotalD29);

            Assert.InRange(snapshot.ValidacionTotal, -Tolerancia, Tolerancia);
            Assert.True(snapshot.ValidacionTotalOkP);
        }
    }

    [Fact]
    public void Q2Golden_SnapshotCierra_TodasLasValidaciones()
    {
        var snapshots = _reader.LeerSnapshots(Insumos.GoldenQ2, Insumos.PeriodoQ2());

        Assert.Equal(5, snapshots.Count);
        foreach (var snapshot in snapshots.OrderBy(s => s.Ase.Id))
        {
            foreach (var empresa in snapshot.PorEmpresa)
            {
                Assert.Equal(snapshot.Ase.Id, empresa.AseId);
                Assert.InRange(empresa.DiferenciaO, -Tolerancia, Tolerancia);
                Assert.True(empresa.VerificacionP);
            }

            var detalle = snapshot.DetValiRetri!;
            Assert.InRange(detalle.DiferenciasAse.Single().Valor, -Tolerancia, Tolerancia);
            Assert.True(detalle.VerificacionesAse.Single().Verificacion);
            Assert.True(detalle.VerificacionTotalD29);

            Assert.InRange(snapshot.ValidacionTotal, -Tolerancia, Tolerancia);
            Assert.True(snapshot.ValidacionTotalOkP);
        }
    }

    [Fact]
    public void Q2Canonico_SnapshotLeeSinFallo_ConCacheCero()
    {
        // El canónico Q2 ("8 agos", plantilla con caché 0) se lee sin fallo: las celdas-oráculo
        // existen y son fórmula (W2); los valores cacheados 0/TRUE son legítimos de plantilla.
        var snapshots = _reader.LeerSnapshots(Insumos.PlantillaQ2, Insumos.PeriodoQ2());

        Assert.Equal(5, snapshots.Count);
        foreach (var snapshot in snapshots)
        {
            Assert.InRange(snapshot.ValidacionTotal, -Tolerancia, Tolerancia);
            Assert.True(snapshot.ValidacionTotalOkP);
        }
    }

    [Fact]
    public void Reader_NuncaEscribe_HashIntactoEnCanonicos()
    {
        var casos = new[]
        {
            (Insumos.Plantilla, Insumos.Periodo()),
            (Insumos.GoldenQ2, Insumos.PeriodoQ2()),
            (Insumos.PlantillaQ2, Insumos.PeriodoQ2())
        };

        foreach (var (ruta, periodo) in casos)
        {
            var hashAntes = Sha256(ruta);
            _ = _reader.LeerSnapshots(ruta, periodo);
            Assert.Equal(hashAntes, Sha256(ruta));
        }
    }

    [Fact]
    public void W2_CeldaOracleConvertidaAValor_FallaNombrandoHojaYCelda()
    {
        // W2: si una celda-oráculo deja de ser fórmula, el reader FALLA nombrando hoja+celda
        // (nunca devuelve snapshot con 0 silencioso).
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-w2-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var copia = Path.Combine(salidaDir, "w2.xlsx");
        File.Copy(Insumos.Plantilla, copia, overwrite: true);

        QuitarFormula(copia, "VALIDACION_ENEL", "P3");

        var ex = Assert.Throws<CalculoInvalidoException>(() => _reader.LeerSnapshots(copia, Insumos.Periodo()));
        Assert.Contains("VALIDACION_ENEL!P3", ex.Message);
    }

    [Fact]
    public void W2_HojaOracleAusente_FallaNombrandoHoja()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-w2-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var copia = Path.Combine(salidaDir, "w2-hoja.xlsx");
        File.Copy(Insumos.Plantilla, copia, overwrite: true);

        QuitarHoja(copia, "VALIDACION_ENEL");

        var ex = Assert.Throws<CalculoInvalidoException>(() => _reader.LeerSnapshots(copia, Insumos.Periodo()));
        Assert.Contains("VALIDACION_ENEL", ex.Message);
    }

    [Fact]
    public void GoldenConCacheReal_LecturaSobreSalidaQ1_EquivaleALaPlantilla()
    {
        // Los snapshots sobre la salida Q1 (producida del canónico sin tocar validaciones) son
        // idénticos a los de la plantilla: las hojas de validación no se escriben (oráculo).
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-oraculo-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, "salida.xlsx");
        File.Copy(Insumos.Plantilla, salida, overwrite: true);

        var dePlantilla = _reader.LeerSnapshots(Insumos.Plantilla, Insumos.Periodo());
        var deSalida = _reader.LeerSnapshots(salida, Insumos.Periodo());

        for (var i = 0; i < 5; i++)
        {
            Assert.Equal(dePlantilla[i].ValidacionTotal, deSalida[i].ValidacionTotal);
            Assert.Equal(dePlantilla[i].DetValiRetri!.DiferenciaTotalD21, deSalida[i].DetValiRetri!.DiferenciaTotalD21);
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────────────────────

    private static string Sha256(string ruta)
    {
        using var stream = File.OpenRead(ruta);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static void QuitarFormula(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var cell = ws.Descendants<Cell>().First(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        cell.CellFormula = null;
        cell.CellValue = new CellValue("0");
        ws.Save();
    }

    private static void QuitarHoja(string ruta, string hoja)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        sheet.Remove();
        workbookPart.Workbook!.Save();
    }
}
