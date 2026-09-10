using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-14 (PR2/PR4): cierre de la deuda HU-13 verificado contra el código real (§0.1).
/// W-1 (sub-bloques VALIDACION_TOTAL con assert + gate TRUE exacto, amparo T0-0.4 HU-13),
/// S-1 (booleano estricto: "false" → FALSE), S-2 (gate TOTAL único: 1 error, no 5),
/// S-3 (AseFactory fuente única), S-4 (caché exigido en celdas de gate) y W-2 (pinning del
/// placement en ValidadorBasico, sin mover código — G4).
/// </summary>
public sealed class RobustezDeudaHu13Tests
{
    // ── W-1: sub-bloques booleanos de VALIDACION_TOTAL ────────────────────────────────────────

    [Theory]
    [InlineData("Q1", 1)]
    [InlineData("Q2", 2)]
    public void W1_SubBloquesValidacionTotal_TRUE_ConFormula_EnAmbosGoldens(string etiqueta, int quincena)
    {
        var (ruta, periodo) = quincena == 1
            ? (Insumos.Plantilla, Insumos.Periodo())
            : (Insumos.GoldenQ2, Insumos.PeriodoQ2());

        // Assert protegido: las celdas están en el mapa con presencia de <f> (W2).
        var protegidas = WorkbookLeafCellMapValidaciones.ProtegidasValidacionesParaPeriodo(quincena);
        foreach (var celda in WorkbookLeafCellMapValidaciones.SubBloquesValidacionTotal)
        {
            Assert.Contains(protegidas, p => p.Hoja == WorkbookLeafCellMapValidaciones.HojaValidacionTotal && p.Celda == celda);
            Assert.True(CeldaEsFormula(ruta, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, celda),
                $"{WorkbookLeafCellMapValidaciones.HojaValidacionTotal}!{celda} debió ser fórmula en {etiqueta}.");
        }

        // Gate TRUE exacto contra los goldens reales (read-only) — amparo T0-0.4 HU-13.
        var snapshots = new ValidacionOracleReader().LeerSnapshots(ruta, periodo);
        foreach (var snapshot in snapshots)
        {
            foreach (var celda in WorkbookLeafCellMapValidaciones.SubBloquesValidacionTotal)
            {
                Assert.True(snapshot.SubBloquesValidacionTotal[celda],
                    $"ASE {snapshot.Ase.Id} · VALIDACION_TOTAL · {celda} debió ser TRUE en {etiqueta}.");
            }
        }
    }

    [Fact]
    public void W1_SubBloqueFalso_ErrorNombraAseValidacionYCelda()
    {
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var snapshots = ValidacionesCruzadasTests.CrearSnapshotsValidos();
        snapshots[2].SubBloquesValidacionTotal = new Dictionary<string, bool>
        {
            ["C15"] = true, ["D25"] = true, ["O25"] = true, ["D34"] = false, ["F34"] = true
        };

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        var error = Assert.Single(errores);
        Assert.Contains("ASE 3", error, StringComparison.Ordinal);
        Assert.Contains("VALIDACION_TOTAL", error, StringComparison.Ordinal);
        Assert.Contains("D34", error, StringComparison.Ordinal);
        Assert.Contains("[ERR-VALIDACION]", error, StringComparison.Ordinal);
    }

    // ── S-1: booleano estricto ("false" → FALSE en toda rama) ─────────────────────────────────

    [Theory]
    [InlineData("false")]
    [InlineData("FALSE")]
    public void S1_LiteralFalse_VerificacionP_False_NoTrue(string literal)
    {
        var copia = CopiarPlantilla("s1-false");
        CambiarValorCache(copia, "VALIDACION_ENEL", "P3", literal);

        var snapshots = new ValidacionOracleReader().LeerSnapshots(copia, Insumos.Periodo());

        // Hoy (pre-HU-14) esto daba TRUE: "false" != "0" → TRUE. Debe ser FALSE (S-1).
        Assert.False(snapshots[0].PorEmpresa.Single(e => e.Empresa == "Enel").VerificacionP);
    }

    [Fact]
    public void S1_BooleanoTipoB_LiteralFalse_EsFalse()
    {
        // Con t="b" y texto "false": S-1 exige FALSE "en todas las ramas" (nunca TRUE).
        var copia = CopiarPlantilla("s1-tb");
        CambiarValorCache(copia, "VALIDACION_ENEL", "P3", "false", tipoBooleano: true);

        var snapshots = new ValidacionOracleReader().LeerSnapshots(copia, Insumos.Periodo());

        Assert.False(snapshots[0].PorEmpresa.Single(e => e.Empresa == "Enel").VerificacionP);
    }

    [Fact]
    public void S1_BooleanoTipoB_Basura_FallaNombrandoCelda()
    {
        var copia = CopiarPlantilla("s1-tb-basura");
        CambiarValorCache(copia, "VALIDACION_ENEL", "P3", "basura", tipoBooleano: true);

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            new ValidacionOracleReader().LeerSnapshots(copia, Insumos.Periodo()));
        Assert.Contains("VALIDACION_ENEL!P3", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void S1_ValorBasura_FallaNombrandoHojaYCelda()
    {
        var copia = CopiarPlantilla("s1-basura");
        CambiarValorCache(copia, "VALIDACION_ENEL", "P3", "basura");

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            new ValidacionOracleReader().LeerSnapshots(copia, Insumos.Periodo()));
        Assert.Contains("VALIDACION_ENEL!P3", ex.Message, StringComparison.Ordinal);
    }

    // ── S-2: gate VALIDACION_TOTAL único (1 error, no 5) ─────────────────────────────────────

    [Fact]
    public void S2_O9Roto_ExactamenteUnError_NoCinco()
    {
        var copia = CopiarPlantilla("s2-o9");
        CambiarValorCache(copia, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "O9", "2");

        var snapshots = new ValidacionOracleReader().LeerSnapshots(copia, Insumos.Periodo());
        Assert.All(snapshots, s => Assert.Equal(2m, s.ValidacionTotal)); // el reader pobló los 5 con una sola lectura

        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        var error = Assert.Single(errores);
        Assert.Contains("VALIDACION_TOTAL", error, StringComparison.Ordinal);
        Assert.Contains("O9", error, StringComparison.Ordinal);
    }

    [Fact]
    public void S2_DivergenciaEntreSnapshots_UnErrorDeIgualdad()
    {
        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var snapshots = ValidacionesCruzadasTests.CrearSnapshotsValidos();
        snapshots[3].ValidacionTotal = 1.5m; // diverge del resto (bug del reader simulado)

        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        var error = Assert.Single(errores);
        Assert.Contains("VALIDACION_TOTAL", error, StringComparison.Ordinal);
        Assert.Contains("diverge", error, StringComparison.OrdinalIgnoreCase);
    }

    // ── S-4: caché <v> exigido en celdas de gate; informativos conservan el 0 ─────────────────

    [Fact]
    public void S4_CeldaDeGateSinCache_FallaNombrandoHojaYCelda_NoCeroSilencioso()
    {
        var copia = CopiarPlantilla("s4-gate");
        QuitarValorCache(copia, "VALIDACION_ENEL", "O3"); // fórmula sin <v> (workbook nunca recalculado)

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            new ValidacionOracleReader().LeerSnapshots(copia, Insumos.Periodo()));
        Assert.Contains("VALIDACION_ENEL!O3", ex.Message, StringComparison.Ordinal);
        Assert.Contains("recalcule", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void S4_ControlInformativoSinCache_ConservaElCeroTolerado()
    {
        var copia = CopiarPlantilla("s4-info");
        QuitarValorCache(copia, "Valida -Remunera", "D9");

        var snapshots = new ValidacionOracleReader().LeerSnapshots(copia, Insumos.Periodo());

        Assert.Equal(0m, snapshots[0].Controles["Valida -Remunera!D9"]); // informativo: 0 tolerado
        Assert.Equal(5, snapshots.Count); // la lectura completa no falla por controles
    }

    // ── S-3: AseFactory fuente única ──────────────────────────────────────────────────────────

    [Theory]
    [InlineData(1, "1-Promoambiental")]
    [InlineData(2, "2-Lime")]
    [InlineData(3, "3-Ciudad Limpia")]
    [InlineData(4, "4-Bogotá Limpia")]
    [InlineData(5, "5-Área Limpia")]
    public void S3_AseFactory_DesdeId_CoincideConCarpetasAse(int id, string prefijoEsperado)
    {
        var ase = AseFactory.DesdeId(id);
        Assert.Equal(id, ase.Id);
        Assert.Equal(id, ase.NumeroCarpeta);
        var guion = prefijoEsperado.IndexOf('-');
        var nombreEsperado = prefijoEsperado[(guion + 1)..];
        Assert.Equal(nombreEsperado, ase.NombreCompleto);
        Assert.Equal(nombreEsperado.ToUpperInvariant(), ase.NombreCorto);
        Assert.Equal(CarpetasAse.Prefijos[id - 1], prefijoEsperado);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(6)]
    [InlineData(-1)]
    public void S3_AseFactory_IdFueraDeRango_FallaFast(int id)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => AseFactory.DesdeId(id));
    }

    [Fact]
    public void S3_ReaderOracle_UsaLaMismaFactory_SnapshotsCanonicos()
    {
        var snapshots = new ValidacionOracleReader().LeerSnapshots(Insumos.Plantilla, Insumos.Periodo());
        foreach (var snapshot in snapshots)
        {
            var esperado = AseFactory.DesdeId(snapshot.Ase.Id);
            Assert.Equal(esperado.NombreCompleto, snapshot.Ase.NombreCompleto);
            Assert.Equal(esperado.NombreCorto, snapshot.Ase.NombreCorto);
            Assert.Equal(esperado.NumeroCarpeta, snapshot.Ase.NumeroCarpeta);
        }
    }

    // ── W-2: pinning del placement (sin mover código, G4) ─────────────────────────────────────

    [Fact]
    public void W2_GateCruzadas_ViveEnValidadorBasico_NoEnWorkbookLeafCoherence()
    {
        // El gate de validaciones cruzadas vive en ValidadorBasico (agregación por lista),
        // NO en WorkbookLeafCoherence (throw al primer fallo): moverlo cambiaría la semántica
        // de agregación de TODOS los gates HU-08..HU-12 (veredicto W-2, Plan 14 §9.7).
        var gate = typeof(ValidadorBasico).GetMethod(
            "ValidarGatesValidacionesCruzadasPorAse",
            BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(gate);

        // WorkbookLeafCoherence es internal (Infrastructure); se pinnea por nombre de tipo y
        // miembros vía reflexión: ningún miembro de validaciones cruzadas debe vivir allí.
        var tipoCoherence = typeof(ValidacionOracleReader).Assembly
            .GetType("Remuneracion.Infrastructure.Excel.WorkbookLeafCoherence");
        Assert.NotNull(tipoCoherence);
        var miembrosCoherence = tipoCoherence.GetMembers(
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance);
        Assert.DoesNotContain(miembrosCoherence, m =>
            m.Name.Contains("Cruzada", StringComparison.OrdinalIgnoreCase)
            || m.Name.Contains("ValidacionCruzada", StringComparison.OrdinalIgnoreCase));
    }

    // ── Helpers OpenXML (copias en temp; nunca tocan los goldens) ─────────────────────────────

    private static string CopiarPlantilla(string sufijo)
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), $"remuneracion-hu14-{sufijo}-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var copia = Path.Combine(salidaDir, "copia.xlsx");
        File.Copy(Insumos.Plantilla, copia, overwrite: true);
        return copia;
    }

    private static void CambiarValorCache(string ruta, string hoja, string celda, string valor, bool tipoBooleano = false)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var cell = ws.Descendants<Cell>().First(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        cell.CellValue = new CellValue(valor);
        if (tipoBooleano)
        {
            cell.DataType = CellValues.Boolean;
        }

        ws.Save();
    }

    private static void QuitarValorCache(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var cell = ws.Descendants<Cell>().First(c =>
            string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        cell.CellValue = null;
        ws.Save();
    }

    private static bool CeldaEsFormula(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var ws = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        return ws.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?.CellFormula is not null;
    }
}
