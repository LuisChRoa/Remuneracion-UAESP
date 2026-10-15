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
/// Plan 40 (T1, R-E-1/R-E-2/R-E-3/R-E-4 — D-A/D-C/D-D): el gate W-1 EVALÚA los sub-bloques
/// booleanos de <c>VALIDACION_TOTAL</c> desde las celdas efectivamente escritas (nunca desde el
/// caché <c>&lt;v&gt;</c>, que era la falsa verificación heredada del incidente 2026-10-14).
///
/// Casos obligatorios:
///   (i)  un FALSE GENUINO (divergencia real inyectada) hace FALLAR el gate nombrando ASE·celda
///        (el test que habría pasado con el lector viejo por herencia: su caché dice TRUE);
///   (ii) un TRUE evaluado desde lo escrito PASA aunque el caché heredado diga FALSE;
///   (iii) una tercera forma de texto <c>&lt;f&gt;</c> → ERR-PLANTILLA (hoja+celda+esperado+observado).
/// </summary>
public sealed class ValidacionTotalEvaluaT1Tests
{
    private const string Hoja = "VALIDACION_TOTAL";

    // ── (ii) TRUE evaluado desde lo escrito (aunque el caché heredado diga FALSE) ────────────────

    [Fact]
    public void T1_C15_TrueEvaluadoAunqueCacheHeredadoFalse_EnSalidaReal()
    {
        // Resultado5 (salida real de hoy): el caché heredado de C15 es 0 (FALSE) — el lector viejo
        // fallaba en falso. El gate nuevo EVALÚA C15 = (C14 = C9) desde lo escrito (F10 y el
        // recaudo) → TRUE.
        var ruta = Insumos.ResultadoJulioQ2(5);
        var snapshots = new ValidacionOracleReader().LeerSnapshots(ruta, Insumos.PeriodoQ2());

        Assert.Equal("0", ValorCache(ruta, Hoja, "C15")); // el caché heredado miente (FALSE)

        var evidencia = snapshots[0].SubBloquesValidacionTotalEvidencia["C15"];
        Assert.True(evidencia.Computable, "C15 debe ser evaluable desde lo escrito.");
        Assert.True(FormasValidacionTotal.Igualdad(evidencia.OperandoIzquierdo, evidencia.OperandoDerecho),
            $"C15 evaluado: {evidencia.OperandoIzquierdo} vs {evidencia.OperandoDerecho} (debe cerrar ±0.5).");
        Assert.All(snapshots, s => Assert.True(s.SubBloquesValidacionTotal["C15"]));
    }

    [Fact]
    public void T1_SubBloques_EvaluadosHonestamente_EnSalidaReal()
    {
        // Plan 40 (enmienda E-1): los 5 sub-bloques se EVALÚAN desde lo escrito. En la salida real
        // (Resultado5 == corrida fresca) C15/D25 (recaudo) cierran TRUE; O25/D34/F34 (espejo de
        // remuneración) DIVERGEN → FALSE (defecto real E1-2, no "no evaluable" ni verde heredado).
        var snapshots = new ValidacionOracleReader().LeerSnapshots(Insumos.ResultadoJulioQ2(5), Insumos.PeriodoQ2());

        Assert.All(snapshots, s => Assert.True(s.SubBloquesValidacionTotal["C15"]));
        Assert.All(snapshots, s => Assert.True(s.SubBloquesValidacionTotal["D25"]));
        Assert.All(snapshots, s => Assert.False(s.SubBloquesValidacionTotal["O25"]));
        Assert.All(snapshots, s => Assert.False(s.SubBloquesValidacionTotal["D34"]));
        Assert.All(snapshots, s => Assert.False(s.SubBloquesValidacionTotal["F34"]));

        // Los 5 son EVALUABLES (no "no evaluable"): la cadena del espejo se resuelve por read-back.
        foreach (var (celda, _) in FormasValidacionTotal.SubBloques)
        {
            Assert.All(snapshots, s => Assert.True(s.SubBloquesValidacionTotalEvidencia[celda].Computable,
                $"{celda} debe ser resoluble desde lo escrito (E1-1)."));
        }
    }

    [Fact]
    public void T1_ManualRecalculado_LosCincoCierranTrue()
    {
        // El manual del administrativo (recalculado por Excel) SÍ cierra los 5: es el oráculo de
        // que el FALSE de la salida de la app es divergencia real y no un artefacto del evaluador.
        var snapshots = new ValidacionOracleReader().LeerSnapshots(Insumos.GoldenQ2, Insumos.PeriodoQ2());
        foreach (var (celda, _) in FormasValidacionTotal.SubBloques)
        {
            Assert.All(snapshots, s => Assert.True(s.SubBloquesValidacionTotal[celda],
                $"VALIDACION_TOTAL · {celda} debe cerrar TRUE en el manual recalculado."));
        }
    }

    // ── (i) FALSE GENUINO: divergencia real inyectada → gate FAIL nombrando ASE·celda ────────────

    [Fact]
    public void T1_FalsoGenuino_DivergenciaInyectada_FallaNombrandoC15()
    {
        using var temp = new DirectorioTemporal("t1-falso-genuino-");
        var copia = Path.Combine(temp.Ruta, "copia.xlsx");
        File.Copy(Insumos.ResultadoJulioQ2(5), copia);

        // Divergencia real: F10 (que alimenta C14) deja de cerrar contra C9 = Σ recaudo.
        CambiarValor(copia, "Valida - Control Recaudo", "F10", "1");
        // Y el caché heredado de C15 queda en TRUE (1) — el lector viejo habría PASADO (falso-verde).
        CambiarValor(copia, Hoja, "C15", "1", booleano: true);

        var snapshots = new ValidacionOracleReader().LeerSnapshots(copia, Insumos.PeriodoQ2());

        var evidencia = snapshots[0].SubBloquesValidacionTotalEvidencia["C15"];
        Assert.True(evidencia.Computable);
        Assert.False(FormasValidacionTotal.Igualdad(evidencia.OperandoIzquierdo, evidencia.OperandoDerecho),
            "La divergencia inyectada debe romper C14 = C9.");
        Assert.All(snapshots, s => Assert.False(s.SubBloquesValidacionTotal["C15"]));

        var (resultado, leafs) = ValidacionesCruzadasTests.CrearCasoValido();
        var errores = new ValidadorBasico().Validar(resultado, leafs, snapshots);

        Assert.Contains(errores, e =>
            e.Contains("[ERR-VALIDACION]", StringComparison.Ordinal)
            && e.Contains("ASE 1", StringComparison.Ordinal)
            && e.Contains("VALIDACION_TOTAL · C15", StringComparison.Ordinal));
    }

    // ── (iii) tercera forma de <f> → ERR-PLANTILLA ───────────────────────────────────────────────

    [Fact]
    public void T1_TerceraForma_ErrPlantilla_NombraEsperadoYObservado()
    {
        using var temp = new DirectorioTemporal("t1-tercera-forma-");
        var copia = Path.Combine(temp.Ruta, "copia.xlsx");
        File.Copy(Insumos.ResultadoJulioQ2(5), copia);

        CambiarFormula(copia, Hoja, "C15", "C14<>C9");

        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            new ValidacionOracleReader().LeerSnapshots(copia, Insumos.PeriodoQ2()));

        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Contains("VALIDACION_TOTAL!C15", ex.Message, StringComparison.Ordinal);
        Assert.Contains("C14=C9", ex.Message, StringComparison.Ordinal);   // esperado
        Assert.Contains("C14<>C9", ex.Message, StringComparison.Ordinal); // observado
    }

    // ── Textos de los 5 sub-bloques == tabla (T0) ────────────────────────────────────────────────

    [Fact]
    public void T1_TextosFormula_CoincidenTabla_EnSalidaRealYManual()
    {
        foreach (var (ruta, periodo) in new[]
        {
            (Insumos.ResultadoJulioQ2(5), Insumos.PeriodoQ2()),
            (Insumos.GoldenQ2, Insumos.PeriodoQ2()),
            (Insumos.Plantilla, Insumos.Periodo()),
        })
        {
            var snapshot = new ValidacionOracleReader().LeerSnapshots(ruta, periodo)[0];
            foreach (var (celda, textoEsperado) in FormasValidacionTotal.SubBloques)
            {
                Assert.Equal(textoEsperado, snapshot.SubBloquesValidacionTotalEvidencia[celda].TextoFormula);
            }
        }
    }

    // ── R-E-3: tolerancia ±0.5 en el borde ───────────────────────────────────────────────────────

    [Theory]
    [InlineData(0.4, true)]
    [InlineData(0.5, true)]
    [InlineData(0.6, false)]
    [InlineData(-0.4, true)]
    [InlineData(-0.6, false)]
    public void T1_Tolerancia_Borde(decimal delta, bool esperado)
    {
        Assert.Equal(esperado, FormasValidacionTotal.Igualdad(1000m, 1000m + delta));
    }

    // ── Helpers ─────────────────────────────────────────────────────────────────────────────────

    private static string? ValorCache(string ruta, string hoja, string celda)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, false);
        return Celda(workbook, hoja, celda).CellValue?.InnerText;
    }

    private static void CambiarValor(string ruta, string hoja, string celda, string valor, bool booleano = false)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var (c, worksheet) = CeldaYHoja(workbook, hoja, celda);
        c.CellValue = new CellValue(valor);
        c.DataType = booleano ? CellValues.Boolean : CellValues.Number;
        worksheet.Save();
    }

    private static void CambiarFormula(string ruta, string hoja, string celda, string formula)
    {
        using var workbook = SpreadsheetDocument.Open(ruta, true);
        var (c, worksheet) = CeldaYHoja(workbook, hoja, celda);
        c.CellValue = null;
        c.DataType = null;
        c.CellFormula = new CellFormula(formula);
        worksheet.Save();
    }

    private static Cell Celda(SpreadsheetDocument workbook, string hoja, string celda) => CeldaYHoja(workbook, hoja, celda).Cell;

    private static (Cell Cell, Worksheet Worksheet) CeldaYHoja(SpreadsheetDocument workbook, string hoja, string celda)
    {
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var sheet = workbookPart.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var c = worksheet.Descendants<Cell>().First(x =>
            string.Equals(x.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        return (c, worksheet);
    }

    private sealed class DirectorioTemporal : IDisposable
    {
        public DirectorioTemporal(string prefijo) =>
            Ruta = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), prefijo + Guid.NewGuid().ToString("N"))).FullName;

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
