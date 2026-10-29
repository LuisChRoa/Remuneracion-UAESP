using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 62 (H5 — T5.1, D-J9): el comparador <see cref="ComparadorSalidaVsManual"/> deja de excluir
/// <c>DetValiRetri!J</c> (AJUSTE A LA DECENA) y las divergencias que SÍ son legítimas frente al
/// manual del administrativo viven en el catálogo versionado
/// <see cref="DivergenciasDeclaradasComparador"/> — con Id, cita, fecha y condición de retiro.
///
/// <para><b>Evidencia medida (H5, sobre <c>Docs/Prueba Septiembre 1</c>):</b> corrida 5-ASE de
/// 2026091 (fixture compartido <c>EspecialesR4H2Plan61Tests.Generacion2026091</c>) comparada contra
/// el manual <c>Resultados/Administrativo/Remuneracion 202609-1 Total.xlsx</c>. En la SUPERFICIE DEL
/// PEGADO (<c>DetRetri</c> A3:O19, <c>DetValiRetri</c> A3:O14) quedan exactamente tres divergencias,
/// las tres declaradas: el pie de interventoría de <c>C16</c> y el peso de la errata del ASE 3 en
/// <c>D11</c>/<c>D14</c>. Cero inesperadas.</para>
///
/// <para><b>Alcance (honesto):</b> las fórmulas de las filas 20+ de esas hojas son geometría del
/// manual heredado (mesh R1/R2/R4 por anclas distintas), NO superficie del pegado: pertenecen a la
/// clase pre-existente R-TRAMPA-ANCLA y siguen en rojo en los goldens del comparador con manual; este
/// fixture NO las declara (declararlas taparía una regresión de geometría).</para>
/// </summary>
public sealed class ComparadorPegadoH5Plan62Tests
{
    private const string CodigoPeriodo = "2026091";
    private const string HojaDetRetri = "DetRetri2026091";
    private const string HojaDetValiRetri = "DetValiRetri2026091";

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T5.1 — Formato del catálogo (R-FALSO-POSITIVO: sin los 6 datos no hay entrada)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void T5_1_Catalogo_TodaEntradaTraeIdCitaFechaYCondicionDeRetiro()
    {
        Assert.NotEmpty(DivergenciasDeclaradasComparador.Entradas);

        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var entrada in DivergenciasDeclaradasComparador.Entradas)
        {
            Assert.True(entrada.Completa, $"Entrada incompleta en el catálogo: '{entrada.Id}'.");
            Assert.StartsWith(DivergenciasDeclaradasComparador.PrefijoId, entrada.Id, StringComparison.Ordinal);
            Assert.True(ids.Add(entrada.Id), $"Id duplicado en el catálogo: '{entrada.Id}'.");
            Assert.Contains("Plan 62", entrada.Cita, StringComparison.Ordinal);
            Assert.True(DateOnly.TryParseExact(entrada.Fecha, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out _), $"Fecha no ISO en '{entrada.Id}': '{entrada.Fecha}'.");
            Assert.False(string.IsNullOrWhiteSpace(entrada.CondicionRetiro));
        }
    }

    [Fact]
    public void T5_1_Catalogo_SoloDeclaraCeldasDeLasDosHojasPegadas()
    {
        foreach (var entrada in DivergenciasDeclaradasComparador.Entradas)
        {
            Assert.True(
                entrada.Hoja is DivergenciasDeclaradasComparador.HojaDetRetri or DivergenciasDeclaradasComparador.HojaDetValiRetri,
                $"'{entrada.Id}' declara una hoja fuera del pegado: '{entrada.Hoja}'.");

            foreach (var celda in entrada.Celdas)
            {
                Assert.Matches("^[A-Z]{1,2}[0-9]{1,3}$", celda);
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T5.1 — La exclusión SALE (D-J9) y la declaración es EXCLUSIVA (no tapa una regresión)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void T5_1_DetValiRetriJ_YaNoEsExclusionVersionada(int quincena)
    {
        var comparador = new ComparadorSalidaVsManual(quincena);

        for (var fila = 9; fila <= 14; fila++)
        {
            var celda = $"J{fila}";
            Assert.False(
                comparador.EsExclusionVersionada($"DetValiRetri202609{quincena}", celda),
                $"Plan 62 (D-J9): DetValiRetri!{celda} ya NO se excluye del comparador.");
        }

        // Anti-vacuidad: las exclusiones que SÍ quedan siguen vigentes.
        Assert.True(comparador.EsExclusionVersionada("CONSOLIDADO_TOTAL RECAUDO", "N3"));
        Assert.True(comparador.EsExclusionVersionada("Recaudo ENEL", quincena == 2 ? "D21" : "F21"));
    }

    [Fact]
    public void T5_1_Declaracion_EsExclusivaDelParMedido()
    {
        const string c16App = "* Costos de Interventoría: REMUNERACIÓN 202609-1";
        const string c16Manual = "* Costos de Interventoría: REMUNERACIÓN 202604-1";

        Assert.True(ComparadorSalidaVsManual.EsDivergenciaDeclarada(HojaDetRetri, "C16", c16App, c16Manual));

        // Otro texto en C16 NO está declarado (la entrada cita el pie de interventoría, no la celda).
        Assert.False(ComparadorSalidaVsManual.EsDivergenciaDeclarada(
            HojaDetRetri, "C16", "* Costos de Interventoría: REMUNERACIÓN 202609-1", "Firma autorizada"));

        // Errata: Δ = 1 peso exacto y entero. Δ = 2 o Δ fraccionario NO están declarados.
        Assert.True(ComparadorSalidaVsManual.EsDivergenciaDeclarada(HojaDetRetri, "D11", "10891338739", "10891338740"));
        Assert.False(ComparadorSalidaVsManual.EsDivergenciaDeclarada(HojaDetRetri, "D11", "10891338739", "10891338741"));
        Assert.False(ComparadorSalidaVsManual.EsDivergenciaDeclarada(HojaDetRetri, "D11", "10891338739", "10891338739,5"));

        // La declaración es de CELDA: el mismo par en otra celda no se declara.
        Assert.False(ComparadorSalidaVsManual.EsDivergenciaDeclarada(HojaDetRetri, "D12", c16App, c16Manual));
        Assert.False(ComparadorSalidaVsManual.EsDivergenciaDeclarada(HojaDetValiRetri, "D11", "10891338739", "10891338740"));
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // T5.1 — Septiembre-1: la superficie del pegado solo trae las divergencias declaradas
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void T5_1_Septiembre1_LaSuperficieDelPegadoSoloTraeLasDeclaradas()
    {
        var (resultado, _) = EspecialesR4H2Plan61Tests.Generacion2026091.Value;
        var manual = RutaManualAdministrativo();

        var comparador = new ComparadorSalidaVsManual(1);
        var inesperadas = comparador.Comparar(resultado.RutaSalida, manual)
            .Where(d => EnSuperficieDelPegado(d.Hoja, d.Celda))
            .ToList();

        Assert.True(
            inesperadas.Count == 0,
            "Divergencias INESPERADAS en la superficie del pegado (DetRetri A3:O19 / DetValiRetri A3:O14):\n"
            + string.Join("\n", inesperadas.Select(d => $"{d.Hoja}!{d.Celda} app=[{d.ValorApp}] manual=[{d.ValorManual}] ({d.Motivo})")));

        // Anti-vacuidad: las tres declaradas SE APLICAN (una entrada que no aplica es una declaración
        // muerta: la divergencia que documenta ya no existe y la cita quedó obsoleta).
        var conDeclaradas = comparador.Comparar(resultado.RutaSalida, manual, incluirDivergenciasDeclaradas: true)
            .Where(d => EnSuperficieDelPegado(d.Hoja, d.Celda))
            .ToList();

        var declaradasAplicadas = conDeclaradas
            .Where(d => ComparadorSalidaVsManual.EsDivergenciaDeclarada(d.Hoja, d.Celda, d.ValorApp, d.ValorManual))
            .Select(d => DivergenciasDeclaradasComparador.Buscar(d.Hoja, d.Celda, d.ValorApp, d.ValorManual)!.Id)
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToList();

        Assert.Equal(
            new[]
            {
                "P62-DETRETRI-C16-PIE-INTERVENTORIA",
                "P62-DETRETRI-D11-ERRATA-PESO-ASE3",
                "P62-DETRETRI-D14-TOTAL-ERRATA-PESO"
            },
            declaradasAplicadas);
    }

    /// <summary>
    /// D-J9 (R-J9-FALSO-VERDE): el comparador COMPARA <c>DetValiRetri!J</c>. Se injerta un valor
    /// distinto en una COPIA del manual: el comparador debe reportarlo (si la exclusión hubiera
    /// vuelto, este fixture seguiría verde y el AJUSTE A LA DECENA quedaría sin malla).
    /// </summary>
    [Fact]
    public void T5_1_DetValiRetriJ_InjertoEnElManual_ElComparadorLoReporta()
    {
        var (resultado, _) = EspecialesR4H2Plan61Tests.Generacion2026091.Value;
        var carpeta = Path.Combine(Path.GetTempPath(), "remuneracion-p62-h5-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(carpeta);
        var manualMutado = Path.Combine(carpeta, "manual-j-mutado.xlsx");

        try
        {
            File.Copy(RutaManualAdministrativo(), manualMutado);
            MutarCelda(manualMutado, HojaDetValiRetri, "J9", "-46494");

            var divergencias = new ComparadorSalidaVsManual(1).Comparar(resultado.RutaSalida, manualMutado);
            var j9 = divergencias.SingleOrDefault(d =>
                ComparadorSalidaVsManual.NormalizarHoja(d.Hoja) == DivergenciasDeclaradasComparador.HojaDetValiRetri
                && d.Celda == "J9");

            Assert.NotNull(j9);
            Assert.Equal("literal numérico fuera de tolerancia (±0.5)", j9!.Motivo);
            Assert.False(ComparadorSalidaVsManual.EsDivergenciaDeclarada(j9.Hoja, j9.Celda, j9.ValorApp, j9.ValorManual));
        }
        finally
        {
            try
            {
                Directory.Delete(carpeta, recursive: true);
            }
            catch (IOException)
            {
                // Limpieza best-effort (temp).
            }
        }
    }

    // ── Helpers ─────────────────────────────────────────────────────────────────────────────────

    private static bool EnSuperficieDelPegado(string hoja, string celda)
    {
        var fila = int.Parse(new string(celda.SkipWhile(char.IsLetter).ToArray()), CultureInfo.InvariantCulture);
        return ComparadorSalidaVsManual.NormalizarHoja(hoja) switch
        {
            DivergenciasDeclaradasComparador.HojaDetRetri => fila is >= 3 and <= 19,
            DivergenciasDeclaradasComparador.HojaDetValiRetri => fila is >= 3 and <= 14,
            _ => false
        };
    }

    private static string RutaManualAdministrativo() => Path.Combine(
        RaizRepositorio(), "Docs", "Prueba Septiembre 1", "Resultados", "Administrativo",
        "Remuneracion 202609-1 Total.xlsx");

    private static void MutarCelda(string ruta, string hoja, string celda, string valor)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var workbookPart = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart nulo.");
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook nulo.");
        var nombre = workbook.Descendants<Sheet>()
            .Single(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(nombre.Id!.Value!);
        var worksheet = worksheetPart.Worksheet
            ?? throw new InvalidOperationException($"La hoja '{hoja}' no tiene Worksheet.");
        var destino = worksheet.Descendants<Cell>()
            .Single(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));

        destino.CellFormula = null;
        destino.DataType = CellValues.Number;
        destino.CellValue = new CellValue(valor);
        worksheet.Save();
    }

    private static string RaizRepositorio()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !Directory.Exists(Path.Combine(dir.FullName, "plans")))
        {
            dir = dir.Parent;
        }

        return dir?.FullName ?? throw new InvalidOperationException("Raíz del repositorio no encontrada.");
    }
}
