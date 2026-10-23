using System.Globalization;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 54 (H1): evidencia de implementación de la ruta ÚNICA de lectura.
/// - S2.3 (D1.1 enmendada): `Mes2`/`Lmes2` opcionales → 0 (julio-1 ASE1 sin aborto y SIN cambio de
///   valor); ASE3-octubre con 0 filas `Aplicacion` resuelve EXTEMP 0; `F30/F10/L10` constantes `0m`.
/// - S2.5 (D6): la divergencia estructural ASE5-F674 está declarada en el catálogo Plan 42 y el slot
///   queda en el 0 del template.
/// - S2.6 (H1.7): `R1Q2ProtectedPorAse` declara las 10 `<f>` TOT_OPT/EXTEMP (incluida `F682`).
///
/// Rutas explícitas (sin <c>Insumos.Raiz()</c>), igual que el fixture T0 del mismo plan.
/// </summary>
public sealed class ContratoUnicoLecturaR1H1Tests
{
    private const decimal Tolerancia = 0.5m;

    private static string PlantillaV2 =>
        Ruta("Docs", "Plantilla estandarizada", "v2", "Formato Remuneracion 2026xx-x Total vf.xlsx");

    // ════════════════════════════════════════════════════════════════════════════════════════
    // S2.3 — Mes2 opcional (enmienda 2) sin cambio de valor
    // ════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Julio-1 ASE1 trae 2 filas <c>Mes/Total</c>; el mapa declara 3 slots (Mes2 → F70). Con la
    /// enmienda 2 la lectura RESUELVE (antes abortaba) y el valor no cambia:
    /// <c>TOT_OPT = Σ F(Mes) − Σ Esp(Mes sin la última) = 16.704.332.434,57</c> y
    /// <c>EXTEMP = Σ F(Aplic) − Esp(Aplic[0]) = 11.673.020,00</c>; D104 = 16.758.167.584,60 ⇒
    /// ROUND = 16.758.167.585 = <c>DetRetri2026071!D9</c> (R10 oficial, evidencia T0.3/T0.6).
    /// </summary>
    [Fact]
    public void S2_3_Ase1Julio1_DosFilasMes_ResuelveConMes2EnCero()
    {
        var (r1, r2, r4) = Rutas("Julio-1", 1);
        var leaf = new ExcelDataReaderWorkbookLeafInputReader()
            .LeerLeafInputs(AseFactory.DesdeId(1), Periodo.Parse("2026071"), r1, r2, r4);

        Assert.Equal(0m, leaf.R1.CeldasPorAse["F70"]); // Mes2 sin fila de origen → 0 del template
        Assert.InRange(leaf.R1.TotalOportunoEsperadoPorAse - 16704332434.57m, -Tolerancia, Tolerancia);
        Assert.InRange(leaf.R1.ExtemporaneoEsperadoPorAse - 11673020.00m, -Tolerancia, Tolerancia);
    }

    /// <summary>
    /// ASE3-octubre trae CERO filas <c>Aplicacion</c> (el caso rojo del plan): la lectura resuelve con
    /// EXTEMP 0 y <c>TOT_OPT = 8.811.472.160,34</c> (el slot <c>L330 = 19.305.717,66</c> sí se resta).
    /// </summary>
    [Fact]
    public void S2_3_Ase3Octubre_CeroAplic_ResuelveConExtempCero()
    {
        var (r1, r2, r4) = Rutas("Octubre-1", 3);
        var leaf = new ExcelDataReaderWorkbookLeafInputReader()
            .LeerLeafInputs(AseFactory.DesdeId(3), Periodo.Parse("2026101"), r1, r2, r4);

        Assert.Equal(0m, leaf.R1.ExtemporaneoEsperadoPorAse);
        Assert.Equal(0m, leaf.R1.CeldasPorAse["F315"]);
        Assert.Equal(0m, leaf.R1.CeldasPorAse["F342"]);
        Assert.Equal(0m, leaf.R1.CeldasPorAse["L315"]);
        Assert.InRange(leaf.R1.TotalOportunoEsperadoPorAse - 8811472160.34m, -Tolerancia, Tolerancia);
    }

    /// <summary>
    /// D2: <c>F30/F10/L10</c> son constantes <c>0m</c> en AMBAS quincenas (el término v1
    /// <c>subsidio[0] + Aplic[0]</c> desaparece). Misma fuente, dos <see cref="Periodo"/>.
    /// </summary>
    [Fact]
    public void S2_3_F30F10L10_SonCeroConstanteEnAmbasQuincenas()
    {
        var (r1, r2, r4) = Rutas("Julio-1", 1);
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var ase = AseFactory.DesdeId(1);

        var q1 = reader.LeerLeafInputs(ase, Periodo.Parse("2026071"), r1, r2, r4);
        var q2 = reader.LeerLeafInputs(ase, Periodo.Parse("2026072"), r1, r2, r4);

        foreach (var leaf in new[] { q1, q2 })
        {
            Assert.Equal(0m, leaf.R1.F30);
            Assert.Equal(0m, leaf.R1.F10);
            Assert.Equal(0m, leaf.R1.L10);
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // S2.6 / H1.7 — simetría de las 10 <f> TOT_OPT/EXTEMP (incluida F682 de ASE5)
    // ════════════════════════════════════════════════════════════════════════════════════════

    public static TheoryData<int, string, string[]> FormulasProtegidas => new()
    {
        { 1, "F75", ["F47", "F70", "F16", "L16", "L47"] },
        { 1, "F77", ["F59", "F32", "L32"] },
        { 2, "F254", ["F171", "F208", "L171", "F120", "L120"] },
        { 2, "F256", ["F190", "F146", "L146"] },
        { 3, "F415", ["F330", "F353", "F299", "L299", "L330"] },
        { 3, "F417", ["F342", "F315", "L315"] },
        { 4, "F568", ["F511", "F548", "F460", "L460", "L511"] },
        { 4, "F570", ["F530", "F486", "L486"] },
        { 5, "F680", ["F674", "F648", "F613", "L613", "L648"] },
        { 5, "F682", ["F660", "F629", "L629"] }
    };

    [Theory]
    [MemberData(nameof(FormulasProtegidas))]
    public void S2_6_ProtegidasR1_DeclaranLaFormulaPorBytes(int aseId, string celda, string[] fragmentos)
    {
        var declarada = WorkbookLeafCellMapQ2.ObtenerR1Q2Protegidos(aseId)
            .Single(p => string.Equals(p.Celda, celda, StringComparison.OrdinalIgnoreCase));

        Assert.Equal(fragmentos, declarada.Fragmentos);
    }

    /// <summary>Simetría: exactamente 2 visibles protegidos (TOT_OPT + EXTEMP) por ASE = 10.</summary>
    [Fact]
    public void S2_6_HayDiezFormulasProtegidas_DosPorAse()
    {
        var total = Enumerable.Range(1, 5)
            .SelectMany(aseId => WorkbookLeafCellMapQ2.ObtenerR1Q2Protegidos(aseId))
            .Count();

        Assert.Equal(10, total);
        Assert.Contains(WorkbookLeafCellMapQ2.ObtenerR1Q2Protegidos(5), p => p.Celda == "F682");
    }

    // ════════════════════════════════════════════════════════════════════════════════════════
    // S2.5 / H1.5 — ASE5-F674: divergencia ESTRUCTURAL declarada + slot en el 0 del template
    // ════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void S2_5_Ase5F674_EstaDeclaradaEnElCatalogoPlan42()
    {
        var entrada = DivergenciasDeclaradasCatalogo.Entradas
            .Single(e => e.Id == "ASE5-F674-MES2-AUSENTE");

        Assert.Equal(new[] { "F674", "F680" }, entrada.Celdas);
        Assert.Equal(0m, entrada.Delta); // estructural: no hay diferencia material
        Assert.False(string.IsNullOrWhiteSpace(entrada.OrigenMonto));
        Assert.Contains("2 filas", entrada.OrigenMonto, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("UAESP", entrada.Dueno, StringComparison.Ordinal);
        Assert.Contains("Docs/T0-Plan54", entrada.CitaEvidencia, StringComparison.Ordinal);
        Assert.Equal("2026-10-22", entrada.Fecha);
        Assert.False(string.IsNullOrWhiteSpace(entrada.CondicionRetiro));
        Assert.False(entrada.Retirada);

        // Δ = 0 ⇒ el matcher NO la reconoce como discrepancia material (no hay verde silencioso ni
        // divergencia fantasma); la fila documenta el estado, no aplica.
        Assert.False(entrada.Vigente);
    }

    /// <summary>
    /// La divergencia se PINCHA: la fuente de ASE5 trae 2 filas <c>Mes</c>, el slot <c>F674</c> queda
    /// en 0 y el TOT_OPT del dominio = <c>7.544.253.279,94</c> (con <c>L613 = 12.379.938,06</c>
    /// restado), lo que cierra el R10 oficial al peso.
    /// </summary>
    [Fact]
    public void S2_5_Ase5Octubre_DosFilasMes_TotOptDelDominioYF674EnCero()
    {
        var (r1, r2, r4) = Rutas("Octubre-1", 5);
        var leaf = new ExcelDataReaderWorkbookLeafInputReader()
            .LeerLeafInputs(AseFactory.DesdeId(5), Periodo.Parse("2026101"), r1, r2, r4);

        Assert.DoesNotContain(WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(5), p => p.Celda == "F674");
        Assert.False(leaf.R1.CeldasPorAse.ContainsKey("F674"));
        Assert.InRange(leaf.R1.TotalOportunoEsperadoPorAse - 7544253279.94m, -Tolerancia, Tolerancia);
        Assert.Equal(0m, leaf.R1.ExtemporaneoEsperadoPorAse);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════

    private static (string R1, string R2, string R4) Rutas(string carpetaPeriodo, int aseId)
    {
        var carpeta = Ruta("Docs", $"Prueba {carpetaPeriodo}", "Insumos", CarpetasAse(aseId));

        return (
            Unico(carpeta, "Recaudoporcomponente"),
            Unico(carpeta, "RerpoteDetalleSaldosaFavor"),
            Unico(carpeta, "ReversiónPorComponente"));
    }

    private static string CarpetasAse(int aseId) => aseId switch
    {
        1 => "1-Promoambiental",
        2 => "2-Lime",
        3 => "3-Ciudad Limpia",
        4 => "4-Bogota Limpia",
        5 => "5-Área Limpia",
        _ => throw new ArgumentOutOfRangeException(nameof(aseId))
    };

    private static string Unico(string carpeta, string prefijo)
    {
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .SingleOrDefault(f => Path.GetFileName(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta '{prefijo}*' en {carpeta}.");
        return archivo!;
    }

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }

    private static string Ruta(params string[] partes) => Path.Combine(partes.Prepend(Raiz()).ToArray());
}
