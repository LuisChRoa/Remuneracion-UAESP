using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Fixtures compartidos de los insumos reales Q1 (Docs/Insumos). Nombres reales de carpeta
/// (cuidado con la tilde de "5-Área Limpia") y archivos prefix-based (mismo patrón del locator).
/// </summary>
internal static class Insumos
{
    public const decimal Tolerancia = 0.5m;

    public static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")) && Directory.Exists(Path.Combine(dir.FullName, "Docs", "Insumos")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio con Docs/Insumos.");
    }

    public static string Plantilla => Path.Combine(Raiz(), "Docs", "Insumos", "Remuneracion 202607-1 Total.xlsx");

    public static string CarpetaPeriodo => Path.Combine(Raiz(), "Docs", "Insumos", "REMUNERACION 2026071");

    public static string[] CarpetasAse => new[]
    {
        Path.Combine(CarpetaPeriodo, "1-Promoambiental"),
        Path.Combine(CarpetaPeriodo, "2-Lime"),
        Path.Combine(CarpetaPeriodo, "3-Ciudad Limpia"),
        Path.Combine(CarpetaPeriodo, "4-Bogota Limpia"),
        Path.Combine(CarpetaPeriodo, "5-Área Limpia")
    };

    public static string R1(int aseId) => Buscar(CarpetasAse[aseId - 1], "Recaudoporcomponente", $"R1 ASE{aseId}");

    public static string R2(int aseId) => Buscar(CarpetasAse[aseId - 1], "RerpoteDetalleSaldosaFavor", $"R2 ASE{aseId}");

    public static string R4(int aseId)
    {
        var carpeta = CarpetasAse[aseId - 1];
        var r4 = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("ReversiónPorComponente", StringComparison.OrdinalIgnoreCase))
            ?? Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
                .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("ReversionPorComponente", StringComparison.OrdinalIgnoreCase));

        return r4 ?? throw new FileNotFoundException($"No se encontró R4 ASE{aseId} en {carpeta}.");
    }

    /// <summary>
    /// HU-09 (2.3, V13): ruta del <c>ReportePagosxBanco_*.xlsx</c> del ASE (naming verificado por T0).
    /// </summary>
    public static string ReporteBanco(int aseId) => Buscar(CarpetasAse[aseId - 1], "ReportePagosxBanco", $"ReportePagosxBanco ASE{aseId}");

    /// <summary>
    /// HU-10 (2.4, V8): ruta del <c>R4-BalanceSubsidioyContribuciones_*.xlsx</c> del ASE.
    /// Mismo criterio del locator: prefijo base (con guion bajo, excluye la variante
    /// <c>-Optimizado_</c>) primero y variante <c>R4-BalanceSubsidioyContribuciones-Optimizado_</c>
    /// como fallback. El archivo <c>Reca_BalanceSubsidiosyContribuciones_*</c> de ASE5-Q1 es otro
    /// reporte y NO matchea ninguno de los dos prefijos.
    /// </summary>
    public static string Balance(int aseId)
    {
        var carpeta = CarpetasAse[aseId - 1];
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("R4-BalanceSubsidioyContribuciones_", StringComparison.OrdinalIgnoreCase))
            ?? Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
                .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("R4-BalanceSubsidioyContribuciones-Optimizado_", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Balance ASE{aseId} en {carpeta}");
        return archivo!;
    }

    /// <summary>
    /// HU-14 (S-3): helper migrado a la factoría única <see cref="AseFactory.DesdeId"/>
    /// (fuente única desde <see cref="Remuneracion.Core.Constants.CarpetasAse.Prefijos"/>).
    /// </summary>
    public static Ase Ase(int id) => AseFactory.DesdeId(id);

    public static Periodo Periodo() => new() { CodigoAAAAMM = "202607", NumeroQuincena = 1 };

    // ── HU-11 (2.5): insumos Q2 ──────────────────────────────────────────────────────────────

    /// <summary>
    /// HU-11 (2.5, T0-0.1): golden canónico Q2 = "Plantilla 8 agos 2026" (coincide con la
    /// referencia en los valores cacheados de validación). SHA256 95825422… La otra plantilla
    /// queda como control, NUNCA oráculo (G6/D8).
    /// </summary>
    public static string PlantillaQ2 => Path.Combine(Raiz(), "Docs", "Insumos", "REMUNERACION 2026072", "Plantilla 8 agos 2026 _ Remuneracion 202607-2 Total.xlsx");

    /// <summary>
    /// HU-11 (2.5, T0-0.1): golden de VALORES Q2 (referencia completada) usado como caché en la
    /// Capa A (D85:D89, D104:D109, AJUSTES-SF-T). SHA256 58431010…
    /// </summary>
    public static string GoldenQ2 => Path.Combine(Raiz(), "Docs", "Insumos", "Remuneracion 202607-2 Total.xlsx");

    /// <summary>
    /// HU-11 (2.5): la otra plantilla 202607-2 = control/estructura, nunca oráculo de merge.
    /// </summary>
    public static string PlantillaQ2Control => Path.Combine(Raiz(), "Docs", "Insumos", "REMUNERACION 2026072", "Plantilla  _ Remuneracion 202607-2 Total.xlsx");

    public static string CarpetaPeriodoQ2 => Path.Combine(Raiz(), "Docs", "Insumos", "REMUNERACION 2026072");

    public static string[] CarpetasAseQ2 => new[]
    {
        Path.Combine(CarpetaPeriodoQ2, "1-Promoambiental"),
        Path.Combine(CarpetaPeriodoQ2, "2-Lime"),
        Path.Combine(CarpetaPeriodoQ2, "3-Ciudad Limpia"),
        Path.Combine(CarpetaPeriodoQ2, "4-Bogota Limpia"),
        Path.Combine(CarpetaPeriodoQ2, "5-Área Limpia")
    };

    public static Periodo PeriodoQ2() => new() { CodigoAAAAMM = "202607", NumeroQuincena = 2 };

    /// <summary>
    /// HU-11 (2.5): ruta del <c>Recaudoporcomponente_*.xlsx</c> del ASE Q2 (layout ASE1-4
    /// paritario con Q1; ASE5 diverge — recorte T0-0.6).
    /// </summary>
    public static string R1Q2(int aseId) => Buscar(CarpetasAseQ2[aseId - 1], "Recaudoporcomponente", $"R1 Q2 ASE{aseId}");

    /// <summary>
    /// HU-12 (2.6 ampliada, V0.3): ruta del R1-Q2 de ASE5 (variante 2 filas Mes/Total — el
    /// dispatch Q2 del reader la soporta). Alias explícito para los tests de la variante.
    /// </summary>
    public static string R1Q2Ase5() => R1Q2(5);

    /// <summary>
    /// HU-11 (2.5): ruta del <c>RerpoteDetalleSaldosaFavor_*.xlsx</c> del ASE Q2.
    /// </summary>
    public static string R2Q2(int aseId) => Buscar(CarpetasAseQ2[aseId - 1], "RerpoteDetalleSaldosaFavor", $"R2 Q2 ASE{aseId}");

    /// <summary>
    /// HU-11 (2.5): ruta del <c>ReversiónPorComponente_*.xlsx</c> del ASE Q2 (diacríticos en disco).
    /// </summary>
    public static string R4Q2(int aseId)
    {
        var carpeta = CarpetasAseQ2[aseId - 1];
        var r4 = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Normalizar(Path.GetFileNameWithoutExtension(f)).StartsWith("reversionporcomponente", StringComparison.Ordinal));
        Assert.True(r4 is not null, $"Falta R4 Q2 ASE{aseId} en {carpeta}");
        return r4!;
    }

    /// <summary>
    /// HU-11 (2.5): ruta del <c>R4-BalanceSubsidioyContribuciones_*</c> del ASE Q2 (ASE5 solo
    /// variante -Optimizado; fallback obligatorio, V4).
    /// </summary>
    public static string BalanceQ2(int aseId)
    {
        var carpeta = CarpetasAseQ2[aseId - 1];
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("R4-BalanceSubsidioyContribuciones_", StringComparison.OrdinalIgnoreCase))
            ?? Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
                .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("R4-BalanceSubsidioyContribuciones-Optimizado_", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Balance Q2 ASE{aseId} en {carpeta}");
        return archivo!;
    }

    /// <summary>
    /// HU-11 (2.5): ruta del <c>ReportePagosxBanco_*.xlsx</c> del ASE Q2.
    /// </summary>
    public static string ReporteBancoQ2(int aseId) => Buscar(CarpetasAseQ2[aseId - 1], "ReportePagosxBanco", $"ReportePagosxBanco Q2 ASE{aseId}");

    /// <summary>
    /// HU-11 (2.5, T0-0.5): ruta del <c>SaldosaFavorAplicadosPorNotas_*.xlsx</c> del ASE Q2.
    /// </summary>
    public static string SaldosNotas(int aseId) => Buscar(CarpetasAseQ2[aseId - 1], "SaldosaFavorAplicadosPorNotas", $"SaldosNotas ASE{aseId}");

    /// <summary>
    /// HU-11 (2.5, T0-0.4/0.5): ruta del <c>RetribuciónNegativa_*.xlsx</c> del ASE Q2. El nombre
    /// en disco lleva diacríticos (V9): el match usa el stem normalizado sin acentos.
    /// </summary>
    public static string RetribucionNegativa(int aseId)
    {
        var carpeta = CarpetasAseQ2[aseId - 1];
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Normalizar(Path.GetFileNameWithoutExtension(f)).StartsWith("retribucionnegativa", StringComparison.Ordinal));
        Assert.True(archivo is not null, $"Falta RetribuciónNegativa ASE{aseId} en {carpeta}");
        return archivo!;
    }

    /// <summary>
    /// HU-11 (2.5, T0-0.4): normaliza un nombre (minúsculas, sin diacríticos) para el match por prefijo.
    /// </summary>
    private static string Normalizar(string texto)
    {
        var normalizado = texto.Normalize(System.Text.NormalizationForm.FormD);
        var sb = new System.Text.StringBuilder(normalizado.Length);
        foreach (var ch in normalizado)
        {
            if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch) != System.Globalization.UnicodeCategory.NonSpacingMark)
            {
                sb.Append(char.ToLowerInvariant(ch));
            }
        }

        return sb.ToString();
    }

    private static string Buscar(string carpeta, string prefijo, string etiqueta)
    {
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta {etiqueta} en {carpeta}");
        return archivo!;
    }
}
