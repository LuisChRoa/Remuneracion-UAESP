using Remuneracion.Core.Models;
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

    public static Ase Ase(int id)
    {
        var nombres = new[] { "Promoambiental", "Lime", "Ciudad Limpia", "Bogotá Limpia", "Área Limpia" };
        return new Ase
        {
            Id = id,
            NombreCorto = nombres[id - 1].ToUpperInvariant(),
            NombreCompleto = nombres[id - 1],
            NumeroCarpeta = id
        };
    }

    public static Periodo Periodo() => new() { CodigoAAAAMM = "202607", NumeroQuincena = 1 };

    private static string Buscar(string carpeta, string prefijo, string etiqueta)
    {
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta {etiqueta} en {carpeta}");
        return archivo!;
    }
}