using System.IO;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Interfaces;

namespace Remuneracion.Infrastructure.FileSystem;

/// <summary>
/// Localiza los archivos fuente de los ASE dentro de la estructura de carpetas del periodo.
/// </summary>
public class ArchivoFuenteLocator : ILocalizadorArchivosAse
{
    /// <summary>
    /// Obtiene los subdirectorios de ASE que existen dentro de la carpeta del periodo,
    /// filtrando por los prefijos conocidos (1- a 5-).
    /// </summary>
    /// <param name="carpetaPeriodo">Ruta de la carpeta del periodo quincenal.</param>
    /// <returns>Rutas de las carpetas de ASE encontradas.</returns>
    public string[] ObtenerCarpetasAse(string carpetaPeriodo)
    {
        ArgumentNullException.ThrowIfNull(carpetaPeriodo);

        if (!Directory.Exists(carpetaPeriodo))
        {
            return [];
        }

        return Directory.GetDirectories(carpetaPeriodo)
            .Where(dir => CarpetasAse.Prefijos.Any(p => Path.GetFileName(dir).StartsWith(p[..2], StringComparison.OrdinalIgnoreCase)))
            .ToArray();
    }

    /// <summary>
    /// Busca el primer archivo .xlsx cuya raíz de nombre coincida con el prefijo dado.
    /// </summary>
    /// <param name="carpeta">Carpeta donde buscar.</param>
    /// <param name="prefijo">Prefijo (case-insensitive) del nombre del archivo.</param>
    /// <returns>Ruta completa del archivo, o <c>null</c> si no se encuentra.</returns>
    public string? BuscarArchivo(string carpeta, string prefijo)
    {
        ArgumentNullException.ThrowIfNull(carpeta);
        ArgumentNullException.ThrowIfNull(prefijo);

        if (!Directory.Exists(carpeta))
        {
            return null;
        }

        return Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f)
                .StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// HU-08 (2.2, T0-0.6): localiza el archivo de conciliación por empresa dentro de
    /// <c>{carpetaPeriodo}/Consolidado/Conciliaciones/</c> (p. ej. "Conjunta ENEL",
    /// "Directa", "Conjunta Otros"). Devuelve <c>null</c> si no existe.
    /// </summary>
    /// <param name="carpetaPeriodo">Carpeta del período quincenal.</param>
    /// <param name="prefijo">Prefijo del nombre del archivo de conciliación.</param>
    /// <returns>Ruta completa del archivo, o <c>null</c> si no se encuentra.</returns>
    public string? BuscarConciliacion(string carpetaPeriodo, string prefijo)
    {
        ArgumentNullException.ThrowIfNull(carpetaPeriodo);
        ArgumentNullException.ThrowIfNull(prefijo);

        var carpetaConciliaciones = Path.Combine(carpetaPeriodo, "Consolidado", "Conciliaciones");
        if (!Directory.Exists(carpetaConciliaciones))
        {
            return null;
        }

        return Directory.EnumerateFiles(carpetaConciliaciones, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f)
                .StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// HU-09 (2.3, V13): localiza el <c>ReportePagosxBanco_*.xlsx</c> dentro de la carpeta del
    /// ASE (mismo patrón prefix-based de R1/R2/R4; naming verificado en disco por T0).
    /// </summary>
    /// <param name="carpetaAse">Carpeta del ASE (p. ej. "1-Promoambiental").</param>
    /// <returns>Ruta completa del archivo, o <c>null</c> si no se encuentra.</returns>
    public string? BuscarReporteBanco(string carpetaAse) =>
        BuscarArchivo(carpetaAse, "ReportePagosxBanco");

    /// <summary>
    /// HU-10 (2.4, V8): localiza el <c>R4-BalanceSubsidioyContribuciones_*.xlsx</c> dentro de la
    /// carpeta del ASE con los DOS prefijos verificados en disco: base
    /// <c>R4-BalanceSubsidioyContribuciones_</c> primero (el guion bajo excluye la variante
    /// <c>-Optimizado_</c> por prefijo), y la variante <c>R4-BalanceSubsidioyContribuciones-Optimizado_</c>
    /// como fallback (ASE5 en Q2). El archivo <c>Reca_BalanceSubsidiosyContribuciones_*</c> de
    /// ASE5-Q1 es otro reporte (layout por componentes, sin "Total General") y NO matchea ninguno
    /// de los dos prefijos.
    /// </summary>
    /// <param name="carpetaAse">Carpeta del ASE.</param>
    /// <returns>Ruta completa del archivo, o <c>null</c> si no se encuentra.</returns>
    public string? BuscarBalance(string carpetaAse) =>
        BuscarArchivo(carpetaAse, "R4-BalanceSubsidioyContribuciones_")
        ?? BuscarArchivo(carpetaAse, "R4-BalanceSubsidioyContribuciones-Optimizado_");

    /// <summary>
    /// HU-11 (2.5, D4): localiza el <c>SaldosaFavorAplicadosPorNotas_*.xlsx</c> del ASE.
    /// Prefijo sin fechas (V3: los rangos varían por ASE; prohibido codificarlos) y el nombre
    /// en disco no lleva diacríticos (V9). Matcher normalizado: lowercase + sin diacríticos.
    /// </summary>
    /// <param name="carpetaAse">Carpeta del ASE.</param>
    /// <returns>Ruta completa del archivo, o <c>null</c> si no se encuentra.</returns>
    public string? BuscarSaldosNotas(string carpetaAse) =>
        BuscarArchivoNormalizado(carpetaAse, "saldosafavoraplicadospornotas");

    /// <summary>
    /// HU-11 (2.5, D4): localiza el <c>RetribuciónNegativa_*.xlsx</c> del ASE. El nombre en disco
    /// lleva diacríticos (<c>RetribuciónNegativa_…</c>, V9) y el rango de fechas NO es homogéneo
    /// (ASE4: 16072026–31072026, resto 0107–3107, V3) → el matcher normaliza y usa el stem sin
    /// acento <c>retribucionnegativa</c>; prohibido literal con acento frágil o rango de fechas.
    /// </summary>
    /// <param name="carpetaAse">Carpeta del ASE.</param>
    /// <returns>Ruta completa del archivo, o <c>null</c> si no se encuentra.</returns>
    public string? BuscarRetribucionNegativa(string carpetaAse) =>
        BuscarArchivoNormalizado(carpetaAse, "retribucionnegativa");

    /// <summary>
    /// Busca el primer archivo .xlsx cuyo nombre normalizado (lowercase, sin diacríticos) comienza
    /// con el prefijo normalizado. Agnóstico a rango de fechas y a diacríticos (D4/V3/V9).
    /// </summary>
    private static string? BuscarArchivoNormalizado(string carpeta, string prefijoNormalizado)
    {
        ArgumentNullException.ThrowIfNull(carpeta);
        ArgumentNullException.ThrowIfNull(prefijoNormalizado);

        if (!Directory.Exists(carpeta))
        {
            return null;
        }

        return Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => NormalizarParaMatch(Path.GetFileNameWithoutExtension(f))
                .StartsWith(prefijoNormalizado, StringComparison.Ordinal));
    }

    /// <summary>
    /// Normaliza un nombre para el match por prefijo: minúsculas y sin diacríticos (V9).
    /// "RetribuciónNegativa" → "retribucionnegativa".
    /// </summary>
    private static string NormalizarParaMatch(string texto)
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
}