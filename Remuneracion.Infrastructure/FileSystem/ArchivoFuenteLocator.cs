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
}