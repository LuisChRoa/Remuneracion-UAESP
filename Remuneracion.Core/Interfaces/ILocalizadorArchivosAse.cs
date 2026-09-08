namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Contrato para localizar carpetas y archivos fuente de los ASE dentro de la estructura del período.
/// Implementado en Infrastructure (<c>ArchivoFuenteLocator</c>); Core orquesta contra este contrato.
/// </summary>
public interface ILocalizadorArchivosAse
{
    /// <summary>
    /// Obtiene los subdirectorios de ASE que existen dentro de la carpeta del período,
    /// filtrando por los prefijos conocidos (1- a 5-).
    /// </summary>
    /// <param name="carpetaPeriodo">Ruta de la carpeta del periodo quincenal.</param>
    /// <returns>Rutas de las carpetas de ASE encontradas.</returns>
    string[] ObtenerCarpetasAse(string carpetaPeriodo);

    /// <summary>
    /// Busca el primer archivo .xlsx cuya raíz de nombre coincida con el prefijo dado.
    /// </summary>
    /// <param name="carpeta">Carpeta donde buscar.</param>
    /// <param name="prefijo">Prefijo (case-insensitive) del nombre del archivo.</param>
    /// <returns>Ruta completa del archivo, o <c>null</c> si no se encuentra.</returns>
    string? BuscarArchivo(string carpeta, string prefijo);

    /// <summary>
    /// HU-08 (2.2, T0-0.6): localiza el archivo de conciliación por empresa dentro de
    /// <c>{carpetaPeriodo}/Consolidado/Conciliaciones/</c>.
    /// </summary>
    /// <param name="carpetaPeriodo">Carpeta del período quincenal.</param>
    /// <param name="prefijo">Prefijo del nombre del archivo de conciliación.</param>
    /// <returns>Ruta completa del archivo, o <c>null</c> si no se encuentra.</returns>
    string? BuscarConciliacion(string carpetaPeriodo, string prefijo);
}