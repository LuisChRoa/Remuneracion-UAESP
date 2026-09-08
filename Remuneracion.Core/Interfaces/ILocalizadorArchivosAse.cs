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

    /// <summary>
    /// HU-09 (2.3, V13): localiza el <c>ReportePagosxBanco_*.xlsx</c> dentro de la carpeta del
    /// ASE (prefijo nuevo del locator; naming verificado en disco por T0).
    /// </summary>
    /// <param name="carpetaAse">Carpeta del ASE (p. ej. "1-Promoambiental").</param>
    /// <returns>Ruta completa del archivo, o <c>null</c> si no se encuentra.</returns>
    string? BuscarReporteBanco(string carpetaAse);

    /// <summary>
    /// HU-10 (2.4, V8): localiza el <c>R4-BalanceSubsidioyContribuciones_*.xlsx</c> dentro de la
    /// carpeta del ASE. Acepta los DOS prefijos verificados en disco: el base
    /// <c>R4-BalanceSubsidioyContribuciones_</c> (ASE1..5 Q1) y la variante
    /// <c>R4-BalanceSubsidioyContribuciones-Optimizado_</c> (ASE5 en Q2). Base primero; el
    /// archivo <c>Reca_BalanceSubsidiosyContribuciones_*</c> de ASE5-Q1 es OTRO reporte (layout
    /// por componentes, sin "Total General") y NO debe resolverse por ninguno de los dos prefijos.
    /// </summary>
    /// <param name="carpetaAse">Carpeta del ASE.</param>
    /// <returns>Ruta completa del archivo, o <c>null</c> si no se encuentra.</returns>
    string? BuscarBalance(string carpetaAse);
}