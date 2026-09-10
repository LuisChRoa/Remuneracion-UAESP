using Remuneracion.Core.Exceptions;

namespace Remuneracion.Core.Errors;

/// <summary>
/// HU-14 (3.1, D2): mensajes UX centralizados por categoría del catálogo. La UI consume
/// <see cref="Para"/> para mostrar título de MessageBox + guía accionable; el detalle técnico
/// completo queda en el log (nunca <c>ex.Message</c> crudo como texto principal del box).
/// Testeable in-memory (sin WinForms).
/// </summary>
public static class CatalogoErrores
{
    /// <summary>
    /// Devuelve (título del MessageBox, guía accionable) para un código del catálogo.
    /// Código desconocido → <see cref="CodigoError.Inesperado"/> (guía genérica + RunId).
    /// </summary>
    public static (string Titulo, string Guia) Para(string codigo, Exception? ex)
    {
        var detalle = ex?.Message ?? string.Empty;
        return codigo switch
        {
            CodigoError.FuenteNoEncontrada => (
                "Archivo fuente faltante",
                $"Falta un archivo fuente del período: {detalle}. Verifique la carpeta indicada, genere el reporte faltante y reintente."),
            CodigoError.Plantilla => (
                "Plantilla o ruta de salida no válida",
                $"La plantilla o la ruta de salida no son válidas: {detalle}. Use una plantilla existente y una ruta de salida distinta a la plantilla."),
            CodigoError.FormatoFuente => (
                "Formato de archivo fuente no compatible",
                $"Un archivo fuente no tiene la estructura esperada: {detalle}. Revise el archivo, la hoja y la celda indicadas."),
            CodigoError.Validacion => (
                "La validación no cierra",
                $"La validación no cierra dentro de la tolerancia ±0.5: {detalle}."),
            CodigoError.Escritura => (
                "Error al generar el archivo de salida",
                $"No se pudo escribir el workbook: {detalle}. No quedó archivo parcial; reintente."),
            CodigoError.CanceladoPorUsuario => (
                "Proceso cancelado",
                $"El proceso se canceló por decisión del usuario: {detalle}."),
            _ => (
                "Error inesperado",
                "Ocurrió un error inesperado durante la ejecución. Búsquelo en el log (remuneracion_log_*.txt) con el RunId de esta ejecución.")
        };
    }

    /// <summary>
    /// Mapa código → <see cref="CodigosSalida"/> (contrato HU-15). Código desconocido → Inesperado.
    /// </summary>
    public static int CodigoSalidaPara(string codigo) => codigo switch
    {
        CodigoError.Validacion => CodigosSalida.Validacion,
        CodigoError.FuenteNoEncontrada or CodigoError.Plantilla or CodigoError.FormatoFuente => CodigosSalida.FuenteOPlantilla,
        CodigoError.Escritura => CodigosSalida.Escritura,
        CodigoError.CanceladoPorUsuario => CodigosSalida.CanceladoPorUsuario,
        _ => CodigosSalida.Inesperado
    };

    /// <summary>
    /// HU-15 (D6): código del catálogo para una excepción, sin UI. Las 2 excepciones de dominio
    /// portan <c>Codigo</c> (D1); cualquier otra = <see cref="CodigoError.Inesperado"/>.
    /// <c>Form1.ObtenerCodigoError</c> delega aquí y el CLI consume el mismo mapeo (una sola
    /// fuente; testeable in-memory).
    /// </summary>
    public static string CodigoDe(Exception ex) => ex switch
    {
        ArchivoFuenteNoEncontradoException archivo => archivo.Codigo,
        CalculoInvalidoException calculo => calculo.Codigo,
        _ => CodigoError.Inesperado
    };
}
