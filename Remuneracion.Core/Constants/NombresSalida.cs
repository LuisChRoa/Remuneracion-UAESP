namespace Remuneracion.Core.Constants;

/// <summary>
/// Plan 43 (O-4): convención de nombres de los archivos de salida. Una corrida que NO cierra bien
/// NUNCA produce el nombre definitivo: el archivo queda con una marca visible y determinista de
/// "sin validar" (misma entrada → misma ruta). Solo una corrida que supera las validaciones produce
/// el nombre limpio. El administrativo distingue a simple vista un entregable certificado de un
/// descarte.
/// </summary>
public static class NombresSalida
{
    /// <summary>
    /// Sufijo visible y determinista del archivo de una corrida que todavía no cerró bien. Se
    /// inserta ANTES de la extensión (p. ej. <c>… Total (sin validar).xlsx</c>).
    /// </summary>
    public const string SufijoSinValidar = " (sin validar)";

    /// <summary>
    /// Ruta "marcada" de una salida que aún no superó las validaciones. Determinista: la misma
    /// entrada produce siempre la misma ruta (mismo directorio, mismo nombre base, mismo sufijo).
    /// </summary>
    /// <param name="rutaSalida">Ruta definitiva solicitada por el frontend.</param>
    /// <returns>Ruta con la marca de "sin validar" antes de la extensión.</returns>
    public static string RutaSinValidar(string rutaSalida)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rutaSalida);
        var absoluta = Path.GetFullPath(rutaSalida);
        var directorio = Path.GetDirectoryName(absoluta) ?? ".";
        var nombreBase = Path.GetFileNameWithoutExtension(absoluta);
        var extension = Path.GetExtension(absoluta);
        return Path.Combine(directorio, $"{nombreBase}{SufijoSinValidar}{extension}");
    }
}
