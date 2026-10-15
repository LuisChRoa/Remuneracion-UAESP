namespace Remuneracion.Core.Constants;

/// <summary>
/// Plan 43 (O-3): vocabulario de la huella de quién y por qué canal se reconoció una divergencia
/// declarada. El reporte durable y la traza en log registran SIEMPRE ambos datos (además de celda,
/// observado, esperado, Δ, dueño, evidencia, fecha, condición de retiro y RunId). Cuando el
/// frontend no declara el canal o el sistema no expone el usuario, se registra de forma HONESTA
/// (<see cref="CanalIndefinido"/> / <see cref="UsuarioNoDisponible"/>): nunca se inventa UI/CLI ni
/// un nombre de usuario.
/// </summary>
public static class HuellaAceptacion
{
    /// <summary>Canal de aceptación cuando el usuario opera desde la interfaz gráfica.</summary>
    public const string CanalUi = "UI";

    /// <summary>Canal de aceptación cuando el usuario opera desde la línea de comandos.</summary>
    public const string CanalCli = "CLI";

    /// <summary>Canal no declarado por el frontend: se registra tal cual, sin inventar UI/CLI.</summary>
    public const string CanalIndefinido = "No especificado";

    /// <summary>Usuario no disponible en el entorno: se registra tal cual, sin inventar un nombre.</summary>
    public const string UsuarioNoDisponible = "No disponible";
}
