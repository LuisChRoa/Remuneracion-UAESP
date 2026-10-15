namespace Remuneracion.Core.Models;

/// <summary>
/// Solicitud para ejecutar la liquidación de un período completo (los 5 ASE) en un solo proceso.
/// </summary>
public sealed class SolicitudProcesoPeriodo
{
    /// <summary>
    /// Periodo quincenal a liquidar.
    /// </summary>
    public Periodo Periodo { get; set; } = new();

    /// <summary>
    /// Carpeta del período (ej. "REMUNERACION 2026071") que contiene las 5 carpetas de ASE.
    /// </summary>
    public string CarpetaPeriodo { get; set; } = string.Empty;

    /// <summary>
    /// Ruta de la plantilla de origen.
    /// </summary>
    public string RutaPlantilla { get; set; } = string.Empty;

    /// <summary>
    /// Ruta de salida (carpeta + <see cref="Periodo.NombreArchivo"/>; igual que HU-06).
    /// </summary>
    public string RutaSalida { get; set; } = string.Empty;

    /// <summary>
    /// HU-15 (W-2.1, D4): RunId inyectado por el frontend (UI/CLI) para correlacionar TODOS
    /// los eventos de la ejecución (UI → procesador → writer) en el log. <c>null</c> = el
    /// procesador genera uno (compat HU-14: <c>Guid.NewGuid()</c>).
    /// </summary>
    public Guid? RunId { get; set; }

    /// <summary>
    /// Plan 42 (T1/T3, D-D/D-G): ids de las entradas del catálogo de divergencias declaradas que el
    /// usuario reconoció explícitamente en ESTA corrida (UI: acto consciente; CLI:
    /// <c>--aceptar-divergencia-declarada</c>). No es persistente ni "por costumbre": cada corrida
    /// re-presenta la divergencia y exige su propio reconocimiento. Sin el id de la entrada vigente,
    /// la declarada bloquea con <see cref="Errors.CodigoError.Validacion"/> (S3).
    /// </summary>
    public IReadOnlyList<string> DivergenciasDeclaradasAceptadas { get; set; } = [];

    /// <summary>
    /// Plan 43 (O-3): canal por el que el usuario aceptó la divergencia declarada
    /// (<see cref="Constants.HuellaAceptacion.CanalUi"/> o <see cref="Constants.HuellaAceptacion.CanalCli"/>).
    /// Aditivo: lo declara cada frontend. <c>null</c>/vacío = <see cref="Constants.HuellaAceptacion.CanalIndefinido"/>
    /// (nunca se inventa UI/CLI).
    /// </summary>
    public string? Canal { get; set; }

    /// <summary>
    /// Plan 43 (O-3): usuario del sistema que aceptó la divergencia declarada. Aditivo: si es
    /// <c>null</c>/vacío se usa <see cref="Environment.UserName"/> (o
    /// <see cref="Constants.HuellaAceptacion.UsuarioNoDisponible"/> si el entorno no lo expone).
    /// </summary>
    public string? UsuarioAceptacion { get; set; }
}
