using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Plan 61 (H4.1 / S3.3, §2-D-H4): contrato del REPORTE DE NOVEDADES de la corrida. Mientras el
/// reporte de divergencias declaradas (Plan 42) documenta una deuda reconocida, este reporte
/// documenta <b>qué hay que revisar antes de dar la corrida por buena</b>: es el artefacto que
/// permite al administrativo tomar el Excel marcado <c>(sin validar)</c> y compararlo.
///
/// <para><b>Un solo reporte, con secciones.</b> Cuando la corrida además cierra divergencias
/// declaradas, este reporte las absorbe como sección final (prohibido tener dos TXT que se
/// contradigan, R-DOBLE-REPORTE). Compatibilidad ratificada (D-TXT): si SOLO hay declaradas, se
/// conserva el nombre anterior <c> - Divergencias declaradas.txt</c> y este writer no se invoca.</para>
///
/// <para>La implementación vive en Infrastructure (I/O); Core solo la invoca con las novedades ya
/// clasificadas.</para>
/// </summary>
public interface IReporteNovedades
{
    /// <summary>
    /// Emite el reporte de novedades junto a la salida y devuelve su ruta absoluta. El texto es de
    /// NEGOCIO y determinista: la misma corrida reescribe el mismo archivo. UTF-8 sin BOM
    /// (precedente Plan 42).
    ///
    /// <para><b>Reporte escueto (D-UX-TXT, 2026-10-28).</b> El documento trae solo el veredicto, la
    /// ficha de cada novedad y la guía de revisión. NO imprime bloque de datos de corrida, ni
    /// explicación general del estado, ni reglas de comparación: eso vive en el manual de usuario y
    /// la trazabilidad (RunId, usuario, canal, fecha) en el log y en la salida de UI/CLI. Los
    /// parámetros se conservan por estabilidad del contrato y para esa trazabilidad.</para>
    /// </summary>
    /// <param name="rutaSalida">
    /// Ruta DEFINITIVA solicitada (sin la marca de "sin validar"): gobierna el nombre del reporte,
    /// de modo que la misma corrida produzca siempre el mismo archivo.
    /// </param>
    /// <param name="rutaExcelEntregado">Archivo realmente entregado (puede llevar la marca de "sin validar"). No se imprime en el reporte.</param>
    /// <param name="periodo">Período de la corrida (gobierna el encabezado y el nombre).</param>
    /// <param name="runId">RunId de la corrida (correlación con el log). No se imprime en el reporte.</param>
    /// <param name="novedades">Novedades tipadas, en lenguaje de negocio. Lista vacía = no hay nada que reportar.</param>
    /// <param name="declaradas">Divergencias declaradas de la corrida (sección final); <c>null</c> = sin declaradas.</param>
    /// <param name="usuario">Usuario que ejecutó la corrida. No se imprime en el reporte.</param>
    /// <param name="canal">Canal de la corrida (UI/CLI). No se imprime en el reporte.</param>
    /// <param name="fechaHora">Fecha y hora de la corrida. No se imprime en el reporte.</param>
    /// <returns>Ruta absoluta del reporte emitido.</returns>
    string Emitir(
        string rutaSalida,
        string rutaExcelEntregado,
        Periodo periodo,
        Guid runId,
        IReadOnlyList<NovedadEntrega> novedades,
        ClasificacionDivergenciasDeclaradas? declaradas = null,
        string? usuario = null,
        string? canal = null,
        DateTime? fechaHora = null);
}
