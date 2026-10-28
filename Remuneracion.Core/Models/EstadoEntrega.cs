namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 61 (H3.1 / S3.1, §2-D-H3): estado de entrega de una corrida del período. Su DEFAULT es
/// <see cref="Certificada"/> para que el modelo sea ADITIVO: toda composición que no conozca el
/// concepto conserva el comportamiento anterior.
///
/// <para><b>Política del Ingeniero (2026-10-28).</b> Ante una novedad SIEMPRE se entrega el Excel
/// revisable (marcado <c>(sin validar)</c>) más el reporte de novedades: el usuario puede tomar la
/// salida y ajustar mientras la aplicación se corrige. Por eso "no certificar" ya NO significa "no
/// escribir": significa <see cref="EntregadaConNovedades"/>, que nunca es un cierre en verde.</para>
/// </summary>
public enum EstadoEntrega
{
    /// <summary>
    /// La corrida superó todos los gates de datos: el archivo tomó el nombre definitivo y no hay
    /// reporte de novedades. Comportamiento anterior intacto.
    /// </summary>
    Certificada = 0,

    /// <summary>
    /// La corrida tiene al menos una novedad de datos: el Excel quedó marcado
    /// <c>(sin validar)</c>, se emitió el reporte de novedades y el código de salida NO es 0.
    /// Nunca produce el nombre limpio ni se reporta como <c>OK</c>.
    /// </summary>
    EntregadaConNovedades = 1,
}
