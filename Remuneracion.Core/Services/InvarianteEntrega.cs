using Remuneracion.Core.Constants;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Services;

/// <summary>
/// Plan 61 (H3.4 / CA-C3, §2-D-H3-iii): guarda de la PROHIBICIÓN DE FALSO VERDE. Una corrida con
/// novedades no puede presentarse de ninguna forma como certificada, y una corrida certificada no
/// puede apuntar a un archivo marcado. El procesador la invoca antes de devolver y los tests la
/// ejercitan con injertos: es el punto único donde un estado de entrega inconsistente se vuelve
/// imposible (rojo) en vez de viajar al usuario.
///
/// <para><b>Invariantes</b></para>
/// <list type="number">
/// <item><see cref="ResultadoProcesoPeriodo.Novedades"/> no vacía ⇒ estado
/// <see cref="EstadoEntrega.EntregadaConNovedades"/> (nunca <see cref="EstadoEntrega.Certificada"/>).</item>
/// <item>… y la ruta entregada termina en <see cref="NombresSalida.SufijoSinValidar"/> (nunca el nombre limpio).</item>
/// <item>… y existe el reporte de novedades (sin TXT no hay entrega: no cierra en silencio).</item>
/// <item>Estado <see cref="EstadoEntrega.EntregadaConNovedades"/> ⇒ hay al menos una novedad
/// (un estado "con novedades" vacío también sería un reporte que miente).</item>
/// <item>Estado <see cref="EstadoEntrega.Certificada"/> ⇒ sin novedades y sin marca de "sin validar".</item>
/// </list>
/// </summary>
public static class InvarianteEntrega
{
    /// <summary>
    /// Verifica las invariantes de entrega del resultado. Lanza
    /// <see cref="CalculoInvalidoException"/> con <see cref="CodigoError.Inesperado"/> (nunca un
    /// cierre silencioso) nombrando la invariante violada.
    /// </summary>
    /// <param name="resultado">Resultado del período a verificar.</param>
    public static void Garantizar(ResultadoProcesoPeriodo resultado)
    {
        ArgumentNullException.ThrowIfNull(resultado);

        var tieneNovedades = resultado.Novedades.Count > 0;

        if (tieneNovedades && resultado.EstadoEntrega != EstadoEntrega.EntregadaConNovedades)
        {
            throw new CalculoInvalidoException(
                CodigoError.Inesperado,
                $"Entrega inconsistente: la corrida trae {resultado.Novedades.Count} novedad(es) y su estado es '{resultado.EstadoEntrega}'. Una corrida con novedades NUNCA se certifica (prohibido el falso verde).");
        }

        if (tieneNovedades && !TerminaEnMarcaSinValidar(resultado.RutaSalida))
        {
            throw new CalculoInvalidoException(
                CodigoError.Inesperado,
                $"Entrega inconsistente: la corrida trae {resultado.Novedades.Count} novedad(es) y la ruta entregada '{resultado.RutaSalida}' no está marcada con '{NombresSalida.SufijoSinValidar}'. Una corrida con novedades NUNCA toma el nombre limpio.");
        }

        if (tieneNovedades && (string.IsNullOrWhiteSpace(resultado.RutaReporteNovedades) || !File.Exists(resultado.RutaReporteNovedades)))
        {
            throw new CalculoInvalidoException(
                CodigoError.Inesperado,
                $"Entrega inconsistente: la corrida trae {resultado.Novedades.Count} novedad(es) y no existe el reporte de novedades ('{resultado.RutaReporteNovedades}'). Sin reporte no hay entrega: el usuario no puede revisar ni comparar.");
        }

        if (!tieneNovedades && resultado.EstadoEntrega == EstadoEntrega.EntregadaConNovedades)
        {
            throw new CalculoInvalidoException(
                CodigoError.Inesperado,
                "Entrega inconsistente: el resultado declara 'entregado con novedades' pero no trae ninguna novedad que revisar.");
        }

        if (!tieneNovedades && TerminaEnMarcaSinValidar(resultado.RutaSalida))
        {
            throw new CalculoInvalidoException(
                CodigoError.Inesperado,
                $"Entrega inconsistente: la corrida no trae novedades y aun así la ruta entregada '{resultado.RutaSalida}' conserva la marca de '{NombresSalida.SufijoSinValidar}': un archivo marcado no se certifica.");
        }
    }

    /// <summary>Verdadero si la ruta termina en la marca determinista de "sin validar".</summary>
    private static bool TerminaEnMarcaSinValidar(string ruta)
    {
        if (string.IsNullOrWhiteSpace(ruta))
        {
            return false;
        }

        var nombre = Path.GetFileNameWithoutExtension(ruta);
        return nombre.EndsWith(NombresSalida.SufijoSinValidar, StringComparison.OrdinalIgnoreCase);
    }
}
