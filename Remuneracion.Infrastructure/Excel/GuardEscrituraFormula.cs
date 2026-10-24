using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 58 (H1.1) — <b>predicado único</b> del escritor values-only: ninguna rutina de escritura
/// puede escribir sobre una celda que trae <c>&lt;f&gt;</c>.
///
/// La condición se evalúa sobre el CONTENIDO de la celda (<see cref="Cell.CellFormula"/> de la
/// plantilla), nunca sobre el nombre de la hoja, la celda, el ASE, la empresa, la quincena ni el
/// período: una plantilla futura queda protegida sin agregar código.
///
/// Por qué existe este tipo: los guards anti-fórmula estaban duplicados en cada escritor
/// (<c>OpenXmlPlantillaWriter.EscribirValorNumerico</c>, <c>EscribirValorTexto</c>,
/// <c>EscritorValoresSlotR1.EscribirValor</c>). Duplicado ⇒ un camino nuevo puede olvidarlo. Acá
/// hay una sola puerta: todo camino que escriba un valor la cruza. La otra puerta del paquete es
/// <see cref="ValidadorIntegridadPaquete"/>, que verifica el resultado por bytes.
/// </summary>
internal static class GuardEscrituraFormula
{
    /// <summary>
    /// Lanza <c>ERR-PLANTILLA</c> si <paramref name="celda"/> ya trae una fórmula en la plantilla.
    /// </summary>
    /// <param name="celda">Celda destino del writer (con o sin fórmula).</param>
    /// <param name="hoja">Nombre de la hoja (solo para el mensaje).</param>
    /// <param name="referencia">Referencia A1 de la celda (solo para el mensaje).</param>
    /// <param name="operacion">Rutina/contrato que intentó escribir (solo para el mensaje).</param>
    internal static void ExigirSinFormula(Cell celda, string hoja, string referencia, string operacion)
    {
        ArgumentNullException.ThrowIfNull(celda);

        if (celda.CellFormula is null)
        {
            return;
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"El escritor values-only no puede escribir en '{hoja}!{referencia}' ({operacion}): la celda trae <f> en la plantilla. " +
            "Las fórmulas de la plantilla se conservan intactas por construcción (Plan 58); " +
            "si el contrato necesita un valor ahí, la plantilla cambió de forma.");
    }

    /// <summary>
    /// <see langword="true"/> si la celda destino es efectivamente una fórmula de la plantilla.
    /// </summary>
    internal static bool EsFormula(Cell celda) => celda.CellFormula is not null;
}
