using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 28 (Unidad S, Alternativa A): sanea la cadena de cálculo del workbook de salida.
///
/// El espejo estructural R1 (<see cref="OpenXmlEspejoR1Mutador"/>) inserta/borra filas pero la
/// <c>CalculationChainPart</c> de la plantilla se preserva intacta; al abrir la salida Excel
/// detecta entradas de la cadena que ya no corresponden a fórmulas (agosto: 1003 entradas
/// inconsistentes) y muestra el diálogo de reparación. Saneamiento:
///   1. elimina <c>CalculationChainPart</c> si existe (tolerante: la plantilla puede no traerla);
///   2. setea <c>FullCalculationOnLoad=true</c> en <c>CalculationProperties</c> para que Excel
///      recalcule una vez al abrir con las fórmulas intactas.
///
/// NUNCA toca <c>&lt;f&gt;</c> ni <c>&lt;v&gt;</c>: los valores cacheados quedan y el recálculo
/// lo dispara Excel (fullCalcOnLoad). Se invoca en el punto único de guardado: los 2 <c>Save()</c>
/// de <see cref="OpenXmlPlantillaWriter.GenerarWorkbook(string, string, Remuneracion.Core.Models.ResultadoRemuneracion, Remuneracion.Core.Models.WorkbookLeafInputs)"/>
/// y el guardado del espejo standalone en <see cref="OpenXmlEspejoR1Mutador.Ajustar"/>.
/// </summary>
internal static class SaneadorCadenaCalculo
{
    /// <summary>
    /// Aplica el saneamiento sobre un <see cref="WorkbookPart"/> abierto en modo escritura.
    /// Idempotente: sin <c>CalculationChainPart</c> solo asegura el flag.
    /// </summary>
    public static void Sanear(WorkbookPart workbookPart)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);

        // Tolerante a plantillas sin cadena: borra si existe, nunca falla por ausencia.
        if (workbookPart.CalculationChainPart is not null)
        {
            workbookPart.DeletePart(workbookPart.CalculationChainPart);
        }

        var workbook = workbookPart.Workbook
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                "El workbook a sanear no tiene metadata Workbook válida.");

        // Preserva calcId si la plantilla trae <calcPr>; crea el elemento solo si faltara.
        var propiedades = workbook.CalculationProperties;
        if (propiedades is null)
        {
            propiedades = new CalculationProperties();
            workbook.CalculationProperties = propiedades;
        }

        propiedades.FullCalculationOnLoad = true;
    }
}
