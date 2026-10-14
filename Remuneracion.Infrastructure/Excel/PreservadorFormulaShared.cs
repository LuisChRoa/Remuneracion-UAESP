using DocumentFormat.OpenXml.Spreadsheet;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 33 (T2, D-A/D-B) — doctrina de preservación de <c>shared formula</c>, <b>intacta</b> tras el
/// retiro del motor geométrico del espejo R1 (Plan 36/T2, R-B-3).
///
/// Escribe el TEXTO de una fórmula preservando los atributos shared (<c>t</c>/<c>ref</c>/<c>si</c>)
/// del master de la plantilla cuando la celda YA traía <see cref="CellFormula"/>; solo crea un
/// <c>new CellFormula</c> cuando la celda nace sin fórmula (literales). La regla es por EXISTENCIA,
/// no por lista de celdas: un master nuevo futuro se preserva sin código.
///
/// Nota de vecindad (Plan 36/T2): el path de producción R1 (<see cref="EscritorValoresSlotR1"/>)
/// NO escribe <c>&lt;f&gt;</c> — pega solo valores y las fórmulas del template quedan byte-idénticas,
/// por lo que esta costura ya no tiene consumidor en el flujo R1; se conserva como la disciplina
/// shared única del proyecto (cualquier path futuro que mute texto de fórmula la reutiliza) y la
/// audita el gate <c>SharedFormulaR1GateTests</c>.
/// </summary>
internal static class PreservadorFormulaShared
{
    /// <summary>
    /// Escribe <paramref name="texto"/> como texto de fórmula de <paramref name="celda"/> sin perder
    /// los atributos shared del master existente (mutación in-place) y creando la fórmula plana
    /// cuando la celda no traía ninguna.
    /// </summary>
    internal static void EscribirPreservando(Cell celda, string texto)
    {
        ArgumentNullException.ThrowIfNull(celda);
        ArgumentNullException.ThrowIfNull(texto);

        if (celda.CellFormula is not null)
        {
            celda.CellFormula.Text = texto;
        }
        else
        {
            celda.CellFormula = new CellFormula(texto);
        }
    }
}
