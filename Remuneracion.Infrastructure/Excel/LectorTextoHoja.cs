using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 59 (H1.1): lectura de TEXTO de celda compartida por el delimitador de bloques
/// (<see cref="DelimitadorBloquesHoja"/>) y el escritor values-only
/// (<see cref="EscritorValoresSlotR1"/>).
///
/// Única implementación de la regla de strings compartidos / inline strings: las etiquetas que
/// delimitan y las que resuelven destino son DATOS leídos del workbook, así que ambas rutas deben
/// leerlas exactamente igual (duplicar la regla permitiría que una etiqueta se interprete distinto
/// en la delimitación que en la escritura). Cero coordenadas: recibe la celda o la referencia ya
/// calculada por el llamador.
/// </summary>
internal static class LectorTextoHoja
{
    /// <summary>
    /// Texto de la celda <paramref name="referencia"/> (p. ej. la que arma el llamador a partir de su
    /// columna y su fila) dentro de <paramref name="fila"/>; vacío si la celda no existe.
    /// </summary>
    internal static string Texto(WorkbookPart parte, Row fila, string referencia)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(fila);

        var celda = fila.Elements<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        return celda is null ? string.Empty : Texto(parte, celda);
    }

    /// <summary>
    /// Texto de la celda de la <paramref name="columna"/> en la propia fila (la fila aporta su
    /// índice; el llamador no escribe ninguna fila literal).
    /// </summary>
    internal static string TextoDeColumna(WorkbookPart parte, Row fila, string columna)
    {
        ArgumentNullException.ThrowIfNull(fila);
        return fila.RowIndex is null
            ? string.Empty
            : Texto(parte, fila, $"{columna}{fila.RowIndex.Value.ToString(CultureInfo.InvariantCulture)}");
    }

    /// <summary>Texto de una celda ya localizada (shared string, inline string o valor simple).</summary>
    internal static string Texto(WorkbookPart parte, Cell celda)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(celda);

        if (celda.DataType?.Value == CellValues.SharedString
            && celda.CellValue is not null
            && int.TryParse(celda.CellValue.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indice))
        {
            var item = parte.SharedStringTablePart?.SharedStringTable?
                .Elements<SharedStringItem>()
                .ElementAtOrDefault(indice);
            return item?.InnerText ?? string.Empty;
        }

        if (celda.DataType?.Value == CellValues.InlineString)
        {
            return celda.InlineString?.InnerText ?? string.Empty;
        }

        return celda.CellValue?.InnerText ?? string.Empty;
    }
}
