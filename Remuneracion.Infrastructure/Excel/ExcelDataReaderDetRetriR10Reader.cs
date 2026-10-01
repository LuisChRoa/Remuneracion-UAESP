using System.Globalization;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// HU-20 (G3): lector read-only del oráculo <c>R10_Remuneracion_{AAAAMMQ}.xlsx</c>
/// (hoja <c>DetRetri{AAAAMMQ}</c>).
///
/// T0c (evidencia en disco): la hoja trae el encabezado en la fila 8 (<c>B8=ASE</c>,
/// <c>D8=RECAUDO TOTAL</c>); las filas 9..13 son ASE 1..5 y la 14 es <c>Total</c>. Los valores de
/// la columna D están almacenados como TEXTO con formato monetario es-CO (p. ej.
/// <c>"$58.210.094.822"</c>), por eso se normalizan antes de convertirlos a decimal. El rango del
/// período vive en <c>G7</c> ("Fecha Desde") y <c>J7</c> ("Fecha Hasta").
///
/// La fila de encabezado se localiza DINÁMICAMENTE por sus etiquetas (nunca por número fijo).
/// Fail-fast: si la hoja, el encabezado, un ASE del 1..5 o la fila Total faltan, lanza
/// <c>CalculoInvalidoException</c> que nombra el archivo y la hoja/celda — nunca 0 silencioso.
/// El reader NUNCA escribe.
/// </summary>
public sealed class ExcelDataReaderDetRetriR10Reader : IDetRetriR10Reader
{
    /// <inheritdoc />
    public DetRetriInputs LeerDetRetri(Periodo periodo, string rutaR10)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(rutaR10);

        var nombreArchivo = Path.GetFileName(rutaR10);
        var hoja = $"DetRetri{periodo.CodigoCompleto}";
        var filas = ExcelWorksheetNavigator.LeerFilas(rutaR10, hoja);
        if (filas.Count == 0)
        {
            throw new CalculoInvalidoException(
                $"R10 {nombreArchivo}: la hoja '{hoja}' está vacía; no se puede leer el DetRetri del período.");
        }

        var indiceHeader = filas.FindIndex(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1)).Equals("ASE", StringComparison.OrdinalIgnoreCase)
            && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(3))
                .StartsWith("RECAUDO TOTAL", StringComparison.OrdinalIgnoreCase));
        if (indiceHeader < 0)
        {
            throw new CalculoInvalidoException(
                $"R10 {nombreArchivo}: no se encontró el encabezado 'ASE / RECAUDO TOTAL' en la hoja '{hoja}' (T0c: fila 8).");
        }

        var porAse = new Dictionary<int, decimal>();
        decimal? total = null;
        for (var i = indiceHeader + 1; i < filas.Count; i++)
        {
            var fila = filas[i];
            var aseTexto = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(1));
            var concepto = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(2));

            if (int.TryParse(aseTexto, NumberStyles.Integer, CultureInfo.InvariantCulture, out var aseId)
                && aseId is >= 1 and <= 5)
            {
                porAse[aseId] = ParsearNumeroMonetario(fila.ElementAtOrDefault(3));
                continue;
            }

            if (concepto.Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                total = ParsearNumeroMonetario(fila.ElementAtOrDefault(3));
                break;
            }
        }

        var faltantes = Enumerable.Range(1, 5).Where(id => !porAse.ContainsKey(id)).ToArray();
        if (faltantes.Length > 0)
        {
            throw new CalculoInvalidoException(
                $"R10 {nombreArchivo}: la hoja '{hoja}' no trae el DetRetri de los ASE {string.Join(", ", faltantes)} en la columna D (T0c: D9:D13).");
        }

        if (total is null)
        {
            throw new CalculoInvalidoException(
                $"R10 {nombreArchivo}: la hoja '{hoja}' no trae la fila 'Total' del DetRetri (T0c: D14).");
        }

        return new DetRetriInputs
        {
            CodigoRemuneracion = LeerCodigoRemuneracion(filas, periodo),
            FechaDesde = BuscarFecha(filas, "Fecha Desde", columnaEtiqueta: 5, columnaValor: 6) ?? default,
            FechaHasta = BuscarFecha(filas, "Fecha Hasta", columnaEtiqueta: 8, columnaValor: 9) ?? default,
            DetRetriPorAse = porAse,
            Total = total.Value
        };
    }

    /// <summary>
    /// Código de remuneración declarado en la hoja (M3 = "2026071"); si no se puede leer, cae al
    /// código del período (nunca falla por un dato informativo).
    /// </summary>
    private static string LeerCodigoRemuneracion(List<object?[]> filas, Periodo periodo)
    {
        foreach (var fila in filas)
        {
            var etiqueta = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(11));
            if (Normalizar(etiqueta).StartsWith("codigoremuneracion", StringComparison.Ordinal))
            {
                var valor = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(12));
                if (!string.IsNullOrWhiteSpace(valor))
                {
                    return valor;
                }
            }
        }

        return periodo.CodigoCompleto;
    }

    /// <summary>
    /// Busca la celda de fecha cuya etiqueta (columna <paramref name="columnaEtiqueta"/>) sea igual
    /// al texto normalizado esperado y devuelve el valor de la columna contigua
    /// (<paramref name="columnaValor"/>). Devuelve <c>null</c> si no existe/no parsea.
    /// </summary>
    private static DateTime? BuscarFecha(List<object?[]> filas, string etiquetaEsperada, int columnaEtiqueta, int columnaValor)
    {
        var esperado = Normalizar(etiquetaEsperada);
        foreach (var fila in filas)
        {
            if (!Normalizar(ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(columnaEtiqueta)))
                .Equals(esperado, StringComparison.Ordinal))
            {
                continue;
            }

            return ParsearFecha(fila.ElementAtOrDefault(columnaValor));
        }

        return null;
    }

    private static DateTime? ParsearFecha(object? valor)
    {
        if (valor is DateTime fecha)
        {
            return fecha;
        }

        if (valor is double oa)
        {
            return DateTime.FromOADate(oa);
        }

        var texto = ExcelWorksheetNavigator.CeldaTexto(valor);
        if (string.IsNullOrWhiteSpace(texto))
        {
            return null;
        }

        var formatos = new[] { "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy" };
        return DateTime.TryParseExact(texto, formatos, CultureInfo.GetCultureInfo("es-CO"), DateTimeStyles.None, out var exacta)
            ? exacta
            : DateTime.TryParse(texto, CultureInfo.GetCultureInfo("es-CO"), DateTimeStyles.None, out var general) ? general : null;
    }

    /// <summary>
    /// Convierte un valor del R10 a decimal. Los montos vienen como TEXTO con formato monetario
    /// es-CO (<c>"$58.210.094.822"</c>, negativo <c>"-$5.011.515.654"</c>): se quita el símbolo,
    /// se interpreta "." como separador de miles y "," como decimal.
    /// </summary>
    private static decimal ParsearNumeroMonetario(object? valor)
    {
        if (valor is null)
        {
            return 0m;
        }

        if (valor is decimal dec)
        {
            return dec;
        }

        if (valor is double doble)
        {
            return Convert.ToDecimal(doble, CultureInfo.InvariantCulture);
        }

        if (valor is int entero)
        {
            return entero;
        }

        var texto = ExcelWorksheetNavigator.CeldaTexto(valor);
        if (string.IsNullOrWhiteSpace(texto))
        {
            return 0m;
        }

        var negativo = texto.Contains('-', StringComparison.Ordinal);
        var limpio = new string(texto.Where(char.IsDigit).ToArray());
        if (limpio.Length == 0)
        {
            return 0m;
        }

        var numero = decimal.Parse(limpio, NumberStyles.Integer, CultureInfo.InvariantCulture);
        return negativo ? -numero : numero;
    }

    /// <summary>Normaliza una etiqueta (minúsculas, solo letras/dígitos) para el match por título.</summary>
    private static string Normalizar(string texto)
    {
        var sb = new System.Text.StringBuilder(texto.Length);
        foreach (var ch in texto)
        {
            if (char.IsLetterOrDigit(ch))
            {
                sb.Append(char.ToLowerInvariant(ch));
            }
        }

        return sb.ToString();
    }
}
