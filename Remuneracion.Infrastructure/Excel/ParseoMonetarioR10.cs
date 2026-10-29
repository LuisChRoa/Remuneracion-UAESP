using System.Globalization;
using System.Text;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 62 (T1.2, D-PARSE-UNICO): conversor ÚNICO de los montos del insumo
/// <c>R10_Remuneracion_{AAAAMMQ}.xlsx</c>.
///
/// <para>
/// <b>Por qué vive en Infrastructure (y no en Core):</b> el dato de entrada es la representación de
/// CELDA que entrega <see cref="ExcelWorksheetNavigator.CeldaTexto"/> al leer el workbook
/// (ExcelDataReader). Llevarlo a Core obligaría a duplicar esa normalización de celda en el dominio,
/// que debe quedar sin convenciones de Excel. Es puro en el sentido de D-H1.3: estático, sin I/O,
/// sin estado y sin dependencias de ExcelDataReader más allá del helper de texto ya existente.
/// </para>
///
/// <para>
/// <b>Semántica decidida (resuelve DOCUMENTO-vs-CÓDIGO del hallazgo §T0/T1.2):</b> el R10 guarda los
/// montos como TEXTO con la convención es-CO — <c>'.'</c> es separador de MILES y <c>','</c> es el
/// separador DECIMAL (p. ej. <c>"$1.234,56"</c> = 1.234,56). La implementación anterior documentaba
/// esa convención pero ejecutaba <c>texto.Where(char.IsDigit)</c>, con lo que <c>"$1.234,56"</c> se
/// leía como <c>123456</c> (×100 y sin decimales). Aquí la coma SÍ es decimal y el punto se
/// descarta como separador de miles, que es lo que el documento y todos los insumos del período
/// traen. Formatos con OTRO convenio (p. ej. <c>"1,234.56"</c> anglosajón) no se adivinan: se
/// rechazan (<see cref="TryParseTexto"/> = <c>false</c>) en vez de producir un monto silenciosamente
/// equivocado. Cita: <c>Docs/T0-Plan62-Pegado-R10.md</c> §6 (hallazgo T1.2) y Plan 62 §11
/// ("Riesgo operativo añadido a H1").
/// </para>
///
/// <para>
/// Dos entradas, por diseño: <see cref="TryConvertir"/> es el modo ESTRICTO que consume el lector de
/// bloque (una celda ilegible debe fallar antes de copiar nada) y <see cref="Convertir"/> conserva la
/// semántica histórica del lector del oráculo (celda vacía o ilegible = 0), para no mover un solo
/// valor de los goldens de julio.
/// </para>
/// </summary>
public static class ParseoMonetarioR10
{
    /// <summary>
    /// Convierte el valor de una celda del R10 a decimal con la semántica histórica del oráculo:
    /// celda vacía o ilegible ⇒ <c>0</c>.
    /// </summary>
    /// <param name="valor">Valor de celda (texto monetario, numérico o nulo).</param>
    /// <returns>El monto, o <c>0</c> si la celda está vacía o no es un monto legible.</returns>
    public static decimal Convertir(object? valor) => TryConvertir(valor, out var numero) ? numero : 0m;

    /// <summary>
    /// Intenta convertir el valor de una celda del R10 a decimal SIN redondear.
    /// </summary>
    /// <param name="valor">Valor de celda (texto monetario, numérico o nulo).</param>
    /// <param name="numero">Monto exacto cuando el valor es legible.</param>
    /// <returns>
    /// <c>true</c> si el valor es numérico o un texto monetario legible; <c>false</c> si la celda
    /// está vacía o el texto no es un monto reconocible.
    /// </returns>
    public static bool TryConvertir(object? valor, out decimal numero)
    {
        numero = 0m;

        switch (valor)
        {
            case null:
                return false;
            case decimal dec:
                numero = dec;
                return true;
            case int entero:
                numero = entero;
                return true;
            case long largo:
                numero = largo;
                return true;
            case short corto:
                numero = corto;
                return true;
            case byte octeto:
                numero = octeto;
                return true;
            case double doble:
                numero = Convert.ToDecimal(doble, CultureInfo.InvariantCulture);
                return true;
            case float flotante:
                numero = Convert.ToDecimal(flotante, CultureInfo.InvariantCulture);
                return true;
        }

        return TryParseTexto(ExcelWorksheetNavigator.CeldaTexto(valor), out numero);
    }

    /// <summary>
    /// Interpreta un texto monetario es-CO (<c>"$17.450.250.543"</c>, <c>"-$5.093.259.797"</c>,
    /// <c>"$1.234,56"</c>, <c>"1234"</c>) como decimal exacto, sin redondear.
    /// </summary>
    /// <param name="texto">Texto de la celda.</param>
    /// <param name="numero">Monto exacto cuando el texto es legible.</param>
    /// <returns><c>true</c> si el texto contiene un monto reconocible.</returns>
    public static bool TryParseTexto(string? texto, out decimal numero)
    {
        numero = 0m;

        if (string.IsNullOrWhiteSpace(texto))
        {
            return false;
        }

        // Plan 62 (N-1, hardening del signo): el '-' SOLO es valido en POSICION INICIAL (al frente o
        // detras del simbolo monetario). Un texto con guiones "laterales" (separadores de fecha
        // dd-MM-yyyy, prefijos tipo "R10-2026091", rangos "01-09") NO es un monto: se RECHAZA
        // (false) en vez de adivinar un negativo. Antes bastaba un Contains('-') para negar
        // cualquier texto, asi que "01-09" se leia como -109 (monto inventado).
        var sinEspacios = new StringBuilder(texto.Length);
        foreach (var caracter in texto)
        {
            if (!char.IsWhiteSpace(caracter))
            {
                sinEspacios.Append(caracter);
            }
        }

        var limpio = sinEspacios.ToString();
        var negativo = false;

        if (limpio.StartsWith("-$", StringComparison.Ordinal))
        {
            negativo = true;
            limpio = limpio[2..];
        }
        else if (limpio.StartsWith('-'))
        {
            negativo = true;
            limpio = limpio[1..];
        }
        else if (limpio.StartsWith('$'))
        {
            limpio = limpio[1..];
            if (limpio.StartsWith('-'))
            {
                negativo = true;
                limpio = limpio[1..];
            }
        }

        // Convencion es-CO: '.' es separador de MILES (se descarta) y ',' es el separador DECIMAL.
        // Cualquier otro caracter (un separador de fecha '/', letras, un segundo signo, parentesis)
        // hace que el texto NO sea un monto reconocible: no se adivina.
        // Un texto con la coma ANTES que el punto ("1,234.56") esta en el convenio anglosajon: no
        // se adivina, se rechaza (nunca un monto x1000 en silencio).
        var posicionComa = limpio.IndexOf(',', StringComparison.Ordinal);
        var posicionPunto = limpio.IndexOf('.', StringComparison.Ordinal);
        if (posicionComa >= 0 && posicionPunto > posicionComa)
        {
            return false;
        }

        var soloDigitos = new StringBuilder(limpio.Length);
        foreach (var caracter in limpio)
        {
            if (char.IsAsciiDigit(caracter))
            {
                soloDigitos.Append(caracter);
            }
            else if (caracter == ',')
            {
                soloDigitos.Append('.');
            }
            else if (caracter != '.' && caracter != '$')
            {
                return false;
            }
        }

        if (soloDigitos.Length == 0
            || !decimal.TryParse(soloDigitos.ToString(), NumberStyles.Number, CultureInfo.InvariantCulture, out var valor))
        {
            return false;
        }

        numero = negativo ? -valor : valor;
        return true;
    }
}
