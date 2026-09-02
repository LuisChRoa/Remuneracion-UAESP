using System.Globalization;
using System.Text;
using ExcelDataReader;
using Remuneracion.Core.Exceptions;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Helper compartido para leer y localizar celdas en archivos Excel de la solución.
/// Centraliza la lógica de resolución de filas, textos y números para los readers de HU-02 y HU-05.
/// </summary>
public static class ExcelWorksheetNavigator
{
    static ExcelWorksheetNavigator()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public static List<object?[]> LeerFilas(string rutaArchivo)
    {
        ArgumentNullException.ThrowIfNull(rutaArchivo);

        if (!File.Exists(rutaArchivo))
        {
            throw new ArchivoFuenteNoEncontradoException($"No se encontró el archivo fuente: '{rutaArchivo}'.");
        }

        using var stream = File.Open(rutaArchivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var filas = new List<object?[]>();

        while (reader.Read())
        {
            var fila = new object?[reader.FieldCount];
            for (var i = 0; i < reader.FieldCount; i++)
            {
                fila[i] = reader.GetValue(i);
            }

            filas.Add(fila);
        }

        return filas;
    }

    public static string CeldaTexto(object? valor)
    {
        if (valor is null)
        {
            return string.Empty;
        }

        if (valor is string texto)
        {
            return texto.Trim();
        }

        return Convert.ToString(valor, CultureInfo.InvariantCulture) ?? string.Empty;
    }

    public static decimal CeldaNumero(object? valor)
    {
        if (valor is null)
        {
            return 0m;
        }

        if (valor is string texto)
        {
            texto = texto.Trim();
            if (string.IsNullOrEmpty(texto))
            {
                return 0m;
            }

            return Convert.ToDecimal(texto, CultureInfo.InvariantCulture);
        }

        if (valor is double d)
        {
            return Convert.ToDecimal(d, CultureInfo.InvariantCulture);
        }

        if (valor is decimal dec)
        {
            return dec;
        }

        return Convert.ToDecimal(valor, CultureInfo.InvariantCulture);
    }

    public static string ObtenerNombreAse(string rutaArchivo)
    {
        var carpetaPadre = Path.GetDirectoryName(rutaArchivo);
        if (!string.IsNullOrEmpty(carpetaPadre))
        {
            var nombreCarpeta = Path.GetFileName(carpetaPadre);
            if (!string.IsNullOrEmpty(nombreCarpeta))
            {
                return nombreCarpeta;
            }
        }

        return Path.GetFileNameWithoutExtension(rutaArchivo);
    }

    public static int IndiceColumnaPorEncabezado(IEnumerable<object?[]> filas, string encabezado)
    {
        ArgumentNullException.ThrowIfNull(filas);
        ArgumentNullException.ThrowIfNull(encabezado);

        foreach (var fila in filas)
        {
            for (var i = 0; i < fila.Length; i++)
            {
                if (CeldaTexto(fila[i]).Equals(encabezado, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
        }

        return -1;
    }

    public static List<object?[]> FiltrarFilas(IEnumerable<object?[]> filas, Func<object?[], bool> predicado)
    {
        ArgumentNullException.ThrowIfNull(filas);
        ArgumentNullException.ThrowIfNull(predicado);
        return filas.Where(predicado).ToList();
    }
}
