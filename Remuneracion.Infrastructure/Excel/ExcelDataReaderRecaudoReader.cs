using System.Globalization;
using System.Text;
using ExcelDataReader;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Lector de los reportes R1, R2 y R4 usando ExcelDataReader.
/// Implementación conforme al Plan 02 — HU-02.
/// </summary>
public class ExcelDataReaderRecaudoReader : IRecaudoReader
{
    static ExcelDataReaderRecaudoReader()
    {
        // ExcelDataReader requiere el registro de encoding providers para leer archivos .xlsx.
        // System.Text.Encoding.CodePages ya es dependencia transitiva de ExcelDataReader 3.9.0.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    /// <inheritdoc/>
    public RecaudoComponenteR1 LeerR1(string rutaArchivo)
    {
        ArgumentNullException.ThrowIfNull(rutaArchivo);

        using var reader = AbrirReader(rutaArchivo);
        var filas = LeerFilas(reader);

        if (filas.Count == 0)
            throw new CalculoInvalidoException(
                $"El archivo '{Path.GetFileName(rutaArchivo)}' no contiene datos.");

        // Buscar TotalOportuno: fila donde A=="Componente" y B=="Total" (case-insensitive)
        var filaTotal = filas.FirstOrDefault(f =>
            CeldaTexto(f.ElementAtOrDefault(0)).Equals("Componente", StringComparison.OrdinalIgnoreCase) &&
            CeldaTexto(f.ElementAtOrDefault(1)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(
                $"No se encontró la fila con A='Componente' y B='Total' en '{Path.GetFileName(rutaArchivo)}'. " +
                $"Filas leídas: {filas.Count}");

        var totalOportuno = CeldaNumero(filaTotal.ElementAtOrDefault(5)); // Columna F (index 5)

        // Buscar Extemporaneo: primera fila donde B=="Mes" (case-insensitive)
        var filaMes = filas.FirstOrDefault(f =>
            CeldaTexto(f.ElementAtOrDefault(1)).Equals("Mes", StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(
                $"No se encontró la fila con B='Mes' (Extemporáneo) en '{Path.GetFileName(rutaArchivo)}'. " +
                $"Filas leídas: {filas.Count}");

        var extemporaneo = CeldaNumero(filaMes.ElementAtOrDefault(5)); // Columna F (index 5)

        var nombreAse = ObtenerNombreAse(rutaArchivo);

        return new RecaudoComponenteR1
        {
            NombreAse = nombreAse,
            TotalOportuno = totalOportuno,
            Extemporaneo = extemporaneo
        };
    }

    /// <inheritdoc/>
    public SaldosFavorR2 LeerR2(string rutaArchivo)
    {
        ArgumentNullException.ThrowIfNull(rutaArchivo);

        using var reader = AbrirReader(rutaArchivo);
        var filas = LeerFilas(reader);

        if (filas.Count == 0)
            throw new CalculoInvalidoException(
                $"El archivo '{Path.GetFileName(rutaArchivo)}' no contiene datos.");

        // Localizar fila de encabezados: primera fila que contiene "Total" y "Componente TDF"
        var filaEncabezados = filas.FirstOrDefault(f =>
            f.Any(c => CeldaTexto(c).Equals("Total", StringComparison.OrdinalIgnoreCase)) &&
            f.Any(c => CeldaTexto(c).Equals("Componente TDF", StringComparison.OrdinalIgnoreCase)))
            ?? throw new CalculoInvalidoException(
                $"No se encontró la fila de encabezados con 'Total' y 'Componente TDF' en '{Path.GetFileName(rutaArchivo)}'. " +
                $"Filas leídas: {filas.Count}");

        // Detectar columna "Especiales" (case-insensitive, trimmed)
        int indiceEspeciales = -1;
        for (int i = 0; i < filaEncabezados.Length; i++)
        {
            if (CeldaTexto(filaEncabezados[i]).Equals("Especiales", StringComparison.OrdinalIgnoreCase))
            {
                indiceEspeciales = i;
                break;
            }
        }

        bool tieneColumnaEspeciales = indiceEspeciales >= 0;

        // Buscar GrandTotal: fila donde A=="Total" y B está vacío (case-insensitive)
        var filaGrandTotal = filas.FirstOrDefault(f =>
            CeldaTexto(f.ElementAtOrDefault(0)).Equals("Total", StringComparison.OrdinalIgnoreCase) &&
            string.IsNullOrWhiteSpace(CeldaTexto(f.ElementAtOrDefault(1))))
            ?? throw new CalculoInvalidoException(
                $"No se encontró la fila con A='Total' y B vacío (GrandTotal) en '{Path.GetFileName(rutaArchivo)}'. " +
                $"Filas leídas: {filas.Count}");

        var grandTotal = CeldaNumero(filaGrandTotal.ElementAtOrDefault(4)); // Columna E (index 4)

        // ServEspK: valor de la columna "Especiales" en la fila GrandTotal (0 si no existe)
        decimal servEspK = 0;
        if (tieneColumnaEspeciales && indiceEspeciales < filaGrandTotal.Length)
        {
            servEspK = CeldaNumero(filaGrandTotal[indiceEspeciales]);
        }

        var nombreAse = ObtenerNombreAse(rutaArchivo);

        return new SaldosFavorR2
        {
            NombreAse = nombreAse,
            GrandTotal = grandTotal,
            TieneColumnaEspeciales = tieneColumnaEspeciales,
            ServEspK = servEspK
        };
    }

    /// <inheritdoc/>
    public ReversionR4 LeerR4(string rutaArchivo)
    {
        ArgumentNullException.ThrowIfNull(rutaArchivo);

        using var reader = AbrirReader(rutaArchivo);
        var filas = LeerFilas(reader);

        if (filas.Count == 0)
            throw new CalculoInvalidoException(
                $"El archivo '{Path.GetFileName(rutaArchivo)}' no contiene datos.");

        // Buscar TotalReversiones: fila donde A=="Total" y B está vacío (case-insensitive)
        var filaTotal = filas.FirstOrDefault(f =>
            CeldaTexto(f.ElementAtOrDefault(0)).Equals("Total", StringComparison.OrdinalIgnoreCase) &&
            string.IsNullOrWhiteSpace(CeldaTexto(f.ElementAtOrDefault(1))))
            ?? throw new CalculoInvalidoException(
                $"No se encontró la fila con A='Total' y B vacío (TotalReversiones) en '{Path.GetFileName(rutaArchivo)}'. " +
                $"Filas leídas: {filas.Count}");

        var totalReversiones = CeldaNumero(filaTotal.ElementAtOrDefault(3)); // Columna D (index 3) — ya negativo

        var nombreAse = ObtenerNombreAse(rutaArchivo);

        return new ReversionR4
        {
            NombreAse = nombreAse,
            TotalReversiones = totalReversiones
        };
    }

    #region Helpers privados (§2.2.1)

    /// <summary>
    /// Abre un archivo Excel y retorna un IExcelDataReader.
    /// Valida existencia del archivo y lanza ArchivoFuenteNoEncontradoException si no existe.
    /// </summary>
    private static IExcelDataReader AbrirReader(string rutaArchivo)
    {
        if (!File.Exists(rutaArchivo))
            throw new ArchivoFuenteNoEncontradoException(
                $"No se encontró el archivo fuente: '{rutaArchivo}'. Verifique que la ruta sea correcta.");

        var stream = File.Open(rutaArchivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ExcelReaderFactory.CreateReader(stream);
    }

    /// <summary>
    /// Lee la primera hoja del reader y retorna todas las filas como lista de celdas.
    /// Cada fila es un object?[] de tamaño igual al número de columnas del reader.
    /// </summary>
    private static List<object?[]> LeerFilas(IExcelDataReader reader)
    {
        var filas = new List<object?[]>();

        // Solo leer la primera hoja (ExcelDataReader arranca posicionado en la primera).
        while (reader.Read())
        {
            var fila = new object?[reader.FieldCount];
            for (int i = 0; i < reader.FieldCount; i++)
            {
                fila[i] = reader.GetValue(i);
            }
            filas.Add(fila);
        }

        return filas;
    }

    /// <summary>
    /// Normaliza el valor de una celda a string: null → "", string → Trim(), otro → Convert.ToString.
    /// </summary>
    private static string CeldaTexto(object? valor)
    {
        if (valor is null)
            return string.Empty;
        if (valor is string texto)
            return texto.Trim();
        return Convert.ToString(valor, CultureInfo.InvariantCulture) ?? string.Empty;
    }

    /// <summary>
    /// Convierte el valor de una celda a decimal de forma culture-safe.
    /// null o "" → 0m; double/decimal → Convert.ToDecimal(InvariantCulture).
    /// </summary>
    private static decimal CeldaNumero(object? valor)
    {
        if (valor is null)
            return 0m;
        if (valor is string texto)
        {
            texto = texto.Trim();
            if (string.IsNullOrEmpty(texto))
                return 0m;
            return Convert.ToDecimal(texto, CultureInfo.InvariantCulture);
        }
        if (valor is double d)
            return Convert.ToDecimal(d, CultureInfo.InvariantCulture);
        if (valor is decimal dec)
            return dec;
        return Convert.ToDecimal(valor, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Obtiene el nombre del ASE a partir de la ruta del archivo (carpeta padre).
    /// Patrón esperado: "N-NombreASE". Fallback: nombre base del archivo sin extensión.
    /// </summary>
    private static string ObtenerNombreAse(string rutaArchivo)
    {
        var carpetaPadre = Path.GetDirectoryName(rutaArchivo);
        if (!string.IsNullOrEmpty(carpetaPadre))
        {
            var nombreCarpeta = Path.GetFileName(carpetaPadre);
            if (!string.IsNullOrEmpty(nombreCarpeta))
                return nombreCarpeta;
        }
        return Path.GetFileNameWithoutExtension(rutaArchivo);
    }

    #endregion
}
