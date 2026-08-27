using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Lector de los reportes R1, R2 y R4 usando ExcelDataReader.
/// Implementación pendiente: requiere archivos de plantilla y fuentes de ejemplo.
/// </summary>
public class ExcelDataReaderRecaudoReader : IRecaudoReader
{
    /// <inheritdoc/>
    public RecaudoComponenteR1 LeerR1(string rutaArchivo)
    {
        ArgumentNullException.ThrowIfNull(rutaArchivo);
        throw new NotImplementedException(
            "Pendiente: implementar lectura con datos reales de R1. Requiere archivo de plantilla y fuentes de ejemplo.");
    }

    /// <inheritdoc/>
    public SaldosFavorR2 LeerR2(string rutaArchivo)
    {
        ArgumentNullException.ThrowIfNull(rutaArchivo);
        throw new NotImplementedException(
            "Pendiente: implementar lectura con datos reales de R2. Requiere archivo de plantilla y fuentes de ejemplo.");
    }

    /// <inheritdoc/>
    public ReversionR4 LeerR4(string rutaArchivo)
    {
        ArgumentNullException.ThrowIfNull(rutaArchivo);
        throw new NotImplementedException(
            "Pendiente: implementar lectura con datos reales de R4. Requiere archivo de plantilla y fuentes de ejemplo.");
    }
}
