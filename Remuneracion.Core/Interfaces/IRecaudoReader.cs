using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Define la lectura de los reportes fuente de recaudo (R1, R2, R4) desde archivos Excel.
/// </summary>
public interface IRecaudoReader
{
    /// <summary>
    /// Lee el reporte R1 (Recaudo por componente).
    /// </summary>
    /// <param name="rutaArchivo">Ruta del archivo Excel de origen.</param>
    /// <returns>Datos de recaudo por componente.</returns>
    RecaudoComponenteR1 LeerR1(string rutaArchivo);

    /// <summary>
    /// Lee el reporte R2 (Reporte detalle saldos a favor).
    /// </summary>
    /// <param name="rutaArchivo">Ruta del archivo Excel de origen.</param>
    /// <returns>Datos de saldos a favor.</returns>
    SaldosFavorR2 LeerR2(string rutaArchivo);

    /// <summary>
    /// Lee el reporte R4 (Reversión por componente).
    /// </summary>
    /// <param name="rutaArchivo">Ruta del archivo Excel de origen.</param>
    /// <returns>Datos de reversión.</returns>
    ReversionR4 LeerR4(string rutaArchivo);
}
