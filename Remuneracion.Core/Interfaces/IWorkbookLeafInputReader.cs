using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Contrato para poblar los inputs leaf reales del workbook a partir de las fuentes R1/R2/R4.
/// </summary>
public interface IWorkbookLeafInputReader
{
    /// <summary>
    /// Lee los inputs leaf editables del workbook para el ASE y periodo indicados.
    /// </summary>
    /// <param name="ase">ASE asociado al cálculo.</param>
    /// <param name="periodo">Periodo de la liquidación.</param>
    /// <param name="rutaR1">Ruta del archivo fuente R1.</param>
    /// <param name="rutaR2">Ruta del archivo fuente R2.</param>
    /// <param name="rutaR4">Ruta del archivo fuente R4.</param>
    /// <returns>Modelo <see cref="WorkbookLeafInputs"/> listo para la escritura real del workbook.</returns>
    WorkbookLeafInputs LeerLeafInputs(Ase ase, Periodo periodo, string rutaR1, string rutaR2, string rutaR4);
}
