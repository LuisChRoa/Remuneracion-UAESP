using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Escritor de resultados en la plantilla Excel usando OpenXML SDK.
/// Implementación pendiente: requiere archivo de plantilla real.
/// </summary>
public class OpenXmlPlantillaWriter : IPlantillaWriter
{
    /// <inheritdoc/>
    public void EscribirConsolidado(string rutaPlantilla, ResultadoRemuneracion resultado)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(resultado);
        throw new NotImplementedException(
            "Pendiente: implementar escritura en plantilla con OpenXML SDK. Requiere archivo de plantilla real.");
    }

    /// <inheritdoc/>
    public void EscribirDetalleR1(string rutaPlantilla, Ase ase, RecaudoComponenteR1 datos)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(datos);
        throw new NotImplementedException(
            "Pendiente: implementar escritura en plantilla con OpenXML SDK. Requiere archivo de plantilla real.");
    }

    /// <inheritdoc/>
    public void EscribirDetalleR2(string rutaPlantilla, Ase ase, SaldosFavorR2 datos)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(datos);
        throw new NotImplementedException(
            "Pendiente: implementar escritura en plantilla con OpenXML SDK. Requiere archivo de plantilla real.");
    }

    /// <inheritdoc/>
    public void EscribirDetalleR4(string rutaPlantilla, Ase ase, ReversionR4 datos)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(datos);
        throw new NotImplementedException(
            "Pendiente: implementar escritura en plantilla con OpenXML SDK. Requiere archivo de plantilla real.");
    }
}
