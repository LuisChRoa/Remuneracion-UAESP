using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Define la escritura de resultados en la plantilla Excel de salida.
/// </summary>
public interface IPlantillaWriter
{
    /// <summary>
    /// Escribe el consolidado global en la plantilla.
    /// </summary>
    /// <param name="rutaPlantilla">Ruta de la plantilla Excel.</param>
    /// <param name="resultado">Resultado de la remuneración.</param>
    void EscribirConsolidado(string rutaPlantilla, ResultadoRemuneracion resultado);

    /// <summary>
    /// Escribe el detalle del reporte R1 para un ASE en la plantilla.
    /// </summary>
    /// <param name="rutaPlantilla">Ruta de la plantilla Excel.</param>
    /// <param name="ase">ASE correspondiente.</param>
    /// <param name="datos">Datos del reporte R1.</param>
    void EscribirDetalleR1(string rutaPlantilla, Ase ase, RecaudoComponenteR1 datos);

    /// <summary>
    /// Escribe el detalle del reporte R2 para un ASE en la plantilla.
    /// </summary>
    /// <param name="rutaPlantilla">Ruta de la plantilla Excel.</param>
    /// <param name="ase">ASE correspondiente.</param>
    /// <param name="datos">Datos del reporte R2.</param>
    void EscribirDetalleR2(string rutaPlantilla, Ase ase, SaldosFavorR2 datos);

    /// <summary>
    /// Escribe el detalle del reporte R4 para un ASE en la plantilla.
    /// </summary>
    /// <param name="rutaPlantilla">Ruta de la plantilla Excel.</param>
    /// <param name="ase">ASE correspondiente.</param>
    /// <param name="datos">Datos del reporte R4.</param>
    void EscribirDetalleR4(string rutaPlantilla, Ase ase, ReversionR4 datos);
}
