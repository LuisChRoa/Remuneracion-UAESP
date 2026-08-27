namespace Remuneracion.Core.Models;

/// <summary>
/// Datos de recaudo por componente extraídos del reporte R1 (Recaudo por componente).
/// </summary>
public class RecaudoComponenteR1
{
    /// <summary>
    /// Nombre del ASE al que corresponden los datos.
    /// </summary>
    public string NombreAse { get; set; } = string.Empty;

    /// <summary>
    /// Valor total oportuno (columna F, fila donde col1="Componente" y col2="Total").
    /// </summary>
    public decimal TotalOportuno { get; set; }

    /// <summary>
    /// Valor extemporáneo (columna F, fila de la sección Extemporáneo donde col2="Mes").
    /// </summary>
    public decimal Extemporaneo { get; set; }

    /// <summary>
    /// Desglose del recaudo por cada componente individual.
    /// </summary>
    public List<ComponenteR1> DetallePorComponente { get; set; } = new();

    /// <summary>
    /// Representa el valor de un componente de recaudo individual dentro del reporte R1.
    /// </summary>
    public class ComponenteR1
    {
        /// <summary>
        /// Nombre del componente.
        /// </summary>
        public string Nombre { get; set; } = string.Empty;

        /// <summary>
        /// Valor recaudado para el componente.
        /// </summary>
        public decimal Valor { get; set; }
    }
}
