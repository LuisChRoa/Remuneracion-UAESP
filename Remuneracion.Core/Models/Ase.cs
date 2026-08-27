namespace Remuneracion.Core.Models;

/// <summary>
/// Representa una de las cinco empresas operadoras de aseo (ASE) de Bogotá.
/// </summary>
public class Ase
{
    /// <summary>
    /// Identificador numérico del ASE (1 a 5).
    /// </summary>
    public int Id { get; set; }

    /// <summary>
    /// Nombre corto o clave del ASE (por ejemplo "PROMOAMBIENTAL").
    /// </summary>
    public string NombreCorto { get; set; } = string.Empty;

    /// <summary>
    /// Nombre completo y legible del ASE (por ejemplo "Promoambiental").
    /// </summary>
    public string NombreCompleto { get; set; } = string.Empty;

    /// <summary>
    /// Número de carpeta asignado al ASE (1 a 5) dentro de la estructura de fuentes.
    /// </summary>
    public int NumeroCarpeta { get; set; }
}
