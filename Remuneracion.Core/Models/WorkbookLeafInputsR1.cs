namespace Remuneracion.Core.Models;

/// <summary>
/// Inputs editables reales de la hoja <c>Reporte Componentes R1</c>.
/// Las celdas visibles <c>F46</c> y <c>F48</c> son fórmulas del workbook; los valores de detalle
/// que las alimentan se persisten en las celdas leaf que componen esas fórmulas.
/// </summary>
public sealed class WorkbookLeafInputsR1
{
    /// <summary>
    /// Valor de la celda leaf <c>F25</c>: primera fila <c>Mes/Total</c> col F de la fuente R1
    /// (coincide con <see cref="RecaudoComponenteR1.Extemporaneo"/>). Alimenta <c>F46 = F25 + F41 - L25</c>.
    /// </summary>
    public decimal F25 { get; set; }

    /// <summary>
    /// Valor de la celda leaf <c>F41</c>: segunda fila <c>Mes/Total</c> col F (Subs/Cont de Mes).
    /// Alimenta <c>F46 = F25 + F41 - L25</c>.
    /// </summary>
    public decimal F41 { get; set; }

    /// <summary>
    /// Valor de la celda leaf <c>L25</c>: especiales de la primera fila <c>Mes/Total</c>.
    /// Se resta en <c>F46 = F25 + F41 - L25</c>.
    /// </summary>
    public decimal L25 { get; set; }

    /// <summary>
    /// Valor de la celda leaf <c>F30</c>: primera fila <c>Subsidio(-)/Contribucion(+)</c> col F.
    /// Alimenta <c>F48 = F30 + F10 - L10</c>.
    /// </summary>
    public decimal F30 { get; set; }

    /// <summary>
    /// Valor de la celda leaf <c>F10</c>: fila <c>Aplicacion nuevos x reversion</c> col F.
    /// Alimenta <c>F48 = F30 + F10 - L10</c>.
    /// </summary>
    public decimal F10 { get; set; }

    /// <summary>
    /// Valor de la celda leaf <c>L10</c>. Se resta en <c>F48 = F30 + F10 - L10</c>.
    /// En la plantilla de referencia Promoambiental 202607-1 vale 0.
    /// </summary>
    public decimal L10 { get; set; }

    /// <summary>
    /// Visible esperado de <c>Reporte Componentes R1!F46</c> / consolidado D9:
    /// <c>F46 = F25 + F41 - L25</c>. No coincide con <see cref="RecaudoComponenteR1.TotalOportuno"/>.
    /// </summary>
    public decimal TotalOportunoEsperado => F25 + F41 - L25;

    /// <summary>
    /// Visible esperado de <c>Reporte Componentes R1!F48</c> / consolidado D47:
    /// <c>F48 = F30 + F10 - L10</c>. No coincide con <see cref="RecaudoComponenteR1.Extemporaneo"/>.
    /// </summary>
    public decimal ExtemporaneoEsperado => F30 + F10 - L10;
}
