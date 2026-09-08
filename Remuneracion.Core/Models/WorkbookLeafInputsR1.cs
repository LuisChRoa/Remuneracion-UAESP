namespace Remuneracion.Core.Models;

/// <summary>
/// Inputs editables reales de la hoja <c>Reporte Componentes R1</c>.
/// Las celdas visibles <c>F46</c> y <c>F48</c> son fórmulas del workbook; los valores de detalle
/// que las alimentan se persisten en las celdas leaf que componen esas fórmulas.
///
/// HU-07 (multi-ASE): cada bloque del template tiene operandos en direcciones distintas
/// (congeladas por T0 en <c>plans/07 - HU-07 T0 Evidencia.md</c>). Las propiedades ASE1
/// (F25/F41/L25/F30/F10/L10) se conservan para compatibilidad single-ASE; los operandos por
/// bloque se exponen en <see cref="CeldasPorAse"/> y los visibles esperados por bloque en
/// <see cref="TotalOportunoEsperadoPorAse"/> y <see cref="ExtemporaneoEsperadoPorAse"/>.
/// </summary>
public sealed class WorkbookLeafInputsR1
{
    /// <summary>
    /// Valor de la celda leaf <c>F25</c>: primera fila <c>Mes/Total</c> col F de la fuente R1
    /// (coincide con <see cref="RecaudoComponenteR1.Extemporaneo"/>). Alimenta <c>F46 = F25 + F41 - L25</c>.
    /// En bloques 2..5 es el equivalente "primera fila Mes/Total" (F90/F217/F357/F478).
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
    /// Valor de la celda leaf <c>F30</c>: primer operando extemporáneo col F.
    /// Alimenta <c>F48 = F30 + F10 - L10</c>. Por bloque: ASE1/ASE3 = Subsidio[0], ASE4 = Aplicacion[1].
    /// </summary>
    public decimal F30 { get; set; }

    /// <summary>
    /// Valor de la celda leaf <c>F10</c>: segundo operando extemporáneo col F.
    /// Alimenta <c>F48 = F30 + F10 - L10</c>. Por bloque: ASE1/ASE3 = Aplicacion[0], ASE4 = Aplicacion[0].
    /// </summary>
    public decimal F10 { get; set; }

    /// <summary>
    /// Valor de la celda leaf <c>L10</c>. Se resta en <c>F48 = F30 + F10 - L10</c>.
    /// En la plantilla de referencia Promoambiental 202607-1 vale 0.
    /// </summary>
    public decimal L10 { get; set; }

    /// <summary>
    /// Visible esperado de <c>Reporte Componentes R1!F46</c> / consolidado D9 (ASE1):
    /// <c>F46 = F25 + F41 - L25</c>. No coincide con <see cref="RecaudoComponenteR1.TotalOportuno"/>.
    /// </summary>
    public decimal TotalOportunoEsperado => F25 + F41 - L25;

    /// <summary>
    /// Visible esperado de <c>Reporte Componentes R1!F48</c> / consolidado D47 (ASE1):
    /// <c>F48 = F30 + F10 - L10</c>. No coincide con <see cref="RecaudoComponenteR1.Extemporaneo"/>.
    /// </summary>
    public decimal ExtemporaneoEsperado => F30 + F10 - L10;

    /// <summary>
    /// Operandos leaf del bloque R1 por ASE, keyed por referencia de celda del template
    /// (ej. "F113", "L113", "F90"). Congelado por T0; el writer escribe SOLO estas celdas.
    /// </summary>
    public IReadOnlyDictionary<string, decimal> CeldasPorAse { get; set; } = new Dictionary<string, decimal>();

    /// <summary>
    /// Visible esperado del TOT_OPT por bloque (fórmula T0 §0.3 de <c>plans/07 - HU-07 T0 Evidencia.md</c>).
    /// Para ASE1 coincide con <see cref="TotalOportunoEsperado"/>.
    /// </summary>
    public decimal TotalOportunoEsperadoPorAse { get; set; }

    /// <summary>
    /// Visible esperado del EXTEMP por bloque (fórmula T0 §0.3; 0 para ASE2/ASE5 donde es valor estático).
    /// Para ASE1 coincide con <see cref="ExtemporaneoEsperado"/>.
    /// </summary>
    public decimal ExtemporaneoEsperadoPorAse { get; set; }
}