namespace Remuneracion.Core.Models;

/// <summary>
/// HU-11 (2.5): inputs de un ASE desde la fuente <c>RetribuciónNegativa_*.xlsx</c>
/// (RETRIBUCION NEGATIVA). T0-0.4/0.5: en Q2 los 5 archivos traen SOLO la fila 1 con el rango
/// de fechas (fuente vacía de datos) → totales 0 legítimos confirmados por el golden
/// (D28:D32 del AJUSTES-SF-T = 0 en la referencia). El reader distingue "fuente vacía = 0
/// legítimo" de "header ausente = fallo que nombra ASE + reporte" (Riesgo 6 del plan).
///
/// Veredicto T0-0.3: la fila "Total" del bloque del template (C8/C21/C34/C47/C59) es VALOR
/// editable y el visible (C10/C23/C36/C49/C61) es fórmula <c>Cn-In</c>; la aritmética de dominio
/// <see cref="TotalRetribucionNegativa"/> replica esa fórmula (Total − ServEspK). Los valores de
/// retribución son negativos por componente (los totales cuadran con la fuente ±0.5).
/// </summary>
public sealed class RetribucionNegativaAseInputs
{
    /// <summary>
    /// ASE al que pertenece el bloque.
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Indica si la fuente trae columna "Especiales" (patrón R2). En Q2 la fuente está vacía:
    /// false, ServEspK = 0.
    /// </summary>
    public bool TieneColumnaEspeciales { get; set; }

    /// <summary>
    /// Total (col C) de la fila "Total" del bloque en la fuente (0 legítimo en Q2).
    /// </summary>
    public decimal Total { get; set; }

    /// <summary>
    /// Especiales de la fila "Total" (0 si ausente).
    /// </summary>
    public decimal ServEspK { get; set; }

    /// <summary>
    /// Celdas editables del bloque RETRIBUCION NEGATIVA por ASE, keyed por referencia de celda
    /// del template (ej. "C3", …, "O8"). Congelado por T0-0.7; el writer escribe SOLO estas
    /// celdas. En Q2 (fuente vacía) queda vacío o con ceros demostrables.
    /// </summary>
    public IReadOnlyDictionary<string, decimal> Celdas { get; set; } = new Dictionary<string, decimal>();

    /// <summary>
    /// Total de retribución negativa del ASE = Total − ServEspK (aritmética T0-0.3, visible Cn-In).
    /// </summary>
    public decimal TotalRetribucionNegativa => Total - ServEspK;
}