namespace Remuneracion.Core.Models;

/// <summary>
/// HU-16 (D3a): L-Especiales menores del <c>Reporte Componentes R1</c> por ASE.
///
/// Veredicto T0-0.5/0.6 (Plan 16 §4 Fase 0, probado contra ambos canónicos y las fuentes R1):
/// la columna L del template R1 es el ESPEJO de la columna <c>SERVICIO ESPECIALES</c> de la
/// fuente <c>Recaudoporcomponente_*</c> (misma fila de etiquetas, misma secuencia de bloques).
/// Toda celda L numérica FUERA del set V4 (HU-07 T0) y del mapa HU-08 cierra contra el ESP de
/// su fila fuente ±0.5 → desenlace D3a: entra al mapa editable por <see cref="Ase.Id"/> y el
/// writer la escribe en la MISMA pasada atómica (elimina el riesgo stale del follow-up HU-07).
///
/// Doctrina HU-12 (nunca 0 silencioso en operando): el reader distingue "leído 0" (la fila
/// fuente existe y su ESP vale 0 — valor legítimo) de "slot ausente" (la fila rol del mapa T0
/// no existe en la fuente → fail-fast que nombra ASE + hoja + celda).
/// </summary>
public sealed class LEspecialesMenoresAseInputs
{
    /// <summary>
    /// ASE al que pertenecen las L-menores.
    /// </summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Celdas L del template → valor leído de la columna SERVICIO ESPECIALES de la fuente
    /// (keyed por referencia de celda del template, ej. "L11"). Congelado por T0-0.5.
    /// </summary>
    public IReadOnlyDictionary<string, decimal> Celdas { get; set; } = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Total de las L-menores del ASE (aritmética de dominio; usada por el log/UI y la Capa A).
    /// No alimenta ninguna fórmula del workbook (bloque informativo, §2.4).
    /// </summary>
    public decimal Total => Celdas.Values.Sum();

    /// <summary>
    /// Verdad si el mapa T0 declara celdas L-menores para este ASE (siempre en 1..5).
    /// </summary>
    public bool TieneCeldas => Celdas.Count > 0;
}
