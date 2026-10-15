namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 40 (T1, R-E-1/R-E-2/R-E-3 — D-A/D-D): tabla CERRADA de las formas que el gate W-1 sabe
/// EVALUAR — los 5 sub-bloques booleanos de <c>VALIDACION_TOTAL</c>. El TEXTO <c>&lt;f&gt;</c>
/// esperado es el congelado por el T0 del Plan 40 desde los bytes de la plantilla oficial
/// (fixture <c>GeometriaPlantillaEstandarizadaT0Tests</c>); el gate lo aserta ANTES de evaluar y
/// una tercera forma = <c>ERR-PLANTILLA</c> (hoja+celda+esperado+observado).
///
/// No es un parser de Excel general: cada sub-bloque es una IGUALDAD de dos valores que el
/// reader resuelve desde las celdas efectivamente escritas (nunca desde el caché <c>&lt;v&gt;</c>
/// de una celda de fórmula). Ver <see cref="Igualdad"/> (±0.5, R-E-3).
///
/// <para>
/// <b>Enmienda E-1 (2026-10-14, ratificada por el Ingeniero):</b> los 5 sub-bloques se EVALÚAN
/// sin excepción (E1-1: prohibido <c>SeGatea=false</c>). La cadena del espejo de remuneración
/// (<c>O24</c>/<c>N9</c>, <c>D33</c>/<c>D109</c>, <c>F33</c>/<c>D109+N109</c>) ES resoluble desde lo
/// escrito por el resolutor acotado; si la igualdad no cierra, el veredicto es <c>FALSE</c>
/// (divergencia real) y se reporta como hallazgo (E1-2) — nunca "no evaluable" ni "verde".
/// </para>
/// </summary>
public static class FormasValidacionTotal
{
    /// <summary>Hoja de los 5 sub-bloques.</summary>
    public const string Hoja = "VALIDACION_TOTAL";

    /// <summary>Tolerancia de la comparación aritmética del gate (R-E-3).</summary>
    public const decimal Tolerancia = 0.5m;

    /// <summary>
    /// Los 5 sub-bloques con su texto-<c>&lt;f&gt;</c> esperado (byte-fiel al oficial, Plan 40 T0).
    /// El gate los EVALÚA TODOS desde lo escrito (E1-1): <c>C15</c>/<c>D25</c> = cadena de RECAUDO
    /// (100 % escrita por la app); <c>O25</c>/<c>D34</c>/<c>F34</c> = cadena del espejo de
    /// REMUNERACIÓN, resoluble por el resolutor (referencias + SUM + fórmulas compartidas) y por
    /// tanto EVALUABLE (un FALSE aquí es divergencia real, no "no evaluable").
    /// </summary>
    public static readonly IReadOnlyList<(string Celda, string TextoFormula)> SubBloques =
    [
        ("C15", "C14=C9"),
        ("D25", "D24=C9"),
        ("O25", "O24=N9"),
        ("D34", "D33='CONSOLIDADO_TOTAL RECAUDO'!D109"),
        ("F34", "F33='CONSOLIDADO_TOTAL RECAUDO'!D109+'CONSOLIDADO_TOTAL RECAUDO'!N109"),
    ];

    /// <summary>Texto-<c>&lt;f&gt;</c> esperado de un sub-bloque; <c>null</c> si no está en la tabla.</summary>
    public static string? TextoEsperado(string celda)
    {
        foreach (var (c, t) in SubBloques)
        {
            if (string.Equals(c, celda, StringComparison.OrdinalIgnoreCase))
            {
                return t;
            }
        }

        return null;
    }

    /// <summary>
    /// Igualdad de los dos operandos del sub-bloque con tolerancia ±0.5 (R-E-3). El reader
    /// entrega los operandos ya resueltos desde lo escrito; aquí solo se compara.
    /// </summary>
    public static bool Igualdad(decimal izquierdo, decimal derecho) =>
        Math.Abs(izquierdo - derecho) <= Tolerancia;
}
