namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 42 (T1, D-C — patrón Plan 34 <see cref="SinonimosEmpresaR1"/>): catálogo de
/// DIVERGENCIAS DECLARADAS como DATOS citados, más el matcher puro
/// <see cref="Buscar(IReadOnlyList{Entrada}, string, decimal, decimal)"/>.
///
/// Una divergencia declarada es una diferencia EXTERNA por construcción (discrepancia
/// oficial-vs-manual con dueño fuera de la aplicación). NO es un defecto interno: los defectos
/// de lectura, cálculo, colocación, sellado, naming, preflight o caché NUNCA son declarables
/// (taxonomía T0 del Plan 42). El catálogo es ABIERTO pero DISCIPLINADO: cada entrada exige sus
/// 6 campos citados (celda(s), origen del monto, dueño, cita de evidencia, fecha y condición de
/// retiro); sin ellos la entrada no aplica (jamás un verde silencioso). Añadir una divergencia
/// futura = añadir una FILA citada, nunca una rama de código.
///
/// D-H (Plan 34/40/41/42): el matcher NO recibe período y NO ramifica por empresa; recibe la
/// celda y los montos por parámetro. La primera (y hoy única) entrada es OCCIDENTE ASE5.
/// </summary>
public static class DivergenciasDeclaradasCatalogo
{
    /// <summary>Tolerancia de la comparación del monto con el Δ citado (misma doctrina ±0.5).</summary>
    public const decimal Tolerancia = 0.5m;

    /// <summary>
    /// Entrada del catálogo: una divergencia declarada con sus 6 campos citados más el Δ
    /// (diferencia material) que permite reconocerla en la clasificación.
    /// </summary>
    public sealed record Entrada(
        string Id,
        IReadOnlyList<string> Celdas,
        decimal Delta,
        string OrigenMonto,
        string Dueno,
        string CitaEvidencia,
        string Fecha,
        string CondicionRetiro)
    {
        /// <summary>
        /// Verdad si la entrada trae sus 6 campos citados (celda(s), origen del monto, dueño,
        /// cita de evidencia, fecha, condición de retiro) más un Δ no nulo. Una entrada incompleta
        /// NUNCA aplica (R-CAT-2: sin los 6 campos no hay entrada).
        /// </summary>
        public bool Completa =>
            !string.IsNullOrWhiteSpace(Id)
            && Celdas.Count > 0
            && Celdas.All(c => !string.IsNullOrWhiteSpace(c))
            && Delta != 0m
            && !string.IsNullOrWhiteSpace(OrigenMonto)
            && !string.IsNullOrWhiteSpace(Dueno)
            && !string.IsNullOrWhiteSpace(CitaEvidencia)
            && !string.IsNullOrWhiteSpace(Fecha)
            && !string.IsNullOrWhiteSpace(CondicionRetiro);
    }

    /// <summary>
    /// Catálogo vigente (abierto). Primera entrada: OCCIDENTE ASE5 (oficial 3 términos vs manual
    /// legacy 2 términos; Δ = 417.519,62 = la fila de la 1.ª sección). Se retira por su condición
    /// de retiro (cuando la UAESP resuelva), no adaptando código.
    /// </summary>
    public static IReadOnlyList<Entrada> Entradas { get; } =
    [
        new(
            Id: "OCCIDENTE-ASE5",
            Celdas: ["O25", "D34", "F34"],
            Delta: 417519.62m,
            OrigenMonto:
                "Reporte Componentes R1, ancla de OCCIDENTE ASE5: el subtotal del espejo del oficial " +
                "(3 términos; incluye la fila de la 1.ª sección) supera al del manual legacy " +
                "(2 términos; omite esa fila) en el valor de la 1.ª sección.",
            Dueno: "UAESP",
            CitaEvidencia:
                "Diagnóstico congelado Docs/T0-Plan41-MiniT0-OccidenteAse5-SaldosNota-TablaJ.md §1 " +
                "(ambas fórmulas, valores por término y Δ) y su fixture read-only " +
                "Remuneracion.IntegrationTests/T0FrenteEspejoR1Ase5EnelTests.cs.",
            Fecha: "2026-10-14",
            CondicionRetiro:
                "La UAESP confirme si la remuneración de OCCIDENTE ASE5 incluye la 1.ª sección; " +
                "entonces se elimina esta fila, no se adapta el código.")
    ];

    /// <summary>
    /// Matcher puro sobre el catálogo vigente. Devuelve la entrada cuya lista de celdas contiene
    /// <paramref name="celda"/> y cuyo Δ coincide con |observado − esperado| (±<see cref="Tolerancia"/>);
    /// <c>null</c> si ninguna entrada COMPLETA aplica (divergencia no declarada).
    /// </summary>
    public static Entrada? Buscar(string celda, decimal observado, decimal esperado) =>
        Buscar(Entradas, celda, observado, esperado);

    /// <summary>
    /// Matcher puro sobre un catálogo explícito (testeable). Las entradas incompletas se IGNORAN
    /// (R-CAT-2: nunca un verde silencioso).
    /// </summary>
    public static Entrada? Buscar(IReadOnlyList<Entrada> catalogo, string celda, decimal observado, decimal esperado)
    {
        if (catalogo is null || string.IsNullOrWhiteSpace(celda))
        {
            return null;
        }

        var delta = Math.Abs(observado - esperado);
        foreach (var entrada in catalogo)
        {
            if (!entrada.Completa)
            {
                continue;
            }

            if (!entrada.Celdas.Any(c => string.Equals(c, celda, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            if (Math.Abs(delta - entrada.Delta) <= Tolerancia)
            {
                return entrada;
            }
        }

        return null;
    }
}
