namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 21 (T1, R-E-1/R-E-6): bloque espejo de un ASE para la hoja <c>Reporte Componentes R1</c>.
///
/// Secuencia ORDENADA de <see cref="FilaEspejoR1"/> leída de la fuente <c>Recaudoporcomponente</c>
/// del período actual, más los encabezados de columna detectados dinámicamente. La fuente del
/// período define la forma: la secuencia observada ES la especificación (sin cardinalidades ni
/// ocurrencias congeladas — D-B/R-E-1).
/// </summary>
public sealed class BloqueEspejoAseInputs
{
    /// <summary>ASE al que pertenece el bloque.</summary>
    public Ase Ase { get; set; } = new();

    /// <summary>
    /// Encabezados de columna detectados en la fuente, en su orden físico (p. ej.
    /// <c>"Total"</c>, <c>"Componente TDF"</c>, …, <c>"SERVICIO ESPECIALES"</c>, …).
    /// </summary>
    public IReadOnlyList<string> Encabezados { get; set; } = [];

    /// <summary>
    /// Secuencia ordenada de filas tipadas del bloque (incluye las filas invariantes de cierre).
    /// </summary>
    public IReadOnlyList<FilaEspejoR1> Filas { get; set; } = [];

    /// <summary>
    /// Verdad si la fuente trae la columna <c>SERVICIO ESPECIALES</c> (o su alias
    /// <c>Especiales</c>). Si es falsa, el reader aporta 0 para esa columna (R-E-6).
    /// </summary>
    public bool TieneColumnaEspeciales { get; set; }

    /// <summary>Número de filas de la secuencia (forma del bloque del período).</summary>
    public int TotalFilas => Filas.Count;

    /// <summary>Fila invariante <c>Componente/Total</c>, o <c>null</c> si la fuente no la trae.</summary>
    public FilaEspejoR1? ComponenteTotal => Filas.FirstOrDefault(f => f.EsComponenteTotal);

    /// <summary>Fila invariante <c>Subs/Cont/Total</c>, o <c>null</c> si la fuente no la trae.</summary>
    public FilaEspejoR1? SubsContTotal => Filas.FirstOrDefault(f => f.EsSubsContTotal);

    /// <summary>Fila invariante <c>Total</c> final, o <c>null</c> si la fuente no la trae.</summary>
    public FilaEspejoR1? TotalFinal => Filas.FirstOrDefault(f => f.EsTotalFinal);

    /// <summary>
    /// Verdad si las tres invariantes duras de cierre T0e están presentes
    /// (<c>Componente/Total</c>, <c>Subs/Cont/Total</c>, <c>Total</c> final).
    /// </summary>
    public bool TieneInvariantesDeCierre =>
        ComponenteTotal is not null && SubsContTotal is not null && TotalFinal is not null;

    /// <summary>
    /// Plan 21 (W-7, auditoría PR3): nombres de las invariantes duras T0e AUSENTES (lista vacía si
    /// están las tres). Fuente ÚNICA del chequeo: la consumen el lector espejo
    /// (<c>LeerEspejoR1</c>, mensaje con nombre de archivo) y el orquestador
    /// (<c>ProcesadorPeriodo.ValidarInvariantesEspejoR1</c>, mensaje con reporte); así no hay dos
    /// definiciones divergentes de "invariante de cierre".
    /// </summary>
    public IReadOnlyList<string> FaltantesInvariantesDeCierre()
    {
        var faltantes = new List<string>();
        if (ComponenteTotal is null)
        {
            faltantes.Add("Componente/Total");
        }

        if (SubsContTotal is null)
        {
            faltantes.Add("Subs/Cont/Total");
        }

        if (TotalFinal is null)
        {
            faltantes.Add("Total (A='Total', B vacío)");
        }

        return faltantes;
    }

    // ── Plan 25 (T1, R-F-1/2/3): acceso en orden a los roles R1-Q2 sobre la secuencia espejo ──
    // La firma vive en FilaEspejoR1 (fuente única); aquí solo se exponen las secuencias en orden
    // y el agregado D-C. MapearR1Q2 consume estas mismas firmas (R-DOBLE-FUENTE).

    /// <summary>
    /// Plan 25 (T1, R-F-1): filas de firma <c>Mes/Total</c> en orden de aparición
    /// (0..2; ASE5 trae 2). Rol OBLIGATORIO.
    /// </summary>
    public IReadOnlyList<FilaEspejoR1> FilasMes() =>
        Filas.Where(f => f.EsMesTotal).ToList();

    /// <summary>
    /// Plan 25 (T1, R-F-2): filas de firma <c>Aplicacion nuevos x reversion/Total</c> en orden de
    /// aparición (puede ser 0). Rol OPCIONAL: ausente = 0 explícito.
    /// </summary>
    public IReadOnlyList<FilaEspejoR1> FilasAplicacion() =>
        Filas.Where(f => f.EsAplicacionTotal).ToList();

    /// <summary>
    /// Plan 25 (T1, T0e §6.1): filas de firma <c>Subsidio(-)/Contribucion(+)</c> en orden de
    /// aparición. Predicado de cierre/documentación (sin consumidor Q2 tras D-A).
    /// </summary>
    public IReadOnlyList<FilaEspejoR1> FilasSubsidio() =>
        Filas.Where(f => f.EsSubsidio).ToList();

    /// <summary>
    /// Plan 25 (T1, D-C / R-F-2): suma el valor de la columna <paramref name="encabezado"/> de
    /// TODAS las filas con firma <c>Aplicacion</c> (0 si no hay ninguna). Fuente única del
    /// agregado EXTEMP tras D-A.
    /// </summary>
    public decimal SumarAplicacion(string encabezado) =>
        FilasAplicacion().Sum(f => f.Valor(encabezado) ?? 0m);
}
