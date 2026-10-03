namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 21 (W-2, auditoría PR3): evidencia estructural de una pasada del espejo R1.
///
/// Instrumenta el reanclaje sin cambiar el comportamiento: <see cref="DeltasPorBloque"/> = Δ de
/// filas (fuente − destino) por ASE y <see cref="ReferenciasReancladasPorBloque"/> = cantidad de
/// referencias A1 cuyo número de fila cambió por el reanclaje de ese bloque. Los tests asertan
/// estos valores contra la expectativa derivada de la tabla T0a/T0c (deltas por bloque reales) en
/// lugar de una cota arbitraria.
/// </summary>
internal sealed class EspejoR1MutacionResultado
{
    public EspejoR1MutacionResultado(
        IReadOnlyDictionary<int, int> deltasPorBloque,
        IReadOnlyDictionary<int, int> referenciasReancladasPorBloque)
    {
        DeltasPorBloque = deltasPorBloque;
        ReferenciasReancladasPorBloque = referenciasReancladasPorBloque;
    }

    /// <summary>Δ de filas (filas fuente − filas destino) aplicado a cada bloque ASE.</summary>
    public IReadOnlyDictionary<int, int> DeltasPorBloque { get; }

    /// <summary>Referencias A1 reancladas por bloque ASE (tokens cuyo número de fila cambió).</summary>
    public IReadOnlyDictionary<int, int> ReferenciasReancladasPorBloque { get; }

    /// <summary>Σ de referencias reancladas de la pasada (invariante = suma por bloque).</summary>
    public int TotalReferenciasReancladas => ReferenciasReancladasPorBloque.Values.Sum();
}
