namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 26 (T1, §2.1 paso 2 / D-D): reportes por ASE que el runtime resuelve y que el preflight
/// verifica. La superficie se gobierna por una sola regla —<b>lo que el runtime resuelve es lo que
/// el preflight exige, ni un archivo más ni uno menos</b>—: <c>RecaudosReversados</c> entró a la
/// superficie con el Plan 64 (H3.2, D-07) porque el runtime lo lee en ambas quincenas para el
/// bloque <c>ANT EXT-REV</c>.
/// </summary>
public enum ReporteInsumoAse
{
    /// <summary>Recaudo por componente (R1).</summary>
    R1,

    /// <summary>Detalle de saldos a favor (R2).</summary>
    R2,

    /// <summary>Reversión por componente (R4).</summary>
    R4,

    /// <summary>Reporte de recaudo por banco.</summary>
    Banco,

    /// <summary>Balance de subsidios y contribuciones.</summary>
    Balance,

    /// <summary>Saldos a favor aplicados por notas (solo quincena 2).</summary>
    SaldosNotas,

    /// <summary>Retribución negativa (solo quincena 2).</summary>
    RetribucionNegativa,

    /// <summary>
    /// Recaudos reversados por aplicar (fuente del bloque ANT EXT-REV; se exige en AMBAS quincenas
    /// porque el runtime lo lee en Q1 y Q2 — Plan 64 H3.2/D-07).
    /// </summary>
    RecaudosReversados
}
