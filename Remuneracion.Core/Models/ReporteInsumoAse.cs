namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 26 (T1, §2.1 paso 2 / D-D): reportes por ASE que el runtime resuelve y que el preflight
/// verifica. <c>RecaudosReversados</c> NO forma parte de esta superficie: el runtime no lo
/// consume (§V9), por lo que exigirlo inventaría un requisito.
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
    RetribucionNegativa
}
