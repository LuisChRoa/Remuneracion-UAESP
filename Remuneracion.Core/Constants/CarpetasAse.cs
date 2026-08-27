namespace Remuneracion.Core.Constants;

/// <summary>
/// Constantes de dominio para la remuneración UAESP.
/// </summary>
public static class CarpetasAse
{
    /// <summary>
    /// Prefijos esperados de las carpetas por ASE, en orden (1 a 5).
    /// </summary>
    public static readonly IReadOnlyList<string> Prefijos =
    [
        "1-Promoambiental",
        "2-Lime",
        "3-Ciudad Limpia",
        "4-Bogotá Limpia",
        "5-Área Limpia"
    ];
}
