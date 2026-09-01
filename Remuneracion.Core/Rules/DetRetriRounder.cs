namespace Remuneracion.Core.Rules;

/// <summary>
/// Regla pure-domain para redondear el valor DetRetri siguiendo la semántica de Excel
/// ROUND(valor, 0) con <see cref="MidpointRounding.AwayFromZero"/>.
/// </summary>
public static class DetRetriRounder
{
    public static decimal Round(decimal valor)
    {
        return decimal.Round(valor, 0, MidpointRounding.AwayFromZero);
    }
}
