namespace Remuneracion.Core.Constants;

/// <summary>
/// HU-16 (D2b): INTERVENTORIA como INSUMO EXTERNO DECLARADO (veredicto T0-0.2/0.3/0.4).
///
/// El bloque <c>INTERVENTORIA</c> (header fila 25; ASE por fila 26..30: K = valor oficial mes,
/// L = Id ASE, M = segunda quincena, N = primera quincena; totales 31 en fórmula SUM; gran
/// total K32 = SUM(M31:N31)) es una TABLA ANUAL ESTÁTICA idéntica en ambos canónicos y en el
/// Q2-raíz (control). Búsqueda exhaustiva normalizada <c>*nterventoria*</c> en
/// <c>Docs/Insumos/</c> (ambos períodos): SIN FUENTE → el costo NO se deriva de R1/R2/R4 ni de
/// ningún archivo de insumos. Por eso esta HU NUNCA escribe ni inventa el bloque: la hoja queda
/// PROTEGIDA intacta (assert estructural en el writer) y estos valores son SOLO la declaración
/// auditable que el log/UI muestran por ASE (guía de diligenciamiento manual → HU-17).
/// </summary>
public static class InterventoriaDeclarada
{
    public const string Hoja = "INTERVENTORIA";

    /// <summary>Fila del header del bloque (T0-0.2).</summary>
    public const int FilaHeader = 25;

    /// <summary>Primera fila de datos ASE (T0-0.2: 26..30 = ASE 1..5).</summary>
    public const int FilaPrimerAse = 26;

    /// <summary>Valor oficial de mes (K26..K30) por ASE — tabla anual declarada.</summary>
    public static readonly IReadOnlyDictionary<int, decimal> ValorOficialMesPorAse =
        new Dictionary<int, decimal>
        {
            [1] = 378371975m,
            [2] = 533160511m,
            [3] = 309577071m,
            [4] = 240782166m,
            [5] = 257980891m
        };

    /// <summary>Segunda quincena (M26..M30) por ASE — tabla anual declarada.</summary>
    public static readonly IReadOnlyDictionary<int, decimal> SegundaQuincenaPorAse =
        new Dictionary<int, decimal>
        {
            [1] = 189185988m,
            [2] = 266580256m,
            [3] = 154788536m,
            [4] = 120391083m,
            [5] = 128990446m
        };

    /// <summary>Primera quincena (N26..N30) por ASE — tabla anual declarada.</summary>
    public static readonly IReadOnlyDictionary<int, decimal> PrimeraQuincenaPorAse =
        new Dictionary<int, decimal>
        {
            [1] = 189185987m,
            [2] = 266580255m,
            [3] = 154788535m,
            [4] = 120391083m,
            [5] = 128990445m
        };
}
