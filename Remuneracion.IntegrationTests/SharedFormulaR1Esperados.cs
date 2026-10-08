namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 33 (T1, R-G-3): fixture del mapa shared congelado de <c>Reporte Componentes R1</c>
/// (<c>sheet10.xml</c>), leído del disco (plantilla julio + <c>Resultado1</c> sano, byte-idénticos en
/// atributos y texto — verificado por <see cref="SharedFormulaR1GateTests"/>). 35 masters, <c>si</c>
/// 0..34, 1035 seguidoras.
///
/// Los 35 masters son los <c>&lt;f t="shared" ref="…" si="…"&gt;</c> de la plantilla; el fix de T2
/// (preserve-and-rewrite) debe reescribirlos <b>preservando</b> <c>t</c>/<c>ref</c>/<c>si</c> y el texto
/// (julio-identidad, D-D). Si el disco cambia, <see cref="SharedFormulaR1GateTests"/> delata el drift
/// antes de que el gate se confunda.
///
/// Referencias explícitas del contrato (E3): <c>G53</c> (si=0, ref=G53:AP53, texto G32+G12) y
/// <c>G468</c> (si=20, ref=G468:AO468, texto G422+G384).
/// </summary>
internal static class SharedFormulaR1Esperados
{
    public const string Hoja = "Reporte Componentes R1";

    /// <summary>Master shared congelado: celda + si + rango ref + texto de &lt;f&gt;.</summary>
    public sealed record MasterShared(string Celda, uint Si, string Rango, string Texto);

    /// <summary>Total de masters shared del mapa (E4).</summary>
    public const int TotalMaestros = 35;

    /// <summary>Total de seguidoras shared del mapa (E4).</summary>
    public const int TotalSeguidoras = 1035;

    /// <summary>Mapa shared completo de <c>Reporte Componentes R1</c> (congelado de la plantilla julio).</summary>
    public static IReadOnlyList<MasterShared> Maestros { get; } =
    [
        new("G53", 0, "G53:AP53", "G32+G12"),
        new("H55", 1, "H55:AP55", "H17"),
        new("H63", 2, "H63:AP63", "H21+H8"),
        new("H65", 3, "H65:AP65", "H16"),
        new("G73", 4, "G73:AP73", "G31+G11"),
        new("F83", 5, "F83:AP83", "F13"),
        new("H206", 6, "H206:AP206", "H132+H94"),
        new("H208", 7, "H208:AP208", "H107"),
        new("H211", 8, "H211:AP211", "H121"),
        new("G216", 9, "G216:AA216", "G111+G90"),
        new("H218", 10, "H218:AA218", "H99"),
        new("H226", 11, "H226:AP226", "H131+H93"),
        new("H228", 12, "H228:AP228", "H106"),
        new("F236", 13, "F236:AA236", "F94"),
        new("H343", 14, "H343:AO343", "H265+H244"),
        new("H345", 15, "H345:AO345", "H250"),
        new("H353", 16, "H353:AO353", "H254"),
        new("H355", 17, "H355:AO355", "H249"),
        new("H363", 18, "H363:AO363", "H264+H243"),
        new("G373", 19, "G373:Z373", "G244"),
        new("G468", 20, "G468:AO468", "G422+G384"),
        new("G470", 21, "G470:AA470", "G397"),
        new("G473", 22, "G473:AO473", "G411"),
        new("H475", 23, "H475:Z475", "H392"),
        new("G478", 24, "G478:Z478", "G401+G380"),
        new("H480", 25, "H480:Z480", "H389"),
        new("G488", 26, "G488:AO488", "G421+G383"),
        new("G490", 27, "G490:Z490", "G396"),
        new("F498", 28, "F498:Z498", "F385"),
        new("H558", 29, "H558:AO558", "H531"),
        new("H560", 30, "H560:AO560", "H512"),
        new("H568", 31, "H568:AO568", "H516"),
        new("H573", 32, "H573:AO573", "H520"),
        new("H578", 33, "H578:AO578", "H530"),
        new("H580", 34, "H580:AO580", "H511")
    ];
}
