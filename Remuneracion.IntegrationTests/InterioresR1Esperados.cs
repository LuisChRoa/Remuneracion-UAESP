using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 32 (T1, pieza a — R-F-1): textos de fórmula del <b>INTERIOR de <c>Reporte Componentes R1</c></b>
/// (sub-visibles por empresa), transcritos EXACTAMENTE de los manuales del administrativo (solo lectura
/// ZIP+XML BCL, sin Excel/COM) y con cita de archivo + hoja + celda por valor.
///
/// Fuente agosto: <c>Docs/Prueba2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx</c>
/// Fuente julio:  <c>Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx</c>
///
/// El conjunto se congeló por lectura directa del disco (método zip+XML del T0 §0.1): TODAS las celdas
/// F/G/H/L de las filas-interior clasificadas por firma —
///   SUB_EMP (subtotal-empresa <c>C=&lt;empresa&gt; ∧ D='OPORTUNO' ∧ E='Total'</c>, columna F),
///   SUB_TDF (espejo G/H del subtotal-empresa),
///   SUB_L   (espejo L del subtotal-empresa, caso <c>L569</c>),
///   EXT_INT (EXTEMP-interior <c>D='EXTEMPORANEO' ∧ E='Total'</c>),
///   SUBS    (single-ref <c>E='Subsidio(-)/Contribucion(+)'</c>),
///   AFASEO  (<c>D='AFASEO' ∧ E='Total'</c>) —
/// con <c>&lt;f&gt;</c> no-vacía en el manual, MÁS las celdas clase L-1 (el manual NO trae <c>&lt;f&gt;</c>;
/// <c>Formula = null</c> + <c>ValorLiteral</c>): <c>F217</c> (testigo) y las filas que el manual de
/// agosto eliminó para ASE3 (<c>F327/F328/F337/F338</c> + <c>G327/G337/H327/H337</c>).
///
/// Notas de disco (el disco MANDA sobre el T0 §0.1):
///   - La fila visible TOT_OPT del bloque (<c>C='TOTAL'</c>) NO entra aquí: sus F/G son visibles (Plan 31,
///     <see cref="TotalesR1Esperados"/>) y su espejo H no es un subtotal de empresa.
///   - <c>L569</c> está en el bloque del ASE5 (fila ENERBIT del manual de agosto), no en ASE4 (donde lo
///     ubicaba la tabla T0a).
///   - <c>F464</c>/<c>F584</c>/<c>G584</c> (clase L-2: la app no trae <c>&lt;f&gt;</c>) SÍ están aquí como
///     objetivo (el manual las trae): su texto se congela como cualquier R-1.
///   - Los bordes vacíos del manual (<c>H50</c>, <c>H194</c>, <c>G355</c>, <c>H584</c>-style) NO se
///     congelan (el manual no tiene elemento allí); quedan señalados como brecha del comparador.
///
/// <see cref="InterioresR1EsperadosTests"/> re-lee el disco y delata cualquier drift del manual contra
/// esta congelación (si el disco cambia, el test falla antes de que el gate se confunda).
/// </summary>
internal static class InterioresR1Esperados
{
    /// <summary>Celda interior de R1: celda + etiqueta + texto de fórmula esperado + cita.</summary>
    /// <param name="AseId">ASE (1..5).</param>
    /// <param name="Celda">Referencia en el manual (p. ej. <c>F60</c>).</param>
    /// <param name="Etiqueta">SUB_EMP | SUB_TDF | SUB_L | EXT_INT | SUBS | AFASEO.</param>
    /// <param name="Formula">Texto de <c>&lt;f&gt;</c> esperado; <c>null</c> = el manual trae literal (L-1).</param>
    /// <param name="Cita">Archivo + hoja + celda (trazabilidad de la transcripción).</param>
    /// <param name="ValorLiteral">Literal esperado cuando <paramref name="Formula"/> es <c>null</c> (clase L-1).</param>
    public sealed record InteriorR1(int AseId, string Celda, string Etiqueta, string? Formula, string Cita, decimal? ValorLiteral = null);

    public const string Hoja = "Reporte Componentes R1";
    public const string ManualAgostoRelativo = TotalesR1Esperados.ManualAgostoRelativo;
    public const string ManualJulioRelativo = TotalesR1Esperados.ManualJulioRelativo;

    /// <summary>Manual agosto 202608-2 (interior congelado: R-1 + L-1).</summary>
    public static IReadOnlyList<InteriorR1> Agosto { get; } =
    [
        new(1, "F51", "SUBS", "F45", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F51"),
        new(1, "F52", "EXT_INT", "F34+F14-L14", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F52"),
        new(1, "G52", "EXT_INT", "G14", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G52"),
        new(1, "H52", "EXT_INT", "H14", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H52"),
        new(1, "F53", "SUBS", "F34", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F53"),
        new(1, "F60", "SUB_EMP", "F37+F18-L18", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F60"),
        new(1, "G60", "SUB_TDF", "G18", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G60"),
        new(1, "H60", "SUB_TDF", "H18", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H60"),
        new(1, "F61", "SUBS", "F37", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F61"),
        new(1, "F62", "EXT_INT", "F33+F13-L13", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F62"),
        new(1, "G62", "EXT_INT", "G13", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G62"),
        new(1, "H62", "EXT_INT", "H13", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H62"),
        new(1, "F63", "SUBS", "F33", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F63"),
        new(1, "F70", "SUB_EMP", "F44+F28+F8-L8-L28", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F70"),
        new(1, "G70", "SUB_TDF", "G28+G8", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G70"),
        new(1, "F71", "SUBS", "F44", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F71"),
        new(1, "F80", "AFASEO", "F10", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F80"),
        new(2, "F195", "SUBS", "F148", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F195"),
        new(2, "F196", "EXT_INT", "F130+F101-L101", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F196"),
        new(2, "G196", "EXT_INT", "G101", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G196"),
        new(2, "F197", "SUBS", "F130", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F197"),
        new(2, "F199", "SUB_EMP", "F140+F114-L114", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F199"),
        new(2, "G199", "SUB_TDF", "G114", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G199"),
        new(2, "F200", "SUBS", "F140", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F200"),
        new(2, "F204", "SUB_EMP", "F133+F105+F87--L87-L105", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F204"),
        new(2, "G204", "SUB_TDF", "G105+G87", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G204"),
        new(2, "F205", "SUBS", "F133", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F205"),
        new(2, "F206", "EXT_INT", "F129+F97-L97", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F206"),
        new(2, "G206", "EXT_INT", "G97", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G206"),
        new(2, "F207", "SUBS", "F129", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F207"),
        new(2, "F214", "SUB_EMP", "F147+F124+F92-L92-L124", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F214"),
        new(2, "G214", "SUB_TDF", "G124+G92", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G214"),
        new(2, "F215", "SUBS", "F147", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F215"),
        new(2, "F216", "EXT_INT", "F100-L100", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F216"),
        new(2, "G216", "EXT_INT", "G100", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G216"),
        new(2, "F224", "AFASEO", "F93", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F224"),
        new(3, "F326", "SUBS", "F263", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F326"),
        new(3, "F335", "SUB_EMP", "F255+F240+F231-L231-L240", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F335"),
        new(3, "G335", "SUB_TDF", "G240+G231", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G335"),
        new(3, "H335", "SUB_TDF", "H240+H231", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H335"),
        new(3, "F336", "SUBS", "F255", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F336"),
        new(3, "F345", "SUB_EMP", "F262+F250+F234-L234-L250", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F345"),
        new(3, "G345", "SUB_TDF", "G250+G234", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G345"),
        new(3, "F346", "SUBS", "F262", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F346"),
        new(3, "F355", "AFASEO", "F235", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F355"),
        new(4, "F457", "SUBS", "F436", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F457"),
        new(4, "F458", "EXT_INT", "F418+F382-L382", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F458"),
        new(4, "G458", "EXT_INT", "G382", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G458"),
        new(4, "H458", "EXT_INT", "H382", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H458"),
        new(4, "F459", "SUBS", "F418", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F459"),
        new(4, "F461", "SUB_EMP", "F428+F396-L396", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F461"),
        new(4, "G461", "SUB_TDF", "G396", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G461"),
        new(4, "H461", "SUB_TDF", "H396", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H461"),
        new(4, "F462", "SUBS", "F428", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F462"),
        new(4, "F463", "EXT_INT", "F414+F377-L377", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F463"),
        new(4, "G463", "EXT_INT", "G377", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G463"),
        new(4, "H463", "EXT_INT", "H377", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H463"),
        new(4, "F464", "SUBS", "F414", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F464"),
        new(4, "F466", "SUB_EMP", "F421+F386+F362-L362-L386", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F466"),
        new(4, "G466", "SUB_TDF", "G386+G362", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G466"),
        new(4, "H466", "SUB_TDF", "H386+H362", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H466"),
        new(4, "F467", "SUBS", "F421", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F467"),
        new(4, "F468", "EXT_INT", "F411+F371-L371", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F468"),
        new(4, "G468", "EXT_INT", "G371", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G468"),
        new(4, "H468", "EXT_INT", "H371", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H468"),
        new(4, "F469", "SUBS", "F411", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F469"),
        new(4, "F476", "SUB_EMP", "F435+F406+F365-L365-L406", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F476"),
        new(4, "G476", "SUB_TDF", "G406+G365", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G476"),
        new(4, "H476", "SUB_TDF", "H406+H365", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H476"),
        new(4, "F477", "SUBS", "F435", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F477"),
        new(4, "F478", "EXT_INT", "F417+F381-L381", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F478"),
        new(4, "G478", "EXT_INT", "G381", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G478"),
        new(4, "H478", "EXT_INT", "H381", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H478"),
        new(4, "F479", "SUBS", "F417", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F479"),
        new(4, "F486", "AFASEO", "F367", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F486"),
        new(5, "F555", "SUBS", "F548", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F555"),
        new(5, "F556", "EXT_INT", "F534+F506-L506", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F556"),
        new(5, "G556", "EXT_INT", "G506", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G556"),
        new(5, "H556", "EXT_INT", "H506", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H556"),
        new(5, "F557", "SUBS", "F534", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F557"),
        new(5, "F564", "SUB_EMP", "F537+F510-L510", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F564"),
        new(5, "G564", "SUB_TDF", "G510", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G564"),
        new(5, "H564", "SUB_TDF", "H510", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H564"),
        new(5, "F565", "SUBS", "F537", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F565"),
        new(5, "F569", "SUB_EMP", "F540+F514-L514", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F569"),
        new(5, "G569", "SUB_TDF", "G514", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G569"),
        new(5, "H569", "SUB_TDF", "H514", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H569"),
        new(5, "L569", "SUB_L", "L514", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!L569"),
        new(5, "F570", "SUBS", "F540", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F570"),
        new(5, "F574", "SUB_EMP", "F524+F547+F493-L493-L524", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F574"),
        new(5, "G574", "SUB_TDF", "G524+G493", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G574"),
        new(5, "H574", "SUB_TDF", "H524+H493", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H574"),
        new(5, "F575", "SUBS", "F547", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F575"),
        new(5, "F576", "EXT_INT", "F533+F505-L505", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F576"),
        new(5, "G576", "EXT_INT", "G505", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G576"),
        new(5, "H576", "EXT_INT", "H505", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H576"),
        new(5, "F577", "SUBS", "F533", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F577"),
        new(5, "F584", "AFASEO", "F495", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F584"),
        new(5, "G584", "AFASEO", "G495", "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G584"),
        new(2, "F217", "SUBS", null, "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F217 (literal; sin <f> en el manual)", 0m),
        new(3, "F327", "EXT_INT", null, "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F327 (literal; sin <f> en el manual)", 0m),
        new(3, "F328", "SUBS", null, "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F328 (literal; sin <f> en el manual)", 0m),
        new(3, "F337", "EXT_INT", null, "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F337 (literal; sin <f> en el manual)", 0m),
        new(3, "F338", "SUBS", null, "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F338 (literal; sin <f> en el manual)", 0m),
        new(3, "G327", "EXT_INT", null, "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G327 (literal; sin <f> en el manual)", 0m),
        new(3, "G337", "EXT_INT", null, "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G337 (literal; sin <f> en el manual)", 0m),
        new(3, "H327", "EXT_INT", null, "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H327 (literal; sin <f> en el manual)", 0m),
        new(3, "H337", "EXT_INT", null, "Manual Agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!H337 (literal; sin <f> en el manual)", 0m)
    ];

    /// <summary>Manual julio 202607-2 (interior congelado: julio no tiene clases L; identidad D-E).</summary>
    public static IReadOnlyList<InteriorR1> Julio { get; } =
    [
        new(1, "F54", "SUBS", "F48", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F54"),
        new(1, "F55", "EXT_INT", "F37+F17-L17", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F55"),
        new(1, "G55", "EXT_INT", "G17", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G55"),
        new(1, "H55", "EXT_INT", "H17", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H55"),
        new(1, "F56", "SUBS", "F37", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F56"),
        new(1, "F63", "SUB_EMP", "F40+F21+F8-L8-L21", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F63"),
        new(1, "G63", "SUB_TDF", "G21+G8", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G63"),
        new(1, "H63", "SUB_TDF", "H21+H8", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H63"),
        new(1, "F64", "SUBS", "F40", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F64"),
        new(1, "F65", "EXT_INT", "F36+F16-L16", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F65"),
        new(1, "G65", "EXT_INT", "G16", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G65"),
        new(1, "H65", "EXT_INT", "H16", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H65"),
        new(1, "F66", "SUBS", "F36", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F66"),
        new(1, "F73", "SUB_EMP", "F47+F31+F11-L11-L31", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F73"),
        new(1, "G73", "SUB_TDF", "G31+G11", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G73"),
        new(1, "F74", "SUBS", "F47", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F74"),
        new(1, "F83", "AFASEO", "F13", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F83"),
        new(2, "F207", "SUBS", "F160", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F207"),
        new(2, "F208", "EXT_INT", "F142+F107-L107", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F208"),
        new(2, "G208", "EXT_INT", "G107", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G208"),
        new(2, "H208", "EXT_INT", "H107", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H208"),
        new(2, "F209", "SUBS", "F142", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F209"),
        new(2, "F211", "SUB_EMP", "F152+F121-L121", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F211"),
        new(2, "G211", "SUB_TDF", "G121", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G211"),
        new(2, "H211", "SUB_TDF", "H121", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H211"),
        new(2, "F212", "SUBS", "F152", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F212"),
        new(2, "F216", "SUB_EMP", "F145+F111+F90--L90-L111", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F216"),
        new(2, "G216", "SUB_TDF", "G111+G90", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G216"),
        new(2, "F217", "SUBS", "F145", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F217"),
        new(2, "F218", "EXT_INT", "F136+F99-L99", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F218"),
        new(2, "G218", "EXT_INT", "G99", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G218"),
        new(2, "H218", "EXT_INT", "H99", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H218"),
        new(2, "F219", "SUBS", "F136", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F219"),
        new(2, "F226", "SUB_EMP", "F159+F131+F93-L93-L131", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F226"),
        new(2, "G226", "SUB_TDF", "G131+G93", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G226"),
        new(2, "H226", "SUB_TDF", "H131+H93", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H226"),
        new(2, "F227", "SUBS", "F159", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F227"),
        new(2, "F228", "EXT_INT", "F141+F106-L106", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F228"),
        new(2, "G228", "EXT_INT", "G106", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G228"),
        new(2, "H228", "EXT_INT", "H106", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H228"),
        new(2, "F229", "SUBS", "F141", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F229"),
        new(2, "F236", "AFASEO", "F94", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F236"),
        new(3, "F344", "SUBS", "F281", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F344"),
        new(3, "F345", "EXT_INT", "F270+F250-L250", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F345"),
        new(3, "G345", "EXT_INT", "G250", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G345"),
        new(3, "H345", "EXT_INT", "H250", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H345"),
        new(3, "F346", "SUBS", "F270", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F346"),
        new(3, "F353", "SUB_EMP", "F273+F254-L254", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F353"),
        new(3, "G353", "SUB_TDF", "G254", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G353"),
        new(3, "H353", "SUB_TDF", "H254", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H353"),
        new(3, "F354", "SUBS", "F273", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F354"),
        new(3, "F355", "EXT_INT", "F269+F249-L249", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F355"),
        new(3, "G355", "EXT_INT", "G249", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G355"),
        new(3, "H355", "EXT_INT", "H249", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H355"),
        new(3, "F356", "SUBS", "F269", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F356"),
        new(3, "F363", "SUB_EMP", "F280+F264+F243-L243-L264", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F363"),
        new(3, "G363", "SUB_TDF", "G264+G243", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G363"),
        new(3, "H363", "SUB_TDF", "H264+H243", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H363"),
        new(3, "F364", "SUBS", "F280", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F364"),
        new(3, "F373", "AFASEO", "F244", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F373"),
        new(3, "G373", "AFASEO", "G244", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G373"),
        new(4, "F469", "SUBS", "F448", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F469"),
        new(4, "F470", "EXT_INT", "F430+F397-L397", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F470"),
        new(4, "G470", "EXT_INT", "G397", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G470"),
        new(4, "F471", "SUBS", "F430", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F471"),
        new(4, "F473", "SUB_EMP", "F440+F411-L411", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F473"),
        new(4, "G473", "SUB_TDF", "G411", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G473"),
        new(4, "F474", "SUBS", "F440", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F474"),
        new(4, "F475", "EXT_INT", "F392-L392", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F475"),
        new(4, "G475", "EXT_INT", "G392", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G475"),
        new(4, "H475", "EXT_INT", "H392", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H475"),
        new(4, "F478", "SUB_EMP", "F433+F401+F380-L380-L401", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F478"),
        new(4, "G478", "SUB_TDF", "G401+G380", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G478"),
        new(4, "F479", "SUBS", "F433", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F479"),
        new(4, "F480", "EXT_INT", "F426+F389-L389", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F480"),
        new(4, "G480", "EXT_INT", "G389", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G480"),
        new(4, "H480", "EXT_INT", "H389", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H480"),
        new(4, "F481", "SUBS", "F426", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F481"),
        new(4, "F488", "SUB_EMP", "F447+F421+F383-L383-L421", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F488"),
        new(4, "G488", "SUB_TDF", "G421+G383", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G488"),
        new(4, "F489", "SUBS", "F447", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F489"),
        new(4, "F490", "EXT_INT", "F429+F396-L396", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F490"),
        new(4, "G490", "EXT_INT", "G396", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G490"),
        new(4, "F491", "SUBS", "F429", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F491"),
        new(4, "F498", "AFASEO", "F385", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F498"),
        new(5, "F559", "SUBS", "F552", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F559"),
        new(5, "F560", "EXT_INT", "F538+F512-L512", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F560"),
        new(5, "G560", "EXT_INT", "G512", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G560"),
        new(5, "H560", "EXT_INT", "H512", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H560"),
        new(5, "F561", "SUBS", "F538", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F561"),
        new(5, "F568", "SUB_EMP", "F541+F516-L516", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F568"),
        new(5, "G568", "SUB_TDF", "G516", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G568"),
        new(5, "H568", "SUB_TDF", "H516", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H568"),
        new(5, "F569", "SUBS", "F541", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F569"),
        new(5, "F573", "SUB_EMP", "F544+F520-L520", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F573"),
        new(5, "G573", "SUB_TDF", "G520", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G573"),
        new(5, "H573", "SUB_TDF", "H520", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H573"),
        new(5, "L573", "SUB_L", "L520", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!L573"),
        new(5, "F574", "SUBS", "F544", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F574"),
        new(5, "F578", "SUB_EMP", "F530+F551-L530", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F578"),
        new(5, "G578", "SUB_TDF", "G530", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G578"),
        new(5, "H578", "SUB_TDF", "H530", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H578"),
        new(5, "F579", "SUBS", "F551", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F579"),
        new(5, "F580", "EXT_INT", "F537+F511-L511", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F580"),
        new(5, "G580", "EXT_INT", "G511", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G580"),
        new(5, "H580", "EXT_INT", "H511", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!H580"),
        new(5, "F581", "SUBS", "F537", "Manual Julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F581")
    ];
}

/// <summary>
/// Plan 32 (T1, pieza a): verifica contra el disco que la congelación del INTERIOR R1 sigue vigente.
/// Si un manual cambia, este test lo delata (lectura ZIP+XML BCL, sin Excel/COM).
/// </summary>
public sealed class InterioresR1EsperadosTests
{
    public static IEnumerable<object[]> Agosto =>
        InterioresR1Esperados.Agosto.Select(v => new object[] { v.AseId, v.Celda, v.Etiqueta, v.Formula!, v.ValorLiteral! });

    public static IEnumerable<object[]> Julio =>
        InterioresR1Esperados.Julio.Select(v => new object[] { v.AseId, v.Celda, v.Etiqueta, v.Formula!, v.ValorLiteral! });

    [Theory]
    [MemberData(nameof(Agosto))]
    public void ManualAgosto_InteriorR1_CongeladoCoincideConDisco(int aseId, string celda, string etiqueta, string? formula, decimal? valorLiteral)
    {
        var manual = Path.Combine(Insumos.Raiz(), InterioresR1Esperados.ManualAgostoRelativo);
        AssertInterior(manual, aseId, celda, etiqueta, formula, valorLiteral);
    }

    [Theory]
    [MemberData(nameof(Julio))]
    public void ManualJulio_InteriorR1_CongeladoCoincideConDisco(int aseId, string celda, string etiqueta, string? formula, decimal? valorLiteral)
    {
        var manual = Path.Combine(Insumos.Raiz(), InterioresR1Esperados.ManualJulioRelativo);
        AssertInterior(manual, aseId, celda, etiqueta, formula, valorLiteral);
    }

    private static void AssertInterior(string manual, int aseId, string celda, string etiqueta, string? formula, decimal? valorLiteral)
    {
        Assert.True(File.Exists(manual), $"Falta el manual del administrativo: {manual}");
        var celdas = ValidadorTotalesR1Workbook.LeerCeldas(manual, InterioresR1Esperados.Hoja);

        Assert.True(celdas.TryGetValue(celda, out var interior), $"Falta la celda interior {celda} (ASE {aseId}, {etiqueta}) en el manual.");

        if (formula is null)
        {
            // Clase L-1: el manual trae literal (sin <f>).
            Assert.Null(interior!.Formula);
            if (valorLiteral is { } valor)
            {
                Assert.Equal(valor, interior.Numero ?? 0m);
            }
        }
        else
        {
            Assert.Equal(formula, interior!.Formula);
        }
    }
}
