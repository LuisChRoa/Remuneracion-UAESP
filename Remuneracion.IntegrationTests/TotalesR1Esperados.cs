using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 31 (T1, pieza a — congelar E2): textos de fórmula de los <b>totales visibles del bloque R1</b>
/// (TOT_OPT F, total TDF G, EXTEMP F) por ASE, transcritos EXACTAMENTE de los manuales del
/// administrativo (solo lectura ZIP+XML BCL, sin Excel/COM) y con cita de archivo + celda por valor.
///
/// Fuente agosto: <c>Docs/Prueba Agosto-2/Resultado/Resultado Manual por el administrativo/Remuneracion 202608-2 Total_7721.xlsx</c>
/// Fuente julio:  <c>Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx</c>
///
/// <see cref="TotalesR1EsperadosTests"/> re-lee el disco y delata cualquier drift del manual contra
/// esta congelación (si el disco cambia, el test falla antes de que el gate se confunda).
///
/// Nota de forma (el disco MANDA sobre el plan §2.3): el total TDF se compone como
/// <c>ΣG(Mes) − G(última Mes por orden físico)</c>, NO como <c>ΣG(todas las Mes)</c>; y la resta
/// de especiales del TOT_OPT también excluye la última Mes. El caso ASE3-agosto no trae filas
/// <c>Aplicacion</c>: su EXTEMP visible es el literal 0 (no fórmula).
///
/// Sub-visibles de empresa (RECIPROCIDAD/ENEL/…): el gate T1 no los consume (solo valida TOT_OPT,
/// TDF y EXTEMP del bloque), por eso no se congelan aquí.
/// </summary>
internal static class TotalesR1Esperados
{
    /// <summary>Total visible del bloque R1: fila (columna F/G) + texto de fórmula esperado.</summary>
    /// <param name="AseId">ASE (1..5).</param>
    /// <param name="Celda">Referencia en el manual (p. ej. <c>F50</c>).</param>
    /// <param name="Etiqueta">TOT_OPT | TDF | EXTEMP.</param>
    /// <param name="Formula">Texto de <c>&lt;f&gt;</c> esperado; <c>null</c> = literal (0-Aplic).</param>
    /// <param name="Cita">Archivo + hoja + celda (trazabilidad de la transcripción).</param>
    public sealed record VisibleR1(int AseId, string Celda, string Etiqueta, string? Formula, string Cita);

    public const string Hoja = "Reporte Componentes R1";
    public const string ManualAgostoRelativo = @"Docs\Prueba Agosto-2\Resultado\Resultado Manual por el administrativo\Remuneracion 202608-2 Total_7721.xlsx";
    public const string ManualJulioRelativo = @"Docs\Prueba Julio-2\Resultado\Remuneracion 202607-2 Total Administrativo.xlsx";

    /// <summary>Manual agosto 202608-2 (E2 congelado: TOT_OPT / TDF / EXTEMP por ASE).</summary>
    public static IReadOnlyList<VisibleR1> Agosto { get; } =
    [
        new(1, "F50", "TOT_OPT", "F29+F45+F9-L9-L29", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F50"),
        new(1, "G50", "TDF", "G29+G9", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G50"),
        new(1, "F52", "EXTEMP", "F34+F14-L14", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F52"),
        new(2, "F194", "TOT_OPT", "F125+F148-L125+F93-L93", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F194"),
        new(2, "G194", "TDF", "G125+G93", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G194"),
        new(2, "F196", "EXTEMP", "F130+F101-L101", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F196"),
        new(3, "F325", "TOT_OPT", "F251+F263+F235-L235-L251", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F325"),
        new(3, "G325", "TDF", "G251+G235", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G325"),
        new(3, "F327", "EXTEMP", null, "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F327 (literal 0; sin filas Aplicacion)"),
        new(4, "F456", "TOT_OPT", "F407+F436+F366-L366-L407", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F456"),
        new(4, "G456", "TDF", "G407+G366", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G456"),
        new(4, "F458", "EXTEMP", "F418+F382-L382", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F458"),
        new(5, "F554", "TOT_OPT", "F548+F525+F494-L494-L525", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F554 (ASE5 con 3 Mes)"),
        new(5, "G554", "TDF", "G525+G494", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!G554"),
        new(5, "F556", "EXTEMP", "F534+F506-L506", "Manual agosto 'Remuneracion 202608-2 Total_7721.xlsx' Reporte Componentes R1!F556")
    ];

    /// <summary>Manual julio 202607-2 (E2 congelado: TOT_OPT / TDF / EXTEMP por ASE).</summary>
    public static IReadOnlyList<VisibleR1> Julio { get; } =
    [
        new(1, "F53", "TOT_OPT", "F32+F48+F12-L12-L32", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F53"),
        new(1, "G53", "TDF", "G32+G12", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G53"),
        new(1, "F55", "EXTEMP", "F37+F17-L17", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F55"),
        new(2, "F206", "TOT_OPT", "F132+F160-L132+F94-L94", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F206"),
        new(2, "G206", "TDF", "G132+G94", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G206"),
        new(2, "F208", "EXTEMP", "F142+F107-L107", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F208"),
        new(3, "F343", "TOT_OPT", "F265+F281+F244-L244-L265", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F343"),
        new(3, "G343", "TDF", "G265+G244", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G343"),
        new(3, "F345", "EXTEMP", "F270+F250-L250", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F345"),
        new(4, "F468", "TOT_OPT", "F422+F448+F384-L384-L422", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F468"),
        new(4, "G468", "TDF", "G422+G384", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G468"),
        new(4, "F470", "EXTEMP", "F430+F397-L397", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F470"),
        new(5, "F558", "TOT_OPT", "F552+F531-L531", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F558 (ASE5 con 2 Mes)"),
        new(5, "G558", "TDF", "G531", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!G558"),
        new(5, "F560", "EXTEMP", "F538+F512-L512", "Manual julio 'Remuneracion 202607-2 Total Administrativo.xlsx' Reporte Componentes R1!F560")
    ];
}

/// <summary>
/// Plan 31 (T1, pieza a): verifica contra el disco que la congelación E2 de
/// <see cref="TotalesR1Esperados"/> sigue vigente. Si un manual cambia, este test lo delata
/// (lectura ZIP+XML BCL, sin Excel/COM).
/// </summary>
public sealed class TotalesR1EsperadosTests
{
    public static IEnumerable<object[]> Agosto =>
        TotalesR1Esperados.Agosto.Select(v => new object[] { v.AseId, v.Celda, v.Etiqueta, v.Formula! });

    public static IEnumerable<object[]> Julio =>
        TotalesR1Esperados.Julio.Select(v => new object[] { v.AseId, v.Celda, v.Etiqueta, v.Formula! });

    [Theory]
    [MemberData(nameof(Agosto))]
    public void ManualAgosto_TotalVisibleR1_CongeladoCoincideConDisco(int aseId, string celda, string etiqueta, string? formula)
    {
        var manual = Path.Combine(Insumos.Raiz(), TotalesR1Esperados.ManualAgostoRelativo);
        AssertVisible(manual, aseId, celda, etiqueta, formula);
    }

    [Theory]
    [MemberData(nameof(Julio))]
    public void ManualJulio_TotalVisibleR1_CongeladoCoincideConDisco(int aseId, string celda, string etiqueta, string? formula)
    {
        var manual = Path.Combine(Insumos.Raiz(), TotalesR1Esperados.ManualJulioRelativo);
        AssertVisible(manual, aseId, celda, etiqueta, formula);
    }

    private static void AssertVisible(string manual, int aseId, string celda, string etiqueta, string? formula)
    {
        Assert.True(File.Exists(manual), $"Falta el manual del administrativo: {manual}");
        var celdas = ValidadorTotalesR1Workbook.LeerCeldas(manual, TotalesR1Esperados.Hoja);

        Assert.True(celdas.TryGetValue(celda, out var visible), $"Falta el visible {celda} (ASE {aseId}, {etiqueta}) en el manual.");

        if (formula is null)
        {
            // 0-Aplic explícito (S5): el manual trae el literal 0, no una fórmula.
            Assert.Null(visible!.Formula);
            Assert.Equal(0m, visible.Numero);
        }
        else
        {
            Assert.Equal(formula, visible!.Formula);
        }
    }
}
