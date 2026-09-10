using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// HU-08 (2.2): mapa explícito empresa × ASE de celdas editables, visibles esperados y fórmulas
/// protegidas. Congelado por T0 (dumps OpenXML celda-por-celda del golden
/// <c>Docs/Insumos/Remuneracion 202607-1 Total.xlsx</c> + fuentes Q1 de los 5 ASE).
///
/// REGLA DEL PLAN (G3): sin offsets aritméticos — cada celda se declara explícitamente. Layouts
/// heterogéneos (R1 col C con fórmulas de 3/5 términos según ASE; R2 col B; R4 col A; EAAB-CL
/// con 5 fórmulas menos porque sus filas CIUDAD LIMPIA son valor estático 0).
///
/// La semántica de cada operando (de dónde sale su valor en la fuente) usa el enum
/// <see cref="FuenteR1"/>/<see cref="FuenteR2"/>/<see cref="FuenteR4"/>. El reader resuelve el
/// valor; el writer escribe SOLO estas celdas.
/// </summary>
public static class WorkbookLeafCellMapPorEmpresa
{
    public const string HojaR1 = WorkbookLeafCellMap.HojaR1;
    public const string HojaR2 = WorkbookLeafCellMap.HojaR2;
    public const string HojaR4 = WorkbookLeafCellMap.HojaR4;

    /// <summary>
    /// Semántica de un operando R1 en la fuente (Recaudoporcomponente).
    /// Main = mayor |valor| de la empresa; Menor/Medio = los otros bloques de la misma empresa.
    /// </summary>
    public enum FuenteR1
    {
        /// <summary>Fila Total principal de la empresa (mayor |F|).</summary>
        TotalMain,

        /// <summary>Especiales (L) de la fila Total principal.</summary>
        EspecialesMain,

        /// <summary>Fila Subsidio(-)/Contribucion(+) principal (mayor |F|).</summary>
        SubsMain,

        /// <summary>Primer bloque Total de la empresa (menor |F|).</summary>
        TotalMenor,

        /// <summary>Especiales (L) del primer bloque Total.</summary>
        EspecialesMenor,

        /// <summary>Primer bloque Subsidio(-)/Contribucion(+) (menor |F|).</summary>
        SubsMenor,

        /// <summary>Bloque Total intermedio (caso ASE3/ASE4: 3 filas Total).</summary>
        TotalMedio,

        /// <summary>Especiales (L) del bloque Total intermedio.</summary>
        EspecialesMedio,

        /// <summary>Bloque "5" del nuevo esquema (RECIPROCIDAD ASE2).</summary>
        Block5Total
    }

    /// <summary>
    /// Semántica de un operando R2 en la fuente (RerpoteDetalleSaldosaFavor).
    /// Uniforme en los 5 bloques: visible empresa = Total + SubsCont − Especiales.
    /// </summary>
    public enum FuenteR2
    {
        Total,
        SubsCont,
        Especiales
    }

    /// <summary>
    /// Semántica de un operando R4 en la fuente (ReversiónPorComponente).
    /// Uniforme: visible empresa = Total − P (P = 0 en Q1).
    /// </summary>
    public enum FuenteR4
    {
        Total,
        P
    }

    /// <summary>
    /// Celdas editables R1 por (empresa × ASE). Valor = la celda del template que recibe el
    /// operando y su semántica de fuente. Solo las empresas con detalle no-estático.
    /// </summary>
    public static readonly IReadOnlyDictionary<(int EmpresaId, int AseId), (string Celda, FuenteR1 Fuente)[]> EditablesR1PorEmpresa =
        new Dictionary<(int, int), (string, FuenteR1)[]>
        {
            // ASE1: ENEL F56 = F33+F14-L14; OCCIDENTE F66 = F40+F24-L24 + EXTEMP F68 = F29+F9-L9.
            [(2, 1)] = [( "F33", FuenteR1.SubsMain), ("F14", FuenteR1.TotalMain), ("L14", FuenteR1.EspecialesMain)],
            [(4, 1)] =
            [
                ("F40", FuenteR1.SubsMain), ("F24", FuenteR1.TotalMain), ("L24", FuenteR1.EspecialesMain),
                ("F29", FuenteR1.SubsMenor), ("F9", FuenteR1.TotalMenor), ("L9", FuenteR1.EspecialesMenor)
            ],

            // ASE2: RECIP F181 = F102+F122+F86; ENEL F186 = F117+F95+F83-L95;
            //       OCCIDENTE F196 = F129+F112+F89-L112.
            [(1, 2)] = [("F102", FuenteR1.TotalMain), ("F122", FuenteR1.SubsMain), ("F86", FuenteR1.Block5Total)],
            [(2, 2)] = [("F117", FuenteR1.SubsMain), ("F95", FuenteR1.TotalMain), ("F83", FuenteR1.TotalMenor), ("L95", FuenteR1.EspecialesMain)],
            [(4, 2)] = [("F129", FuenteR1.SubsMain), ("F112", FuenteR1.TotalMain), ("F89", FuenteR1.TotalMenor), ("L112", FuenteR1.EspecialesMain)],

            // ASE3: ENEL F326 = F246+F227+F213-L227; OCCIDENTE F336 = F253+F237+F216-L237
            //       + EXTEMP F338 = F242+F222.
            [(2, 3)] = [("F246", FuenteR1.SubsMain), ("F227", FuenteR1.TotalMain), ("F213", FuenteR1.TotalMenor), ("L227", FuenteR1.EspecialesMain)],
            [(4, 3)] =
            [
                ("F253", FuenteR1.SubsMain), ("F237", FuenteR1.TotalMain), ("F216", FuenteR1.TotalMenor),
                ("L237", FuenteR1.EspecialesMain), ("F242", FuenteR1.SubsMenor), ("F222", FuenteR1.TotalMedio)
            ],

            // ASE4: RECIP F442 = F409+F380-L380; ENEL F447 = F404+F373+F353-L373
            //       + EXTEMP F449 = F395+F362; OCCIDENTE F457 = F416+F390+F356-L390 + EXTEMP F459 = F400+F368.
            [(1, 4)] = [("F409", FuenteR1.SubsMain), ("F380", FuenteR1.TotalMain), ("L380", FuenteR1.EspecialesMain)],
            [(2, 4)] =
            [
                ("F404", FuenteR1.SubsMain), ("F373", FuenteR1.TotalMain), ("F353", FuenteR1.TotalMenor),
                ("L373", FuenteR1.EspecialesMain), ("F395", FuenteR1.SubsMenor), ("F362", FuenteR1.TotalMedio)
            ],
            [(4, 4)] =
            [
                ("F416", FuenteR1.SubsMain), ("F390", FuenteR1.TotalMain), ("F356", FuenteR1.TotalMenor),
                ("L390", FuenteR1.EspecialesMain), ("F400", FuenteR1.SubsMenor), ("F368", FuenteR1.TotalMedio)
            ],

            // ASE5: ENEL F529 = F502+F483+F474-L483; ENERBIT F534 = F505+F487-L487;
            //       OCCIDENTE F539 = F512+F497+F477-L497.
            [(2, 5)] = [("F502", FuenteR1.SubsMain), ("F483", FuenteR1.TotalMain), ("F474", FuenteR1.TotalMenor), ("L483", FuenteR1.EspecialesMain)],
            [(3, 5)] = [("F505", FuenteR1.SubsMain), ("F487", FuenteR1.TotalMain), ("L487", FuenteR1.EspecialesMain)],
            [(4, 5)] = [("F512", FuenteR1.SubsMain), ("F497", FuenteR1.TotalMain), ("F477", FuenteR1.TotalMenor), ("L497", FuenteR1.EspecialesMain)]
        };

    /// <summary>
    /// Celdas editables R2 por (empresa × ASE). Visible = Total + SubsCont − Especiales.
    /// </summary>
    public static readonly IReadOnlyDictionary<(int EmpresaId, int AseId), (string Celda, FuenteR2 Fuente)[]> EditablesR2PorEmpresa =
        new Dictionary<(int, int), (string, FuenteR2)[]>
        {
            [(2, 1)] = [("E18", FuenteR2.SubsCont), ("E6", FuenteR2.Total), ("K6", FuenteR2.Especiales)],
            [(4, 1)] = [("E14", FuenteR2.Total), ("E25", FuenteR2.SubsCont), ("K14", FuenteR2.Especiales)],
            [(1, 2)] = [("E95", FuenteR2.SubsCont), ("E74", FuenteR2.Total), ("K74", FuenteR2.Especiales)],
            [(2, 2)] = [("E88", FuenteR2.SubsCont), ("E64", FuenteR2.Total), ("K64", FuenteR2.Especiales)],
            [(4, 2)] = [("E84", FuenteR2.Total), ("E102", FuenteR2.SubsCont), ("K84", FuenteR2.Especiales)],
            [(2, 3)] = [("E170", FuenteR2.SubsCont), ("E158", FuenteR2.Total), ("K158", FuenteR2.Especiales)],
            [(4, 3)] = [("E166", FuenteR2.Total), ("E177", FuenteR2.SubsCont), ("K166", FuenteR2.Especiales)],
            [(1, 4)] = [("E300", FuenteR2.SubsCont), ("E279", FuenteR2.Total), ("K279", FuenteR2.Especiales)],
            [(2, 4)] = [("E293", FuenteR2.SubsCont), ("E270", FuenteR2.Total), ("K270", FuenteR2.Especiales)],
            [(4, 4)] = [("E289", FuenteR2.Total), ("E307", FuenteR2.SubsCont), ("K289", FuenteR2.Especiales)],
            [(2, 5)] = [("E377", FuenteR2.SubsCont), ("E366", FuenteR2.Total), ("K366", FuenteR2.Especiales)],
            [(4, 5)] = [("E373", FuenteR2.Total), ("E384", FuenteR2.SubsCont), ("K373", FuenteR2.Especiales)]
        };

    /// <summary>
    /// Celdas editables R4 por (empresa × ASE). Visible = Total − P.
    /// </summary>
    public static readonly IReadOnlyDictionary<(int EmpresaId, int AseId), (string Celda, FuenteR4 Fuente)[]> EditablesR4PorEmpresa =
        new Dictionary<(int, int), (string, FuenteR4)[]>
        {
            [(4, 1)] = [("D8", FuenteR4.Total), ("P8", FuenteR4.P)],
            [(1, 2)] = [("D97", FuenteR4.Total), ("P97", FuenteR4.P)],
            [(2, 2)] = [("D92", FuenteR4.Total), ("P92", FuenteR4.P)],
            [(2, 3)] = [("D186", FuenteR4.Total), ("P186", FuenteR4.P)],
            [(4, 3)] = [("D192", FuenteR4.Total), ("P192", FuenteR4.P)],
            [(1, 4)] = [("D229", FuenteR4.Total), ("P229", FuenteR4.P)],
            [(2, 4)] = [("D223", FuenteR4.Total), ("P223", FuenteR4.P)],
            [(4, 4)] = [("D235", FuenteR4.Total), ("P235", FuenteR4.P)],
            [(2, 5)] = [("D337", FuenteR4.Total), ("P337", FuenteR4.P)],
            [(4, 5)] = [("D343", FuenteR4.Total), ("P343", FuenteR4.P)]
        };

    /// <summary>
    /// Celdas de las hojas <c>Recaudo *</c> (zona de datos filas ~3–28) por hoja.
    /// Fuente 1:1: <c>Consolidado/Conciliaciones/Conjunta {prefijo}*.xlsx</c> hoja
    /// <c>RESUMEN MES</c> (T0-0.6). D = valor, E = número de registros.
    /// </summary>
    public static readonly IReadOnlyDictionary<string, string[]> CeldasRecaudoPorHoja =
        new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            ["Recaudo EAAB Reciprocidad"] = CeldasRecaudoBasicas(),
            ["Recaudo ENEL"] = CeldasRecaudoBasicas(),
            ["Recaudo ENERBIT"] = CeldasRecaudoBasicas(),
            ["Recaudo EAAB + Ciud Limp"] = CeldasRecaudoBasicas(),
            ["Recaudo Directa Occidente"] = CeldasRecaudoBasicas()
        };

    /// <summary>
    /// Filas del bloque AJUSTES (D85:D89) que son FÓRMULA por empresa (T0). Las que no están
    /// aquí son valor estático 0 en el golden (RECIP alterna; OCCIDENTE/EAAB-CL todo valor).
    /// Declarado ANTES de <see cref="ProtectedFormulasPorEmpresa"/> por orden de inicialización estática.
    /// </summary>
    private static readonly IReadOnlyDictionary<int, int[]> AjustesFormulaPorEmpresa =
        new Dictionary<int, int[]>
        {
            [1] = [86, 88], // ReciprocidadEaab
            [2] = [85, 87, 89], // Enel
            [3] = [85, 87, 89], // Enerbit
            [4] = [], // OccidenteDirecta (todo valor 0)
            [5] = [] // EaabCiudadLimpia (todo valor 0)
        };

    /// <summary>
    /// Fórmulas protegidas 2.2: hojas <c>REMUNERACION_*</c> (100 % fórmulas, V3), filas de
    /// validación de las hojas <c>Recaudo *</c> (fila 29+), y celdas representativas de
    /// <c>VALIDACION_*</c>/<c>GERENTES_*</c> (V4). El writer falla si alguna deja de ser fórmula.
    /// </summary>
    public static readonly (string Hoja, string Celda, string[] Fragmentos)[] ProtectedFormulasPorEmpresa = CrearProtegidasPorEmpresa();

    /// <summary>
    /// Celdas base de la zona de datos de una hoja Recaudo * (OPORTUNO 3–9, EXTEMP 12–18, TOTAL 21–27).
    /// </summary>
    private static string[] CeldasRecaudoBasicas()
    {
        var celdas = new List<string>();
        foreach (var (inicio, fin) in new[] { (3, 9), (12, 18), (21, 27) })
        {
            for (var fila = inicio; fila <= fin; fila++)
            {
                celdas.Add($"D{fila}");
                celdas.Add($"E{fila}");
            }
        }

        return celdas.ToArray();
    }

    private static (string Hoja, string Celda, string[] Fragmentos)[] CrearProtegidasPorEmpresa()
    {
        var lista = new List<(string, string, string[])>();
        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            // REMUNERACION_*: bloques R1 (D9:D13 + D14 SUM), R2 (D28:D32 + D33 SUM), EXTEMP (D47:D51 + D52),
    // R4 (D66:D70 + D71), AJUSTES (heterogéneo: T0 — RECIP D86/D88, ENEL/ENERBIT D85/D87/D89,
    // OCCIDENTE/EAAB-CL todo valor 0), consolidado (D104:D108 + D109) — solo las celdas fórmula.
            foreach (var fila in Enumerable.Range(9, 5))
            {
                lista.Add((empresa.HojaRemuneracion, $"D{fila}", ["Reporte Componentes R1"]));
            }

            lista.Add((empresa.HojaRemuneracion, "D14", ["SUM"]));

            foreach (var fila in Enumerable.Range(28, 5))
            {
                lista.Add((empresa.HojaRemuneracion, $"D{fila}", ["Rem. Anticipos R2"]));
            }

            lista.Add((empresa.HojaRemuneracion, "D33", ["SUM"]));

            foreach (var fila in Enumerable.Range(47, 5))
            {
                lista.Add((empresa.HojaRemuneracion, $"D{fila}", ["Reporte Componentes R1"]));
            }

            lista.Add((empresa.HojaRemuneracion, "D52", ["SUM"]));

            foreach (var fila in Enumerable.Range(66, 5))
            {
                lista.Add((empresa.HojaRemuneracion, $"D{fila}", ["Reversion Pagos R4"]));
            }

            lista.Add((empresa.HojaRemuneracion, "D71", ["SUM"]));

            foreach (var fila in AjustesFormulaPorEmpresa[empresa.Id])
            {
                lista.Add((empresa.HojaRemuneracion, $"D{fila}", ["AJUSTES - SF-T"]));
            }

            lista.Add((empresa.HojaRemuneracion, "D90", ["SUM"]));

            foreach (var fila in Enumerable.Range(104, 5))
            {
                lista.Add((empresa.HojaRemuneracion, $"D{fila}", ["D"]));
            }

            lista.Add((empresa.HojaRemuneracion, "D109", ["SUM"]));

            // Recaudo *: fila 29+ (validaciones =SUM) protegidas.
            foreach (var celda in new[] { "D31", "E31", "D33", "E33", "D35", "E35" })
            {
                lista.Add((empresa.HojaRecaudo, celda, ["SUM"]));
            }

            // VALIDACION_* / GERENTES_*: celdas representativas (fórmulas puras, V4/T0-0.5).
            lista.Add((empresa.HojaValidacion, "O3", ["H3", "N3"]));
            lista.Add((empresa.HojaGerentes, "D9", ["REMUNERACION"]));
        }

        return lista.ToArray();
    }
}
