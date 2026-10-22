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
    /// Semántica de un operando R1 en la fuente (Recaudoporcomponente). Plan 53 (D1/D2): cada celda del
    /// mapa declara ADEMÁS su <see cref="SeccionEsperadaR1"/> — la sección del bloque que debe
    /// alimentarla (nunca la ocurrencia ni el mayor/menor <c>|F|</c>: Plan 52 §13).
    /// El rol se resuelve con <see cref="ResolvedorOperandosEmpresaR1"/> sobre las secciones ya tipadas
    /// por <see cref="ContratoSeccionPeriodFreeR1.SeccionPorCierre"/> (Plan 46); la sección declarada
    /// ausente deja la celda en 0 y un empate de sección es STOP.
    /// Evidencia celda a celda: <c>Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md</c>.
    /// </summary>
    public enum FuenteR1
    {
        /// <summary>Fila Total del bloque que CIERRA (última sección presente).</summary>
        TotalMain,

        /// <summary>Especiales (L) de la fila Total principal.</summary>
        EspecialesMain,

        /// <summary>Fila Subsidio(-)/Contribucion(+) del bloque que cierra (última sección de Subs).</summary>
        SubsMain,

        /// <summary>Fila Total de la PRIMERA sección del bloque (el "otro" bloque; ausente → 0).</summary>
        TotalMenor,

        /// <summary>Especiales (L) de la fila Total de la primera sección.</summary>
        EspecialesMenor,

        /// <summary>Fila Subsidio(-)/Contribucion(+) de la primera sección de Subs (ausente → 0).</summary>
        SubsMenor,

        /// <summary>Fila Total de la sección intermedia (solo con ≥3 secciones; si no → 0).</summary>
        TotalMedio,

        /// <summary>Especiales (L) de la fila Total intermedia.</summary>
        EspecialesMedio,

        /// <summary>Bloque "5" del nuevo esquema (RECIPROCIDAD ASE2): canal <c>D=5</c>, sin sección.</summary>
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
    /// Celdas editables R1 por (empresa × ASE). Plan 53 (H1.3, D1/D2): cada entrada declara la celda del
    /// template, el operando de la fuente <b>y la sección que debe alimentarla</b>
    /// (<see cref="SeccionEsperadaR1"/>) — el contrato se deriva de la plantilla/manual y está congelado
    /// celda a celda en <c>Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md</c> §3 (56 entradas;
    /// triple evidencia: sección de la fuente + `<f>` consumidora donde existe + oráculo del manual).
    /// Solo las empresas con detalle no-estático.
    /// </summary>
    public static readonly IReadOnlyDictionary<(int EmpresaId, int AseId), (string Celda, FuenteR1 Fuente, SeccionEsperadaR1 Seccion)[]> EditablesR1PorEmpresa =
        new Dictionary<(int, int), (string, FuenteR1, SeccionEsperadaR1)[]>
        {
            // ASE1: ENEL F56 = F33+F14-L14; OCCIDENTE F66 = F40+F24-L24 + EXTEMP F68 = F29+F9-L9.
            [(2, 1)] =
            [
                ("F33", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F14", FuenteR1.TotalMain, SeccionEsperadaR1.UltimaOportuna),
                ("L14", FuenteR1.EspecialesMain, SeccionEsperadaR1.UltimaOportuna)
            ],
            [(4, 1)] =
            [
                ("F40", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F24", FuenteR1.TotalMain, SeccionEsperadaR1.UltimaOportuna),
                ("L24", FuenteR1.EspecialesMain, SeccionEsperadaR1.UltimaOportuna),
                ("F29", FuenteR1.SubsMenor, SeccionEsperadaR1.Primera),
                ("F9", FuenteR1.TotalMenor, SeccionEsperadaR1.Primera),
                ("L9", FuenteR1.EspecialesMenor, SeccionEsperadaR1.Primera)
            ],

            // ASE2: RECIP F181 = F102+F122+F86; ENEL F186 = F117+F95+F83-L95;
            //       OCCIDENTE F196 = F129+F112+F89-L112.
            [(1, 2)] =
            [
                ("F102", FuenteR1.TotalMain, SeccionEsperadaR1.UltimaOportuna),
                ("F122", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F86", FuenteR1.Block5Total, SeccionEsperadaR1.UltimaOportuna)
            ],
            [(2, 2)] =
            [
                ("F117", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F95", FuenteR1.TotalMain, SeccionEsperadaR1.UltimaOportuna),
                ("F83", FuenteR1.TotalMenor, SeccionEsperadaR1.Primera),
                ("L95", FuenteR1.EspecialesMain, SeccionEsperadaR1.UltimaOportuna)
            ],
            [(4, 2)] =
            [
                ("F129", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F112", FuenteR1.TotalMain, SeccionEsperadaR1.UltimaOportuna),
                ("F89", FuenteR1.TotalMenor, SeccionEsperadaR1.Primera),
                ("L112", FuenteR1.EspecialesMain, SeccionEsperadaR1.UltimaOportuna)
            ],

            // ASE3: ENEL F326 = F246+F227+F213-L227; OCCIDENTE F336 = F253+F237+F216-L237
            //       + EXTEMP F338 = F242+F222.
            [(2, 3)] =
            [
                ("F246", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F227", FuenteR1.TotalMain, SeccionEsperadaR1.UltimaOportuna),
                ("F213", FuenteR1.TotalMenor, SeccionEsperadaR1.Primera),
                ("L227", FuenteR1.EspecialesMain, SeccionEsperadaR1.UltimaOportuna)
            ],
            [(4, 3)] =
            [
                ("F253", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F237", FuenteR1.TotalMain, SeccionEsperadaR1.UltimaOportuna),
                ("F216", FuenteR1.TotalMenor, SeccionEsperadaR1.Primera),
                ("L237", FuenteR1.EspecialesMain, SeccionEsperadaR1.UltimaOportuna),
                ("F242", FuenteR1.SubsMenor, SeccionEsperadaR1.Primera),
                ("F222", FuenteR1.TotalMedio, SeccionEsperadaR1.Intermedia)
            ],

            // ASE4: RECIP F442 = F409+F380-L380; ENEL F447 = F404+F373+F353-L373
            //       + EXTEMP F449 = F395+F362; OCCIDENTE F457 = F416+F390+F356-L390 + EXTEMP F459 = F400+F368.
            [(1, 4)] =
            [
                ("F409", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F380", FuenteR1.TotalMain, SeccionEsperadaR1.UltimaOportuna),
                ("L380", FuenteR1.EspecialesMain, SeccionEsperadaR1.UltimaOportuna)
            ],
            [(2, 4)] =
            [
                ("F404", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F373", FuenteR1.TotalMain, SeccionEsperadaR1.UltimaOportuna),
                ("F353", FuenteR1.TotalMenor, SeccionEsperadaR1.Primera),
                ("L373", FuenteR1.EspecialesMain, SeccionEsperadaR1.UltimaOportuna),
                ("F395", FuenteR1.SubsMenor, SeccionEsperadaR1.Primera),
                ("F362", FuenteR1.TotalMedio, SeccionEsperadaR1.Intermedia)
            ],
            [(4, 4)] =
            [
                ("F416", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F390", FuenteR1.TotalMain, SeccionEsperadaR1.UltimaOportuna),
                ("F356", FuenteR1.TotalMenor, SeccionEsperadaR1.Primera),
                ("L390", FuenteR1.EspecialesMain, SeccionEsperadaR1.UltimaOportuna),
                ("F400", FuenteR1.SubsMenor, SeccionEsperadaR1.Primera),
                ("F368", FuenteR1.TotalMedio, SeccionEsperadaR1.Intermedia)
            ],

            // ASE5: ENEL F529 = F502+F483+F474-L483; ENERBIT F534 = F505+F487-L487;
            //       OCCIDENTE F539 = F512+F497+F477-L497.
            // Plan 53 (D5): las celdas Main/Especiales de ASE5 declaran `UnicaRecaudo` — el T0 probó que
            // en octubre-1 las 3 empresas traen UNA sola fila en la única sección de recaudo; el rol
            // prohíbe la lectura "única primera ⇒ Menor" (D3-del-52, derogada) y mantiene la identidad
            // `F483/F487/F497` (§9 #13–#17). `F474/F477` (Primera) quedan en 0 donde la fuente no trae
            // una primera sección distinta.
            [(2, 5)] =
            [
                ("F502", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F483", FuenteR1.TotalMain, SeccionEsperadaR1.UnicaRecaudo),
                ("F474", FuenteR1.TotalMenor, SeccionEsperadaR1.Primera),
                ("L483", FuenteR1.EspecialesMain, SeccionEsperadaR1.UnicaRecaudo)
            ],
            [(3, 5)] =
            [
                ("F505", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F487", FuenteR1.TotalMain, SeccionEsperadaR1.UnicaRecaudo),
                ("L487", FuenteR1.EspecialesMain, SeccionEsperadaR1.UnicaRecaudo)
            ],
            [(4, 5)] =
            [
                ("F512", FuenteR1.SubsMain, SeccionEsperadaR1.UltimaOportuna),
                ("F497", FuenteR1.TotalMain, SeccionEsperadaR1.UnicaRecaudo),
                ("F477", FuenteR1.TotalMenor, SeccionEsperadaR1.Primera),
                ("L497", FuenteR1.EspecialesMain, SeccionEsperadaR1.UnicaRecaudo)
            ]
        };

    /// <summary>
    /// Celdas editables R2 por (empresa × ASE). Visible = Total + SubsCont − Especiales.
    /// </summary>
    public static readonly IReadOnlyDictionary<(int EmpresaId, int AseId), (string Celda, FuenteR2 Fuente)[]> EditablesR2PorEmpresa =
        new Dictionary<(int, int), (string, FuenteR2)[]>
        {
            // Plan 38 (T2-continuación, R-ATRIBUCION): re-congeladas desde los bytes de la plantilla
            // estandarizada única (mini-T0). Por empresa/ASE: la celda "Total" es la fila-dato de la
            // empresa en el sub-bloque Componente y la "SubsCont" en el sub-bloque Subs/Cont; la
            // "Especiales" es la col K de la fila Total. Derivadas de las fórmulas OPORTUNO del
            // template (p. ej. ASE1 ENEL E53=E24+E10+E6-K6-K10 → Total E10, SubsCont E24, Esp K10).
            [(2, 1)] = [("E24", FuenteR2.SubsCont), ("E10", FuenteR2.Total), ("K10", FuenteR2.Especiales)],
            [(4, 1)] = [("E20", FuenteR2.Total), ("E31", FuenteR2.SubsCont), ("K20", FuenteR2.Especiales)],
            [(1, 2)] = [("E104", FuenteR2.SubsCont), ("E83", FuenteR2.Total), ("K83", FuenteR2.Especiales)],
            [(2, 2)] = [("E97", FuenteR2.SubsCont), ("E74", FuenteR2.Total), ("K74", FuenteR2.Especiales)],
            [(4, 2)] = [("E93", FuenteR2.Total), ("E111", FuenteR2.SubsCont), ("K93", FuenteR2.Especiales)],
            [(2, 3)] = [("E185", FuenteR2.SubsCont), ("E171", FuenteR2.Total), ("K171", FuenteR2.Especiales)],
            [(4, 3)] = [("E181", FuenteR2.Total), ("E192", FuenteR2.SubsCont), ("K181", FuenteR2.Especiales)],
            [(1, 4)] = [("E319", FuenteR2.SubsCont), ("E285", FuenteR2.Total), ("K285", FuenteR2.Especiales)],
            [(2, 4)] = [("E312", FuenteR2.SubsCont), ("E289", FuenteR2.Total), ("K289", FuenteR2.Especiales)],
            [(4, 4)] = [("E308", FuenteR2.Total), ("E326", FuenteR2.SubsCont), ("K308", FuenteR2.Especiales)],
            [(2, 5)] = [("E407", FuenteR2.SubsCont), ("E389", FuenteR2.Total), ("K389", FuenteR2.Especiales)],
            [(4, 5)] = [("E403", FuenteR2.Total), ("E417", FuenteR2.SubsCont), ("K403", FuenteR2.Especiales)]
        };

    /// <summary>
    /// Celdas editables R4 por (empresa × ASE). Visible = Total − P.
    /// </summary>
    public static readonly IReadOnlyDictionary<(int EmpresaId, int AseId), (string Celda, FuenteR4 Fuente)[]> EditablesR4PorEmpresa =
        new Dictionary<(int, int), (string, FuenteR4)[]>
        {
            // Plan 38 (T2-continuación, R-ATRIBUCION): direcciones re-congeladas desde los bytes de
            // la plantilla estandarizada única (mini-T0, GeometriaPlantillaEstandarizadaT0Tests).
            // Cada celda D es la fila-dato "Total" de la empresa en el bloque de detalle del ASE,
            // la misma que referencian las fórmulas OPORTUNO del bloque (p. ej. ASE3 ENEL
            // D217=D199-P199). El inventario de empresas por ASE lo manda la FUENTE R4 (Q1):
            // ASE1/3/5 ENEL+OCCIDENTE; ASE2 ENEL+OCCIDENTE; ASE4 ENEL+NUEVO ESQUEMA+OCCIDENTE.
            [(4, 1)] = [("D14", FuenteR4.Total), ("P14", FuenteR4.P)],
            [(1, 2)] = [("D104", FuenteR4.Total), ("P104", FuenteR4.P)],
            [(2, 2)] = [("D98", FuenteR4.Total), ("P98", FuenteR4.P)],
            [(2, 3)] = [("D199", FuenteR4.Total), ("P199", FuenteR4.P)],
            [(4, 3)] = [("D205", FuenteR4.Total), ("P205", FuenteR4.P)],
            [(1, 4)] = [("D242", FuenteR4.Total), ("P242", FuenteR4.P)],
            [(2, 4)] = [("D236", FuenteR4.Total), ("P236", FuenteR4.P)],
            [(4, 4)] = [("D248", FuenteR4.Total), ("P248", FuenteR4.P)],
            [(2, 5)] = [("D351", FuenteR4.Total), ("P351", FuenteR4.P)],
            [(4, 5)] = [("D357", FuenteR4.Total), ("P357", FuenteR4.P)]
        };

    /// <summary>
    /// Celdas de las hojas <c>Recaudo *</c> (zona de datos filas ~3–28) por hoja.
    /// Fuente 1:1: <c>{periodo}/Conciliaciones/Conjunta {prefijo}*.xlsx</c> hoja
    /// <c>RESUMEN MES</c> (HU-20/G1-D1). Las letras destino dependen de la quincena
    /// (D/E en Q1, F/G en Q2; G2-D2) y las construye el reader.
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
            [3] = [89], // Enerbit
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
