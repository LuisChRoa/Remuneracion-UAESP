namespace Remuneracion.Infrastructure.Excel;

using Remuneracion.Core.Models;

/// <summary>
/// HU-12 (2.6 ampliada): mapa de celdas Q2 por ASE/período congelado por T0 réplica-HU-07
/// (evidencia en <c>plans/12 - HU-12-detretri-Q2-r1-ase5.md</c> §4 Fase 0 + dumps OpenXML del
/// canónico <c>Docs/Insumos/REMUNERACION 2026072/Plantilla 8 agos 2026 _ Remuneracion 202607-2
/// Total.xlsx</c> y del caché golden <c>Docs/Insumos/Remuneracion 202607-2 Total.xlsx</c>).
///
/// Veredictos T0 (probados, no asumidos):
/// - V0.1: CONSOLIDADO Q2 D9:D13 → R1!F53/F206/F343/F468/F558; D47:D51 → R1!F55/F208/F345/F470/F560;
///   D66:D70 → R4!D73/D168/D205/D320/D355; D28:D32 → R2!E43/E139/E256/E358/E438;
///   D85:D89 → 'AJUSTES - SF-T'!D47:D51 (igual forma que Q1).
/// - T0-0.2: todos los operandos de los visibles Q2 son VALORES editables (cero en el template);
///   los visibles son FÓRMULA con las cadenas exactas (F53=F32+F48+F12-L12-L32, …).
/// - T0-0.3: mapeo fuente→template por rol de fila (Mes0/Mes1/Mes2, Aplic0/Aplic1) verificado
///   contra el caché golden en los 5 ASE (R1/R2/R4) y contra V0.3 en ASE5 (F558=F552+F531-L531 =
///   12033011685.71 = D13).
/// - T0-0.4: hojas DetRetri2026072/DetValiRetri2026072 (sheets 36/37): D9:D14 = VALORES editables
///   (0 en template; ROUND de D104:D109 en el golden); D23:D28/D32:D36 (DetRetri) y
///   D16:D21/D24:D29 (DetValiRetri) = FÓRMULA → protegidas.
/// - T0-0.5: M1 estructural: ProtegidasBceParaPeriodo(true) y el mapa Q2 matchean celdas reales
///   del canónico (BCE H3:H7 = DetRetri2026072!J9..J13, CONSOLIDADO J9:M13, INTERVENTORIA,
///   ANT EXT-REV, VALIDACION_*, GERENTES_*, REMUNERACION_* genéricas).
///
/// Regla de oro (plan D1/G2): mapa hermano explícito por (<see cref="Ase.Id"/>, período);
/// PROHIBIDOS offsets aritméticos; el mapa Q1 (<see cref="WorkbookLeafCellMapPorAse"/>) queda
/// INTACTO y solo se activa con <see cref="Periodo.NumeroQuincena"/> == 1.
/// </summary>
public static class WorkbookLeafCellMapQ2
{
    public const string HojaConsolidado = WorkbookLeafCellMap.HojaConsolidado;
    public const string HojaR1 = WorkbookLeafCellMap.HojaR1;
    public const string HojaR2 = WorkbookLeafCellMap.HojaR2;
    public const string HojaR4 = WorkbookLeafCellMap.HojaR4;
    public const string HojaInterventoria = "INTERVENTORIA";
    public const string HojaAntExtRev = "ANT EXT-REV";

    /// <summary>
    /// Hoja DetRetri del período (D-A, Plan 30): se compone del dominio vía
    /// <see cref="NombresHojaPeriodo.DetRetri(Periodo)"/>, nunca se detecta enumerando hojas.
    /// </summary>
    public static string HojaDetRetri(Periodo periodo) => NombresHojaPeriodo.DetRetri(periodo);

    /// <summary>
    /// Hoja DetValiRetri del período (D-A, Plan 30): se compone del dominio vía
    /// <see cref="NombresHojaPeriodo.DetValiRetri(Periodo)"/>.
    /// </summary>
    public static string HojaDetValiRetri(Periodo periodo) => NombresHojaPeriodo.DetValiRetri(periodo);

    /// <summary>
    /// Rol de la fila de la fuente R1-Q2 que alimenta una celda operando del template (T0-0.3).
    /// Mes0/Mes1/Mes2 = filas "Mes/Total" en orden de aparición (ASE5 solo trae 2 → Mes2 ausente).
    /// Aplic0/Aplic1 = filas "Aplicacion nuevos x reversion" en orden (opcionales); L* = columna
    /// Especiales (SERVICIO ESPECIALES) de la fila.
    ///
    /// Plan 25 (D-A, T2): se retira <c>Subs0</c> — <c>F37/F270</c> eran filas <c>Aplicacion</c>
    /// mal bindeadas (coincidían con <c>Subs0</c> en julio por bloques monocompañía). La resolución
    /// por firma vive en el lector; aquí solo se declaran celdas → rol.
    /// </summary>
    public enum R1Q2Fuente
    {
        Mes0, Mes1, Mes2, Lmes0, Lmes1, Lmes2, Aplic0, Aplic1, LAplic0
    }

    /// <summary>
    /// Operandos editables R1-Q2 por ASE: celda del template → rol de la fuente (T0-0.2/0.3).
    /// Congelado del dump OpenXML: F53=F32+F48+F12-L12-L32, F55=F37+F17-L17, F206=…, F560=…
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Celda, R1Q2Fuente Fuente)[]> R1Q2EditablesPorAse =
        new Dictionary<int, (string, R1Q2Fuente)[]>
        {
            // Plan 38 (T2-continuación, D-B): direcciones re-congeladas desde los bytes de la
            // plantilla estandarizada única (mini-T0). Las celdas son los operandos-dato de la
            // fórmula TOT_OPT/EXTEMP del template por ASE (p. ej. ASE1 F75=F47+F70+F16-L16-L47 →
            // Mes0/1/2 = F16/F47/F70; EXTEMP F77=F59+F32-L32 → Aplic0/1 = F32/F59). El rol
            // "no-referenciado" por su L (p. ej. F70/L70) es el "último" Mes (Mes2), coherente con
            // la regla del reader (Σ Mes − Especiales de todas menos la última).
            [1] =
            [
                ("F16", R1Q2Fuente.Mes0), ("L16", R1Q2Fuente.Lmes0),
                ("F47", R1Q2Fuente.Mes1), ("L47", R1Q2Fuente.Lmes1),
                ("F70", R1Q2Fuente.Mes2),
                ("F32", R1Q2Fuente.Aplic0), ("F59", R1Q2Fuente.Aplic1), ("L32", R1Q2Fuente.LAplic0)
            ],
            [2] =
            [
                ("F120", R1Q2Fuente.Mes0), ("L120", R1Q2Fuente.Lmes0),
                ("F171", R1Q2Fuente.Mes1), ("L171", R1Q2Fuente.Lmes1),
                ("F208", R1Q2Fuente.Mes2),
                ("F146", R1Q2Fuente.Aplic0), ("F190", R1Q2Fuente.Aplic1), ("L146", R1Q2Fuente.LAplic0)
            ],
            [3] =
            [
                ("F299", R1Q2Fuente.Mes0), ("L299", R1Q2Fuente.Lmes0),
                ("F330", R1Q2Fuente.Mes1), ("L330", R1Q2Fuente.Lmes1),
                ("F353", R1Q2Fuente.Mes2),
                ("F315", R1Q2Fuente.Aplic0), ("F342", R1Q2Fuente.Aplic1), ("L315", R1Q2Fuente.LAplic0)
            ],
            [4] =
            [
                ("F460", R1Q2Fuente.Mes0), ("L460", R1Q2Fuente.Lmes0),
                ("F511", R1Q2Fuente.Mes1), ("L511", R1Q2Fuente.Lmes1),
                ("F548", R1Q2Fuente.Mes2),
                ("F486", R1Q2Fuente.Aplic0), ("F530", R1Q2Fuente.Aplic1), ("L486", R1Q2Fuente.LAplic0)
            ],
            [5] =
            [
                ("F613", R1Q2Fuente.Mes0), ("L613", R1Q2Fuente.Lmes0),
                ("F648", R1Q2Fuente.Mes1),
                ("F629", R1Q2Fuente.Aplic0), ("F660", R1Q2Fuente.Aplic1), ("L629", R1Q2Fuente.LAplic0)
            ]
        };

    /// <summary>
    /// Fórmulas visibles R1-Q2 protegidas por ASE (TOT_OPT + EXTEMP) con sus fragmentos exactos
    /// (T0-0.2). El writer falla si alguna deja de ser fórmula. F46/F48/F519/F521 del mapa Q1
    /// NO están aquí: en Q2 son VALORES 0 (V0.2) y se excluyen de protegidas-fórmula (D5).
    ///
    /// Plan 36 (T2, R-B-3/D-B): <b>la derogación por desplazamiento queda sin objeto</b>. El espejo
    /// dimensional (que movía filas con Δ≠0 y dejaba obsoletas estas direcciones congeladas a julio)
    /// fue RETIRADO; el escritor values-only por slot (<c>EscritorValoresSlotR1</c>) pega solo
    /// <c>&lt;v&gt;</c> y la geometría/fórmulas del template quedan byte-idénticas, así que la
    /// validación protegida por dirección absoluta corre SIEMPRE. La vigencia de estas direcciones
    /// contra la plantilla estandarizada única se re-arbitra en el T0/regresión del Plan 36.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Celda, string[] Fragmentos)[]> R1Q2ProtectedPorAse =
        new Dictionary<int, (string, string[])[]>
        {
            [1] =
            [
                ("F75", ["F47", "F70", "F16", "L16", "L47"]),
                ("F77", ["F59", "F32", "L32"])
            ],
            [2] =
            [
                ("F254", ["F171", "F208", "L171", "F120", "L120"]),
                ("F256", ["F190", "F146", "L146"])
            ],
            [3] =
            [
                ("F415", ["F330", "F353", "F299", "L299", "L330"]),
                ("F417", ["F342", "F315", "L315"])
            ],
            [4] =
            [
                ("F568", ["F511", "F548", "F460", "L460", "L511"]),
                ("F570", ["F530", "F486", "L486"])
            ],
            [5] =
            [
                ("F680", ["F674", "F648", "F613", "L613", "L648"])
            ]
        };

    /// <summary>
    /// Celdas editables R2-Q2 por ASE (T0-0.2): Componente/Total, Subs/Cont/Total y Especiales
    /// del template. La semántica es la misma que Q1 (e15/e26/k15 de la fuente por labels);
    /// solo cambian las direcciones (E43=E17+E28-K17, E139=E89+E107-K89, …, E438=E410+E396-K396).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Componente, string SubsCont, string Especiales)> R2Q2EditablesPorAse =
        new Dictionary<int, (string, string, string)>
        {
            [1] = ("E21", "E32", "K21"),
            [2] = ("E94", "E112", "K94"),
            [3] = ("E182", "E193", "K182"),
            [4] = ("E309", "E327", "K309"),
            [5] = ("E404", "E418", "K404")
        };

    /// <summary>
    /// Fórmulas visibles R2-Q2 protegidas por ASE (T0-0.2).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Celda, string[] Fragmentos)[]> R2Q2ProtectedPorAse =
        new Dictionary<int, (string, string[])[]>
        {
            [1] = [("E47", ["E21", "E32", "K21"])],
            [2] = [("E144", ["E94", "E112", "K94"])],
            [3] = [("E262", ["E182", "E193", "K182"])],
            [4] = [("E362", ["E309", "E327", "K309"])],
            [5] = [("E446", ["E418", "E404", "K404"])]
        };

    /// <summary>
    /// Celdas editables R4-Q2 por ASE (T0-0.2): Total y P del bloque (D73=D15-P15, …,
    /// D355=D352-P352). P = 0 en el golden (sin análogo en la fuente, patrón Q1).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Total, string P)> R4Q2EditablesPorAse =
        new Dictionary<int, (string, string)>
        {
            [1] = ("D15", "P15"),
            [2] = ("D111", "P111"),
            [3] = ("D206", "P206"),
            [4] = ("D249", "P249"),
            [5] = ("D358", "P358")
        };

    /// <summary>
    /// Fórmulas visibles R4-Q2 protegidas por ASE (T0-0.2).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Celda, string[] Fragmentos)[]> R4Q2ProtectedPorAse =
        new Dictionary<int, (string, string[])[]>
        {
            [1] = [("D73", ["D15", "P15"])],
            [2] = [("D174", ["D111", "P111"])],
            [3] = [("D211", ["D206", "P206"])],
            [4] = [("D326", ["D249", "P249"])],
            [5] = [("D361", ["D358", "P358"])]
        };

    /// <summary>
    /// Filas CONSOLIDADO Q2 protegidas (V0.1): D9:D13 → R1!F53/…, D28:D32 → R2!E43/…,
    /// D47:D51 → R1!F55/…, D66:D70 → R4!D73/…, D85:D89 → AJUSTES-SF-T (igual forma que Q1),
    /// D104:D108 sumas + D109 SUM. D106 = shared follower del maestro D104 (solo presencia).
    /// </summary>
    public static readonly (string Hoja, string Celda, string[] Fragmentos)[] ConsolidadoProtected =
    [
        (HojaConsolidado, "D9", ["F75", "Reporte Componentes R1"]),
        (HojaConsolidado, "D10", ["F254", "Reporte Componentes R1"]),
        (HojaConsolidado, "D11", ["F415", "Reporte Componentes R1"]),
        (HojaConsolidado, "D12", ["F568", "Reporte Componentes R1"]),
        (HojaConsolidado, "D13", ["F680", "Reporte Componentes R1"]),
        (HojaConsolidado, "D28", ["E47", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D29", ["E144", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D30", ["E262", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D31", ["E362", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D32", ["E446", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D47", ["F77", "Reporte Componentes R1"]),
        (HojaConsolidado, "D48", ["F256", "Reporte Componentes R1"]),
        (HojaConsolidado, "D49", ["F417", "Reporte Componentes R1"]),
        (HojaConsolidado, "D50", ["F570", "Reporte Componentes R1"]),
        (HojaConsolidado, "D51", ["F682", "Reporte Componentes R1"]),
        (HojaConsolidado, "D66", ["D73", "Reversion Pagos R4"]),
        (HojaConsolidado, "D67", ["D174", "Reversion Pagos R4"]),
        (HojaConsolidado, "D68", ["D211", "Reversion Pagos R4"]),
        (HojaConsolidado, "D69", ["D326", "Reversion Pagos R4"]),
        (HojaConsolidado, "D70", ["D361", "Reversion Pagos R4"]),
        (HojaConsolidado, "D85", ["AJUSTES", "D47"]),
        (HojaConsolidado, "D86", ["AJUSTES", "D48"]),
        (HojaConsolidado, "D87", ["AJUSTES", "D49"]),
        (HojaConsolidado, "D88", ["AJUSTES", "D50"]),
        (HojaConsolidado, "D89", ["AJUSTES", "D51"]),
        (HojaConsolidado, "D104", ["D9", "D28", "D47", "D66", "D85"]),
        (HojaConsolidado, "D105", ["D10", "D29", "D48", "D67", "D86"]),
        (HojaConsolidado, "D106", []), // shared follower del maestro D104
        (HojaConsolidado, "D107", ["D12", "D31", "D50", "D69", "D88"]),
        (HojaConsolidado, "D108", ["D13", "D32", "D51", "D70", "D89"]),
        (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
    ];

    /// <summary>
    /// Celdas destino DetRetri-Q2 (V0.4): columna D de <c>DetRetri2026072</c> por ASE
    /// (D9..D13 = ASE1..5) + D14 = total (ROUND de D109). En el canónico son VALORES 0 editables;
    /// en el caché golden = ROUND(D104:D109). Composición congelada: <c>D(8+ase) = ROUND(D104:D108)</c>.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, string> DetRetriDestinoPorAse =
        new Dictionary<int, string>
        {
            [1] = "D9", [2] = "D10", [3] = "D11", [4] = "D12", [5] = "D13"
        };

    public const string DetRetriTotal = "D14";

    /// <summary>
    /// Fórmulas protegidas DetRetri/DetValiRetri Q2 (T0-0.4, Requirement 5): diferencias vs
    /// CONSOLIDADO (D23:D28 / D16:D21) y verificaciones booleanas de composición (D32:D36 /
    /// D24:D29). Nunca se escriben; el writer falla si alguna deja de ser fórmula.
    /// </summary>
    public static (string Hoja, string Celda, string[] Fragmentos)[] DetRetriProtected(Periodo periodo)
    {
        var hojaDetRetri = HojaDetRetri(periodo);
        var hojaDetValiRetri = HojaDetValiRetri(periodo);
        return
        [
            (hojaDetRetri, "D23", ["D9", "CONSOLIDADO_TOTAL RECAUDO", "D104"]),
            (hojaDetRetri, "D24", ["D10", "CONSOLIDADO_TOTAL RECAUDO", "D105"]),
            (hojaDetRetri, "D25", ["D11", "CONSOLIDADO_TOTAL RECAUDO", "D106"]),
            (hojaDetRetri, "D26", ["D12", "CONSOLIDADO_TOTAL RECAUDO", "D107"]),
            (hojaDetRetri, "D27", ["D13", "CONSOLIDADO_TOTAL RECAUDO", "D108"]),
            (hojaDetRetri, "D28", ["D14", "CONSOLIDADO_TOTAL RECAUDO", "D109"]),
            (hojaDetRetri, "D32", ["D23", "F75", "E47", "D73", "F77", "AJUSTES", "D47", "D9"]),
            (hojaDetRetri, "D33", ["D24", "F254", "E144", "D174", "F256", "AJUSTES", "D48", "D10"]),
            (hojaDetRetri, "D34", ["D25", "F415", "F417", "E262", "D211", "AJUSTES", "D49", "D11"]),
            (hojaDetRetri, "D35", ["D26", "F568", "F570", "E362", "D326", "AJUSTES", "D50", "D12"]),
            (hojaDetRetri, "D36", ["D27", "F680", "F682", "E446", "D361", "AJUSTES", "D51", "D13"]),
            (hojaDetValiRetri, "D16", ["D9", "CONSOLIDADO_TOTAL RECAUDO", "U104"]),
            (hojaDetValiRetri, "D17", ["D10", "CONSOLIDADO_TOTAL RECAUDO", "U105"]),
            (hojaDetValiRetri, "D18", ["D11", "CONSOLIDADO_TOTAL RECAUDO", "U106"]),
            (hojaDetValiRetri, "D19", ["D12", "CONSOLIDADO_TOTAL RECAUDO", "U107"]),
            (hojaDetValiRetri, "D20", ["D13", "CONSOLIDADO_TOTAL RECAUDO", "U108"]),
            (hojaDetValiRetri, "D21", ["D14", "CONSOLIDADO_TOTAL RECAUDO", "U109"]),
            (hojaDetValiRetri, "D24", ["M75", "M77", "L47", "J73", "D16", "AJUSTES", "U47", "D9"]),
            (hojaDetValiRetri, "D25", ["M254", "L144", "J174", "M256", "D17", "AJUSTES", "U48", "D10"]),
            (hojaDetValiRetri, "D26", ["M415", "M417", "L262", "J211", "D18", "AJUSTES", "U49", "D11"]),
            (hojaDetValiRetri, "D27", ["M568", "M570", "L362", "J326", "D19", "AJUSTES", "U50", "D12"]),
            (hojaDetValiRetri, "D28", ["M680", "M682", "L446", "J361", "D20", "AJUSTES", "U51", "D13"]),
            (hojaDetValiRetri, "D29", ["M75", "L47", "J73", "M254", "L144", "J174", "M415", "L262", "J211", "M568", "L362", "J326", "M680", "L446", "J361", "D21", "AJUSTES", "U52", "D14"])
        ];
    }

    /// <summary>
    /// Fórmulas protegidas adicionales Q2 (Requirement 5, T0-0.5): REMUNERACION_* por empresa (bloques
    /// genéricos; el bloque AJUSTES D85:D89 NO se valida en Q2 porque son VALORES — veredicto
    /// T0-0.5), VALIDACION_* O3, GERENTES_* D9, INTERVENTORIA K31/M31/N31 y ANT EXT-REV
    /// C3/M3/N3/C5. La cadena HU-08 (Recaudo * fila 29+) queda fuera en Q2 (recorte 1 HU-11:
    /// conciliación no se reescribe en Q2).
    /// </summary>
    public static readonly (string Hoja, string Celda, string[] Fragmentos)[] ProtegidasAdicionalesQ2 =
        CrearProtegidasAdicionalesQ2();

    /// <summary>
    /// HU-09 (2.3) parametrizado al período (D5): el mapa banco del Q1 referencia el TOTAL RECAUDO
    /// de la fila 81 (C81/D81) que NO existe en el template Q2 (la hoja REPORTE RECAUDO x BANCO
    /// termina en la fila 79 — veredicto T0-0.5). El resto de celdas del mapa es idéntico en Q2
    /// (verificado con shared-formula awareness: I17/J17/G55 son followers con <c>si</c>).
    /// </summary>
    public static readonly (string Hoja, string Celda, string[] Fragmentos)[] BancoProtegidasQ2 =
        WorkbookLeafCellMapReporteBanco.Protegidas
            .Where(p => !string.Equals(p.Celda, "C81", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(p.Celda, "D81", StringComparison.OrdinalIgnoreCase))
            .ToArray();

    private static (string Hoja, string Celda, string[] Fragmentos)[] CrearProtegidasAdicionalesQ2()
    {
        var lista = new List<(string, string, string[])>();

        // REMUNERACION_* por empresa: bloques genéricos que SÍ son fórmula en Q2 (T0-0.5).
        // Se omiten D85:D89 (valores en Q2) y D90 se incluye (SUM, fórmula verificada).
        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
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
            lista.Add((empresa.HojaRemuneracion, "D90", ["SUM"]));
            foreach (var fila in Enumerable.Range(104, 5))
            {
                lista.Add((empresa.HojaRemuneracion, $"D{fila}", ["D"]));
            }

            lista.Add((empresa.HojaRemuneracion, "D109", ["SUM"]));
        }

        // VALIDACION_* / GERENTES_* / INTERVENTORIA / ANT EXT-REV (2.7 protegidas, T0-0.5).
        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            lista.Add((empresa.HojaValidacion, "O3", ["H3", "N3"]));
            lista.Add((empresa.HojaGerentes, "D9", ["REMUNERACION"]));
        }

        lista.Add((HojaInterventoria, "K31", ["SUM", "K26", "K30"]));
        lista.Add((HojaInterventoria, "M31", ["SUM", "M26", "M30"]));
        lista.Add((HojaInterventoria, "N31", ["SUM", "N26", "N30"]));
        lista.Add((HojaAntExtRev, "C3", ["C11", "C19", "C27", "C35", "C43"]));
        lista.Add((HojaAntExtRev, "M3", ["C3", "E3", "G3", "I3", "K3"]));
        lista.Add((HojaAntExtRev, "N3", ["D3", "F3", "H3", "J3", "L3"]));
        lista.Add((HojaAntExtRev, "C5", ["SUM", "C3", "C4"]));

        return lista.ToArray();
    }

    /// <summary>
    /// Obtiene los operandos R1-Q2 editables de un ASE; lanza si no está soportado.
    /// </summary>
    public static (string Celda, R1Q2Fuente Fuente)[] ObtenerR1Q2Editables(int aseId) =>
        R1Q2EditablesPorAse.TryGetValue(aseId, out var mapa)
            ? mapa
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay cell-map R1-Q2 para el ASE {aseId}.");

    /// <summary>
    /// Obtiene las fórmulas visibles R1-Q2 protegidas de un ASE.
    /// </summary>
    public static (string Celda, string[] Fragmentos)[] ObtenerR1Q2Protegidos(int aseId) =>
        R1Q2ProtectedPorAse.TryGetValue(aseId, out var mapa)
            ? mapa
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay fórmulas protegidas R1-Q2 para el ASE {aseId}.");

    /// <summary>
    /// Obtiene las celdas editables R2-Q2 de un ASE.
    /// </summary>
    public static (string Componente, string SubsCont, string Especiales) ObtenerR2Q2Editables(int aseId) =>
        R2Q2EditablesPorAse.TryGetValue(aseId, out var mapa)
            ? mapa
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay cell-map R2-Q2 para el ASE {aseId}.");

    /// <summary>
    /// Obtiene las fórmulas visibles R2-Q2 protegidas de un ASE.
    /// </summary>
    public static (string Celda, string[] Fragmentos)[] ObtenerR2Q2Protegidos(int aseId) =>
        R2Q2ProtectedPorAse.TryGetValue(aseId, out var mapa)
            ? mapa
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay fórmulas protegidas R2-Q2 para el ASE {aseId}.");

    /// <summary>
    /// Obtiene las celdas editables R4-Q2 de un ASE.
    /// </summary>
    public static (string Total, string P) ObtenerR4Q2Editables(int aseId) =>
        R4Q2EditablesPorAse.TryGetValue(aseId, out var mapa)
            ? mapa
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay cell-map R4-Q2 para el ASE {aseId}.");

    /// <summary>
    /// Obtiene las fórmulas visibles R4-Q2 protegidas de un ASE.
    /// </summary>
    public static (string Celda, string[] Fragmentos)[] ObtenerR4Q2Protegidos(int aseId) =>
        R4Q2ProtectedPorAse.TryGetValue(aseId, out var mapa)
            ? mapa
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay fórmulas protegidas R4-Q2 para el ASE {aseId}.");

    /// <summary>
    /// Celda destino DetRetri-Q2 de un ASE (D9..D13); lanza si el ASE no está soportado.
    /// </summary>
    public static string ObtenerDetRetriDestino(int aseId) =>
        DetRetriDestinoPorAse.TryGetValue(aseId, out var celda)
            ? celda
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay destino DetRetri-Q2 para el ASE {aseId}.");
}
