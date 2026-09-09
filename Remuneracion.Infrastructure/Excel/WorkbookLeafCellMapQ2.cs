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
/// - T0-0.3: mapeo fuente→template por rol de fila (Mes0/Mes1/Mes2, Subs0, Aplic0/Aplic1) verificado
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
    public const string HojaDetRetri = "DetRetri2026072";
    public const string HojaDetValiRetri = "DetValiRetri2026072";
    public const string HojaInterventoria = "INTERVENTORIA";
    public const string HojaAntExtRev = "ANT EXT-REV";

    /// <summary>
    /// Rol de la fila de la fuente R1-Q2 que alimenta una celda operando del template (T0-0.3).
    /// Mes0/Mes1/Mes2 = filas "Mes/Total" en orden de aparición (ASE5 solo trae 2 → Mes2 ausente).
    /// Subs0 = primera fila "Subsidio(-)/Contribucion(+)"; Aplic0/Aplic1 = filas
    /// "Aplicacion nuevos x reversion" en orden; L* = columna Especiales (SERVICIO ESPECIALES) de la fila.
    /// </summary>
    public enum R1Q2Fuente
    {
        Mes0, Mes1, Mes2, Lmes0, Lmes1, Lmes2, Subs0, Aplic0, Aplic1, LAplic0
    }

    /// <summary>
    /// Operandos editables R1-Q2 por ASE: celda del template → rol de la fuente (T0-0.2/0.3).
    /// Congelado del dump OpenXML: F53=F32+F48+F12-L12-L32, F55=F37+F17-L17, F206=…, F560=…
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Celda, R1Q2Fuente Fuente)[]> R1Q2EditablesPorAse =
        new Dictionary<int, (string, R1Q2Fuente)[]>
        {
            [1] =
            [
                ("F12", R1Q2Fuente.Mes0), ("L12", R1Q2Fuente.Lmes0),
                ("F32", R1Q2Fuente.Mes1), ("L32", R1Q2Fuente.Lmes1),
                ("F48", R1Q2Fuente.Mes2),
                ("F37", R1Q2Fuente.Subs0), ("F17", R1Q2Fuente.Aplic0), ("L17", R1Q2Fuente.LAplic0)
            ],
            [2] =
            [
                ("F94", R1Q2Fuente.Mes0), ("L94", R1Q2Fuente.Lmes0),
                ("F132", R1Q2Fuente.Mes1), ("L132", R1Q2Fuente.Lmes1),
                ("F160", R1Q2Fuente.Mes2),
                ("F142", R1Q2Fuente.Aplic1), ("F107", R1Q2Fuente.Aplic0), ("L107", R1Q2Fuente.LAplic0)
            ],
            [3] =
            [
                ("F244", R1Q2Fuente.Mes0), ("L244", R1Q2Fuente.Lmes0),
                ("F265", R1Q2Fuente.Mes1), ("L265", R1Q2Fuente.Lmes1),
                ("F281", R1Q2Fuente.Mes2),
                ("F270", R1Q2Fuente.Subs0), ("F250", R1Q2Fuente.Aplic0), ("L250", R1Q2Fuente.LAplic0)
            ],
            [4] =
            [
                ("F384", R1Q2Fuente.Mes0), ("L384", R1Q2Fuente.Lmes0),
                ("F422", R1Q2Fuente.Mes1), ("L422", R1Q2Fuente.Lmes1),
                ("F448", R1Q2Fuente.Mes2),
                ("F430", R1Q2Fuente.Aplic1), ("F397", R1Q2Fuente.Aplic0), ("L397", R1Q2Fuente.LAplic0)
            ],
            [5] =
            [
                ("F531", R1Q2Fuente.Mes0), ("L531", R1Q2Fuente.Lmes0),
                ("F552", R1Q2Fuente.Mes1),
                ("F538", R1Q2Fuente.Aplic1), ("F512", R1Q2Fuente.Aplic0), ("L512", R1Q2Fuente.LAplic0)
            ]
        };

    /// <summary>
    /// Fórmulas visibles R1-Q2 protegidas por ASE (TOT_OPT + EXTEMP) con sus fragmentos exactos
    /// (T0-0.2). El writer falla si alguna deja de ser fórmula. F46/F48/F519/F521 del mapa Q1
    /// NO están aquí: en Q2 son VALORES 0 (V0.2) y se excluyen de protegidas-fórmula (D5).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Celda, string[] Fragmentos)[]> R1Q2ProtectedPorAse =
        new Dictionary<int, (string, string[])[]>
        {
            [1] =
            [
                ("F53", ["F32", "F48", "F12", "L12", "L32"]),
                ("F55", ["F37", "F17", "L17"])
            ],
            [2] =
            [
                ("F206", ["F132", "F160", "L132", "F94", "L94"]),
                ("F208", ["F142", "F107", "L107"])
            ],
            [3] =
            [
                ("F343", ["F265", "F281", "F244", "L244", "L265"]),
                ("F345", ["F270", "F250", "L250"])
            ],
            [4] =
            [
                ("F468", ["F422", "F448", "F384", "L384", "L422"]),
                ("F470", ["F430", "F397", "L397"])
            ],
            [5] =
            [
                ("F558", ["F552", "F531", "L531"]),
                ("F560", ["F538", "F512", "L512"])
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
            [1] = ("E17", "E28", "K17"),
            [2] = ("E89", "E107", "K89"),
            [3] = ("E176", "E187", "K176"),
            [4] = ("E303", "E323", "K303"),
            [5] = ("E396", "E410", "K396")
        };

    /// <summary>
    /// Fórmulas visibles R2-Q2 protegidas por ASE (T0-0.2).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Celda, string[] Fragmentos)[]> R2Q2ProtectedPorAse =
        new Dictionary<int, (string, string[])[]>
        {
            [1] = [("E43", ["E17", "E28", "K17"])],
            [2] = [("E139", ["E89", "E107", "K89"])],
            [3] = [("E256", ["E176", "E187", "K176"])],
            [4] = [("E358", ["E303", "E323", "K303"])],
            [5] = [("E438", ["E410", "E396", "K396"])]
        };

    /// <summary>
    /// Celdas editables R4-Q2 por ASE (T0-0.2): Total y P del bloque (D73=D15-P15, …,
    /// D355=D352-P352). P = 0 en el golden (sin análogo en la fuente, patrón Q1).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Total, string P)> R4Q2EditablesPorAse =
        new Dictionary<int, (string, string)>
        {
            [1] = ("D15", "P15"),
            [2] = ("D105", "P105"),
            [3] = ("D200", "P200"),
            [4] = ("D243", "P243"),
            [5] = ("D352", "P352")
        };

    /// <summary>
    /// Fórmulas visibles R4-Q2 protegidas por ASE (T0-0.2).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (string Celda, string[] Fragmentos)[]> R4Q2ProtectedPorAse =
        new Dictionary<int, (string, string[])[]>
        {
            [1] = [("D73", ["D15", "P15"])],
            [2] = [("D168", ["D105", "P105"])],
            [3] = [("D205", ["D200", "P200"])],
            [4] = [("D320", ["D243", "P243"])],
            [5] = [("D355", ["D352", "P352"])]
        };

    /// <summary>
    /// Filas CONSOLIDADO Q2 protegidas (V0.1): D9:D13 → R1!F53/…, D28:D32 → R2!E43/…,
    /// D47:D51 → R1!F55/…, D66:D70 → R4!D73/…, D85:D89 → AJUSTES-SF-T (igual forma que Q1),
    /// D104:D108 sumas + D109 SUM. D106 = shared follower del maestro D104 (solo presencia).
    /// </summary>
    public static readonly (string Hoja, string Celda, string[] Fragmentos)[] ConsolidadoProtected =
    [
        (HojaConsolidado, "D9", ["F53", "Reporte Componentes R1"]),
        (HojaConsolidado, "D10", ["F206", "Reporte Componentes R1"]),
        (HojaConsolidado, "D11", ["F343", "Reporte Componentes R1"]),
        (HojaConsolidado, "D12", ["F468", "Reporte Componentes R1"]),
        (HojaConsolidado, "D13", ["F558", "Reporte Componentes R1"]),
        (HojaConsolidado, "D28", ["E43", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D29", ["E139", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D30", ["E256", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D31", ["E358", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D32", ["E438", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D47", ["F55", "Reporte Componentes R1"]),
        (HojaConsolidado, "D48", ["F208", "Reporte Componentes R1"]),
        (HojaConsolidado, "D49", ["F345", "Reporte Componentes R1"]),
        (HojaConsolidado, "D50", ["F470", "Reporte Componentes R1"]),
        (HojaConsolidado, "D51", ["F560", "Reporte Componentes R1"]),
        (HojaConsolidado, "D66", ["D73", "Reversion Pagos R4"]),
        (HojaConsolidado, "D67", ["D168", "Reversion Pagos R4"]),
        (HojaConsolidado, "D68", ["D205", "Reversion Pagos R4"]),
        (HojaConsolidado, "D69", ["D320", "Reversion Pagos R4"]),
        (HojaConsolidado, "D70", ["D355", "Reversion Pagos R4"]),
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
    public static readonly (string Hoja, string Celda, string[] Fragmentos)[] DetRetriProtected =
    [
        (HojaDetRetri, "D23", ["D9", "CONSOLIDADO_TOTAL RECAUDO", "D104"]),
        (HojaDetRetri, "D24", ["D10", "CONSOLIDADO_TOTAL RECAUDO", "D105"]),
        (HojaDetRetri, "D25", ["D11", "CONSOLIDADO_TOTAL RECAUDO", "D106"]),
        (HojaDetRetri, "D26", ["D12", "CONSOLIDADO_TOTAL RECAUDO", "D107"]),
        (HojaDetRetri, "D27", ["D13", "CONSOLIDADO_TOTAL RECAUDO", "D108"]),
        (HojaDetRetri, "D28", ["D14", "CONSOLIDADO_TOTAL RECAUDO", "D109"]),
        (HojaDetRetri, "D32", ["D23", "F53", "E43", "D73", "F55", "AJUSTES", "D47", "D9"]),
        (HojaDetRetri, "D33", ["D24", "F206", "E139", "D168", "F208", "AJUSTES", "D48", "D10"]),
        (HojaDetRetri, "D34", ["D25", "F343", "F345", "E256", "D205", "AJUSTES", "D49", "D11"]),
        (HojaDetRetri, "D35", ["D26", "F468", "F470", "E358", "D320", "AJUSTES", "D50", "D12"]),
        (HojaDetRetri, "D36", ["D27", "F558", "F560", "E438", "D355", "AJUSTES", "D51", "D13"]),
        (HojaDetValiRetri, "D16", ["D9", "CONSOLIDADO_TOTAL RECAUDO", "U104"]),
        (HojaDetValiRetri, "D17", ["D10", "CONSOLIDADO_TOTAL RECAUDO", "U105"]),
        (HojaDetValiRetri, "D18", ["D11", "CONSOLIDADO_TOTAL RECAUDO", "U106"]),
        (HojaDetValiRetri, "D19", ["D12", "CONSOLIDADO_TOTAL RECAUDO", "U107"]),
        (HojaDetValiRetri, "D20", ["D13", "CONSOLIDADO_TOTAL RECAUDO", "U108"]),
        (HojaDetValiRetri, "D21", ["D14", "CONSOLIDADO_TOTAL RECAUDO", "U109"]),
        (HojaDetValiRetri, "D24", ["M53", "M55", "L43", "J73", "D16", "AJUSTES", "U47", "D9"]),
        (HojaDetValiRetri, "D25", ["M206", "L139", "J168", "M208", "D17", "AJUSTES", "U48", "D10"]),
        (HojaDetValiRetri, "D26", ["M343", "M345", "L256", "J205", "D18", "AJUSTES", "U49", "D11"]),
        (HojaDetValiRetri, "D27", ["M468", "M470", "L358", "J320", "D19", "AJUSTES", "U50", "D12"]),
        (HojaDetValiRetri, "D28", ["M558", "M560", "L438", "J355", "D20", "AJUSTES", "U51", "D13"]),
        (HojaDetValiRetri, "D29", ["M53", "L43", "J73", "M206", "L139", "J168", "M343", "L256", "J205", "M468", "L358", "J320", "M558", "L438", "J355", "D21", "AJUSTES", "U52", "D14"])
    ];

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