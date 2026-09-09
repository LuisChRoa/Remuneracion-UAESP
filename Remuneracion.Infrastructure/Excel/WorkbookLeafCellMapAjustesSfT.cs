namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// HU-11 (2.5): mapa de celdas de la cadena AJUSTES-SF-T congelado por T0-0.7
/// (evidencia en <c>plans/11 - HU-11-desbloqueo-Q2-ajustes-sf-t.md</c> §4 Fase 0 + dumps OpenXML
/// de la plantilla canónica <c>Docs/Insumos/REMUNERACION 2026072/Plantilla 8 agos 2026 _ Remuneracion
/// 202607-2 Total.xlsx</c> y del golden <c>Docs/Insumos/Remuneracion 202607-2 Total.xlsx</c>).
///
/// Veredictos T0 (probados, no asumidos):
/// - T0-0.1 golden canónico Q2 = "Plantilla 8 agos 2026" (coincide con la referencia en los
///   valores cacheados de validación; la otra plantilla = control, NUNCA oráculo).
/// - T0-0.2 valor-vs-fórmula de la cadena: las hojas SALDOS POR NOTA / RETRIBUCION NEGATIVA
///   traen BLOQUES DE VALORES editables por ASE (patrón R2: filas de conceptos + fila Total
///   Cn-In en fórmula) → desenlace D2(a): el writer escribe los operandos editables T0.
/// - T0-0.3 composición D85:D89 = <c>'AJUSTES - SF-T'!D47..D51</c>, y en AJUSTES-SF-T
///   <c>D47..D51 = D9..D13 + D28..D32</c> (SALDOS POR NOTA + RETRIBUCION NEGATIVA), probada
///   contra el golden ±0.5 en los 5 ASE: [973693.46, 216025.77, 104231.83, 35954.44, 0].
/// - T0-0.5 headers fuente: ninguna fuente SALDOS-NOTAS Q2 trae columna "Especiales"
///   (el header de esa posición es "Componente TCS") → <c>TieneColumnaEspeciales=false</c> en
///   los 5; ASE2 agrega "Deb/Cred" (col N) y ASE4/ASE1/ASE3 no; RETRIBUCION-NEGATIVA Q2 trae
///   SOLO la fila 1 (fuente vacía = 0 legítimo).
/// - T0-0.7 filas editables por ASE (plantilla canónica):
///   SALDOS POR NOTA: ASE1 filas 3..7, ASE2 15..20, ASE3 28..32, ASE4 40..45, ASE5 53..58.
///   RETRIBUCION NEGATIVA: ASE1 filas 3..8, ASE2 16..21, ASE3 29..34, ASE4 42..47, ASE5 55..59.
///   Columnas C..O (C=Total … O=Deb/Cred); la col I del template es "Especiales" → 0 en fuente.
///   INTERVENTORIA L25:N31 y ANT EXT-REV entran al mapa protegido (no se escriben).
/// </summary>
public static class WorkbookLeafCellMapAjustesSfT
{
    public const string HojaSaldosNotas = "SALDOS POR NOTA";
    public const string HojaRetribucionNegativa = "RETRIBUCION NEGATIVA";
    public const string HojaAjustesSfT = "AJUSTES - SF-T";
    public const string HojaInterventoria = "INTERVENTORIA";

    /// <summary>
    /// Filas editables del bloque SALDOS POR NOTA por ASE (plantilla canónica, T0-0.7).
    /// Cada fila del rango es una fila de conceptos/Total en VALORES (la fila Cn-In siguiente
    /// es fórmula protegida y NO se incluye aquí).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (int Desde, int Hasta)> FilasSaldosNotasPorAse =
        new Dictionary<int, (int, int)>
        {
            [1] = (3, 7),
            [2] = (15, 20),
            [3] = (28, 32),
            [4] = (40, 45),
            [5] = (53, 58)
        };

    /// <summary>
    /// Filas editables del bloque RETRIBUCION NEGATIVA por ASE (plantilla canónica, T0-0.7).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (int Desde, int Hasta)> FilasRetribucionNegativaPorAse =
        new Dictionary<int, (int, int)>
        {
            [1] = (3, 8),
            [2] = (16, 21),
            [3] = (29, 34),
            [4] = (42, 47),
            [5] = (55, 59)
        };

    /// <summary>
    /// Mapeo fuente→template POR TÍTULO de columna (D3/G4, T0-0.5): cada header de la fuente
    /// SALDOS-NOTAS va a la columna del template con el MISMO título. La columna I del template
    /// ("Especiales") NO existe en las fuentes Q2 → valor 0 demostrado (no está en este mapa).
    /// "Componente TCS" de la fuente (col I) va al template J; "CCSA Prest. No Aprov." (col M)
    /// va al template N.
    /// </summary>
    public static readonly (string HeaderFuente, string ColumnaTemplate)[] MapeoColumnasPorTitulo =
    [
        ("Total", "C"),
        ("Componente TDF", "D"),
        ("Componente TTL", "E"),
        ("Componente TVIAT", "F"),
        ("Aprovechamiento", "G"),
        ("CCSA Prest.Aprov.", "H"),
        ("Componente TCS", "J"),
        ("Componente TLU", "K"),
        ("Componente TBL", "L"),
        ("Componente TRT", "M"),
        ("CCSA Prest. No Aprov.", "N")
    ];

    /// <summary>
    /// Columna OPCIONAL de la fuente ("Deb/Cred", solo ASE2 — T0-0.5). Si la fuente no la trae,
    /// el template O queda en 0 (valor ya presente en la plantilla; el writer no escribe nada).
    /// </summary>
    public const string HeaderDebCredOpcional = "Deb/Cred";
    public const string ColumnaTemplateDebCred = "O";

    /// <summary>
    /// Fila (1-based) de la etiqueta "Total" dentro de cada bloque SALDOS-NOTAS del template.
    /// Es la fila cuyo valor col C alimenta el visible Cn-In (T0-0.3): la fila "Total" del rango.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, int> FilaTotalSaldosPorAse =
        new Dictionary<int, int>
        {
            [1] = 7, [2] = 20, [3] = 32, [4] = 45, [5] = 58
        };

    /// <summary>
    /// Fila (1-based) de la etiqueta "Total" dentro de cada bloque RETRIBUCION-NEGATIVA del template.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, int> FilaTotalRetribucionPorAse =
        new Dictionary<int, int>
        {
            [1] = 8, [2] = 21, [3] = 34, [4] = 47, [5] = 59
        };

    /// <summary>
    /// Concepto fuente por fila del template en SALDOS POR NOTA (mapeo explícito por Ase.Id;
    /// prohibidos offsets, D7). La fila es la del template (rango T0-0.7); el concepto es la
    /// etiqueta que se busca en la fuente (col A/B) para poblar esa fila.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (int Fila, string Concepto)[]> ConceptosSaldosPorAse =
        new Dictionary<int, (int, string)[]>
        {
            [1] =
            [
                (3, "Vlr Servicio"), (4, "Componente"), (5, "Subsidio(-)/Contribucion(+)"),
                (6, "Subs/Cont"), (7, "Total")
            ],
            [2] =
            [
                (15, "Vlr Servicio"), (16, "Vlr Intereses"), (17, "Componente"),
                (18, "Subsidio(-)/Contribucion(+)"), (19, "Subs/Cont"), (20, "Total")
            ],
            [3] =
            [
                (28, "Vlr Servicio"), (29, "Componente"), (30, "Subsidio(-)/Contribucion(+)"),
                (31, "Subs/Cont"), (32, "Total")
            ],
            [4] =
            [
                (40, "Vlr Servicio"), (41, "Vlr Intereses"), (42, "Componente"),
                (43, "Subsidio(-)/Contribucion(+)"), (44, "Subs/Cont"), (45, "Total")
            ],
            [5] =
            [
                (53, "Vlr Servicio"), (54, "Vlr Intereses"), (55, "Componente"),
                (56, "Subsidio(-)/Contribucion(+)"), (57, "Subs/Cont"), (58, "Total")
            ]
        };

    /// <summary>
    /// Concepto fuente por fila del template en RETRIBUCION NEGATIVA (mapeo explícito por Ase.Id).
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (int Fila, string Concepto)[]> ConceptosRetribucionPorAse =
        new Dictionary<int, (int, string)[]>
        {
            [1] =
            [
                (3, "Vlr Intereses"), (4, "Vlr Servicio"), (5, "Componente"),
                (6, "Subsidio(-)/Contribucion(+)"), (7, "Subs/Cont"), (8, "Total")
            ],
            [2] =
            [
                (16, "Vlr Intereses"), (17, "Vlr Servicio"), (18, "Componente"),
                (19, "Subsidio(-)/Contribucion(+)"), (20, "Subs/Cont"), (21, "Total")
            ],
            [3] =
            [
                (29, "Vlr Intereses"), (30, "Vlr Servicio"), (31, "Componente"),
                (32, "Subsidio(-)/Contribucion(+)"), (33, "Subs/Cont"), (34, "Total")
            ],
            [4] =
            [
                (42, "Vlr Intereses"), (43, "Vlr Servicio"), (44, "Componente"),
                (45, "Subsidio(-)/Contribucion(+)"), (46, "Subs/Cont"), (47, "Total")
            ],
            [5] =
            [
                (55, "Vlr Servicio"), (56, "Componente"), (57, "Subsidio(-)/Contribucion(+)"),
                (58, "Subs/Cont"), (59, "Total")
            ]
        };

    /// <summary>
    /// Fórmulas protegidas de la cadena 2.5 (nunca se escriben; D6): visibles SALDOS-NOTAS
    /// (Cn-In), visibles RETRIBUCION-NEGATIVA (Cn-In), AJUSTES-SF-T (D9:D13/D28:D32/D47:D51 +
    /// totales D14/D33/D52), CONSOLIDADO D85:D89 (ya protegidas por el mapa HU-07, aquí se
    /// re-aseguran con la referencia explícita a AJUSTES), INTERVENTORIA L31:M31:N31 (SUM) y
    /// ANT EXT-REV (filas totales). Fragmentos con la hoja normalizada (sin comillas/espacios).
    /// </summary>
    public static readonly (string Hoja, string Celda, string[] Fragmentos)[] Protegidas =
    [
        // Visibles SALDOS-NOTAS: Cn-In (C9/C22/C34/C47/C60) con su fila Total del bloque.
        (HojaSaldosNotas, "C9", ["C7", "I7"]),
        (HojaSaldosNotas, "C22", ["C20", "I20"]),
        (HojaSaldosNotas, "C34", ["C32", "I32"]),
        (HojaSaldosNotas, "C47", ["C45", "I45"]),
        (HojaSaldosNotas, "C60", ["C58", "I58"]),
        // Visibles RETRIBUCION-NEGATIVA: Cn-In (C10/C23/C36/C49/C61).
        (HojaRetribucionNegativa, "C10", ["C8", "I8"]),
        (HojaRetribucionNegativa, "C23", ["C21", "I21"]),
        (HojaRetribucionNegativa, "C36", ["C34", "I34"]),
        (HojaRetribucionNegativa, "C49", ["C47", "I47"]),
        (HojaRetribucionNegativa, "C61", ["C59", "I59"]),
        // AJUSTES-SF-T: sección SALDOS (D9:D13) y RETRIBUCION (D28:D32) + consolidado D47:D51.
        (HojaAjustesSfT, "D9", ["SALDOSPORNOTA", "C9"]),
        (HojaAjustesSfT, "D10", ["SALDOSPORNOTA", "C22"]),
        (HojaAjustesSfT, "D11", ["SALDOSPORNOTA", "C34"]),
        (HojaAjustesSfT, "D12", ["SALDOSPORNOTA", "C47"]),
        (HojaAjustesSfT, "D13", ["SALDOSPORNOTA", "C60"]),
        (HojaAjustesSfT, "D28", ["RETRIBUCIONNEGATIVA", "C10"]),
        (HojaAjustesSfT, "D29", ["RETRIBUCIONNEGATIVA", "C23"]),
        (HojaAjustesSfT, "D30", ["RETRIBUCIONNEGATIVA", "C36"]),
        (HojaAjustesSfT, "D31", ["RETRIBUCIONNEGATIVA", "C49"]),
        (HojaAjustesSfT, "D32", ["RETRIBUCIONNEGATIVA", "C61"]),
        (HojaAjustesSfT, "D47", ["D9", "D28"]),
        (HojaAjustesSfT, "D48", ["D10", "D29"]),
        (HojaAjustesSfT, "D49", []), // shared follower de D48
        (HojaAjustesSfT, "D50", []), // shared follower de D48
        (HojaAjustesSfT, "D51", []), // shared follower de D48
        // Totales de sección AJUSTES-SF-T.
        (HojaAjustesSfT, "D14", ["SUM", "D9", "D13"]),
        (HojaAjustesSfT, "D33", ["SUM", "D28", "D32"]),
        (HojaAjustesSfT, "D52", ["SUM", "D47", "D51"]),
        // CONSOLIDADO D85:D89 → AJUSTES (la fórmula es 'AJUSTES - SF-T'!D47..D51).
        ("CONSOLIDADO_TOTAL RECAUDO", "D85", ["AJUSTES", "D47"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "D86", ["AJUSTES", "D48"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "D87", ["AJUSTES", "D49"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "D88", ["AJUSTES", "D50"]),
        ("CONSOLIDADO_TOTAL RECAUDO", "D89", ["AJUSTES", "D51"]),
        // INTERVENTORIA: totales por quincena en fórmula (K31/M31/N31; L31 no es fórmula en el
        // template — los valores de referencia L26:L30/N26:N30 jamás se escriben).
        (HojaInterventoria, "K31", ["SUM", "K26", "K30"]),
        (HojaInterventoria, "M31", ["SUM", "M26", "M30"]),
        (HojaInterventoria, "N31", ["SUM", "N26", "N30"]),
        // ANT EXT-REV: totales por bloque en fórmula (nunca se escriben; Requirement 7).
        // NOTA HU-12 (bug blocker Q2): C3 = C11+C19+C27+C35+C43 (suma aritmética, NO SUM()).
        // El fragmento "SUM" declarado en HU-11 nunca se ejercitó (writer Q2 bloqueado por el
        // recorte T0-0.6); la fórmula real del canónico no lo contiene → se corrige el fragmento.
        ("ANT EXT-REV", "C3", ["C11", "C19", "C27", "C35", "C43"]),
        ("ANT EXT-REV", "M3", ["C3", "E3", "G3", "I3", "K3"]),
        ("ANT EXT-REV", "N3", ["D3", "F3", "H3", "J3", "L3"]),
        ("ANT EXT-REV", "C5", ["SUM", "C3", "C4"])
    ];

    /// <summary>
    /// Obtiene el rango de filas editables de SALDOS POR NOTA para un ASE; lanza si no está soportado.
    /// </summary>
    public static (int Desde, int Hasta) ObtenerFilasSaldos(int aseId) =>
        FilasSaldosNotasPorAse.TryGetValue(aseId, out var filas)
            ? filas
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay cell-map de SALDOS POR NOTA para el ASE {aseId}.");

    /// <summary>
    /// Obtiene el rango de filas editables de RETRIBUCION NEGATIVA para un ASE; lanza si no está soportado.
    /// </summary>
    public static (int Desde, int Hasta) ObtenerFilasRetribucion(int aseId) =>
        FilasRetribucionNegativaPorAse.TryGetValue(aseId, out var filas)
            ? filas
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay cell-map de RETRIBUCION NEGATIVA para el ASE {aseId}.");

    /// <summary>
    /// Obtiene los conceptos por fila del template de SALDOS POR NOTA para un ASE.
    /// </summary>
    public static (int Fila, string Concepto)[] ObtenerConceptosSaldos(int aseId) =>
        ConceptosSaldosPorAse.TryGetValue(aseId, out var conceptos)
            ? conceptos
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay conceptos de SALDOS POR NOTA para el ASE {aseId}.");

    /// <summary>
    /// Obtiene los conceptos por fila del template de RETRIBUCION NEGATIVA para un ASE.
    /// </summary>
    public static (int Fila, string Concepto)[] ObtenerConceptosRetribucion(int aseId) =>
        ConceptosRetribucionPorAse.TryGetValue(aseId, out var conceptos)
            ? conceptos
            : throw new ArgumentOutOfRangeException(nameof(aseId), $"No hay conceptos de RETRIBUCION NEGATIVA para el ASE {aseId}.");
}