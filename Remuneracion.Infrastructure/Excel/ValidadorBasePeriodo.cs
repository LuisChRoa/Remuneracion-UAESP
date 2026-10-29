using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 36 (T3, R-B-5 — §2.3-A / D-G): preflight de la BASE del período antes de escribir.
///
/// Doctrina (Plan 36 D-G): la base canónica de cada período sale del PROCEDIMIENTO offline
/// (<c>Herramientas/PreparadorBasePeriodo</c>, opción A): ahí se reescriben UNA vez los tokens de
/// período de las <c>&lt;f&gt;</c>, se actualiza <c>docProps/app.xml</c> y se sella
/// <c>N3</c> (código de período de la base). El RUNTIME nunca reescribe tokens de período
/// (Planes 28/30 intactos): solo VALIDA que la base que recibió corresponde al período que se va a
/// procesar, con fail-fast ANTES de escribir un solo valor (nunca una salida sellada con la base
/// de otro período).
///
/// Contrato verificado (todos por <see cref="NombresHojaPeriodo"/>, nunca enumerando hojas):
///   (1) existen las hojas <c>DetRetri{AAAAMMQ}</c>, <c>DetValiRetri{AAAAMMQ}</c> e
///       <c>Informe AFaseo Recaudo {AAAAMM}-{Q}</c> del período;
///   (2) <c>CONSOLIDADO_TOTAL RECAUDO!N3 == periodo.CodigoCompleto</c> (código de período de la base);
///   (3) los sellos base <c>CONSOLIDADO!G7/K7/D6</c> y <c>DetRetri/DetValiRetri!D6/D7</c> existen y
///       NO son fórmula (el guard anti-fórmula de los sellos de runtime es la misma regla);
///   (4) todas las referencias a <c>DetRetri/DetValiRetri</c> dentro de <c>&lt;f&gt;</c> apuntan al
///       período solicitado (0 tokens stale de otro período) y hay al menos una (la base las trae).
///
/// Cero literales de período (D-H): todos los valores esperados salen del dominio
/// (<see cref="Periodo.CodigoCompleto"/> / <see cref="NombresHojaPeriodo"/>).
/// </summary>
internal static class ValidadorBasePeriodo
{
    /// <summary>Hoja de metadatos/sellos base.</summary>
    internal const string HojaConsolidado = WorkbookLeafCellMap.HojaConsolidado;

    /// <summary>
    /// Referencias a hojas de período dentro del texto de una fórmula
    /// (<c>DetRetri{AAAAMMQ}</c> / <c>DetValiRetri{AAAAMMQ}</c>), citadas o no.
    /// </summary>
    private static readonly Regex ReferenciaHojaPeriodo = new(
        "(?:DetRetri|DetValiRetri)(?<periodo>\\d{7})",
        RegexOptions.Compiled);

    /// <summary>
    /// Valida la base del período y lanza <c>ERR-PLANTILLA</c> (nombrando hoja + celda + esperado)
    /// si no corresponde al período que se va a procesar.
    /// </summary>
    internal static void Exigir(WorkbookPart workbookPart, Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);
        ArgumentNullException.ThrowIfNull(periodo);

        foreach (var hoja in new[]
                 {
                     NombresHojaPeriodo.DetRetri(periodo),
                     NombresHojaPeriodo.DetValiRetri(periodo),
                     NombresHojaPeriodo.InformeAFaseo(periodo),
                     HojaConsolidado,
                 })
        {
            ExigirHoja(workbookPart, hoja, periodo);
        }

        var consolidado = ResolverHoja(workbookPart, HojaConsolidado);
        var n3 = ExigirCeldaSinFormula(consolidado, HojaConsolidado, "N3", periodo, "sello del código de período de la base");
        var codigoN3 = ValorTextoDeCelda(n3);
        if (!string.Equals(codigoN3, periodo.CodigoCompleto, StringComparison.Ordinal))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La base no corresponde al período {periodo.CodigoCompleto}: '{HojaConsolidado}!N3' = '{codigoN3}'. " +
                "Prepare la base del período con Herramientas/PreparadorBasePeriodo (Plan 36/T3) antes de procesar; nunca se escribe sobre la base de otro período.");
        }

        foreach (var celda in new[] { "G7", "K7", "D6" })
        {
            ExigirCeldaSinFormula(consolidado, HojaConsolidado, celda, periodo, "sello base");
        }

        // Plan 39 (T2, R-B-2/R-INTERV-Q): L9:L13 debe ser EXACTAMENTE la variante de la quincena
        // del run (Q1->N, Q2->M). Una tercera forma = ERR-PLANTILLA (hoja + celda + variante esperada).
        ExigirVarianteL(consolidado, periodo);

        // Plan 40 (T3, E1-4 + F-NEW-1): las 3 familias por-columna (recaudo de VALIDACION_TOTAL y
        // los auto-chequeos que hardcodean INTERVENTORIA!M26:M31) deben traer la variante de la
        // quincena del run. Una base sin la variante = ERR-PLANTILLA (fail-fast antes de escribir).
        ExigirVariantesDeQuincena(workbookPart, periodo);

        var hojasPeriodo = new[]
        {
            (NombresHojaPeriodo.DetRetri(periodo), new[] { "D6", "D7" }),
            (NombresHojaPeriodo.DetValiRetri(periodo), new[] { "D6", "D7" }),
        };

        foreach (var (hoja, celdas) in hojasPeriodo)
        {
            var worksheet = ResolverHoja(workbookPart, hoja);
            foreach (var celda in celdas)
            {
                ExigirCeldaSinFormula(worksheet, hoja, celda, periodo, "sello de proceso");
            }
        }

        ExigirTokensDePeriodo(workbookPart, periodo);
    }

    /// <summary>
    /// Plan 39 (T2, R-B-2/R-INTERV-Q): <c>CONSOLIDADO_TOTAL RECAUDO!L9:L13</c> debe ser EXACTAMENTE
    /// la variante de la quincena del run (Q1→N, Q2→M). Una tercera forma = <c>ERR-PLANTILLA</c>
    /// nombrando hoja + celda + variante esperada.
    /// </summary>
    private static void ExigirVarianteL(Worksheet consolidado, Periodo periodo)
    {
        foreach (var (celda, formulaEsperada) in PreparadorBasePipeline.FormulasLEsperadas(periodo))
        {
            var cell = consolidado.Descendants<Cell>()
                .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
                ?? throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La base del período {periodo.CodigoCompleto} no trae '{HojaConsolidado}!{celda}' " +
                    $"(variante L esperada '{formulaEsperada}', R-INTERV-Q). Fail-fast antes de escribir.");

            var formula = cell.CellFormula?.Text;
            if (string.IsNullOrWhiteSpace(formula) || !string.Equals(formula, formulaEsperada, StringComparison.Ordinal))
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La base del período {periodo.CodigoCompleto} no trae la variante L de su quincena: " +
                    $"'{HojaConsolidado}!{celda}' = '{formula ?? "(valor)"}' != '{formulaEsperada}' " +
                    $"(R-INTERV-Q, Q{periodo.NumeroQuincena}). Fail-fast antes de escribir.");
            }
        }
    }

    /// <summary>
    /// Plan 40 (T3, E1-4 + F-NEW-1): la base del período debe traer la variante de la quincena del
    /// run en las 3 familias por-columna: <c>VALIDACION_TOTAL!C3:C8</c> (columna del bloque
    /// <c>'Recaudo *'</c>: D en Q1 / F en Q2) y los auto-chequeos de <c>DetRetri/DetValiRetri</c>
    /// que hardcodean <c>INTERVENTORIA!M26:M31</c> (N en Q1 / M en Q2, misma tabla que R-INTERV-Q).
    /// Plan 62 (T4.2, D-H4): se suma <c>CONSOLIDADO_TOTAL RECAUDO!E142:E146</c> (extensión E1-4 al
    /// bloque 142–147, filas 21:25 de cada canal). Cualquier otra forma = <c>ERR-PLANTILLA</c>
    /// nombrando hoja + celda + quincena.
    /// </summary>
    private static void ExigirVariantesDeQuincena(WorkbookPart workbookPart, Periodo periodo)
    {
        var columnaRecaudo = PreparadorBasePipeline.ColumnaRecaudoDe(periodo);
        var hojasRecaudo = EmpresaFacturacion.Catalogo.Select(e => e.HojaRecaudo).ToHashSet(StringComparer.Ordinal);

        // E1-4: VALIDACION_TOTAL!C3:C8 + VALIDACION_*!C3:C8 deben sumar 'Recaudo *' en la columna
        // de la quincena del run.
        var hojasValidacion = new List<string> { WorkbookLeafCellMapValidaciones.HojaValidacionTotal };
        hojasValidacion.AddRange(EmpresaFacturacion.Catalogo.OrderBy(e => e.Id).Select(e => e.HojaValidacion));
        foreach (var nombreHoja in hojasValidacion)
        {
            ExigirRecaudoDeQuincena(
                ResolverHoja(workbookPart, nombreHoja),
                nombreHoja,
                PreparadorBasePipeline.CeldasRecaudoDeValidacionTotal,
                hojasRecaudo,
                columnaRecaudo,
                periodo,
                "E1-4");
        }

        // Plan 62 (T4.2): extensión E1-4 al bloque 142–147 del CONSOLIDADO (misma regla, misma tabla
        // de columnas por quincena de dominio).
        ExigirRecaudoDeQuincena(
            ResolverHoja(workbookPart, HojaConsolidado),
            HojaConsolidado,
            PreparadorBasePipeline.CeldasRecaudoDeConsolidado142,
            hojasRecaudo,
            columnaRecaudo,
            periodo,
            "E1-4/142-147");

        foreach (var (hoja, celda, fragmento) in PreparadorBasePipeline.FragmentosAutoChequeoEsperados(periodo))
        {
            var worksheet = ResolverHoja(workbookPart, hoja);
            var cell = worksheet.Descendants<Cell>()
                .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
                ?? throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La base del período {periodo.CodigoCompleto} no trae '{hoja}!{celda}' (F-NEW-1, variante '{fragmento}'). Fail-fast antes de escribir.");

            var formula = cell.CellFormula?.Text;
            if (string.IsNullOrWhiteSpace(formula) || !formula.Contains(fragmento, StringComparison.Ordinal))
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La base del período {periodo.CodigoCompleto} no trae la variante de su quincena: " +
                    $"'{hoja}!{celda}' = '{formula ?? "(valor)"}' no contiene '{fragmento}' (F-NEW-1, Q{periodo.NumeroQuincena}). Fail-fast antes de escribir.");
            }
        }

        // Plan 41 (T3): el auto-chequeo S26:S31 exige INTERVENTORIA!R26 == la tabla de la quincena
        // del run (Q2 → J15; Q1 → N26). Una base con la tabla de la otra quincena = ERR-PLANTILLA.
        var esperadoTablaJ = PreparadorBasePipeline.TablaJMaestroDe(periodo);
        var hojaInterventoria = ResolverHoja(workbookPart, WorkbookLeafCellMapAjustesSfT.HojaInterventoria);
        var celdaR26 = hojaInterventoria.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, PreparadorBasePipeline.CeldaTablaJMaestro, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La base del período {periodo.CodigoCompleto} no trae '{WorkbookLeafCellMapAjustesSfT.HojaInterventoria}!{PreparadorBasePipeline.CeldaTablaJMaestro}' (T3, tabla J '{esperadoTablaJ}'). Fail-fast antes de escribir.");
        var formulaR26 = celdaR26.CellFormula?.Text;
        if (string.IsNullOrWhiteSpace(formulaR26) || !string.Equals(formulaR26, esperadoTablaJ, StringComparison.Ordinal))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La base del período {periodo.CodigoCompleto} no trae la tabla J de su quincena: " +
                $"'{WorkbookLeafCellMapAjustesSfT.HojaInterventoria}!{PreparadorBasePipeline.CeldaTablaJMaestro}' = '{formulaR26 ?? "(valor)"}' != '{esperadoTablaJ}' (T3, Q{periodo.NumeroQuincena}). Fail-fast antes de escribir.");
        }
    }

    /// <summary>
    /// Plan 40/62 (E1-4 + extensión 142–147): cada celda de <paramref name="celdas"/> debe ser la
    /// suma de referencias al bloque <c>'Recaudo *'</c> (hojas del catálogo abierto) en su fila y en
    /// la columna de la quincena del run. Cualquier otra forma = <c>ERR-PLANTILLA</c> nombrando hoja +
    /// celda + quincena. Una sola implementación para las dos tablas (cero duplicación).
    /// </summary>
    private static void ExigirRecaudoDeQuincena(
        Worksheet worksheet,
        string nombreHoja,
        IReadOnlyList<(string Celda, int FilaRecaudo)> celdas,
        IReadOnlySet<string> hojasRecaudo,
        string columnaRecaudo,
        Periodo periodo,
        string familia)
    {
        foreach (var (celda, fila) in celdas)
        {
            var cell = worksheet.Descendants<Cell>()
                .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
                ?? throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La base del período {periodo.CodigoCompleto} no trae '{nombreHoja}!{celda}' ({familia}). Fail-fast antes de escribir.");

            var formula = cell.CellFormula?.Text;
            if (string.IsNullOrWhiteSpace(formula))
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La base del período {periodo.CodigoCompleto} no trae la variante de recaudo de su quincena: " +
                    $"'{nombreHoja}!{celda}' no es fórmula ({familia}, Q{periodo.NumeroQuincena}). Fail-fast antes de escribir.");
            }

            foreach (var termino in PreparadorBasePipeline.SepararTerminosSuma(formula))
            {
                var m = Regex.Match(termino, "^'(?<hoja>[^']+)'!(?<col>[A-Z]+)(?<fila>\\d+)$");
                var ok = m.Success
                    && hojasRecaudo.Contains(m.Groups["hoja"].Value)
                    && string.Equals(m.Groups["fila"].Value, fila.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
                    && string.Equals(m.Groups["col"].Value, columnaRecaudo, StringComparison.Ordinal);
                if (!ok)
                {
                    throw new CalculoInvalidoException(
                        CodigoError.Plantilla,
                        $"La base del período {periodo.CodigoCompleto} no trae la variante de recaudo de su quincena: " +
                        $"'{nombreHoja}!{celda}' = '{formula}' no es Σ 'Recaudo *'!{columnaRecaudo}{fila} " +
                        $"({familia}, Q{periodo.NumeroQuincena}). Fail-fast antes de escribir.");
                }
            }
        }
    }

    /// <summary>
    /// Plan 36 (T3, R-B-5): todas las referencias <c>DetRetri/DetValiRetri</c> de las <c>&lt;f&gt;</c>
    /// del workbook apuntan al período solicitado (0 stale) y hay al menos una.
    /// </summary>
    private static void ExigirTokensDePeriodo(WorkbookPart workbookPart, Periodo periodo)
    {
        var esperados = new HashSet<string>(StringComparer.Ordinal)
        {
            NombresHojaPeriodo.DetRetri(periodo),
            NombresHojaPeriodo.DetValiRetri(periodo),
        };

        var total = 0;
        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var worksheet = worksheetPart.Worksheet;
            if (worksheet is null)
            {
                continue;
            }

            var nombreHoja = NombreDeHoja(workbookPart, worksheetPart);
            foreach (var celda in worksheet.Descendants<Cell>())
            {
                var formula = celda.CellFormula;
                if (formula is null || string.IsNullOrEmpty(formula.Text))
                {
                    continue;
                }

                foreach (Match referencia in ReferenciaHojaPeriodo.Matches(formula.Text))
                {
                    var detectada = referencia.Value;
                    if (!esperados.Contains(detectada))
                    {
                        throw new CalculoInvalidoException(
                            CodigoError.Plantilla,
                            $"La base no corresponde al período {periodo.CodigoCompleto}: la fórmula '{nombreHoja}!{celda.CellReference?.Value}' " +
                            $"referencia la hoja '{detectada}' de otro período. Regenere la base con Herramientas/PreparadorBasePeriodo (Plan 36/T3).");
                    }

                    total++;
                }
            }
        }

        if (total == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La base del período {periodo.CodigoCompleto} no trae ninguna <f> que referencie sus hojas de período " +
                $"('{NombresHojaPeriodo.DetRetri(periodo)}' / '{NombresHojaPeriodo.DetValiRetri(periodo)}'); no es una base preparada (Plan 36/T3).");
        }
    }

    private static void ExigirHoja(WorkbookPart workbookPart, string nombreHoja, Periodo periodo)
    {
        var existe = workbookPart.Workbook?.Descendants<Sheet>()
            .Any(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase)) ?? false;
        if (!existe)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La base no corresponde al período {periodo.CodigoCompleto}: falta la hoja '{nombreHoja}'. " +
                "Prepare la base del período con Herramientas/PreparadorBasePeriodo (Plan 36/T3).");
        }
    }

    private static Cell ExigirCeldaSinFormula(Worksheet worksheet, string nombreHoja, string celda, Periodo periodo, string que)
    {
        var cell = worksheet.Descendants<Cell>()
                .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La base del período {periodo.CodigoCompleto} no trae '{nombreHoja}!{celda}' ({que}). Fail-fast antes de escribir.");

        if (cell.CellFormula is not null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La celda {que} '{nombreHoja}!{celda}' es fórmula en la base del período {periodo.CodigoCompleto}; " +
                "los sellos base son valores (Plan 28/31). Fail-fast antes de escribir.");
        }

        return cell;
    }

    private static string ValorTextoDeCelda(Cell celda)
    {
        if (celda.CellValue is null)
        {
            return string.Empty;
        }

        return celda.CellValue.InnerText?.Trim() ?? string.Empty;
    }

    private static Worksheet ResolverHoja(WorkbookPart workbookPart, string nombreHoja)
    {
        var sheet = workbookPart.Workbook?.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"El workbook no trae la hoja '{nombreHoja}'.");

        if (workbookPart.GetPartById(sheet.Id!) is not WorksheetPart worksheetPart || worksheetPart.Worksheet is null)
        {
            throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{nombreHoja}' no tiene WorksheetPart válido.");
        }

        return worksheetPart.Worksheet;
    }

    private static string NombreDeHoja(WorkbookPart workbookPart, WorksheetPart worksheetPart)
    {
        var sheet = workbookPart.Workbook?.Descendants<Sheet>()
            .FirstOrDefault(s => s.Id?.Value == workbookPart.GetIdOfPart(worksheetPart));
        return sheet?.Name?.Value ?? string.Empty;
    }
}
