using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// HU-13 (2.7): lector-oráculo de validaciones cruzadas (D1) — OpenXML READ-ONLY (nunca
/// escribe; A4 hash intacto). Lee el caché &lt;v&gt; de las hojas de validación protegidas y
/// verifica la presencia de fórmula &lt;f&gt; donde T0 lo exige (W2/D7): hoja o celda ausente →
/// fallo que NOMBRA hoja+celda; nunca 0 silencioso.
///
/// HU-14 (3.1 + deuda HU-13): S-1 booleano estricto ("false" literal → FALSE en toda rama),
/// S-2 O9/P9 leídos UNA vez (fuera del loop por ASE; el validador gatea una vez + igualdad ×5),
/// S-4 caché &lt;v&gt; exigido en celdas de gate (<see cref="LeerValorNumericoExigido"/>) y
/// W-1 sub-bloques booleanos de VALIDACION_TOTAL (C15/D25/O25/D34/F34) leídos por ASE.
///
/// El validador de dominio nunca abre .xlsx; este reader vive en Infrastructure y produce
/// <see cref="ValidacionCruzadaSnapshot"/> por ASE (matcheo estricto por Ase.Id).
/// </summary>
public sealed class ValidacionOracleReader : IValidacionOracleReader
{
    public IReadOnlyList<ValidacionCruzadaSnapshot> LeerSnapshots(string rutaWorkbook, Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(periodo);

        if (string.IsNullOrWhiteSpace(rutaWorkbook) || !File.Exists(rutaWorkbook))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.FuenteNoEncontrada, $"No se encontró el workbook para el oráculo de validaciones: '{rutaWorkbook}'.");
        }

        var snapshots = new List<ValidacionCruzadaSnapshot>(5);

        using (var workbook = SpreadsheetDocument.Open(rutaWorkbook, false))
        {
            var workbookPart = workbook.WorkbookPart
                ?? throw new CalculoInvalidoException(CodigoError.FormatoFuente, "El workbook del oráculo no tiene WorkbookPart válido.");

            var hojaDetValiRetri = WorkbookLeafCellMapValidaciones.HojaDetValiRetri(periodo.NumeroQuincena);

            // W2: la hoja DetValiRetri y la fila Total (D21/D29) se verifican UNA vez (compartidas).
            _ = ObtenerHoja(workbookPart, hojaDetValiRetri, "oráculo");
            _ = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, "D21", "oráculo");
            _ = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, "D29", "oráculo");

            // HU-14 (S-2, D8): VALIDACION_TOTAL O9/P9 se leen UNA vez (fuera del loop por ASE)
            // y se pueblan los 5 snapshots con el mismo valor; el validador gatea una vez
            // (snapshot[0]) y aserta igualdad ×5 (divergencia = bug del reader, fail-fast).
            var celdaTotalO = ObtenerCeldaFormula(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "O9", "oráculo");
            var celdaTotalP = ObtenerCeldaFormula(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "P9", "oráculo");
            var valorTotalO = LeerValorNumericoExigido(celdaTotalO, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "O9");
            var valorTotalP = LeerValorBooleano(celdaTotalP, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "P9");

            for (var aseId = 1; aseId <= 5; aseId++)
            {
                snapshots.Add(LeerSnapshotAse(workbookPart, aseId, periodo, hojaDetValiRetri, valorTotalO, valorTotalP));
            }
        }

        return snapshots;
    }

    private static ValidacionCruzadaSnapshot LeerSnapshotAse(
        WorkbookPart workbookPart,
        int aseId,
        Periodo periodo,
        string hojaDetValiRetri,
        decimal valorTotalO,
        bool valorTotalP)
    {
        var ase = CrearAse(aseId);
        var porEmpresa = new List<ValidacionEmpresaSnapshot>(5);
        var fila = WorkbookLeafCellMapValidaciones.FilaValidacionPorAse(aseId);

        foreach (var empresa in EmpresaFacturacion.Catalogo.OrderBy(e => e.Id))
        {
            // W2: celda O/P de la fila-ASE de la empresa debe existir y ser fórmula.
            var celdaO = ObtenerCeldaFormula(workbookPart, empresa.HojaValidacion, $"O{fila}", "oráculo");
            var celdaP = ObtenerCeldaFormula(workbookPart, empresa.HojaValidacion, $"P{fila}", "oráculo");

            porEmpresa.Add(new ValidacionEmpresaSnapshot
            {
                Empresa = empresa.Nombre,
                AseId = aseId,
                DiferenciaO = LeerValorNumericoExigido(celdaO, empresa.HojaValidacion, $"O{fila}"),
                VerificacionP = LeerValorBooleano(celdaP, empresa.HojaValidacion, $"P{fila}")
            });
        }

        // DetValiRetri: diferencias D(15+ase) y verificaciones D(23+ase); totales D21/D29 ya
        // verificados en presencia (W2) a nivel hoja.
        var diferencias = new List<DetValiRetriCeldaValor>(1);
        var filaDif = WorkbookLeafCellMapValidaciones.FilaDiferenciaDetValiRetri(aseId);
        var celdaDiferencia = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, $"D{filaDif}", "oráculo");
        diferencias.Add(new DetValiRetriCeldaValor
        {
            Celda = $"D{filaDif}",
            Valor = LeerValorNumericoExigido(celdaDiferencia, hojaDetValiRetri, $"D{filaDif}")
        });

        var verificaciones = new List<DetValiRetriCeldaVerificacion>(1);
        var filaVerif = WorkbookLeafCellMapValidaciones.FilaVerificacionDetValiRetri(aseId);
        var celdaVerificacion = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, $"D{filaVerif}", "oráculo");
        verificaciones.Add(new DetValiRetriCeldaVerificacion
        {
            Celda = $"D{filaVerif}",
            Verificacion = LeerValorBooleano(celdaVerificacion, hojaDetValiRetri, $"D{filaVerif}")
        });

        var celdaD21 = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, "D21", "oráculo");
        var celdaD29 = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, "D29", "oráculo");

        var detValiRetri = new DetValiRetriSnapshot
        {
            Ase = ase,
            DiferenciasAse = diferencias,
            VerificacionesAse = verificaciones,
            DiferenciaTotalD21 = LeerValorNumericoExigido(celdaD21, hojaDetValiRetri, "D21"),
            VerificacionTotalD29 = LeerValorBooleano(celdaD29, hojaDetValiRetri, "D29")
        };

        // HU-14 (W-1): sub-bloques booleanos de VALIDACION_TOTAL (C15/D25/O25/D34/F34).
        // Assert de presencia <f> (mapa protegido, W2) + lectura estricta; el gate TRUE
        // exacto por ASE vive en el validador (amparo T0-0.4 HU-13).
        var subBloques = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        foreach (var celdaSub in WorkbookLeafCellMapValidaciones.SubBloquesValidacionTotal)
        {
            var celdaSubBloque = ObtenerCeldaFormula(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, celdaSub, "oráculo");
            subBloques[celdaSub] = LeerValorBooleano(celdaSubBloque, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, celdaSub);
        }

        // Controles Valida -* (informativo; sin gate numérico — T0-0.5; S-4 conserva el 0).
        var controles = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
        {
            ["Valida -Remunera!D9"] = LeerValorNumericoSiExiste(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidaRemunera, "D9"),
            ["Valida - Anticipos!D6"] = LeerValorNumericoSiExiste(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidaAnticipos, "D6"),
            ["Valida - Control Recaudo!F10"] = LeerValorNumericoSiExiste(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidaControlRecaudo, "F10")
        };

        return new ValidacionCruzadaSnapshot
        {
            Ase = ase,
            PorEmpresa = porEmpresa,
            DetValiRetri = detValiRetri,
            ValidacionTotal = valorTotalO,
            ValidacionTotalOkP = valorTotalP,
            Controles = controles,
            SubBloquesValidacionTotal = subBloques
        };
    }

    // ── Helpers OpenXML (read-only; duplican navegación mínima del writer, sin escritura) ─────

    private static Cell ObtenerCeldaFormula(WorkbookPart workbookPart, string hoja, string celda, string operacion)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"La celda-oráculo '{hoja}!{celda}' no existe. (W2: sin asserts no hay snapshot; nunca 0 silencioso.)");

        if (cell.CellFormula is null)
        {
            throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"La celda-oráculo '{hoja}!{celda}' debería ser fórmula protegida y se detectó un valor fijo. (W2)");
        }

        return cell;
    }

    /// <summary>
    /// Lectura de controles informativos <c>Valida -*</c> (S-4): celda ausente o sin caché →
    /// 0 tolerado documentado (nunca gate). Solo un valor no numérico presente es fallo.
    /// </summary>
    private static decimal LeerValorNumericoSiExiste(WorkbookPart workbookPart, string hoja, string celda)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, "oráculo");
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        if (cell is null || cell.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
        {
            return 0m; // control sin valor o sin caché en esta plantilla (informativo)
        }

        if (decimal.TryParse(cell.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor))
        {
            return valor;
        }

        throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"La celda-oráculo '{hoja}!{celda}' tiene un valor no numérico: '{cell.CellValue.InnerText}'.");
    }

    /// <summary>
    /// HU-14 (S-4): lectura EXIGIDA para celdas de gate. <c>&lt;v&gt;</c> ausente (fórmula sin
    /// caché = workbook nunca recalculado) → fallo que NOMBRA hoja+celda; nunca 0 silencioso
    /// en gates (doctrina W2 HU-13, extendida por S-4).
    /// </summary>
    private static decimal LeerValorNumericoExigido(Cell cell, string hoja, string celda)
    {
        if (cell.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
        {
            throw new CalculoInvalidoException(
                CodigoError.FormatoFuente,
                $"La celda-oráculo '{hoja}!{celda}' es fórmula sin caché (<v> ausente); recalcule el workbook en Excel y reintente. (S-4: nunca 0 silencioso en gates.)");
        }

        if (decimal.TryParse(cell.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor))
        {
            return valor;
        }

        throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"La celda-oráculo '{hoja}!{celda}' tiene un valor no numérico: '{cell.CellValue.InnerText}'.");
    }

    /// <summary>
    /// HU-14 (S-1): lectura booleana ESTRICTA.
    /// "1"/"true" → TRUE y "0"/"false" → FALSE en TODAS las ramas (case-insensitive; el
    /// literal "false" YA NO se interpreta como TRUE — S-1).
    /// - Con t="b": cualquier otro texto es serialización booleana inválida → fail-fast nombrado.
    /// - Sin t="b": numérico ≠ 0 → TRUE; vacío → FALSE (caché ausente; conserva S-4 para
    ///   informativos); resto → fail-fast nombrado (hoja+celda+valor).
    /// </summary>
    private static bool LeerValorBooleano(Cell cell, string hoja, string celda)
    {
        if (cell.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
        {
            return false; // caché ausente → FALSE (el gate lo reporta nombrando la celda)
        }

        var texto = cell.CellValue.InnerText.Trim();

        if (string.Equals(texto, "1", StringComparison.OrdinalIgnoreCase)
            || string.Equals(texto, "true", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (string.Equals(texto, "0", StringComparison.OrdinalIgnoreCase)
            || string.Equals(texto, "false", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (cell.DataType is not null && cell.DataType.Value == CellValues.Boolean)
        {
            throw new CalculoInvalidoException(
                CodigoError.FormatoFuente,
                $"La celda-oráculo '{hoja}!{celda}' tiene un booleano t=\"b\" no reconocido: '{texto}' (se esperaba 1/0 o true/false). (S-1)");
        }

        if (decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor))
        {
            return valor != 0m;
        }

        throw new CalculoInvalidoException(
            CodigoError.FormatoFuente,
            $"La celda-oráculo '{hoja}!{celda}' tiene un valor booleano no reconocido: '{texto}' (se esperaba 1/0 o true/false). (S-1)");
    }

    private static Worksheet ObtenerHoja(WorkbookPart workbookPart, string nombreHoja, string operacion)
    {
        var workbook = workbookPart.Workbook ?? throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"El workbook para {operacion} no tiene metadata Workbook válida.");
        var sheet = workbook.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"La hoja-oráculo '{nombreHoja}' no existe en el workbook para {operacion}. (W2: sin asserts no hay snapshot.)");

        var worksheetPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart
            ?? throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"No se pudo resolver la hoja '{nombreHoja}' en el workbook para {operacion}.");

        return worksheetPart.Worksheet ?? throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"La hoja '{nombreHoja}' no tiene Worksheet válido.");
    }

    /// <summary>
    /// HU-14 (S-3): delega a <see cref="AseFactory.DesdeId"/> (fuente única desde
    /// <see cref="Constants.CarpetasAse.Prefijos"/>); el array hardcodeado desaparece.
    /// </summary>
    private static Ase CrearAse(int id) => AseFactory.DesdeId(id);
}