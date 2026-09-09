using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// HU-13 (2.7): lector-oráculo de validaciones cruzadas (D1) — OpenXML READ-ONLY (nunca
/// escribe; A4 hash intacto). Lee el caché &lt;v&gt; de las hojas de validación protegidas y
/// verifica la presencia de fórmula &lt;f&gt; donde T0 lo exige (W2/D7): hoja o celda ausente →
/// fallo que NOMBRA hoja+celda; nunca 0 silencioso.
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
            throw new ArchivoFuenteNoEncontradoException($"No se encontró el workbook para el oráculo de validaciones: '{rutaWorkbook}'.");
        }

        var snapshots = new List<ValidacionCruzadaSnapshot>(5);

        using (var workbook = SpreadsheetDocument.Open(rutaWorkbook, false))
        {
            var workbookPart = workbook.WorkbookPart
                ?? throw new CalculoInvalidoException("El workbook del oráculo no tiene WorkbookPart válido.");

            var hojaDetValiRetri = WorkbookLeafCellMapValidaciones.HojaDetValiRetri(periodo.NumeroQuincena);

            // W2: la hoja DetValiRetri y la fila Total (D21/D29) se verifican UNA vez (compartidas).
            _ = ObtenerHoja(workbookPart, hojaDetValiRetri, "oráculo");
            _ = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, "D21", "oráculo");
            _ = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, "D29", "oráculo");

            for (var aseId = 1; aseId <= 5; aseId++)
            {
                snapshots.Add(LeerSnapshotAse(workbookPart, aseId, periodo, hojaDetValiRetri));
            }
        }

        return snapshots;
    }

    private static ValidacionCruzadaSnapshot LeerSnapshotAse(
        WorkbookPart workbookPart,
        int aseId,
        Periodo periodo,
        string hojaDetValiRetri)
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
                DiferenciaO = LeerValorNumerico(celdaO, empresa.HojaValidacion, $"O{fila}"),
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
            Valor = LeerValorNumerico(celdaDiferencia, hojaDetValiRetri, $"D{filaDif}")
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
            DiferenciaTotalD21 = LeerValorNumerico(celdaD21, hojaDetValiRetri, "D21"),
            VerificacionTotalD29 = LeerValorBooleano(celdaD29, hojaDetValiRetri, "D29")
        };

        // VALIDACION_TOTAL: O9/P9 (fila Σ) — misma semántica por ASE (plan §2.4).
        var celdaTotal = ObtenerCeldaFormula(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "O9", "oráculo");
        var celdaTotalP = ObtenerCeldaFormula(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "P9", "oráculo");

        // Controles Valida -* (informativo; sin gate numérico — T0-0.5).
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
            ValidacionTotal = LeerValorNumerico(celdaTotal, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "O9"),
            ValidacionTotalOkP = LeerValorBooleano(celdaTotalP, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "P9"),
            Controles = controles
        };
    }

    // ── Helpers OpenXML (read-only; duplican navegación mínima del writer, sin escritura) ─────

    private static Cell ObtenerCeldaFormula(WorkbookPart workbookPart, string hoja, string celda, string operacion)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException($"La celda-oráculo '{hoja}!{celda}' no existe. (W2: sin asserts no hay snapshot; nunca 0 silencioso.)");

        if (cell.CellFormula is null)
        {
            throw new CalculoInvalidoException($"La celda-oráculo '{hoja}!{celda}' debería ser fórmula protegida y se detectó un valor fijo. (W2)");
        }

        return cell;
    }

    private static decimal LeerValorNumericoSiExiste(WorkbookPart workbookPart, string hoja, string celda)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, "oráculo");
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        if (cell is null)
        {
            return 0m; // celda no presente: control sin valor en esta plantilla (informativo)
        }

        return LeerValorNumerico(cell, hoja, celda);
    }

    private static decimal LeerValorNumerico(Cell cell, string hoja, string celda)
    {
        if (cell.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
        {
            return 0m; // fórmula sin caché (workbook nunca recalculado): OpenXML no recalcula
        }

        if (decimal.TryParse(cell.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor))
        {
            return valor;
        }

        throw new CalculoInvalidoException($"La celda-oráculo '{hoja}!{celda}' tiene un valor no numérico: '{cell.CellValue.InnerText}'.");
    }

    private static bool LeerValorBooleano(Cell cell, string hoja, string celda)
    {
        if (cell.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
        {
            return false;
        }

        var texto = cell.CellValue.InnerText.Trim();
        if (cell.DataType is not null && cell.DataType.Value == CellValues.Boolean)
        {
            return texto != "0";
        }

        return decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor) && valor != 0m;
    }

    private static Worksheet ObtenerHoja(WorkbookPart workbookPart, string nombreHoja, string operacion)
    {
        var workbook = workbookPart.Workbook ?? throw new CalculoInvalidoException($"El workbook para {operacion} no tiene metadata Workbook válida.");
        var sheet = workbook.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException($"La hoja-oráculo '{nombreHoja}' no existe en el workbook para {operacion}. (W2: sin asserts no hay snapshot.)");

        var worksheetPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart
            ?? throw new CalculoInvalidoException($"No se pudo resolver la hoja '{nombreHoja}' en el workbook para {operacion}.");

        return worksheetPart.Worksheet ?? throw new CalculoInvalidoException($"La hoja '{nombreHoja}' no tiene Worksheet válido.");
    }

    private static Ase CrearAse(int id)
    {
        var nombres = new[] { "Promoambiental", "Lime", "Ciudad Limpia", "Bogotá Limpia", "Área Limpia" };
        return new Ase
        {
            Id = id,
            NombreCorto = nombres[id - 1].ToUpperInvariant(),
            NombreCompleto = nombres[id - 1],
            NumeroCarpeta = id
        };
    }
}
