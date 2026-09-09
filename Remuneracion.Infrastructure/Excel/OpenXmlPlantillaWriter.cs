using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Writer OpenXML de la plantilla real.
/// HU-04: validación estructural no destructiva (<see cref="IPlantillaWriter"/>).
/// HU-05: escritura real de celdas leaf sobre una copia (<see cref="IWorkbookLeafWriter"/>).
/// HU-07: overload multi-ASE con mapa por bloque (<see cref="WorkbookLeafCellMapPorAse"/>),
/// una sola copia, validación pre/post ampliada y borrado de parcial ante fallo.
/// HU-11 (2.5): en Q2 escribe los operandos editables de SALDOS POR NOTA / RETRIBUCION NEGATIVA
/// en la MISMA pasada atómica (D2a: T0-0.7 demostró bloques de valores editables) y amplía la
/// validación protegida a la cadena AJUSTES-SF-T (mapa <see cref="WorkbookLeafCellMapAjustesSfT"/>
/// + HU-10 parametrizado al sufijo de hoja 2026072).
/// HU-12 (2.6 ampliada): en Q2 escribe los leafs R1/R2/R4-Q2 (mapa <see cref="WorkbookLeafCellMapQ2"/>,
/// con variante ASE5 de 2 filas V0.3) + DetRetri-Q2 (V0.4: ROUND(D104:D108,0) vía
/// <see cref="DetRetriRounder"/>) en la MISMA pasada; la validación protegida se parametriza por
/// período (D5): Q1 exige las fórmulas HU-07, Q2 exige el mapa T0 (F53…, D73…, DetValiRetri/
/// VALIDACION_*/INTERVENTORIA/ANT EXT-REV protegidas) y M1 queda ejercitado contra el canónico.
/// </summary>
public class OpenXmlPlantillaWriter : IPlantillaWriter, IWorkbookLeafWriter
{
    private const string HojaConsolidado = WorkbookLeafCellMap.HojaConsolidado;
    private const string HojaR1 = WorkbookLeafCellMap.HojaR1;
    private const string HojaR2 = WorkbookLeafCellMap.HojaR2;
    private const string HojaR4 = WorkbookLeafCellMap.HojaR4;
    private const string HojaBanco = WorkbookLeafCellMapReporteBanco.HojaBanco;
    private const string HojaBce = WorkbookLeafCellMapBalanceSc.HojaBce;
    private const string HojaSaldosNotas = WorkbookLeafCellMapAjustesSfT.HojaSaldosNotas;
    private const string HojaRetribucionNegativa = WorkbookLeafCellMapAjustesSfT.HojaRetribucionNegativa;
    private const string SufijoHojasQ1 = "2026071";
    private const string SufijoHojasQ2 = "2026072";

    public void EscribirConsolidado(string rutaPlantilla, ResultadoRemuneracion resultado)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(resultado);

        var workbookPath = ValidarArchivo(rutaPlantilla, nameof(EscribirConsolidado));
        using var workbook = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = workbook.WorkbookPart ?? throw new CalculoInvalidoException("El workbook abierto no tiene WorkbookPart válido.");
        var worksheet = ObtenerHoja(workbook, HojaConsolidado, nameof(EscribirConsolidado));

        ValidarConsolidadoFormulario(workbookPart, worksheet);
        ValidarCadenaFormulaR1R2R4(workbookPart);
    }

    public void EscribirDetalleR1(string rutaPlantilla, Ase ase, RecaudoComponenteR1 datos)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(datos);

        var workbookPath = ValidarArchivo(rutaPlantilla, nameof(EscribirDetalleR1));
        using var workbook = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = workbook.WorkbookPart ?? throw new CalculoInvalidoException("El workbook abierto no tiene WorkbookPart válido.");
        var worksheet = ObtenerHoja(workbook, HojaR1, nameof(EscribirDetalleR1));

        ValidarHojaConLabelsEsperados(workbookPart, HojaR1, ["Componente", "Total", "Mes"]);
        ValidarCeldaTieneFormula(workbookPart, worksheet, "F46", ["F25", "F41", "L25"], HojaR1, nameof(EscribirDetalleR1));
        ValidarCeldaTieneFormula(workbookPart, worksheet, "F48", ["F30", "F10", "L10"], HojaR1, nameof(EscribirDetalleR1));

        LanzarPayloadInsuficiente(nameof(EscribirDetalleR1), HojaR1,
            "Se requiere un contrato de WorkbookLeafInputs que exponga los inputs reales detrás de F25,F41,L25,F30,F10,L10; el payload actual solo expone agregados TotOpt/Extemp y detalle parcial.");
    }

    public void EscribirDetalleR2(string rutaPlantilla, Ase ase, SaldosFavorR2 datos)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(datos);

        var workbookPath = ValidarArchivo(rutaPlantilla, nameof(EscribirDetalleR2));
        using var workbook = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = workbook.WorkbookPart ?? throw new CalculoInvalidoException("El workbook abierto no tiene WorkbookPart válido.");
        var worksheet = ObtenerHoja(workbook, HojaR2, nameof(EscribirDetalleR2));

        ValidarHojaConLabelsEsperados(workbookPart, HojaR2, ["Total", "Componente TDF"]);
        ValidarCeldaTieneFormula(workbookPart, worksheet, "E41", ["E15", "E26", "K15"], HojaR2, nameof(EscribirDetalleR2));

        LanzarPayloadInsuficiente(nameof(EscribirDetalleR2), HojaR2,
            "Se requiere un contrato de WorkbookLeafInputs que exponga los inputs reales detrás de E15,E26,K15; el modelo actual solo expone GrandTotal y ServEspK.");
    }

    public void EscribirDetalleR4(string rutaPlantilla, Ase ase, ReversionR4 datos)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(datos);

        var workbookPath = ValidarArchivo(rutaPlantilla, nameof(EscribirDetalleR4));
        using var workbook = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = workbook.WorkbookPart ?? throw new CalculoInvalidoException("El workbook abierto no tiene WorkbookPart válido.");
        var worksheet = ObtenerHoja(workbook, HojaR4, nameof(EscribirDetalleR4));

        ValidarCeldaTieneFormula(workbookPart, worksheet, "D67", ["D9", "P9"], HojaR4, nameof(EscribirDetalleR4));

        LanzarPayloadInsuficiente(nameof(EscribirDetalleR4), HojaR4,
            "Se requiere un contrato de WorkbookLeafInputs que exponga los inputs reales detrás de D9 y P9; el modelo actual solo expone TotalReversiones agregado.");
    }

    /// <inheritdoc />
    public void GenerarWorkbook(string rutaPlantillaOrigen, string rutaSalida, ResultadoRemuneracion resultado, WorkbookLeafInputs leafInputs)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantillaOrigen);
        ArgumentNullException.ThrowIfNull(rutaSalida);
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leafInputs);

        var origen = ValidarArchivo(rutaPlantillaOrigen, nameof(GenerarWorkbook));
        ValidarRutaSalida(rutaSalida);
        ValidarNoInPlace(origen, rutaSalida);

        WorkbookLeafCoherence.ValidarContraResultado(leafInputs, resultado);

        var directorioSalida = Path.GetDirectoryName(rutaSalida);
        if (!string.IsNullOrWhiteSpace(directorioSalida))
        {
            Directory.CreateDirectory(directorioSalida);
        }

        File.Copy(origen, rutaSalida, overwrite: true);

        try
        {
            using (var workbook = SpreadsheetDocument.Open(rutaSalida, true))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException("El workbook abierto no tiene WorkbookPart válido.");
                ValidarFormulasProtegidas(workbookPart, nameof(GenerarWorkbook));
                EscribirCeldasLeaf(workbookPart, leafInputs);
                var workbookXml = workbookPart.Workbook
                    ?? throw new CalculoInvalidoException("El workbook abierto no tiene metadata Workbook válida.");
                workbookXml.Save();
            }

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, false))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException("El workbook generado no tiene WorkbookPart válido.");
                ValidarFormulasProtegidas(workbookPart, nameof(GenerarWorkbook));
            }
        }
        catch
        {
            if (File.Exists(rutaSalida))
            {
                File.Delete(rutaSalida);
            }

            throw;
        }
    }

    /// <inheritdoc />
    public void GenerarWorkbook(string rutaPlantillaOrigen, string rutaSalida, ResultadoRemuneracion resultado, IReadOnlyList<WorkbookLeafInputs> leafInputs)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantillaOrigen);
        ArgumentNullException.ThrowIfNull(rutaSalida);
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leafInputs);

        if (leafInputs.Count == 0)
        {
            throw new CalculoInvalidoException("La lista de leafs del período está vacía; no hay nada que escribir.");
        }

        var origen = ValidarArchivo(rutaPlantillaOrigen, nameof(GenerarWorkbook));
        ValidarRutaSalida(rutaSalida);
        ValidarNoInPlace(origen, rutaSalida);

        WorkbookLeafCoherence.ValidarContraResultadoMultiAse(leafInputs, resultado);

        // HU-08 (2.2): gate Σ empresas = visible de bloque por ASE antes de escribir (D4/G5).
        foreach (var leaf in leafInputs)
        {
            WorkbookLeafCoherence.ValidarSigmaEmpresas(leaf.Conciliacion, leaf);
        }

        // HU-11 (2.5): el sufijo de las hojas DetRetri/DetValiRetri depende del período.
        // Q1 = 2026071 (mapa HU-10 intacto); Q2 = 2026072 (mapa parametrizado + mapa 2.5).
        var esQuincena2 = leafInputs.Any(l => l.Periodo.NumeroQuincena == 2);

        var directorioSalida = Path.GetDirectoryName(rutaSalida);
        if (!string.IsNullOrWhiteSpace(directorioSalida))
        {
            Directory.CreateDirectory(directorioSalida);
        }

        File.Copy(origen, rutaSalida, overwrite: true);

        try
        {
            using (var workbook = SpreadsheetDocument.Open(rutaSalida, true))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException("El workbook abierto no tiene WorkbookPart válido.");
                ValidarFormulasProtegidasMultiAse(workbookPart, nameof(GenerarWorkbook), esQuincena2);
                foreach (var leaf in leafInputs.OrderBy(l => l.Ase.Id))
                {
                    EscribirCeldasLeafPorAse(workbookPart, leaf, esQuincena2);
                    EscribirCeldasEmpresa(workbookPart, leaf);
                    EscribirCeldasBanco(workbookPart, leaf);
                    EscribirCeldasBalanceSc(workbookPart, leaf);
                    EscribirCeldasAjustesSfT(workbookPart, leaf);
                }

                // HU-12 (2.6 ampliada, V0.4): DetRetri-Q2 (enteros por ASE + total) en la MISMA pasada.
                EscribirCeldasDetRetriQ2(workbookPart, leafInputs);

                var workbookXml = workbookPart.Workbook
                    ?? throw new CalculoInvalidoException("El workbook abierto no tiene metadata Workbook válida.");
                workbookXml.Save();
            }

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, false))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException("El workbook generado no tiene WorkbookPart válido.");
                ValidarFormulasProtegidasMultiAse(workbookPart, nameof(GenerarWorkbook), esQuincena2);
            }
        }
        catch
        {
            if (File.Exists(rutaSalida))
            {
                File.Delete(rutaSalida);
            }

            throw;
        }
    }

    private static void ValidarConsolidadoFormulario(WorkbookPart workbookPart, Worksheet worksheet)
    {
        var reglas = new Dictionary<string, string[]>
        {
            ["D9"] = ["F46", "Reporte Componentes R1"],
            ["D28"] = ["E41", "Rem. Anticipos R2"],
            ["D47"] = ["F48", "Reporte Componentes R1"],
            ["D66"] = ["D67", "Reversion Pagos R4"],
            ["D85"] = ["D47", "AJUSTES"],
            ["D104"] = ["D9", "D28", "D47", "D66", "D85"],
            ["D109"] = ["D104", "D108", "SUM"]
        };

        foreach (var (celda, fragmentos) in reglas)
        {
            var formula = ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, HojaConsolidado, nameof(EscribirConsolidado));
            if (string.IsNullOrWhiteSpace(formula))
            {
                throw new CalculoInvalidoException($"La celda '{celda}' de '{HojaConsolidado}' no contiene una fórmula válida.");
            }
        }
    }

    private static void ValidarCadenaFormulaR1R2R4(WorkbookPart workbookPart)
    {
        var hojaR1 = ObtenerHoja(workbookPart, HojaR1, nameof(EscribirConsolidado));
        var hojaR2 = ObtenerHoja(workbookPart, HojaR2, nameof(EscribirConsolidado));
        var hojaR4 = ObtenerHoja(workbookPart, HojaR4, nameof(EscribirConsolidado));

        ValidarCeldaTieneFormula(workbookPart, hojaR1, "F46", ["F25", "F41", "L25"], HojaR1, nameof(EscribirConsolidado));
        ValidarCeldaTieneFormula(workbookPart, hojaR1, "F48", ["F30", "F10", "L10"], HojaR1, nameof(EscribirConsolidado));
        ValidarCeldaTieneFormula(workbookPart, hojaR2, "E41", ["E15", "E26", "K15"], HojaR2, nameof(EscribirConsolidado));
        ValidarCeldaTieneFormula(workbookPart, hojaR4, "D67", ["D9", "P9"], HojaR4, nameof(EscribirConsolidado));
    }

    private static void ValidarHojaConLabelsEsperados(WorkbookPart workbookPart, string hoja, params string[] labelsEsperados)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(EscribirConsolidado));
        var textos = worksheet.Descendants<Cell>()
            .Select(c => LeerTextoCelda(c, workbookPart))
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Select(t => t.Trim())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var label in labelsEsperados)
        {
            var coincidencia = textos.Any(t => t.Contains(label, StringComparison.OrdinalIgnoreCase));
            if (!coincidencia)
            {
                throw new CalculoInvalidoException($"La hoja '{hoja}' no incluye el label esperado '{label}' y no es compatible con la semántica actual del workbook.");
            }
        }
    }

    private static string ValidarCeldaTieneFormula(WorkbookPart workbookPart, Worksheet worksheet, string celda, string[] fragmentosEsperados, string hoja, string operacion)
    {
        var cell = ObtenerCelda(worksheet, celda)
            ?? throw new CalculoInvalidoException($"La celda '{celda}' no existe en la hoja '{hoja}' para {operacion}. El workbook no es compatible con la estructura esperada.");

        if (cell.CellFormula is null)
        {
            throw new CalculoInvalidoException($"La celda '{celda}' de '{hoja}' debería seguir siendo fórmula; se detectó un valor fijo. No se puede continuar con la validación del workbook derivado.");
        }

        var formula = NormalizarFormula(LeerTextoCelda(cell, workbookPart));

        // Shared-formula awareness (HU-07 task 2.3): un follower con SharedIndex y texto vacío
        // es una réplica del maestro; se valida por presencia de fórmula, no por fragmentos.
        var esSharedFollower = string.IsNullOrWhiteSpace(formula) && cell.CellFormula.SharedIndex is not null;
        if (string.IsNullOrWhiteSpace(formula) && !esSharedFollower)
        {
            throw new CalculoInvalidoException($"La celda '{celda}' de '{hoja}' no tiene una fórmula resoluble por OpenXML.");
        }

        if (fragmentosEsperados.Length > 0 && !esSharedFollower)
        {
            var formulaNormalizada = NormalizarFormula(formula);
            var faltan = fragmentosEsperados
                .Select(NormalizarFormula)
                .Where(fragmento => !string.IsNullOrWhiteSpace(fragmento))
                .Where(fragmento => !formulaNormalizada.Contains(fragmento, StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (faltan.Length > 0)
            {
                var texto = cell.CellFormula.Text ?? formula;
                throw new CalculoInvalidoException($"La celda '{celda}' de '{hoja}' no mantiene todas las referencias esperadas: faltan [{string.Join(", ", faltan)}]. Fórmula actual: '{texto}'.");
            }
        }

        return formula;
    }

    private static string ValidarArchivo(string rutaPlantilla, string operacion)
    {
        if (string.IsNullOrWhiteSpace(rutaPlantilla))
        {
            throw new ArchivoFuenteNoEncontradoException($"La ruta de plantilla para {operacion} es requerida.");
        }

        if (!File.Exists(rutaPlantilla))
        {
            throw new ArchivoFuenteNoEncontradoException($"No se encontró la plantilla para {operacion}: '{rutaPlantilla}'.");
        }

        return rutaPlantilla;
    }

    private static void ValidarRutaSalida(string rutaSalida)
    {
        if (string.IsNullOrWhiteSpace(rutaSalida))
        {
            throw new ArchivoFuenteNoEncontradoException("La ruta de salida del workbook es requerida.");
        }
    }

    private static void ValidarNoInPlace(string origen, string rutaSalida)
    {
        if (string.Equals(Path.GetFullPath(origen), Path.GetFullPath(rutaSalida), StringComparison.OrdinalIgnoreCase))
        {
            throw new CalculoInvalidoException("La escritura real no puede mutar la plantilla original in-place. Use una ruta de salida distinta.");
        }
    }

    private static Worksheet ObtenerHoja(WorkbookPart workbookPart, string nombreHoja, string operacion)
    {
        var workbook = workbookPart.Workbook ?? throw new CalculoInvalidoException($"El workbook para {operacion} no tiene metadata Workbook válida.");
        var sheet = workbook.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException($"La hoja '{nombreHoja}' no existe en el workbook para {operacion}.");

        var worksheetPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart
            ?? throw new CalculoInvalidoException($"No se pudo resolver la hoja '{nombreHoja}' en el workbook para {operacion}.");

        return worksheetPart.Worksheet ?? throw new CalculoInvalidoException($"La hoja '{nombreHoja}' no tiene Worksheet válido.");
    }

    private static Worksheet ObtenerHoja(SpreadsheetDocument workbook, string nombreHoja, string operacion)
    {
        var workbookPart = workbook.WorkbookPart ?? throw new CalculoInvalidoException($"El workbook para {operacion} no tiene WorkbookPart.");
        return ObtenerHoja(workbookPart, nombreHoja, operacion);
    }

    private static void LanzarPayloadInsuficiente(string operacion, string hoja, string detalle)
    {
        throw new CalculoInvalidoException($"La escritura funcional de '{operacion}' sobre '{hoja}' no está soportada con el payload actual. {detalle} Requiere un nuevo contrato WorkbookLeafInputs para la hoja real.");
    }

    private static Cell? ObtenerCelda(Worksheet worksheet, string cellReference)
    {
        var sheetData = worksheet.Elements<SheetData>().FirstOrDefault();
        if (sheetData is null)
        {
            return null;
        }

        return sheetData.Descendants<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));
    }

    private static string LeerTextoCelda(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.CellFormula is not null)
        {
            return cell.CellFormula.Text ?? string.Empty;
        }

        if (cell.CellValue is null)
        {
            return string.Empty;
        }

        if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString)
        {
            var sharedStrings = workbookPart.SharedStringTablePart;
            if (sharedStrings is not null && sharedStrings.SharedStringTable is not null
                && int.TryParse(cell.CellValue.Text, out var index) && index >= 0
                && index < sharedStrings.SharedStringTable.Count())
            {
                var item = sharedStrings.SharedStringTable.ElementAt(index);
                var text = item.InnerText;
                return text ?? string.Empty;
            }
        }

        if (cell.DataType is not null && cell.DataType.Value == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText ?? string.Empty;
        }

        return cell.CellValue.InnerText;
    }

    private static string NormalizarFormula(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return string.Empty;
        }

        var value = formula.Trim();
        value = value.Replace("'", string.Empty, StringComparison.Ordinal);
        value = value.Replace(" ", string.Empty, StringComparison.Ordinal);
        value = value.Replace("_xlfn.", string.Empty, StringComparison.OrdinalIgnoreCase);
        return value;
    }

    private static void ValidarFormulasProtegidas(WorkbookPart workbookPart, string operacion)
    {
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMap.ProtectedFormulas)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }
    }

    /// <summary>
    /// Valida el mapa ampliado: visibles de cada bloque ASE (R1/R2/R4) + filas CONSOLIDADO
    /// D9:D13/D28:D32/D47:D51/D66:D70/D85:D89/D104:D108/D109 con shared-formula awareness,
    /// + mapa 2.2 (HU-08): <c>REMUNERACION_*</c>, <c>VALIDACION_*</c>, <c>GERENTES_*</c>,
    /// <c>Recaudo *</c> fila 29+ (D6) + HU-09 (banco) + HU-10 (BCE, parametrizado al período).
    /// HU-11 (2.5): en Q2 además valida la cadena AJUSTES-SF-T (<see cref="WorkbookLeafCellMapAjustesSfT.Protegidas"/>).
    /// HU-12 (2.6 ampliada, D5): validación PROTEGIDA parametrizada por período — Q1 exige las
    /// fórmulas HU-07 (F46/F176/…/D67, invariante Q1); Q2 exige el mapa T0 (<see cref="WorkbookLeafCellMapQ2"/>:
    /// F53…, E43…, D73…, DetValiRetri/VALIDACION_*/INTERVENTORIA/ANT EXT-REV siempre protegidas).
    /// </summary>
    private static void ValidarFormulasProtegidasMultiAse(WorkbookPart workbookPart, string operacion, bool esQuincena2)
    {
        if (esQuincena2)
        {
            ValidarFormulasProtegidasMultiAseQ2(workbookPart, operacion);
            return;
        }

        foreach (var aseId in WorkbookLeafCellMapPorAse.EditableLeafCellsPorAse.Keys.OrderBy(k => k))
        {
            var bloque = WorkbookLeafCellMapPorAse.ProtectedFormulasPorAse[aseId];
            foreach (var (hoja, celda, fragmentos) in bloque)
            {
                if (string.Equals(hoja, HojaConsolidado, StringComparison.OrdinalIgnoreCase))
                {
                    continue; // las filas CONSOLIDADO se validan una sola vez abajo
                }

                var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
                ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
            }
        }

        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapPorAseProtectedConsolidado)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapPorEmpresa.ProtectedFormulasPorEmpresa)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-09 (2.3, D6): mapa de fórmulas protegidas del reporte por banco (consolidado 1–7,
        // Total de bloque, verificación I/J, validación 59–80 y TOTAL RECAUDO). Nunca se escriben.
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapReporteBanco.Protegidas)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-10 (2.4, D6): mapa de fórmulas protegidas 2.4 (BCE F/I/H + filas 9/10/11/12/13 +
        // bloque 18–24 + CONSOLIDADO J/K/M + refs DetRetri/DetValiRetri). Jamás se escriben.
        // HU-11 (2.5): el sufijo de las hojas DetRetri/DetValiRetri depende del período; el mapa
        // HU-10 queda intacto (Q1 = 2026071) y en Q2 se parametriza al sufijo 2026072.
        foreach (var (hoja, celda, fragmentos) in ProtegidasBceParaPeriodo(esQuincena2))
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-11 (2.5, Requirement 7): cadena AJUSTES-SF-T protegida (visibles SALDOS/RETRIBUCION,
        // AJUSTES D9:D13/D28:D32/D47:D51, CONSOLIDADO D85:D89, INTERVENTORIA, ANT EXT-REV).
        if (esQuincena2)
        {
            foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapAjustesSfT.Protegidas)
            {
                var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
                ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
            }
        }

        // HU-13 (2.7, Requirement 4/D4): mapa protegido extendido de TODAS las hojas de
        // validación (VALIDACION_*, VALIDACION_TOTAL, DetRetri/DetValiRetri col D, Valida -*,
        // GERENTES_*). El writer falla si alguna deja de ser fórmula donde T0 lo exige.
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapValidaciones.ProtegidasValidacionesParaPeriodo(esQuincena2 ? 2 : 1))
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }
    }

    /// <summary>
    /// HU-12 (2.6 ampliada, D5): validación protegida Q2 contra el mapa T0 congelado
    /// (<see cref="WorkbookLeafCellMapQ2"/>). Q1 NO se valida aquí (rama por período): las celdas
    /// F46/F48/F519/F521 del mapa HU-07 son VALORES en Q2 (V0.2) y quedan EXCLUIDAS de
    /// protegidas-fórmula. Incluye: visibles R1/R2/R4-Q2 por ASE, CONSOLIDADO Q2, banco HU-09,
    /// BCE parametrizado 2026072 (M1), cadena AJUSTES-SF-T, DetRetri/DetValiRetri Q2 y las
    /// protegidas adicionales (REMUNERACION_* genéricas + VALIDACION_*/GERENTES_*/INTERVENTORIA/
    /// ANT EXT-REV — Requirement 5). El writer falla si alguna deja de ser fórmula.
    /// </summary>
    private static void ValidarFormulasProtegidasMultiAseQ2(WorkbookPart workbookPart, string operacion)
    {
        foreach (var aseId in WorkbookLeafCellMapQ2.R1Q2EditablesPorAse.Keys.OrderBy(k => k))
        {
            foreach (var (celda, fragmentos) in WorkbookLeafCellMapQ2.ObtenerR1Q2Protegidos(aseId))
            {
                var worksheet = ObtenerHoja(workbookPart, HojaR1, operacion);
                ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, HojaR1, operacion);
            }

            foreach (var (celda, fragmentos) in WorkbookLeafCellMapQ2.ObtenerR2Q2Protegidos(aseId))
            {
                var worksheet = ObtenerHoja(workbookPart, HojaR2, operacion);
                ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, HojaR2, operacion);
            }

            foreach (var (celda, fragmentos) in WorkbookLeafCellMapQ2.ObtenerR4Q2Protegidos(aseId))
            {
                var worksheet = ObtenerHoja(workbookPart, HojaR4, operacion);
                ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, HojaR4, operacion);
            }
        }

        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.ConsolidadoProtected)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-09 (2.3): mapa banco parametrizado al período — el Q2 excluye el TOTAL RECAUDO fila 81
        // (C81/D81) que no existe en el template Q2 (T0-0.5; la hoja termina en la fila 79).
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.BancoProtegidasQ2)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-10 (2.4): BCE parametrizado al sufijo 2026072 (M1: matchea hojas/celdas reales,
        // incl. DetRetri2026072/DetValiRetri2026072).
        foreach (var (hoja, celda, fragmentos) in ProtegidasBceParaPeriodo(true))
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-11 (2.5): cadena AJUSTES-SF-T (Q2).
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapAjustesSfT.Protegidas)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-12 (2.6 ampliada, Requirement 5): DetRetri/DetValiRetri Q2 (T0-0.4) + protegidas
        // adicionales (REMUNERACION_* genéricas, VALIDACION_*, GERENTES_*, INTERVENTORIA, ANT EXT-REV).
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.DetRetriProtected)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.ProtegidasAdicionalesQ2)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-13 (2.7, Requirement 4/D4): mapa protegido extendido 2.7 en Q2 (VALIDACION_* O/P
        // filas 3..7, VALIDACION_TOTAL O/P, Valida -*, GERENTES_* SUM). DetRetri/DetValiRetri Q2
        // ya cubiertos por WorkbookLeafCellMapQ2.DetRetriProtected (no se duplica).
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapValidaciones.ProtegidasValidacionesParaPeriodo(2))
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }
    }

    /// <summary>
    /// HU-11 (2.5): devuelve el mapa protegido BCE con el sufijo de hojas DetRetri/DetValiRetri
    /// correcto para el período (Q1 = 2026071 intacto; Q2 = 2026072). El mapa HU-10 NO se toca
    /// (regla del plan): solo se parametriza su interpretación por período.
    /// </summary>
    private static IEnumerable<(string Hoja, string Celda, string[] Fragmentos)> ProtegidasBceParaPeriodo(bool esQuincena2)
    {
        if (!esQuincena2)
        {
            return WorkbookLeafCellMapBalanceSc.Protegidas;
        }

        return WorkbookLeafCellMapBalanceSc.Protegidas
            .Select(p => (
                ReemplazarSufijo(p.Hoja),
                p.Celda,
                p.Fragmentos.Select(ReemplazarSufijo).ToArray()));
    }

    private static string ReemplazarSufijo(string texto) =>
        texto.Replace(SufijoHojasQ1, SufijoHojasQ2, StringComparison.Ordinal);

    /// <summary>
    /// Filas CONSOLIDADO del mapa ampliado multi-ASE (plan §2.3): D9:D13, D28:D32, D47:D51,
    /// D66:D70, D85:D89, D104:D108 (maestro + réplicas shared), D109. Fórmulas jamás se escriben.
    /// </summary>
    private static readonly (string Hoja, string Celda, string[] Fragmentos)[] WorkbookLeafCellMapPorAseProtectedConsolidado =
    [
        (HojaConsolidado, "D9", ["F46", "Reporte Componentes R1"]),
        (HojaConsolidado, "D10", ["F176", "Reporte Componentes R1"]),
        (HojaConsolidado, "D11", ["F316", "Reporte Componentes R1"]),
        (HojaConsolidado, "D12", ["F437", "Reporte Componentes R1"]),
        (HojaConsolidado, "D13", ["F519", "Reporte Componentes R1"]),
        (HojaConsolidado, "D28", ["E41", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D29", ["E135", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D30", ["E247", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D31", ["E343", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D32", ["E413", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D47", ["F48", "Reporte Componentes R1"]),
        (HojaConsolidado, "D48", ["F178", "Reporte Componentes R1"]),
        (HojaConsolidado, "D49", ["F318", "Reporte Componentes R1"]),
        (HojaConsolidado, "D50", ["F439", "Reporte Componentes R1"]),
        (HojaConsolidado, "D51", ["F521", "Reporte Componentes R1"]),
        (HojaConsolidado, "D66", ["D67", "Reversion Pagos R4"]),
        (HojaConsolidado, "D67", ["D161", "Reversion Pagos R4"]),
        (HojaConsolidado, "D68", ["D198", "Reversion Pagos R4"]),
        (HojaConsolidado, "D69", ["D312", "Reversion Pagos R4"]),
        (HojaConsolidado, "D70", ["D347", "Reversion Pagos R4"]),
        (HojaConsolidado, "D85", ["D47", "AJUSTES"]),
        (HojaConsolidado, "D86", ["D48", "AJUSTES"]),
        (HojaConsolidado, "D87", ["D49", "AJUSTES"]),
        (HojaConsolidado, "D88", ["D50", "AJUSTES"]),
        (HojaConsolidado, "D89", ["D51", "AJUSTES"]),
        (HojaConsolidado, "D104", ["D9", "D28", "D47", "D66", "D85"]),
        (HojaConsolidado, "D105", ["D10", "D29", "D48", "D67", "D86"]),
        (HojaConsolidado, "D106", []), // shared follower del maestro D104
        (HojaConsolidado, "D107", ["D12", "D31", "D50", "D69", "D88"]),
        (HojaConsolidado, "D108", ["D13", "D32", "D51", "D70", "D89"]),
        (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
    ];

    private static void EscribirCeldasLeaf(WorkbookPart workbookPart, WorkbookLeafInputs leafInputs)
    {
        var valores = new Dictionary<(string Hoja, string Celda), decimal>(new LeafCellComparer())
        {
            [(WorkbookLeafCellMap.HojaR1, "F25")] = leafInputs.R1.F25,
            [(WorkbookLeafCellMap.HojaR1, "F41")] = leafInputs.R1.F41,
            [(WorkbookLeafCellMap.HojaR1, "L25")] = leafInputs.R1.L25,
            [(WorkbookLeafCellMap.HojaR1, "F30")] = leafInputs.R1.F30,
            [(WorkbookLeafCellMap.HojaR1, "F10")] = leafInputs.R1.F10,
            [(WorkbookLeafCellMap.HojaR1, "L10")] = leafInputs.R1.L10,
            [(WorkbookLeafCellMap.HojaR2, "E15")] = leafInputs.R2.E15,
            [(WorkbookLeafCellMap.HojaR2, "E26")] = leafInputs.R2.E26,
            [(WorkbookLeafCellMap.HojaR2, "K15")] = leafInputs.R2.K15,
            [(WorkbookLeafCellMap.HojaR4, "D9")] = leafInputs.R4.D9,
            [(WorkbookLeafCellMap.HojaR4, "P9")] = leafInputs.R4.P9
        };

        foreach (var (hoja, celda, nombre) in WorkbookLeafCellMap.EditableLeafCells)
        {
            if (WorkbookLeafCellMap.ProtectedFormulas.Any(p =>
                    string.Equals(p.Hoja, hoja, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(p.Celda, celda, StringComparison.OrdinalIgnoreCase)))
            {
                throw new CalculoInvalidoException(
                    $"El cell-map intentó escribir '{nombre}' sobre la fórmula protegida {hoja}!{celda}.");
            }

            if (!valores.TryGetValue((hoja, celda), out var valor))
            {
                throw new CalculoInvalidoException($"No hay valor leaf mapeado para {nombre} ({hoja}!{celda}).");
            }

            EscribirValorNumerico(workbookPart, hoja, celda, valor, nombre);
        }
    }

    /// <summary>
    /// Escribe SOLO las celdas del bloque del ASE (mapa por ASE congelado por T0).
    /// Los valores provienen de <see cref="WorkbookLeafInputs"/> (CeldasPorAse por hoja).
    /// HU-12 (2.6 ampliada, D1/G2): en Q2 escribe el mapa hermano <see cref="WorkbookLeafCellMapQ2"/>
    /// (R1/R2/R4-Q2 por ASE, incluida la variante ASE5 de 2 filas V0.3); Q1 queda intacto.
    /// </summary>
    private static void EscribirCeldasLeafPorAse(WorkbookPart workbookPart, WorkbookLeafInputs leaf, bool esQuincena2)
    {
        if (esQuincena2)
        {
            EscribirCeldasLeafPorAseQ2(workbookPart, leaf);
            return;
        }

        var editables = WorkbookLeafCellMapPorAse.ObtenerEditables(leaf.Ase.Id);

        foreach (var (hoja, celda, nombre) in editables)
        {
            if (WorkbookLeafCellMapPorAse.ProtectedFormulasPorAse[leaf.Ase.Id].Any(p =>
                    string.Equals(p.Hoja, hoja, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(p.Celda, celda, StringComparison.OrdinalIgnoreCase)))
            {
                throw new CalculoInvalidoException(
                    $"El cell-map del ASE {leaf.Ase.Id} intentó escribir '{nombre}' sobre la fórmula protegida {hoja}!{celda}.");
            }

            var valor = ObtenerValorLeaf(leaf, hoja, celda, nombre);
            EscribirValorNumerico(workbookPart, hoja, celda, valor, nombre);
        }
    }

    private static decimal ObtenerValorLeaf(WorkbookLeafInputs leaf, string hoja, string celda, string nombre)
    {
        IReadOnlyDictionary<string, decimal> celdas = hoja switch
        {
            _ when string.Equals(hoja, HojaR1, StringComparison.OrdinalIgnoreCase) => leaf.R1.CeldasPorAse,
            _ when string.Equals(hoja, HojaR2, StringComparison.OrdinalIgnoreCase) => leaf.R2.CeldasPorAse,
            _ when string.Equals(hoja, HojaR4, StringComparison.OrdinalIgnoreCase) => leaf.R4.CeldasPorAse,
            _ => throw new CalculoInvalidoException($"Hoja '{hoja}' no tiene valores leaf mapeados para '{nombre}'.")
        };

        if (!celdas.TryGetValue(celda, out var valor))
        {
            throw new CalculoInvalidoException($"No hay valor leaf mapeado para {nombre} ({hoja}!{celda}) del ASE {leaf.Ase.Id}.");
        }

        return valor;
    }

    /// <summary>
    /// HU-12 (2.6 ampliada, T0-0.2/0.3): escribe SOLO las celdas Q2 del bloque del ASE
    /// (R1/R2/R4 del mapa <see cref="WorkbookLeafCellMapQ2"/>), incluida la variante ASE5 de 2
    /// filas (V0.3). Fail-fast si falta un valor mapeado (nombra ASE + hoja + celda; nunca 0
    /// silencioso). El guard de <see cref="EscribirValorNumerico"/> impide tocar fórmulas.
    /// </summary>
    private static void EscribirCeldasLeafPorAseQ2(WorkbookPart workbookPart, WorkbookLeafInputs leaf)
    {
        foreach (var (celda, _) in WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(leaf.Ase.Id))
        {
            if (!leaf.R1.CeldasPorAse.TryGetValue(celda, out var valorR1))
            {
                throw new CalculoInvalidoException($"No hay valor leaf R1-Q2 mapeado para {celda} del ASE {leaf.Ase.Id}.");
            }

            EscribirValorNumerico(workbookPart, HojaR1, celda, valorR1, $"R1-Q2.ASE{leaf.Ase.Id}.{celda}");
        }

        var r2 = WorkbookLeafCellMapQ2.ObtenerR2Q2Editables(leaf.Ase.Id);
        foreach (var celda in new[] { r2.Componente, r2.SubsCont, r2.Especiales })
        {
            if (!leaf.R2.CeldasPorAse.TryGetValue(celda, out var valorR2))
            {
                throw new CalculoInvalidoException($"No hay valor leaf R2-Q2 mapeado para {celda} del ASE {leaf.Ase.Id}.");
            }

            EscribirValorNumerico(workbookPart, HojaR2, celda, valorR2, $"R2-Q2.ASE{leaf.Ase.Id}.{celda}");
        }

        var r4 = WorkbookLeafCellMapQ2.ObtenerR4Q2Editables(leaf.Ase.Id);
        foreach (var celda in new[] { r4.Total, r4.P })
        {
            if (!leaf.R4.CeldasPorAse.TryGetValue(celda, out var valorR4))
            {
                throw new CalculoInvalidoException($"No hay valor leaf R4-Q2 mapeado para {celda} del ASE {leaf.Ase.Id}.");
            }

            EscribirValorNumerico(workbookPart, HojaR4, celda, valorR4, $"R4-Q2.ASE{leaf.Ase.Id}.{celda}");
        }
    }

    /// <summary>
    /// HU-12 (2.6 ampliada, V0.4): escribe DetRetri-Q2 (hoja <c>DetRetri2026072</c>) en la MISMA
    /// pasada atómica: D9:D13 = ROUND(D104:D108) por ASE (vía <see cref="DetRetriRounder"/>, única
    /// regla) y D14 = ROUND(Σ D104:D108). <c>leaf.DetRetriQ2 == null</c> para todos = Q1 (no-op).
    /// La composición está CONGELADA (probada 5/5 contra el golden, V0.4); nunca se inventa.
    /// </summary>
    private static void EscribirCeldasDetRetriQ2(WorkbookPart workbookPart, IReadOnlyList<WorkbookLeafInputs> leafInputs)
    {
        var conDetalle = leafInputs
            .Where(l => l.DetRetriQ2 is not null)
            .OrderBy(l => l.Ase.Id)
            .ToList();
        if (conDetalle.Count == 0)
        {
            return; // Q1 puro: sin escrituras 2.6.
        }

        foreach (var leaf in conDetalle)
        {
            var celda = WorkbookLeafCellMapQ2.ObtenerDetRetriDestino(leaf.Ase.Id);
            EscribirValorNumericoEnCeldaExistente(workbookPart, WorkbookLeafCellMapQ2.HojaDetRetri, celda, leaf.DetRetriQ2!.Detalle, $"DetRetri-Q2.ASE{leaf.Ase.Id}.{celda}");
        }

        var totalD104 = conDetalle.Sum(l => l.DetRetriQ2!.TotalD104);
        EscribirValorNumericoEnCeldaExistente(workbookPart, WorkbookLeafCellMapQ2.HojaDetRetri, WorkbookLeafCellMapQ2.DetRetriTotal, DetRetriRounder.Round(totalD104), "DetRetri-Q2.D14");
    }

    /// <summary>
    /// HU-13 (2.7, W1-guarda D8): escribe un valor SOLO si la celda destino YA EXISTE en la
    /// plantilla (nunca crea filas ni celdas: si una plantilla futura trae más/menos filas ASE en
    /// DetRetri D9:D14, el fail-fast NOMBRA la hoja; jamás truncado silencioso ni insert/delete).
    /// Usado por la escritura DetRetri-Q2 (V0.4: D9:D14 son VALORES editables del template).
    /// </summary>
    private static void EscribirValorNumericoEnCeldaExistente(WorkbookPart workbookPart, string hoja, string celda, decimal valor, string nombre)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));

        if (ObtenerCelda(worksheet, celda) is null)
        {
            throw new CalculoInvalidoException(
                $"Guarda de capacidad: la celda '{hoja}!{celda}' no existe en la plantilla para escribir '{nombre}'. La plantilla futura cambió la capacidad D9:D14 del ASE — fail-fast, nunca truncado silencioso.");
        }

        EscribirValorNumerico(workbookPart, hoja, celda, valor, nombre);
    }

    /// <summary>
    /// HU-08 (2.2): escribe en la MISMA pasada los operandos por empresa (R1/R2/R4) y las hojas
    /// <c>Recaudo *</c> en valores. Solo celdas del mapa congelado; el guard de fórmulas
    /// (<see cref="EscribirValorNumerico"/>) impide tocar cualquier celda con <c>&lt;f&gt;</c>.
    /// </summary>
    private static void EscribirCeldasEmpresa(WorkbookPart workbookPart, WorkbookLeafInputs leaf)
    {
        foreach (var conc in leaf.Conciliacion)
        {
            EscribirCeldasConciliacion(workbookPart, leaf.Ase.Id, conc);
        }

        foreach (var recaudo in leaf.Recaudos)
        {
            foreach (var (celda, valor) in recaudo.Celdas)
            {
                EscribirValorNumerico(workbookPart, recaudo.HojaRecaudo, celda, valor, $"Recaudo.{recaudo.Empresa.Nombre}.{celda}");
            }
        }
    }

    private static void EscribirCeldasConciliacion(WorkbookPart workbookPart, int aseId, ConciliacionEmpresaInputs conc)
    {
        if (WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.TryGetValue((conc.Empresa.Id, aseId), out var r1))
        {
            foreach (var (celda, _) in r1)
            {
                if (!conc.CeldasR1.TryGetValue(celda, out var valor))
                {
                    throw new CalculoInvalidoException($"No hay valor R1 por empresa para {conc.Empresa.Nombre} ({HojaR1}!{celda}) del ASE {aseId}.");
                }

                EscribirValorNumerico(workbookPart, HojaR1, celda, valor, $"R1.{conc.Empresa.Nombre}.{celda}");
            }
        }

        if (WorkbookLeafCellMapPorEmpresa.EditablesR2PorEmpresa.TryGetValue((conc.Empresa.Id, aseId), out var r2))
        {
            foreach (var (celda, _) in r2)
            {
                if (!conc.CeldasR2.TryGetValue(celda, out var valor))
                {
                    throw new CalculoInvalidoException($"No hay valor R2 por empresa para {conc.Empresa.Nombre} ({HojaR2}!{celda}) del ASE {aseId}.");
                }

                EscribirValorNumerico(workbookPart, HojaR2, celda, valor, $"R2.{conc.Empresa.Nombre}.{celda}");
            }
        }

        if (WorkbookLeafCellMapPorEmpresa.EditablesR4PorEmpresa.TryGetValue((conc.Empresa.Id, aseId), out var r4))
        {
            foreach (var (celda, _) in r4)
            {
                if (!conc.CeldasR4.TryGetValue(celda, out var valor))
                {
                    throw new CalculoInvalidoException($"No hay valor R4 por empresa para {conc.Empresa.Nombre} ({HojaR4}!{celda}) del ASE {aseId}.");
                }

                EscribirValorNumerico(workbookPart, HojaR4, celda, valor, $"R4.{conc.Empresa.Nombre}.{celda}");
            }
        }
    }

    /// <summary>
    /// HU-09 (2.3): escribe SOLO las celdas del mapa banco del ASE (bloques 9–58 en valores
    /// + C59 = quincena) en la MISMA pasada atómica HU-07/HU-08 (G4/D3). <c>ReporteBanco == null</c>
    /// = comportamiento HU-08 puro. El guard de <see cref="EscribirValorNumerico"/> impide tocar
    /// cualquier celda con &lt;f&gt; (filas 1–7, Total de bloque, 59–80 y TOTAL quedan intactos).
    /// </summary>
    private static void EscribirCeldasBanco(WorkbookPart workbookPart, WorkbookLeafInputs leaf)
    {
        if (leaf.ReporteBanco is null)
        {
            return; // HU-08 puro (lista 2.3 vacía = comportamiento existente intacto).
        }

        var bloque = leaf.ReporteBanco.Ases.SingleOrDefault(b => b.Ase.Id == leaf.Ase.Id)
            ?? throw new CalculoInvalidoException(
                $"El leaf del ASE {leaf.Ase.Id} no trae su bloque de reporte banco para escribir.");

        var editables = WorkbookLeafCellMapReporteBanco.ObtenerEditables(leaf.Ase.Id);

        foreach (var empresa in bloque.Empresas)
        {
            if (!editables.TryGetValue(empresa.Empresa, out var celdas))
            {
                throw new CalculoInvalidoException(
                    $"El bloque banco del ASE {leaf.Ase.Id} trae la empresa '{empresa.Empresa}' pero el mapa T0-0.7 no la declara para ese ASE.");
            }

            var valores = new[] { empresa.AplicadosFacturacion, empresa.SaldosFavorGenerados, empresa.FinanciacionesNuevas, empresa.RecibosServEspeciales };
            for (var i = 0; i < celdas.Length && i < valores.Length; i++)
            {
                EscribirValorNumerico(workbookPart, HojaBanco, celdas[i], valores[i], $"Banco.{empresa.Empresa}.{celdas[i]}");
            }
        }

        // C59 = quincena (dominio, nunca fuente; Requirement 4).
        EscribirValorNumerico(workbookPart, HojaBanco, "C59", leaf.ReporteBanco.Quincena, "Banco.C59");
    }

    /// <summary>
    /// HU-10 (2.4): escribe SOLO las celdas D/E de la hoja <c>BCE SC POR FACT.</c> del ASE
    /// (Contribución → D, Subsidio → E según veredicto T0-0.2) en la MISMA pasada atómica
    /// HU-07/HU-08/HU-09 (G5/D3). <c>BalanceSc == null</c> = comportamiento HU-09 puro. La
    /// columna H es fórmula (D2(b), T0-0.6) → nunca se escribe; F/I/filas 9/11/18–24 y los
    /// downstream J/K/M/DetRetri/DetValiRetri quedan protegidos por
    /// <see cref="WorkbookLeafCellMapBalanceSc.Protegidas"/> y el guard de
    /// <see cref="EscribirValorNumerico"/> impide tocar cualquier celda con &lt;f&gt;.
    /// </summary>
    private static void EscribirCeldasBalanceSc(WorkbookPart workbookPart, WorkbookLeafInputs leaf)
    {
        if (leaf.BalanceSc is null)
        {
            return; // HU-09 puro (lista 2.4 vacía = comportamiento existente intacto).
        }

        var bloque = leaf.BalanceSc.Ases.SingleOrDefault(b => b.Ase.Id == leaf.Ase.Id)
            ?? throw new CalculoInvalidoException(
                $"El leaf del ASE {leaf.Ase.Id} no trae su fila de balance SC para escribir.");

        var editables = WorkbookLeafCellMapBalanceSc.ObtenerEditables(leaf.Ase.Id);
        EscribirValorNumerico(workbookPart, HojaBce, editables.Contribucion, bloque.Contribucion, $"BCE.Contribucion.ASE{leaf.Ase.Id}");
        EscribirValorNumerico(workbookPart, HojaBce, editables.Subsidio, bloque.Subsidio, $"BCE.Subsidio.ASE{leaf.Ase.Id}");
    }

    /// <summary>
    /// HU-11 (2.5): escribe los operandos editables de SALDOS POR NOTA y RETRIBUCION NEGATIVA
    /// del ASE (bloques de valores T0-0.7) en la MISMA pasada atómica HU-07..HU-10 (D2a/G8).
    /// <c>AjustesSfT == null</c> = comportamiento HU-10 puro (Q1, G3). Los visibles Cn-In de
    /// cada bloque, la hoja AJUSTES-SF-T, INTERVENTORIA y ANT EXT-REV son fórmulas protegidas
    /// (mapa <see cref="WorkbookLeafCellMapAjustesSfT.Protegidas"/> + guard de
    /// <see cref="EscribirValorNumerico"/>) → jamás se escriben. La columna I (Especiales) del
    /// template se escribe con 0 (fuente sin esa columna, T0-0.5).
    /// </summary>
    private static void EscribirCeldasAjustesSfT(WorkbookPart workbookPart, WorkbookLeafInputs leaf)
    {
        if (leaf.AjustesSfT is null)
        {
            return; // HU-10 puro (Q1): sin escrituras 2.5.
        }

        foreach (var (celda, valor) in leaf.AjustesSfT.SaldosNotas.Celdas)
        {
            EscribirValorNumerico(workbookPart, HojaSaldosNotas, celda, valor, $"SaldosNotas.ASE{leaf.Ase.Id}.{celda}");
        }

        foreach (var (celda, valor) in leaf.AjustesSfT.RetribucionNegativa.Celdas)
        {
            EscribirValorNumerico(workbookPart, HojaRetribucionNegativa, celda, valor, $"RetribucionNegativa.ASE{leaf.Ase.Id}.{celda}");
        }
    }

    private static void EscribirValorNumerico(WorkbookPart workbookPart, string hoja, string celda, decimal valor, string nombre)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));
        var cell = ObtenerOCrearCelda(worksheet, celda);

        if (cell.CellFormula is not null)
        {
            throw new CalculoInvalidoException(
                $"La celda leaf {nombre} ({hoja}!{celda}) es fórmula en la plantilla. HU-05 no puede sobrescribir fórmulas.");
        }

        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture));
    }

    private static Cell ObtenerOCrearCelda(Worksheet worksheet, string cellReference)
    {
        var existente = ObtenerCelda(worksheet, cellReference);
        if (existente is not null)
        {
            return existente;
        }

        var sheetData = worksheet.Elements<SheetData>().FirstOrDefault()
            ?? throw new CalculoInvalidoException($"La hoja no tiene SheetData para crear la celda '{cellReference}'.");

        var (_, filaNumero) = ParsearReferencia(cellReference);
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == filaNumero);
        if (row is null)
        {
            row = new Row { RowIndex = filaNumero };
            var siguientes = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value > filaNumero);
            if (siguientes is null)
            {
                sheetData.Append(row);
            }
            else
            {
                sheetData.InsertBefore(row, siguientes);
            }
        }

        var nueva = new Cell { CellReference = cellReference };
        var siguienteCelda = row.Elements<Cell>().FirstOrDefault(c =>
            string.Compare(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase) > 0);
        if (siguienteCelda is null)
        {
            row.Append(nueva);
        }
        else
        {
            row.InsertBefore(nueva, siguienteCelda);
        }

        return nueva;
    }

    private static (string Columna, uint Fila) ParsearReferencia(string cellReference)
    {
        var match = Regex.Match(cellReference, @"^(?<col>[A-Za-z]+)(?<row>\d+)$");
        if (!match.Success)
        {
            throw new CalculoInvalidoException($"Referencia de celda inválida: '{cellReference}'.");
        }

        return (match.Groups["col"].Value.ToUpperInvariant(), uint.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture));
    }

    private sealed class LeafCellComparer : IEqualityComparer<(string Hoja, string Celda)>
    {
        public bool Equals((string Hoja, string Celda) x, (string Hoja, string Celda) y) =>
            string.Equals(x.Hoja, y.Hoja, StringComparison.OrdinalIgnoreCase)
            && string.Equals(x.Celda, y.Celda, StringComparison.OrdinalIgnoreCase);

        public int GetHashCode((string Hoja, string Celda) obj) =>
            HashCode.Combine(obj.Hoja.ToUpperInvariant(), obj.Celda.ToUpperInvariant());
    }
}