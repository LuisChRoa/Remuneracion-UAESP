using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Writer OpenXML de la plantilla real.
/// HU-04: validación estructural no destructiva (<see cref="IPlantillaWriter"/>).
/// HU-05: escritura real de celdas leaf sobre una copia (<see cref="IWorkbookLeafWriter"/>).
/// HU-07: overload multi-ASE con mapa por bloque (<see cref="WorkbookLeafCellMapPorAse"/>),
/// una sola copia, validación pre/post ampliada y borrado de parcial ante fallo.
/// </summary>
public class OpenXmlPlantillaWriter : IPlantillaWriter, IWorkbookLeafWriter
{
    private const string HojaConsolidado = WorkbookLeafCellMap.HojaConsolidado;
    private const string HojaR1 = WorkbookLeafCellMap.HojaR1;
    private const string HojaR2 = WorkbookLeafCellMap.HojaR2;
    private const string HojaR4 = WorkbookLeafCellMap.HojaR4;

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
                ValidarFormulasProtegidasMultiAse(workbookPart, nameof(GenerarWorkbook));
                foreach (var leaf in leafInputs.OrderBy(l => l.Ase.Id))
                {
                    EscribirCeldasLeafPorAse(workbookPart, leaf);
                    EscribirCeldasEmpresa(workbookPart, leaf);
                }

                var workbookXml = workbookPart.Workbook
                    ?? throw new CalculoInvalidoException("El workbook abierto no tiene metadata Workbook válida.");
                workbookXml.Save();
            }

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, false))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException("El workbook generado no tiene WorkbookPart válido.");
                ValidarFormulasProtegidasMultiAse(workbookPart, nameof(GenerarWorkbook));
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
    /// <c>Recaudo *</c> fila 29+ (D6).
    /// </summary>
    private static void ValidarFormulasProtegidasMultiAse(WorkbookPart workbookPart, string operacion)
    {
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
    }

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
    /// </summary>
    private static void EscribirCeldasLeafPorAse(WorkbookPart workbookPart, WorkbookLeafInputs leaf)
    {
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