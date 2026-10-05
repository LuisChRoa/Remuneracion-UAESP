using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 21 (T3): motor de mutación de filas + reanclaje de fórmulas acotado a los bloques espejo
/// de la hoja <c>Reporte Componentes R1</c>.
///
/// Técnica de desplazamiento apilado elegida: <b>procesar los bloques de abajo hacia arriba
/// (ASE 5 → ASE 1)</b> (plan §2.3 punto 2). Así, al mutar el bloque del ASE-N, los bloques de los
/// ASE 1..N-1 están por encima del punto de mutación y no se ven afectados; no hace falta acumular
/// offsets. La otra alternativa (offsets acumulados) obligaría a recomputar la dirección física de
/// cada bloque en cada paso. Una sola técnica, sin mezclar.
///
/// Todas las referencias de la plantilla son A1 puras (T0c §4.3: 0 OFFSET/INDIRECT/rangos con
/// nombre en uso, 0 vínculos externos), por lo que el reanclaje es mecánico: se desplaza el número
/// de fila de toda referencia que apunte a una fila afectada, conservando columna y anclas
/// (<c>$</c>). Se reanclan también los rangos de celdas combinadas, los <c>ref</c> de fórmulas
/// compartidas y el único nombre definido que apunta a R1 (<c>PROMOAMBIENTAL_1</c>).
///
/// Regla de ausencia (D-B/R-E-5): la secuencia observada ES la especificación; una fila ausente se
/// suprime (el bloque se reescribe completo con la secuencia fuente). Los valores se copian por
/// ENCABEZADO de columna, nunca por índice fijo.
/// </summary>
internal static class OpenXmlEspejoR1Mutador
{
    public const string HojaR1 = "Reporte Componentes R1";

    /// <summary>
    /// Referencia A1 con prefijo de hoja opcional (citado o no), rango opcional y anclas ($).
    /// El lookbehind evita capturar sufijos dentro de identificadores; el lookahead evita capturar
    /// llamadas a función (p. ej. LOG10(). Las fórmulas del workbook no tienen literales de texto.
    /// </summary>
    private static readonly Regex RefRegex = new(
        @"(?<![A-Za-z0-9_\.])(?:'(?<qs>(?:[^']|'')+)'|(?<us>[A-Za-z_][A-Za-z0-9_\. ]*))!(?<c1>\$?[A-Za-z]{1,3})(?<r1>\$?\d+)(?::(?<c2>\$?[A-Za-z]{1,3})(?<r2>\$?\d+))?|(?<![A-Za-z0-9_\.])(?<c1>\$?[A-Za-z]{1,3})(?<r1>\$?\d+)(?::(?<c2>\$?[A-Za-z]{1,3})(?<r2>\$?\d+))?(?![A-Za-z0-9_(])",
        RegexOptions.Compiled);

    public static EspejoR1MutacionResultado Ajustar(string rutaPlantillaOrigen, string rutaSalida, IReadOnlyList<BloqueEspejoAseInputs> bloques)
    {
        if (string.IsNullOrWhiteSpace(rutaPlantillaOrigen))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.Plantilla, "La ruta de plantilla espejo R1 es requerida.");
        }

        if (string.IsNullOrWhiteSpace(rutaSalida))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.Plantilla, "La ruta de salida del espejo R1 es requerida.");
        }

        if (bloques.Count == 0)
        {
            throw new CalculoInvalidoException(CodigoError.Plantilla, "El espejo R1 no recibió bloques de ASE; no hay nada que dimensionar.");
        }

        if (!File.Exists(rutaPlantillaOrigen))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.Plantilla, $"No se encontró la plantilla para el espejo R1: '{rutaPlantillaOrigen}'.");
        }

        if (string.Equals(Path.GetFullPath(rutaPlantillaOrigen), Path.GetFullPath(rutaSalida), StringComparison.OrdinalIgnoreCase))
        {
            throw new CalculoInvalidoException(CodigoError.Plantilla, "El espejo R1 no puede mutar la plantilla original in-place. Use una ruta de salida distinta.");
        }

        var directorioSalida = Path.GetDirectoryName(rutaSalida);
        if (!string.IsNullOrWhiteSpace(directorioSalida))
        {
            Directory.CreateDirectory(directorioSalida);
        }

        File.Copy(rutaPlantillaOrigen, rutaSalida, overwrite: true);

        using var workbook = SpreadsheetDocument.Open(rutaSalida, true);
        var workbookPart = workbook.WorkbookPart
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene WorkbookPart válido.");
        var resultado = AjustarEnWorkbook(workbookPart, bloques);

        // Plan 28 (Unidad S): el espejo mueve filas y deja la calcChain de la plantilla
        // inconsistente; se sanea en el punto de guardado del espejo standalone (mismo helper
        // que el writer multi-ASE). No toca <f> ni <v>.
        SaneadorCadenaCalculo.Sanear(workbookPart);
        workbookPart.Workbook?.Save();
        return resultado;
    }

    /// <summary>
    /// Plan 21 (T4): aplica el espejo sobre un <see cref="WorkbookPart"/> ya abierto en modo
    /// escritura (usado por <c>GenerarWorkbook</c> en la misma pasada atómica). Procesa los bloques
    /// 5→1 y persiste la hoja mutada.
    /// </summary>
    public static EspejoR1MutacionResultado AjustarEnWorkbook(WorkbookPart workbookPart, IReadOnlyList<BloqueEspejoAseInputs> bloques)
    {
        var hoja = ObtenerHoja(workbookPart, HojaR1);
        var sheetData = hoja.Elements<SheetData>().FirstOrDefault()
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{HojaR1}' no tiene SheetData para el espejo.");

        var deltasPorBloque = new Dictionary<int, int>();
        var referenciasReancladasPorBloque = new Dictionary<int, int>();

        // Técnica 5→1: de abajo hacia arriba, sin mezclar con offsets acumulados.
        foreach (var bloque in bloques.OrderByDescending(b => b.Ase.Id))
        {
            var (delta, reancladas) = AjustarBloque(workbookPart, hoja, sheetData, bloque);
            deltasPorBloque[bloque.Ase.Id] = delta;
            referenciasReancladasPorBloque[bloque.Ase.Id] = reancladas;
        }

        ActualizarDimension(hoja, sheetData);
        hoja.Save();

        return new EspejoR1MutacionResultado(deltasPorBloque, referenciasReancladasPorBloque);
    }

    /// <summary>
    /// Plan 21 (T4): verdad si algún bloque destino difiere de la forma de la fuente del período
    /// actual (⇒ requiere inserción/borrado de filas y reanclaje). <c>false</c> = el bloque ya está
    /// dimensionado al período (Δ=0 para todos los ASE).
    /// </summary>
    public static bool RequiereAjuste(WorkbookPart workbookPart, IReadOnlyList<BloqueEspejoAseInputs> bloques)
    {
        foreach (var bloque in bloques)
        {
            var (_, _, dataRows) = LocalizarBloque(workbookPart, HojaR1, bloque);
            if (dataRows.Count != bloque.Filas.Count)
            {
                return true;
            }
        }

        return false;
    }

    private static (int Delta, int Reancladas) AjustarBloque(WorkbookPart workbookPart, Worksheet hoja, SheetData sheetData, BloqueEspejoAseInputs bloque)
    {
        var aseId = bloque.Ase.Id;
        var (nameRowIdx, totalRowIdx, dataRows) = LocalizarBloque(workbookPart, HojaR1, bloque);

        var headerRowIdx = (int)(dataRows[0].RowIndex?.Value ?? 0) - 1;
        var headerMap = MapearEncabezadosDestino(workbookPart, sheetData, headerRowIdx, bloque.Encabezados);

        var filasFuente = bloque.Filas;
        if (filasFuente.Count == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"Espejo R1: la fuente del ASE {aseId} no trae filas; no se puede dimensionar el bloque.");
        }

        var delta = filasFuente.Count - dataRows.Count;
        List<Row> dataRowsNuevas;
        var reancladas = 0;

        if (delta > 0)
        {
            var plantillaFila = dataRows.Count >= 2 ? dataRows[^2] : dataRows[0];
            var filaTotal = dataRows[^1];
            var ancla = totalRowIdx;

            DesplazarFilas(sheetData, ancla, delta);
            var nuevas = new List<Row>(delta);
            for (var j = 0; j < delta; j++)
            {
                var nueva = ClonarFilaVacia(plantillaFila, ancla + j);
                sheetData.InsertBefore(nueva, filaTotal);
                nuevas.Add(nueva);
            }

            dataRowsNuevas = new List<Row>(dataRows.Count + delta);
            dataRowsNuevas.AddRange(dataRows.Take(dataRows.Count - 1));
            dataRowsNuevas.AddRange(nuevas);
            dataRowsNuevas.Add(filaTotal);

            reancladas = Reanclar(workbookPart, HojaR1, r => r >= ancla ? r + delta : r);
        }
        else if (delta < 0)
        {
            var m = -delta;
            var filaTotal = dataRows[^1];
            var primeraBorrada = totalRowIdx - m;

            var aBorrar = dataRows.Skip(dataRows.Count - 1 - m).Take(m).ToList();
            foreach (var fila in aBorrar)
            {
                sheetData.RemoveChild(fila);
            }

            DesplazarFilas(sheetData, totalRowIdx, delta);

            dataRowsNuevas = new List<Row>(dataRows.Count - m);
            dataRowsNuevas.AddRange(dataRows.Take(dataRows.Count - 1 - m));
            dataRowsNuevas.Add(filaTotal);

            // Filas >= totalRow suben por m; las referencias a las filas suprimidas se reanclan a
            // la primera fila superviviente (nunca #REF!), preservando la validez estructural.
            //
            // W-3 (auditoría PR3, limitación consciente): con Δ<0 las referencias que apuntaban a
            // una fila suprimida [primeraBorrada, totalRowIdx) NO se redirigen a "su" dato (ya no
            // existe), sino a `primeraBorrada` = fila Total del bloque del período. Es un
            // compromiso deliberado (nunca #REF!) mientras la aritmética de negocio no se recalcula
            // (fuera de alcance: la plantilla conserva sus fórmulas y Excel recalcula al abrir).
            // Riesgo declarado para agosto en plans/21 - T0 Evidencia.md §8.
            reancladas = Reanclar(workbookPart, HojaR1, r => r >= totalRowIdx ? r - m : r >= primeraBorrada ? primeraBorrada : r);
        }
        else
        {
            dataRowsNuevas = dataRows;
        }

        EscribirValores(workbookPart, sheetData, dataRowsNuevas, bloque, headerMap, aseId);
        return (delta, reancladas);
    }

    /// <summary>
    /// Plan 21 (T3): localiza el bloque del ASE en la hoja (fila de nombre, fila Total final y
    /// filas de datos con firma A–E no vacía). Fail-fast con ASE + reporte + fila esperada.
    /// </summary>
    private static (int NameRowIdx, int TotalRowIdx, List<Row> DataRows) LocalizarBloque(
        WorkbookPart workbookPart,
        string nombreHoja,
        BloqueEspejoAseInputs bloque)
    {
        var aseId = bloque.Ase.Id;
        var nombreAse = Normalizar(bloque.Ase.NombreCorto);
        var sheetData = ObtenerSheetData(workbookPart, nombreHoja);

        var nameRowIdx = -1;
        foreach (var fila in sheetData.Elements<Row>().OrderBy(r => r.RowIndex?.Value ?? 0))
        {
            if (Normalizar(TextoDeCelda(workbookPart, fila, "B")).Equals(nombreAse, StringComparison.Ordinal))
            {
                nameRowIdx = (int)(fila.RowIndex?.Value ?? 0);
                break;
            }
        }

        if (nameRowIdx < 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"Espejo R1: no se encontró la fila de nombre del ASE {aseId} (B='{bloque.Ase.NombreCorto}') en la hoja '{nombreHoja}'.");
        }

        var totalRowIdx = -1;
        foreach (var fila in sheetData.Elements<Row>().OrderBy(r => r.RowIndex?.Value ?? 0))
        {
            var idx = (int)(fila.RowIndex?.Value ?? 0);
            if (idx <= nameRowIdx)
            {
                continue;
            }

            if (EsTotalFinal(workbookPart, fila))
            {
                totalRowIdx = idx;
                break;
            }
        }

        if (totalRowIdx < 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"Espejo R1: el bloque del ASE {aseId} no trae la fila invariante 'Total' (A='Total', B vacío) en la hoja '{nombreHoja}'.");
        }

        var dataRows = sheetData.Elements<Row>()
            .Where(r => (int)(r.RowIndex?.Value ?? 0) > nameRowIdx && (int)(r.RowIndex?.Value ?? 0) <= totalRowIdx)
            .OrderBy(r => r.RowIndex?.Value ?? 0)
            .Where(r => FirmaNoVacia(workbookPart, r))
            .ToList();

        if (dataRows.Count == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"Espejo R1: el bloque del ASE {aseId} no tiene filas de datos en la hoja '{nombreHoja}'.");
        }

        return (nameRowIdx, totalRowIdx, dataRows);
    }

    private static SheetData ObtenerSheetData(WorkbookPart workbookPart, string nombreHoja) =>
        ObtenerHoja(workbookPart, nombreHoja).Elements<SheetData>().FirstOrDefault()
        ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{nombreHoja}' no tiene SheetData.");

    /// <summary>
    /// Escribe la secuencia fuente por ENCABEZADO en las filas del bloque ya dimensionado. Las
    /// etiquetas A–E provienen de la firma de la fila (la secuencia observada ES la especificación).
    /// Guard anti-fórmula: una celda de valor que sea fórmula → ERR-PLANTILLA nombrando ASE+celda.
    /// </summary>
    private static void EscribirValores(
        WorkbookPart workbookPart,
        SheetData sheetData,
        IReadOnlyList<Row> dataRows,
        BloqueEspejoAseInputs bloque,
        IReadOnlyDictionary<string, string> headerMap,
        int aseId)
    {
        if (dataRows.Count != bloque.Filas.Count)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"Espejo R1: el bloque del ASE {aseId} quedó con {dataRows.Count} filas y la fuente trae {bloque.Filas.Count}.");
        }

        for (var i = 0; i < dataRows.Count; i++)
        {
            var fila = dataRows[i];
            var fuente = bloque.Filas[i];

            EscribirTexto(fila, "A", fuente.A, workbookPart, sheetData);
            EscribirTexto(fila, "B", fuente.B, workbookPart, sheetData);
            EscribirTexto(fila, "C", fuente.C, workbookPart, sheetData);
            EscribirTexto(fila, "D", fuente.D, workbookPart, sheetData);
            EscribirTexto(fila, "E", fuente.E, workbookPart, sheetData);

            foreach (var (encabezado, valor) in fuente.ValoresPorColumna)
            {
                if (!headerMap.TryGetValue(encabezado, out var columna))
                {
                    // La plantilla no trae esa columna (p. ej. ASE4 sin "Especiales"): se omite.
                    continue;
                }

                EscribirNumeroONulo(fila, columna, valor, workbookPart, sheetData, aseId, encabezado);
            }
        }
    }

    private static void EscribirNumeroONulo(
        Row fila,
        string columna,
        decimal? valor,
        WorkbookPart workbookPart,
        SheetData sheetData,
        int aseId,
        string encabezado)
    {
        var indice = (uint?)fila.RowIndex?.Value ?? 0;
        var referencia = columna + indice.ToString(CultureInfo.InvariantCulture);
        var celda = ObtenerOCrearCelda(fila, referencia, sheetData);

        if (celda.CellFormula is not null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"Espejo R1: la celda de valor {HojaR1}!{referencia} (ASE {aseId}, encabezado '{encabezado}') es fórmula en la plantilla; el espejo no sobrescribe fórmulas.");
        }

        if (valor is null)
        {
            celda.CellValue?.Remove();
            celda.InlineString?.Remove();
            celda.DataType = null;
            return;
        }

        celda.DataType = CellValues.Number;
        celda.InlineString?.Remove();
        celda.CellValue = new CellValue(valor.Value.ToString(CultureInfo.InvariantCulture));
    }

    private static void EscribirTexto(Row fila, string columna, string texto, WorkbookPart workbookPart, SheetData sheetData)
    {
        var indice = (uint?)fila.RowIndex?.Value ?? 0;
        var referencia = columna + indice.ToString(CultureInfo.InvariantCulture);
        var celda = ObtenerOCrearCelda(fila, referencia, sheetData);

        if (celda.CellFormula is not null)
        {
            // Las etiquetas A–E del bloque espejo son texto; una fórmula aquí es plantilla incompatible.
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"Espejo R1: la celda de etiqueta {HojaR1}!{referencia} es fórmula en la plantilla; el espejo no sobrescribe fórmulas.");
        }

        if (string.IsNullOrEmpty(texto))
        {
            celda.CellValue?.Remove();
            celda.InlineString?.Remove();
            celda.DataType = null;
            return;
        }

        celda.DataType = CellValues.InlineString;
        celda.CellValue?.Remove();
        celda.InlineString = new InlineString(new Text(texto));
    }

    // ── Geometría / filas ────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Plan 21 (T3): mapea cada encabezado de la FUENTE a su columna destino en el bloque.
    ///
    /// El template del bloque puede traer, dentro de su corrida inicial, los encabezados núcleo
    /// (F..T) y luego una COLA con columnas derivadas ("Exonera…") y encabezados DUPLICADOS
    /// (p. ej. ASE2/ASE4: "Intereses Cte Financiac. Aseo" en T y Y; "Rec.Intereses" en X y Z).
    /// Regla determinista que reproduce el golden congelado:
    ///   (1) NÚCLEO: corrida inicial de encabezados únicos y no derivados. Cada encabezado fuente
    ///       con match en el núcleo usa esa columna (incluido "Intereses Cte"→T, NO su duplicado Y);
    ///       la fuente puede SALTAR columnas núcleo que no trae (p. ej. ASE4-Q1 sin "Ajuste
    ///       Recargo Aseo").
    ///   (2) COLA: los encabezados fuente sin match en el núcleo (p. ej. "Rec.Intereses") se mapean
    ///       a su ÚLTIMA ocurrencia en la fila destino (Z), nunca al duplicado intermedio (X).
    /// </summary>
    private static IReadOnlyDictionary<string, string> MapearEncabezadosDestino(
        WorkbookPart workbookPart,
        SheetData sheetData,
        int headerRowIdx,
        IReadOnlyList<string> encabezadosFuente)
    {
        var mapa = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var fila = sheetData.Elements<Row>().FirstOrDefault(r => (int)(r.RowIndex?.Value ?? 0) == headerRowIdx);
        if (fila is null)
        {
            return mapa;
        }

        var destino = new List<(string Columna, string Encabezado)>();
        foreach (var celda in fila.Elements<Cell>())
        {
            var referencia = celda.CellReference?.Value;
            if (string.IsNullOrWhiteSpace(referencia))
            {
                continue;
            }

            var texto = LeerTexto(workbookPart, celda).Trim();
            if (texto.Length == 0)
            {
                continue;
            }

            var (columna, _) = ParsearReferencia(referencia);
            destino.Add((columna, texto));
        }

        // (1) Núcleo: corrida inicial de encabezados únicos y no derivados ("Exonera…").
        var vistos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var finNucleo = destino.Count;
        for (var i = 0; i < destino.Count; i++)
        {
            var encabezado = destino[i].Encabezado;
            if (EsEncabezadoDerivado(encabezado) || !vistos.Add(encabezado))
            {
                finNucleo = i;
                break;
            }
        }

        var nucleo = destino.Take(finNucleo).ToList();
        var ultimaPorEncabezado = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (columna, encabezado) in destino)
        {
            ultimaPorEncabezado[encabezado] = columna;
        }

        foreach (var encabezadoFuente in encabezadosFuente)
        {
            var enNucleo = nucleo.FirstOrDefault(d =>
                string.Equals(d.Encabezado, encabezadoFuente, StringComparison.OrdinalIgnoreCase));
            if (enNucleo.Columna is not null)
            {
                mapa[encabezadoFuente] = enNucleo.Columna;
            }
            else if (ultimaPorEncabezado.TryGetValue(encabezadoFuente, out var ultima))
            {
                mapa[encabezadoFuente] = ultima;
            }
        }

        return mapa;
    }

    /// <summary>
    /// Plan 21 (T3): verdad si el encabezado es una columna derivada de la cola
    /// (empieza con "Exonera", sin acentos ni mayúsculas), que la fuente R1 no alimenta.
    /// </summary>
    private static bool EsEncabezadoDerivado(string encabezado) =>
        Normalizar(encabezado).StartsWith("EXONERA", StringComparison.Ordinal);

    private static void DesplazarFilas(SheetData sheetData, int desde, int delta)
    {
        foreach (var fila in sheetData.Elements<Row>())
        {
            var indice = (int)(fila.RowIndex?.Value ?? 0);
            if (indice < desde)
            {
                continue;
            }

            var nuevo = indice + delta;
            fila.RowIndex = (uint)nuevo;
            foreach (var celda in fila.Elements<Cell>())
            {
                var referencia = celda.CellReference?.Value;
                if (string.IsNullOrWhiteSpace(referencia))
                {
                    continue;
                }

                var (columna, _) = ParsearReferencia(referencia);
                celda.CellReference = columna + nuevo.ToString(CultureInfo.InvariantCulture);
            }
        }
    }

    private static Row ClonarFilaVacia(Row plantilla, int nuevoIndice)
    {
        var nueva = (Row)plantilla.CloneNode(true);
        nueva.RowIndex = (uint)nuevoIndice;
        nueva.RemoveAllChildren<Cell>();
        foreach (var celdaOriginal in plantilla.Elements<Cell>())
        {
            var referencia = celdaOriginal.CellReference?.Value;
            if (string.IsNullOrWhiteSpace(referencia))
            {
                continue;
            }

            var (columna, _) = ParsearReferencia(referencia);
            var celda = new Cell
            {
                CellReference = columna + nuevoIndice.ToString(CultureInfo.InvariantCulture),
                StyleIndex = celdaOriginal.StyleIndex?.Value
            };
            nueva.Append(celda);
        }

        return nueva;
    }

    private static void ActualizarDimension(Worksheet hoja, SheetData sheetData)
    {
        var maxFila = sheetData.Elements<Row>().Select(r => (int)(r.RowIndex?.Value ?? 0)).DefaultIfEmpty(1).Max();
        var dimension = hoja.SheetDimension;
        var referencia = dimension?.Reference?.Value;
        var columnaFinal = "AP";
        if (!string.IsNullOrWhiteSpace(referencia) && referencia.Contains(':'))
        {
            var fin = referencia.Split(':').LastOrDefault();
            if (!string.IsNullOrWhiteSpace(fin))
            {
                var letras = new string(fin.Where(char.IsLetter).ToArray());
                if (letras.Length > 0)
                {
                    columnaFinal = letras;
                }
            }
        }

        if (dimension is not null)
        {
            dimension.Reference = $"A1:{columnaFinal}{maxFila}";
        }
    }

    // ── Reanclaje de fórmulas / nombres / merges ─────────────────────────────────────────────

    private static int Reanclar(WorkbookPart workbookPart, string hojaMutada, Func<int, int> mapRow)
    {
        var reancladas = 0;

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var nombreHoja = NombreDeHoja(workbookPart, worksheetPart);
            var worksheet = worksheetPart.Worksheet;
            if (worksheet is null)
            {
                continue;
            }

            foreach (var celda in worksheet.Descendants<Cell>())
            {
                var formula = celda.CellFormula;
                if (formula is null)
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(formula.Text))
                {
                    var (reescrita, cambiadas) = ReescribirReferencias(formula.Text, nombreHoja, hojaMutada, mapRow);
                    if (cambiadas > 0)
                    {
                        formula.Text = reescrita;
                        reancladas += cambiadas;
                    }
                }

                var rangoCompartido = formula.Reference?.Value;
                if (!string.IsNullOrWhiteSpace(rangoCompartido)
                    && string.Equals(nombreHoja, hojaMutada, StringComparison.OrdinalIgnoreCase))
                {
                    formula.Reference = ReescribirRango(rangoCompartido, mapRow);
                }
            }

            if (string.Equals(nombreHoja, hojaMutada, StringComparison.OrdinalIgnoreCase))
            {
                var mergeCells = worksheet.Elements<MergeCells>().FirstOrDefault();
                if (mergeCells is not null)
                {
                    foreach (var merge in mergeCells.Elements<MergeCell>())
                    {
                        var referencia = merge.Reference?.Value;
                        if (!string.IsNullOrWhiteSpace(referencia))
                        {
                            merge.Reference = ReescribirRango(referencia, mapRow);
                        }
                    }
                }
            }
        }

        var definedNames = workbookPart.Workbook?.DefinedNames;
        if (definedNames is not null)
        {
            foreach (var definedName in definedNames.Elements<DefinedName>())
            {
                if (!string.IsNullOrEmpty(definedName.Text))
                {
                    var (reescrito, cambiadas) = ReescribirReferencias(definedName.Text, null, hojaMutada, mapRow);
                    if (cambiadas > 0)
                    {
                        definedName.Text = reescrito;
                        reancladas += cambiadas;
                    }
                }
            }
        }

        return reancladas;
    }

    /// <summary>
    /// Reescribe las referencias A1 de <paramref name="texto"/> y devuelve cuántos tokens de fila
    /// cambiaron de valor (instrumentación W-2: evidencia de reanclaje real, no de mera reescritura).
    /// </summary>
    private static (string Texto, int Reancladas) ReescribirReferencias(string texto, string? hojaContexto, string hojaMutada, Func<int, int> mapRow)
    {
        var reancladas = 0;
        var resultado = RefRegex.Replace(texto, match =>
        {
            var citada = match.Groups["qs"];
            var sinCitar = match.Groups["us"];
            string? hojaRef = null;
            if (citada.Success)
            {
                hojaRef = citada.Value.Replace("''", "'", StringComparison.Ordinal);
            }
            else if (sinCitar.Success)
            {
                hojaRef = sinCitar.Value;
            }

            var aplica = hojaRef is null
                ? string.Equals(hojaContexto, hojaMutada, StringComparison.OrdinalIgnoreCase)
                : string.Equals(hojaRef, hojaMutada, StringComparison.OrdinalIgnoreCase);

            if (!aplica)
            {
                return match.Value;
            }

            var tokenFila1 = match.Groups["r1"].Value;
            var fila1 = RemapearFila(tokenFila1, mapRow);
            if (FilaDeToken(fila1) != FilaDeToken(tokenFila1))
            {
                reancladas++;
            }

            var constructor = new StringBuilder();

            if (citada.Success)
            {
                constructor.Append('\'').Append(citada.Value).Append("'!");
            }
            else if (sinCitar.Success)
            {
                constructor.Append(sinCitar.Value).Append('!');
            }

            constructor.Append(match.Groups["c1"].Value).Append(fila1);
            if (match.Groups["r2"].Success)
            {
                var tokenFila2 = match.Groups["r2"].Value;
                var fila2 = RemapearFila(tokenFila2, mapRow);
                if (FilaDeToken(fila2) != FilaDeToken(tokenFila2))
                {
                    reancladas++;
                }

                constructor.Append(':').Append(match.Groups["c2"].Value).Append(fila2);
            }

            return constructor.ToString();
        });

        return (resultado, reancladas);
    }

    private static int FilaDeToken(string token) =>
        int.Parse(token.StartsWith('$') ? token[1..] : token, CultureInfo.InvariantCulture);

    private static string ReescribirRango(string rango, Func<int, int> mapRow)
    {
        if (string.IsNullOrWhiteSpace(rango) || !rango.Contains(':'))
        {
            return rango;
        }

        var partes = rango.Split(':');
        if (partes.Length != 2)
        {
            return rango;
        }

        var (col1, fila1) = ParsearReferencia(partes[0]);
        var (col2, fila2) = ParsearReferencia(partes[1]);
        var nueva1 = mapRow(fila1);
        var nueva2 = mapRow(fila2);
        if (nueva1 > nueva2)
        {
            (nueva1, nueva2) = (nueva2, nueva1);
        }

        return $"{col1}{nueva1}:{col2}{nueva2}";
    }

    private static string RemapearFila(string token, Func<int, int> mapRow)
    {
        var ancla = token.StartsWith('$');
        var numero = int.Parse(ancla ? token[1..] : token, CultureInfo.InvariantCulture);
        var nuevo = mapRow(numero);
        return ancla ? "$" + nuevo.ToString(CultureInfo.InvariantCulture) : nuevo.ToString(CultureInfo.InvariantCulture);
    }

    // ── Helpers de celda ──────────────────────────────────────────────────────────────────────

    private static Cell ObtenerOCrearCelda(Row fila, string referencia, SheetData sheetData)
    {
        var existente = fila.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        if (existente is not null)
        {
            return existente;
        }

        var nueva = new Cell { CellReference = referencia };
        var siguiente = fila.Elements<Cell>().FirstOrDefault(c =>
            string.Compare(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase) > 0);
        if (siguiente is null)
        {
            fila.Append(nueva);
        }
        else
        {
            fila.InsertBefore(nueva, siguiente);
        }

        return nueva;
    }

    private static bool FirmaNoVacia(WorkbookPart workbookPart, Row fila)
    {
        foreach (var columna in new[] { "A", "B", "C", "D", "E" })
        {
            if (!string.IsNullOrWhiteSpace(TextoDeCelda(workbookPart, fila, columna)))
            {
                return true;
            }
        }

        return false;
    }

    private static bool EsTotalFinal(WorkbookPart workbookPart, Row fila) =>
        string.Equals(TextoDeCelda(workbookPart, fila, "A"), "Total", StringComparison.OrdinalIgnoreCase)
        && string.IsNullOrWhiteSpace(TextoDeCelda(workbookPart, fila, "B"));

    private static string TextoDeCelda(WorkbookPart workbookPart, Row fila, string columna)
    {
        var indice = ((uint?)fila.RowIndex?.Value ?? 0).ToString(CultureInfo.InvariantCulture);
        var celda = fila.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, columna + indice, StringComparison.OrdinalIgnoreCase));
        return celda is null ? string.Empty : LeerTexto(workbookPart, celda);
    }

    private static string LeerTexto(WorkbookPart workbookPart, Cell celda)
    {
        if (celda.CellFormula is not null)
        {
            return string.Empty;
        }

        if (celda.InlineString is not null)
        {
            return celda.InlineString.InnerText ?? string.Empty;
        }

        if (celda.DataType is not null && celda.DataType.Value == CellValues.SharedString && celda.CellValue is not null)
        {
            if (int.TryParse(celda.CellValue.Text, out var indice))
            {
                var tabla = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (tabla is not null && indice >= 0 && indice < tabla.Count())
                {
                    return tabla.ElementAt(indice).InnerText ?? string.Empty;
                }
            }

            return string.Empty;
        }

        return celda.CellValue?.InnerText ?? string.Empty;
    }

    private static Worksheet ObtenerHoja(WorkbookPart workbookPart, string nombreHoja)
    {
        var workbook = workbookPart.Workbook
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook no tiene metadata Workbook válida.");
        var sheet = workbook.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{nombreHoja}' no existe en el workbook del espejo.");
        var worksheetPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"No se pudo resolver la hoja '{nombreHoja}' del espejo.");
        return worksheetPart.Worksheet
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{nombreHoja}' no tiene Worksheet válido.");
    }

    private static string NombreDeHoja(WorkbookPart workbookPart, WorksheetPart worksheetPart)
    {
        var sheet = workbookPart.Workbook?.Descendants<Sheet>()
            .FirstOrDefault(s => s.Id?.Value == workbookPart.GetIdOfPart(worksheetPart));
        return sheet?.Name?.Value ?? string.Empty;
    }

    private static (string Columna, int Fila) ParsearReferencia(string referencia)
    {
        var letras = new string(referencia.Where(char.IsLetter).ToArray());
        var digitos = new string(referencia.Where(char.IsAsciiDigit).ToArray());
        if (letras.Length == 0 || digitos.Length == 0)
        {
            throw new CalculoInvalidoException(CodigoError.Plantilla, $"Referencia de celda inválida en el espejo R1: '{referencia}'.");
        }

        return (letras.ToUpperInvariant(), int.Parse(digitos, CultureInfo.InvariantCulture));
    }

    private static string Normalizar(string texto)
    {
        if (string.IsNullOrWhiteSpace(texto))
        {
            return string.Empty;
        }

        var descompuesto = texto.Normalize(NormalizationForm.FormD);
        var constructor = new StringBuilder(descompuesto.Length);
        foreach (var caracter in descompuesto)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(caracter) != UnicodeCategory.NonSpacingMark)
            {
                constructor.Append(char.ToUpperInvariant(caracter));
            }
        }

        return constructor.ToString();
    }
}
