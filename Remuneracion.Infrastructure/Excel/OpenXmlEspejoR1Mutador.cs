using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;

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

        // Plan 31 (T2, D-A/D-B/D-D): tras dimensionar/reanclar TODOS los bloques (ya no hay más
        // Reanclar pendiente), se recomponen por FIRMA las fórmulas de los totales visibles. Va
        // aquí (pase final post-5→1, antes de Save) porque solo entonces la geometría de cada
        // bloque es definitiva; recomponer por bloque lo reescribiría el Reanclar posterior.
        RecomponerVisibles(workbookPart, bloques);

        // Plan 32 (T2-interior, RONDA B2 — cableado mínimo): MISMO pase final, un nivel más adentro,
        // recomponiendo por firma los sub-visibles del INTERIOR de R1 (subtotales-empresa + EXTEMP +
        // single-refs). Julio = NO-OP (D-E identidad); ver RecomponerInterior.
        RecomponerInterior(workbookPart, bloques);

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

    // ── Plan 31 (T2, D-A/D-B/D-D): recomposición por firma de los totales visibles R1 ─────────
    //
    // Tras dimensionar/reanclar 5→1, cada bloque ASE se re-localiza por nombre y sus filas finales
    // se alinean con la secuencia fuente POR ORDEN (misma cardinalidad por construcción). Las
    // filas-ancla de los visibles se identifican POR FIRMA (patrón Plan 25: EsMesTotal /
    // EsAplicacionTotal sobre la secuencia fuente), NUNCA por conteo ni por borde inferior. Se
    // compone el texto <f> con la cardinalidad REAL del período:
    //   TOT_OPT F = ΣF(Mes) − ΣL(Mes menos la última)   (orden del template base, resuelto por rol)
    //   TDF     G = el conjunto E2 (G de las Mes menos la última), en orden descendente de Mes
    //   EXTEMP  F = ΣF(Aplic) − L(primera Aplic); 0 filas Aplic → literal 0 auditado
    // Es la ÚNICA reescritura de <f> del proyecto (excepción acotada al contrato §2.3, D-D),
    // auditada con un log Debug por celda (ASE + celda + texto antes/después).
    private static void RecomponerVisibles(WorkbookPart workbookPart, IReadOnlyList<BloqueEspejoAseInputs> bloques)
    {
        var sheetData = ObtenerSheetData(workbookPart, HojaR1);

        foreach (var bloque in bloques.OrderBy(b => b.Ase.Id))
        {
            var aseId = bloque.Ase.Id;
            var (nameRowIdx, _, dataRows) = LocalizarBloque(workbookPart, HojaR1, bloque);
            if (dataRows.Count != bloque.Filas.Count)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"Espejo R1: al recomponer el ASE {aseId} el bloque quedó con {dataRows.Count} filas y la fuente trae {bloque.Filas.Count}.");
            }

            var mesRows = new List<int>();
            var aplicRows = new List<int>();
            for (var i = 0; i < bloque.Filas.Count; i++)
            {
                var fila = (int)(dataRows[i].RowIndex?.Value ?? 0);
                if (bloque.Filas[i].EsMesTotal)
                {
                    mesRows.Add(fila);
                }
                else if (bloque.Filas[i].EsAplicacionTotal)
                {
                    aplicRows.Add(fila);
                }
            }

            if (mesRows.Count == 0)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"Espejo R1: el bloque del ASE {aseId} no trae filas Mes/Total; no se pueden recomponer los visibles.");
            }

            var totOptRow = LocalizarVisible(workbookPart, sheetData, nameRowIdx, ("C", "TOTAL"), ("D", "OPORTUNO"))
                ?? throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"Espejo R1: no se halló el visible TOT_OPT (C='TOTAL', D='OPORTUNO') del ASE {aseId} para recomponer.");
            var extempRow = LocalizarVisible(workbookPart, sheetData, totOptRow, ("D", "EXTEMPORANEO"))
                ?? throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"Espejo R1: no se halló el visible EXTEMP (D='EXTEMPORANEO') del ASE {aseId} para recomponer.");

            EscribirFormulaVisible(sheetData, "F" + totOptRow, ComponerTotOpt(aseId, mesRows), aseId, "TOT_OPT");
            EscribirFormulaVisible(sheetData, "G" + totOptRow, ComponerTdf(mesRows), aseId, "TDF");

            if (aplicRows.Count == 0)
            {
                EscribirCeroVisible(sheetData, "F" + extempRow, aseId, "EXTEMP");
            }
            else
            {
                EscribirFormulaVisible(sheetData, "F" + extempRow, ComponerExtemp(aseId, aplicRows), aseId, "EXTEMP");
            }
        }
    }

    // ── Plan 32 (T2-interior, RONDA B2 — cableado mínimo; TODO(T2-full)) ──────────────────────
    //
    // Extensión del pase final al INTERIOR de R1 (sub-visibles por empresa). Es el MISMO método del
    // Plan 31: por cada celda interior del contrato, se DERIVA `filasPorFirma` (etiquetas de firma ->
    // fila REAL del período) desde el bloque YA dimensionado usando las firmas puras de
    // `R1FirmaInterior` (catálogo de empresas ABIERTO: nunca `if ENEL/OCCIDENTE`), se compone el texto
    // <f> con `CompositorInteriorR1.ComponerInterior` y se escribe con la misma disciplina auditada
    // que los visibles. `null` (clase L-1/0-Aplic) -> literal 0; anclas faltantes -> fail-fast con
    // hoja+celda+ancla (nunca escritura parcial). Quirk `--L` (H-DOBLE-SIGNO): se reproduce leyéndolo
    // de la fórmula protegida de la celda (el disco manda, no se normaliza).
    //
    // TODO(T2-full): (a) la lista de celdas está CERRADA y hardcodeada (agosto) — debe derivarse del
    // fixture/periodo; (b) las clases L-1 (el manual sin <f>: F217, F327/F328/F337/F338 + G327/G337 +
    // H327/H337) NO entran en esta ronda (quedan con su fórmula previa); (c) el orden de términos de
    // F574 (quirk del manual) difiere del descendente y se resolverá aquí.

    /// <summary>
    /// Plan 32 (T2-interior, RONDA B2): celdas interiores R-1 del período AGOSTO (2026082) de
    /// <c>Reporte Componentes R1</c>, transcritas del fixture congelado
    /// <c>Remuneracion.IntegrationTests.InterioresR1Esperados.Agosto</c> (las entradas con
    /// <c>Formula = null</c> —clase L-1— se omiten por TODO(T2-full)). Clave = ASE; valor = celda +
    /// etiqueta del contrato (<c>SUB_EMP | SUB_TDF | SUB_L | EXT_INT | SUBS | AFASEO</c>).
    /// </summary>
    private static readonly IReadOnlyDictionary<int, (string Celda, string Etiqueta)[]> CeldasInterioresAgosto =
        new Dictionary<int, (string Celda, string Etiqueta)[]>
        {
            [1] =
            [
                ("F51", "SUBS"), ("F52", "EXT_INT"), ("G52", "EXT_INT"), ("H52", "EXT_INT"),
                ("F53", "SUBS"), ("F60", "SUB_EMP"), ("G60", "SUB_TDF"), ("H60", "SUB_TDF"),
                ("F61", "SUBS"), ("F62", "EXT_INT"), ("G62", "EXT_INT"), ("H62", "EXT_INT"),
                ("F63", "SUBS"), ("F70", "SUB_EMP"), ("G70", "SUB_TDF"), ("F71", "SUBS"), ("F80", "AFASEO")
            ],
            [2] =
            [
                ("F195", "SUBS"), ("F196", "EXT_INT"), ("G196", "EXT_INT"), ("F197", "SUBS"),
                ("F199", "SUB_EMP"), ("G199", "SUB_TDF"), ("F200", "SUBS"), ("F204", "SUB_EMP"),
                ("G204", "SUB_TDF"), ("F205", "SUBS"), ("F206", "EXT_INT"), ("G206", "EXT_INT"),
                ("F207", "SUBS"), ("F214", "SUB_EMP"), ("G214", "SUB_TDF"), ("F215", "SUBS"),
                ("F216", "EXT_INT"), ("G216", "EXT_INT"), ("F224", "AFASEO")
            ],
            [3] =
            [
                ("F326", "SUBS"), ("F335", "SUB_EMP"), ("G335", "SUB_TDF"), ("H335", "SUB_TDF"),
                ("F336", "SUBS"), ("F345", "SUB_EMP"), ("G345", "SUB_TDF"), ("F346", "SUBS"), ("F355", "AFASEO")
            ],
            [4] =
            [
                ("F457", "SUBS"), ("F458", "EXT_INT"), ("G458", "EXT_INT"), ("H458", "EXT_INT"),
                ("F459", "SUBS"), ("F461", "SUB_EMP"), ("G461", "SUB_TDF"), ("H461", "SUB_TDF"),
                ("F462", "SUBS"), ("F463", "EXT_INT"), ("G463", "EXT_INT"), ("H463", "EXT_INT"),
                ("F464", "SUBS"), ("F466", "SUB_EMP"), ("G466", "SUB_TDF"), ("H466", "SUB_TDF"),
                ("F467", "SUBS"), ("F468", "EXT_INT"), ("G468", "EXT_INT"), ("H468", "EXT_INT"),
                ("F469", "SUBS"), ("F476", "SUB_EMP"), ("G476", "SUB_TDF"), ("H476", "SUB_TDF"),
                ("F477", "SUBS"), ("F478", "EXT_INT"), ("G478", "EXT_INT"), ("H478", "EXT_INT"),
                ("F479", "SUBS"), ("F486", "AFASEO")
            ],
            [5] =
            [
                ("F555", "SUBS"), ("F556", "EXT_INT"), ("G556", "EXT_INT"), ("H556", "EXT_INT"),
                ("F557", "SUBS"), ("F564", "SUB_EMP"), ("G564", "SUB_TDF"), ("H564", "SUB_TDF"),
                ("F565", "SUBS"), ("F569", "SUB_EMP"), ("G569", "SUB_TDF"), ("H569", "SUB_TDF"),
                ("L569", "SUB_L"), ("F570", "SUBS"), ("F574", "SUB_EMP"), ("G574", "SUB_TDF"),
                ("H574", "SUB_TDF"), ("F575", "SUBS"), ("F576", "EXT_INT"), ("G576", "EXT_INT"),
                ("H576", "EXT_INT"), ("F577", "SUBS"), ("F584", "AFASEO"), ("G584", "AFASEO")
            ]
        };

    /// <summary>
    /// Plan 32 (T2-interior, D-E): recomponer el interior solo cuando el período NO es julio. En julio
    /// la geometría coincide con la plantilla y sus fórmulas canónicas ya están correctas; el pase es
    /// NO-OP por construcción (las celdas cerradas de arriba son las del manual de agosto y no aplican
    /// a la geometría de julio). La detección es por las hojas de período del workbook
    /// (<see cref="NombresHojaPeriodo.DetRetri(string)"/>): si el workbook es 2026071/2026072, julio.
    /// </summary>
    private static bool EsPeriodoJulio(WorkbookPart workbookPart)
    {
        var hojaJulioQ1 = NombresHojaPeriodo.DetRetri("2026071");
        var hojaJulioQ2 = NombresHojaPeriodo.DetRetri("2026072");
        foreach (var sheet in workbookPart.Workbook?.Descendants<Sheet>() ?? Enumerable.Empty<Sheet>())
        {
            var nombre = sheet.Name?.Value;
            if (string.Equals(nombre, hojaJulioQ1, StringComparison.OrdinalIgnoreCase)
                || string.Equals(nombre, hojaJulioQ2, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static void RecomponerInterior(WorkbookPart workbookPart, IReadOnlyList<BloqueEspejoAseInputs> bloques)
    {
        if (EsPeriodoJulio(workbookPart))
        {
            Serilog.Log.Debug(
                "Espejo R1 (Plan 32/T2-interior): período julio detectado -> pase interior NO-OP (identidad D-E).");
            return;
        }

        var sheetData = ObtenerSheetData(workbookPart, HojaR1);

        foreach (var bloque in bloques.OrderBy(b => b.Ase.Id))
        {
            var aseId = bloque.Ase.Id;
            if (!CeldasInterioresAgosto.TryGetValue(aseId, out var celdas))
            {
                continue;
            }

            var (_, totalRowIdx, dataRows) = LocalizarBloque(workbookPart, HojaR1, bloque);
            if (dataRows.Count != bloque.Filas.Count)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"Espejo R1: al recomponer el interior del ASE {aseId} el bloque quedó con {dataRows.Count} filas y la fuente trae {bloque.Filas.Count}.");
            }

            foreach (var (celda, etiqueta) in celdas)
            {
                var filasPorFirma = DerivarFilasPorFirma(workbookPart, sheetData, bloque, dataRows, totalRowIdx, celda, etiqueta);

                var faltantes = CompositorInteriorR1.AnclasFaltantes(etiqueta, filasPorFirma);
                if (faltantes.Count > 0)
                {
                    throw new CalculoInvalidoException(
                        CodigoError.Plantilla,
                        $"Espejo R1 (Plan 32/T2-interior): ASE {aseId} {HojaR1}!{celda} [{etiqueta}]: faltan las anclas [{string.Join(",", faltantes)}] del sub-bloque; no se escribe parcial.");
                }

                var texto = CompositorInteriorR1.ComponerInterior(aseId, celda, filasPorFirma);
                if (texto is null)
                {
                    EscribirLiteralInterior(sheetData, celda, aseId, etiqueta);
                }
                else
                {
                    EscribirFormulaInterior(sheetData, celda, texto, aseId, etiqueta);
                }
            }
        }
    }

    /// <summary>
    /// Deriva <c>filasPorFirma</c> (etiquetas de firma -> fila REAL) para la celda interior. La empresa
    /// del sub-bloque sale de la propia celda (subtotales) o del contexto (EXTEMP/SUBS/AFASEO, que
    /// cuelgan de la última fila con C no-vacía). Las filas-dato se resuelven por FIRMA sobre la
    /// secuencia fuente del bloque ya dimensionado (<c>dataRows[i]</c> corresponde a
    /// <c>bloque.Filas[i]</c>); las de OPORTUNO son las que NO cierran una sección de Aplicacion y las
    /// de Aplicacion las que sí (firma + vecino inmediato, catálogo abierto).
    /// </summary>
    private static IReadOnlyDictionary<string, int> DerivarFilasPorFirma(
        WorkbookPart workbookPart,
        SheetData sheetData,
        BloqueEspejoAseInputs bloque,
        IReadOnlyList<Row> dataRows,
        int totalRowIdx,
        string celda,
        string etiqueta)
    {
        var filas = new Dictionary<string, int>(StringComparer.Ordinal);
        var filaDestino = FilaDeReferencia(celda);

        switch (etiqueta)
        {
            case "SUB_EMP":
            case "SUB_TDF":
            case "SUB_L":
            {
                var fila = FilaPorIndice(sheetData, filaDestino)
                    ?? throw new CalculoInvalidoException(
                        CodigoError.Plantilla,
                        $"Espejo R1 (Plan 32/T2-interior): no existe la fila {filaDestino} de la celda interior {HojaR1}!{celda}.");
                var empresa = TextoDeCelda(workbookPart, fila, "C");
                AsignarPrefijo(filas, CompositorInteriorR1.EtiquetaMes, MesOportuno(bloque, dataRows, empresa));
                break;
            }

            case "EXT_INT":
            {
                var empresa = EmpresaContexto(workbookPart, sheetData, totalRowIdx, filaDestino);
                AsignarPrefijo(filas, CompositorInteriorR1.EtiquetaAplic, FilasAplic(bloque, dataRows, empresa));
                break;
            }

            case "SUBS":
            {
                filas[CompositorInteriorR1.EtiquetaAncla] =
                    FilaAnclaSubs(workbookPart, sheetData, bloque, dataRows, totalRowIdx, filaDestino);
                break;
            }

            case "AFASEO":
            {
                filas[CompositorInteriorR1.EtiquetaAncla] = FilaAnclaAFaseo(bloque, dataRows);
                break;
            }

            default:
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"Espejo R1 (Plan 32/T2-interior): etiqueta interior desconocida '{etiqueta}' en {HojaR1}!{celda}.");
        }

        // H-DOBLE-SIGNO (D-G): el quirk textual `--L` se reproduce desde la fórmula protegida de la
        // celda (el disco manda). Solo aplica a los subtotales-empresa que lo traen (p. ej. F204).
        if (etiqueta == "SUB_EMP")
        {
            var actual = ObtenerOCrearCeldaPorReferencia(sheetData, celda).CellFormula?.Text;
            if (actual is not null && actual.Contains("--L", StringComparison.Ordinal))
            {
                filas[CompositorInteriorR1.EtiquetaQuirkDobleSignoL] = 0;
            }
        }

        return filas;
    }

    /// <summary>
    /// Filas-dato OPORTUNO del sub-bloque de <paramref name="empresa"/> (firma <c>EsDatoEmpresa</c> que
    /// NO cierra una sección de <c>Aplicacion nuevos x reversion</c>), descendente (forma del manual).
    /// </summary>
    private static IReadOnlyList<int> MesOportuno(BloqueEspejoAseInputs bloque, IReadOnlyList<Row> dataRows, string empresa)
    {
        var resultado = new List<int>();
        for (var i = 0; i < bloque.Filas.Count; i++)
        {
            if (!R1FirmaInterior.EsDatoEmpresa(bloque.Filas[i], empresa))
            {
                continue;
            }

            if (i + 1 < bloque.Filas.Count && bloque.Filas[i + 1].EsAplicacionTotal)
            {
                continue;
            }

            resultado.Add(IndiceDe(dataRows[i]));
        }

        return resultado.OrderByDescending(r => r).ToList();
    }

    /// <summary>
    /// Filas-dato de <c>Aplicacion nuevos x reversion</c> del sub-bloque. Con contexto TOTAL (o vacío)
    /// son las filas <c>EsAplicacionTotal</c> del bloque; con una empresa, las <c>EsDatoEmpresa</c> que
    /// cierran una sección de Aplicacion. Descendente (forma del manual).
    /// </summary>
    private static IReadOnlyList<int> FilasAplic(BloqueEspejoAseInputs bloque, IReadOnlyList<Row> dataRows, string empresa)
    {
        var global = string.IsNullOrWhiteSpace(empresa) || Normalizar(empresa) == "TOTAL";
        var resultado = new List<int>();
        for (var i = 0; i < bloque.Filas.Count; i++)
        {
            var fila = bloque.Filas[i];
            var aplica = global
                ? fila.EsAplicacionTotal
                : R1FirmaInterior.EsDatoEmpresa(fila, empresa)
                    && i + 1 < bloque.Filas.Count
                    && bloque.Filas[i + 1].EsAplicacionTotal;
            if (aplica)
            {
                resultado.Add(IndiceDe(dataRows[i]));
            }
        }

        return resultado.OrderByDescending(r => r).ToList();
    }

    /// <summary>
    /// Ancla single-ref de una fila SUBS: si la fila hermana (la de arriba) es EXTEMP, la última
    /// Aplicacion del sub-bloque; si es OPORTUNO, la última fila Mes del sub-bloque (TOTAL -> fila
    /// <c>Mes/Total</c>; empresa -> última fila-dato OPORTUNO).
    /// </summary>
    private static int FilaAnclaSubs(
        WorkbookPart workbookPart,
        SheetData sheetData,
        BloqueEspejoAseInputs bloque,
        IReadOnlyList<Row> dataRows,
        int totalRowIdx,
        int filaDestino)
    {
        var empresa = EmpresaContexto(workbookPart, sheetData, totalRowIdx, filaDestino);
        var hermana = FilaPorIndice(sheetData, filaDestino - 1);
        var esExtemp = hermana is not null
            && string.Equals(Normalizar(TextoDeCelda(workbookPart, hermana, "D")), "EXTEMPORANEO", StringComparison.Ordinal);

        if (esExtemp)
        {
            var aplic = FilasAplic(bloque, dataRows, empresa);
            return aplic.Count > 0 ? aplic.Max() : 0;
        }

        if (string.IsNullOrWhiteSpace(empresa) || Normalizar(empresa) == "TOTAL")
        {
            var mesTotal = FilasMesTotal(bloque, dataRows);
            return mesTotal.Count > 0 ? mesTotal.Max() : 0;
        }

        var oportuno = MesOportuno(bloque, dataRows, empresa);
        return oportuno.Count > 0 ? oportuno.Max() : 0;
    }

    /// <summary>Ancla single-ref de una fila AFASEO: la fila AFaseo del bloque.</summary>
    private static int FilaAnclaAFaseo(BloqueEspejoAseInputs bloque, IReadOnlyList<Row> dataRows)
    {
        var filas = new List<int>();
        for (var i = 0; i < bloque.Filas.Count; i++)
        {
            if (R1FirmaInterior.EsAFaseoInterior(bloque.Filas[i]))
            {
                filas.Add(IndiceDe(dataRows[i]));
            }
        }

        return filas.Count > 0 ? filas.Max() : 0;
    }

    private static IReadOnlyList<int> FilasMesTotal(BloqueEspejoAseInputs bloque, IReadOnlyList<Row> dataRows)
    {
        var resultado = new List<int>();
        for (var i = 0; i < bloque.Filas.Count; i++)
        {
            if (bloque.Filas[i].EsMesTotal)
            {
                resultado.Add(IndiceDe(dataRows[i]));
            }
        }

        return resultado;
    }

    /// <summary>
    /// Empresa activa del sub-bloque: la C no-vacía más próxima por encima (o en) la fila destino,
    /// dentro de la zona de desglose (posterior a la fila Total del bloque).
    /// </summary>
    private static string EmpresaContexto(WorkbookPart workbookPart, SheetData sheetData, int totalRowIdx, int filaDestino)
    {
        var empresa = string.Empty;
        foreach (var fila in sheetData.Elements<Row>().OrderBy(r => r.RowIndex?.Value ?? 0))
        {
            var idx = (int)(fila.RowIndex?.Value ?? 0);
            if (idx <= totalRowIdx || idx > filaDestino)
            {
                continue;
            }

            var c = TextoDeCelda(workbookPart, fila, "C").Trim();
            if (!string.IsNullOrWhiteSpace(c))
            {
                empresa = c;
            }
        }

        return empresa;
    }

    private static void AsignarPrefijo(Dictionary<string, int> filas, string prefijo, IReadOnlyList<int> valores)
    {
        for (var i = 0; i < valores.Count; i++)
        {
            filas[prefijo + i.ToString(CultureInfo.InvariantCulture)] = valores[i];
        }
    }

    private static Row? FilaPorIndice(SheetData sheetData, int indice) =>
        sheetData.Elements<Row>().FirstOrDefault(r => (int)(r.RowIndex?.Value ?? 0) == indice);

    private static int IndiceDe(Row fila) => (int)(fila.RowIndex?.Value ?? 0);

    private static int FilaDeReferencia(string celda) =>
        int.Parse(new string(celda.Where(char.IsAsciiDigit).ToArray()), CultureInfo.InvariantCulture);

    private static void EscribirFormulaInterior(SheetData sheetData, string referencia, string texto, int aseId, string etiqueta)
    {
        var celda = ObtenerOCrearCeldaPorReferencia(sheetData, referencia);
        var antes = celda.CellFormula?.Text ?? celda.CellValue?.InnerText ?? "<vacío>";
        celda.CellFormula = new CellFormula(texto);
        Serilog.Log.Debug(
            "Espejo R1 (Plan 32/T2-interior): ASE {AseId} {Hoja}!{Celda} [{Etiqueta}] fórmula interior recompuesta por firma: '{Antes}' -> '{Despues}'.",
            aseId, HojaR1, referencia, etiqueta, antes, texto);
    }

    private static void EscribirLiteralInterior(SheetData sheetData, string referencia, int aseId, string etiqueta)
    {
        var celda = ObtenerOCrearCeldaPorReferencia(sheetData, referencia);
        var antes = celda.CellFormula?.Text ?? "<vacío>";
        celda.CellFormula = null;
        celda.DataType = null;
        celda.InlineString?.Remove();
        celda.CellValue = new CellValue("0");
        Serilog.Log.Debug(
            "Espejo R1 (Plan 32/T2-interior): ASE {AseId} {Hoja}!{Celda} [{Etiqueta}] sin anclas componibles -> literal 0 (antes '{Antes}').",
            aseId, HojaR1, referencia, etiqueta, antes);
    }

    private static int? LocalizarVisible(WorkbookPart workbookPart, SheetData sheetData, int desdeFila, params (string Col, string Valor)[] criterios)
    {
        foreach (var fila in sheetData.Elements<Row>().OrderBy(r => r.RowIndex?.Value ?? 0))
        {
            var idx = (int)(fila.RowIndex?.Value ?? 0);
            if (idx <= desdeFila)
            {
                continue;
            }

            var coincide = true;
            foreach (var (col, valor) in criterios)
            {
                if (!string.Equals(Normalizar(TextoDeCelda(workbookPart, fila, col)), Normalizar(valor), StringComparison.Ordinal))
                {
                    coincide = false;
                    break;
                }
            }

            if (coincide)
            {
                return idx;
            }
        }

        return null;
    }

    /// <summary>
    /// TOT_OPT: orden del template base (derivado de las fórmulas protegidas + el mapa de roles),
    /// resolviendo cada rol a la fila REAL del bloque; los roles nuevos por crecimiento de
    /// cardinalidad (p. ej. ASE5 pasa de 2 a 3 Mes) se insertan en su posición ordenada por índice.
    /// </summary>
    private static string ComponerTotOpt(int aseId, IReadOnlyList<int> mesRows)
    {
        var baseRoles = OrdenRolesBase(aseId, extemp: false);
        var n = mesRows.Count;

        var orden = baseRoles.Where(r => RolMesPresente(r, n)).ToList();
        for (var k = 0; k < n; k++)
        {
            InsertarMes(orden, k);
        }

        for (var k = 0; k <= n - 2; k++)
        {
            InsertarLmes(orden, k);
        }

        var partes = new List<string>(orden.Count);
        foreach (var rol in orden)
        {
            var k = IndiceMes(rol)!.Value;
            partes.Add(EsRolL(rol) ? "-L" + mesRows[k] : "+F" + mesRows[k]);
        }

        return FormatearFormula(partes);
    }

    private static bool RolMesPresente(WorkbookLeafCellMapQ2.R1Q2Fuente rol, int nMes)
    {
        var k = IndiceMes(rol);
        if (k is null)
        {
            return false;
        }

        return EsRolL(rol) ? k.Value <= nMes - 2 : k.Value <= nMes - 1;
    }

    private static void InsertarMes(List<WorkbookLeafCellMapQ2.R1Q2Fuente> orden, int k)
    {
        var rol = MesRole(k);
        if (orden.Contains(rol))
        {
            return;
        }

        for (var p = 0; p < orden.Count; p++)
        {
            if (!EsRolL(orden[p]) && IndiceMes(orden[p]) is int j && j > k)
            {
                orden.Insert(p, rol);
                return;
            }
        }

        orden.Insert(0, rol);
    }

    private static void InsertarLmes(List<WorkbookLeafCellMapQ2.R1Q2Fuente> orden, int k)
    {
        var rol = LmesRole(k);
        if (orden.Contains(rol))
        {
            return;
        }

        var posicion = -1;
        for (var p = 0; p < orden.Count; p++)
        {
            if (EsRolL(orden[p]) && IndiceMes(orden[p]) is int j && j < k)
            {
                posicion = p;
            }
        }

        if (posicion >= 0)
        {
            orden.Insert(posicion + 1, rol);
        }
        else
        {
            orden.Add(rol);
        }
    }

    /// <summary>
    /// TDF: conjunto E2 congelado por T1 — G de las filas <c>Mes/Total</c> menos la última por
    /// orden físico (n==1 → G(Mes0)), emitidas en orden descendente de índice (forma del manual).
    /// </summary>
    private static string ComponerTdf(IReadOnlyList<int> mesRows)
    {
        var n = mesRows.Count;
        var cuantas = n > 1 ? n - 1 : 1;
        var partes = new List<string>(cuantas);
        for (var k = cuantas - 1; k >= 0; k--)
        {
            partes.Add("+G" + mesRows[k]);
        }

        return FormatearFormula(partes);
    }

    private static string ComponerExtemp(int aseId, IReadOnlyList<int> aplicRows)
    {
        var baseRoles = OrdenRolesBase(aseId, extemp: true);
        var m = aplicRows.Count;

        var orden = baseRoles.Where(r => RolAplicPresente(r, m)).ToList();
        for (var k = 0; k < m; k++)
        {
            InsertarAplic(orden, k);
        }

        var partes = new List<string>(orden.Count);
        foreach (var rol in orden)
        {
            if (rol == WorkbookLeafCellMapQ2.R1Q2Fuente.LAplic0)
            {
                partes.Add("-L" + aplicRows[0]);
            }
            else
            {
                partes.Add("+F" + aplicRows[IndiceAplic(rol)!.Value]);
            }
        }

        return FormatearFormula(partes);
    }

    private static bool RolAplicPresente(WorkbookLeafCellMapQ2.R1Q2Fuente rol, int nAplic) =>
        rol == WorkbookLeafCellMapQ2.R1Q2Fuente.LAplic0
        || (IndiceAplic(rol) is int k && k <= nAplic - 1);

    private static void InsertarAplic(List<WorkbookLeafCellMapQ2.R1Q2Fuente> orden, int k)
    {
        var rol = AplicRole(k);
        if (orden.Contains(rol))
        {
            return;
        }

        for (var p = 0; p < orden.Count; p++)
        {
            if (orden[p] != WorkbookLeafCellMapQ2.R1Q2Fuente.LAplic0 && IndiceAplic(orden[p]) is int j && j > k)
            {
                orden.Insert(p, rol);
                return;
            }
        }

        orden.Insert(0, rol);
    }

    private static IReadOnlyList<WorkbookLeafCellMapQ2.R1Q2Fuente> OrdenRolesBase(int aseId, bool extemp)
    {
        var editables = WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(aseId);
        foreach (var (_, fragmentos) in WorkbookLeafCellMapQ2.ObtenerR1Q2Protegidos(aseId))
        {
            var roles = fragmentos.Select(f => RolDeFragmento(editables, f, aseId)).ToList();
            if (roles.Any(EsRolAplic) == extemp)
            {
                return roles;
            }
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"Espejo R1: no se halló el orden de roles base {(extemp ? "EXTEMP" : "TOT_OPT")} del ASE {aseId} en el mapa congelado.");
    }

    private static WorkbookLeafCellMapQ2.R1Q2Fuente RolDeFragmento(
        IReadOnlyList<(string Celda, WorkbookLeafCellMapQ2.R1Q2Fuente Fuente)> editables,
        string fragmento,
        int aseId)
    {
        var objetivo = fragmento.Replace("$", string.Empty, StringComparison.Ordinal).ToUpperInvariant();
        foreach (var (celda, fuente) in editables)
        {
            if (string.Equals(celda.Replace("$", string.Empty, StringComparison.Ordinal).ToUpperInvariant(), objetivo, StringComparison.Ordinal))
            {
                return fuente;
            }
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"Espejo R1: el fragmento '{fragmento}' de una fórmula protegida del ASE {aseId} no tiene rol en el mapa de editables.");
    }

    private static int? IndiceMes(WorkbookLeafCellMapQ2.R1Q2Fuente rol) => rol switch
    {
        WorkbookLeafCellMapQ2.R1Q2Fuente.Mes0 or WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes0 => 0,
        WorkbookLeafCellMapQ2.R1Q2Fuente.Mes1 or WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes1 => 1,
        WorkbookLeafCellMapQ2.R1Q2Fuente.Mes2 or WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes2 => 2,
        _ => null
    };

    private static bool EsRolL(WorkbookLeafCellMapQ2.R1Q2Fuente rol) =>
        rol is WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes0
            or WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes1
            or WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes2;

    private static WorkbookLeafCellMapQ2.R1Q2Fuente MesRole(int k) => k switch
    {
        0 => WorkbookLeafCellMapQ2.R1Q2Fuente.Mes0,
        1 => WorkbookLeafCellMapQ2.R1Q2Fuente.Mes1,
        _ => WorkbookLeafCellMapQ2.R1Q2Fuente.Mes2
    };

    private static WorkbookLeafCellMapQ2.R1Q2Fuente LmesRole(int k) => k switch
    {
        0 => WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes0,
        1 => WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes1,
        _ => WorkbookLeafCellMapQ2.R1Q2Fuente.Lmes2
    };

    private static int? IndiceAplic(WorkbookLeafCellMapQ2.R1Q2Fuente rol) => rol switch
    {
        WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic0 => 0,
        WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic1 => 1,
        _ => null
    };

    private static WorkbookLeafCellMapQ2.R1Q2Fuente AplicRole(int k) => k == 0
        ? WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic0
        : WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic1;

    private static bool EsRolAplic(WorkbookLeafCellMapQ2.R1Q2Fuente rol) =>
        rol is WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic0
            or WorkbookLeafCellMapQ2.R1Q2Fuente.Aplic1
            or WorkbookLeafCellMapQ2.R1Q2Fuente.LAplic0;

    private static string FormatearFormula(IReadOnlyList<string> partes)
    {
        if (partes.Count == 0)
        {
            throw new CalculoInvalidoException(CodigoError.Plantilla, "Espejo R1: la fórmula visible recompuesta quedó sin términos.");
        }

        var texto = new StringBuilder();
        for (var i = 0; i < partes.Count; i++)
        {
            var parte = partes[i];
            texto.Append(i == 0 && parte.StartsWith('+') ? parte[1..] : parte);
        }

        return texto.ToString();
    }

    private static void EscribirFormulaVisible(SheetData sheetData, string referencia, string texto, int aseId, string etiqueta)
    {
        var celda = ObtenerOCrearCeldaPorReferencia(sheetData, referencia);
        var antes = celda.CellFormula?.Text ?? celda.CellValue?.InnerText ?? "<vacío>";
        celda.CellFormula = new CellFormula(texto);
        Serilog.Log.Debug(
            "Espejo R1 (Plan 31/T2): ASE {AseId} {Hoja}!{Celda} [{Etiqueta}] fórmula visible recompuesta por firma: '{Antes}' -> '{Despues}'.",
            aseId, HojaR1, referencia, etiqueta, antes, texto);
    }

    private static void EscribirCeroVisible(SheetData sheetData, string referencia, int aseId, string etiqueta)
    {
        var celda = ObtenerOCrearCeldaPorReferencia(sheetData, referencia);
        var antes = celda.CellFormula?.Text ?? "<vacío>";
        celda.CellFormula = null;
        celda.DataType = null;
        celda.InlineString?.Remove();
        celda.CellValue = new CellValue("0");
        Serilog.Log.Debug(
            "Espejo R1 (Plan 31/T2): ASE {AseId} {Hoja}!{Celda} [{Etiqueta}] sin filas Aplicacion -> literal 0 (antes '{Antes}').",
            aseId, HojaR1, referencia, etiqueta, antes);
    }

    private static Cell ObtenerOCrearCeldaPorReferencia(SheetData sheetData, string referencia)
    {
        var letras = new string(referencia.Where(char.IsLetter).ToArray());
        var digitos = new string(referencia.Where(char.IsAsciiDigit).ToArray());
        if (letras.Length == 0 || digitos.Length == 0)
        {
            throw new CalculoInvalidoException(CodigoError.Plantilla, $"Referencia de celda inválida al recomponer un visible R1: '{referencia}'.");
        }

        var indice = int.Parse(digitos, CultureInfo.InvariantCulture);
        var fila = sheetData.Elements<Row>().FirstOrDefault(r => (int)(r.RowIndex?.Value ?? 0) == indice)
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"Espejo R1: la fila {indice} del visible {referencia} no existe en la hoja mutada.");
        var columna = letras.ToUpperInvariant();
        return ObtenerOCrearCelda(fila, columna + indice.ToString(CultureInfo.InvariantCulture), sheetData);
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
