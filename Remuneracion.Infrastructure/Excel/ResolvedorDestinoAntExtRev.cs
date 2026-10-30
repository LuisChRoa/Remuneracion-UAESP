using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 64 (H1.4, §2.5 — D-05/REQ-06): RESOLUCIÓN POR ETIQUETAS del destino de las 100 celdas de valor
/// de la hoja <c>ANT EXT-REV</c>.
///
/// <para><b>El defecto que evita.</b> Congelar <c>C11..L12</c> (o cualquier letra/fila) reproduce la
/// fotografía de la plantilla v4: la próxima plantilla mueve una fila y el dinero cae en otra celda en
/// silencio (M-B del Plan 59, R-01). Aquí la identidad de un valor es
/// <c>(hoja · etiqueta de bloque · ASE · fila APLICADO/ANTICIPOS · reporte 3/4)</c> y la celda destino sale
/// de lo que la PROPIA hoja declara: la etiqueta del bloque en la columna B, los rótulos <c>ASE n</c> /
/// <c>TOTAL</c> en la fila de ASEs, los rótulos <c>REPORTE 3</c>/<c>REPORTE 4</c> en la fila de reportes y
/// los rótulos <c>APLICADO</c>/<c>ANTICIPOS</c>/<c>TOTAL</c> en las sub-filas del bloque.</para>
///
/// <para><b>Validación estructural plan-then-write (falla ANTES de escribir).</b> Si la gramática no se
/// cumple —bloque sin su rejilla, etiqueta duplicada, celda de valor con fórmula, fila TOTAL sin fórmula
/// SUM, banda superior con un valor fijo— se lanza <c>ERR-PLANTILLA</c> nombrando hoja + identidad
/// esperada + celda encontrada. La tabla T0 (<c>Docs/T0-Plan64-Ant-Ext-Rev.md</c>) es EVIDENCIA del
/// fixture, nunca contrato de producción: el resolvedor no contiene ninguna coordenada.</para>
///
/// <para><b>Quién lo consume.</b> El escritor values-only de H2 (F2) y los fixtures de H1/H2: una sola
/// autoridad del destino, de modo que una pérdida deja de ser invisible (mismo patrón que
/// <see cref="ResolvedorDestinoDetalleR2R4"/>).</para>
/// </summary>
internal static class ResolvedorDestinoAntExtRev
{
    /// <summary>
    /// Hoja objetivo, por NOMBRE (Plan 27 D-A). La misma que declaran los mapas históricos; se abre por
    /// nombre, nunca por posición del paquete.
    /// </summary>
    internal const string Hoja = "ANT EXT-REV";

    /// <summary>Rótulo de la sub-fila de valores con los aplicados.</summary>
    internal const string FilaAplicado = "APLICADO";

    /// <summary>Rótulo de la sub-fila de valores con los anticipos.</summary>
    internal const string FilaAnticipos = "ANTICIPOS";

    /// <summary>Rótulo de la sub-fila de cierre del bloque (TOTAL con fórmula SUM por columna).</summary>
    internal const string FilaTotal = "TOTAL";

    /// <summary>Rótulo de la columna de reporte 3.</summary>
    internal const string Reporte3 = "REPORTE 3";

    /// <summary>Rótulo de la columna de reporte 4.</summary>
    internal const string Reporte4 = "REPORTE 4";

    /// <summary>
    /// Destino de UNA celda de valor: la identidad completa (etiqueta de bloque + ASE + fila + reporte) con
    /// la letra de columna y la referencia A1 que la plantilla declara para ella.
    /// </summary>
    /// <param name="EtiquetaBloque">Etiqueta del bloque (dato leído de la columna B).</param>
    /// <param name="AseId">Identificador del ASE (1..5) derivado del rótulo <c>ASE n</c>.</param>
    /// <param name="Fila">Rótulo de la fila de valor (<c>APLICADO</c> / <c>ANTICIPOS</c>).</param>
    /// <param name="Reporte">Rótulo de la columna de reporte (<c>REPORTE 3</c> / <c>REPORTE 4</c>).</param>
    /// <param name="Columna">Letra(s) de la columna destino.</param>
    /// <param name="Celda">Referencia A1 de la celda destino.</param>
    internal sealed record Destino(
        string EtiquetaBloque,
        int AseId,
        string Fila,
        string Reporte,
        string Columna,
        string Celda);

    /// <summary>
    /// Resuelve las celdas de valor de la hoja <c>ANT EXT-REV</c> validando su estructura completa
    /// (plan-then-write). Devuelve una entrada por celda de valor: bloques declarados × ASEs declarados ×
    /// reportes declarados × filas de valor declaradas.
    /// </summary>
    /// <param name="workbookPart">WorkbookPart de la plantilla que se va a escribir.</param>
    /// <exception cref="CalculoInvalidoException"><c>ERR-PLANTILLA</c>: la hoja no existe, un bloque no
    /// declara su rejilla, una etiqueta es ambigua, una celda de valor trae fórmula, una fila TOTAL perdió
    /// su SUM o la banda superior dejó de ser fórmula.</exception>
    internal static IReadOnlyList<Destino> Resolver(WorkbookPart workbookPart)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);

        var hoja = HojaPorNombre(workbookPart);
        var filas = hoja.Elements<SheetData>()
            .SelectMany(tabla => tabla.Elements<Row>())
            .Where(fila => fila.RowIndex is not null)
            .ToDictionary(fila => fila.RowIndex!.Value, fila => fila);

        if (filas.Count == 0)
        {
            throw PlantillaInvalida($"la hoja '{Hoja}' no tiene filas: no hay rejilla que resolver");
        }

        var bloques = LocalizarBloques(workbookPart, filas);
        var destinos = new List<Destino>();

        foreach (var bloque in bloques)
        {
            destinos.AddRange(ResolverBloque(workbookPart, filas, bloque).Destinos);
        }

        // Las celdas de valor son únicas: una rejilla solapada haría que dos conceptos escribieran la misma
        // celda y el orden decidiera quién gana (prohibido: nunca "el que llegue último").
        var repetidas = destinos
            .GroupBy(destino => destino.Celda, StringComparer.OrdinalIgnoreCase)
            .Where(grupo => grupo.Count() > 1)
            .Select(grupo => $"{grupo.Key} ({string.Join(", ", grupo.Select(d => $"{d.EtiquetaBloque}/{d.AseId}/{d.Fila}/{d.Reporte}"))})")
            .ToArray();
        if (repetidas.Length > 0)
        {
            throw PlantillaInvalida(
                $"la rejilla de valores de '{Hoja}' resuelve la misma celda para más de una identidad: " +
                $"{string.Join("; ", repetidas)}. El destino sería ambiguo y no se elige por orden");
        }

        // Banda superior (controles y cuadro de totales): debe seguir siendo FÓRMULA. Un valor fijo ahí es
        // una plantilla manipulada y el gate de fórmulas protegidas de H2 ya no tendría qué proteger.
        ValidarBandaSuperior(workbookPart, filas, bloques[0].FilaEtiqueta, destinos);

        return destinos;
    }

    /// <summary>
    /// Celda destino de una identidad concreta; <c>null</c> = la plantilla no declara ese bloque/ASE/fila/
    /// reporte (el consumidor decide: fallo duro si trae dinero, omisión si no).
    /// </summary>
    /// <param name="destinos">Rejilla resuelta por <see cref="Resolver"/>.</param>
    /// <param name="etiquetaBloque">Etiqueta de bloque del catálogo.</param>
    /// <param name="aseId">Identificador del ASE.</param>
    /// <param name="fila">Rótulo de la fila de valor.</param>
    /// <param name="reporte">Rótulo de la columna de reporte.</param>
    internal static Destino? Buscar(
        IReadOnlyList<Destino> destinos,
        string etiquetaBloque,
        int aseId,
        string fila,
        string reporte)
    {
        ArgumentNullException.ThrowIfNull(destinos);

        var filaCatalogo = CatalogoFuentesAntExtRev.BuscarPorEtiquetaBloque(etiquetaBloque);
        foreach (var destino in destinos)
        {
            if (destino.AseId != aseId
                || !string.Equals(destino.Fila, fila, StringComparison.Ordinal)
                || !string.Equals(destino.Reporte, reporte, StringComparison.Ordinal))
            {
                continue;
            }

            var filaDestino = CatalogoFuentesAntExtRev.BuscarPorEtiquetaBloque(destino.EtiquetaBloque);
            if (filaCatalogo is not null && filaDestino is not null
                && CatalogoFuentesAntExtRev.EsMismaFuente(filaCatalogo, filaDestino))
            {
                return destino;
            }
        }

        return null;
    }

    /// <summary>
    /// Bloque declarado por la hoja: su etiqueta, la fila donde vive y las columnas que declara para cada
    /// ASE (par de reportes) más la columna del rótulo TOTAL.
    /// </summary>
    private sealed record Bloque(
        string Etiqueta,
        uint FilaEtiqueta,
        uint FilaAses,
        uint FilaReportes,
        IReadOnlyList<(int AseId, string ColumnaReporte3, string ColumnaReporte4)> Ases,
        string ColumnaTotalAses);

    /// <summary>Resultado de resolver un bloque: sus celdas de valor.</summary>
    private sealed record RejillaBloque(IReadOnlyList<Destino> Destinos);

    /// <summary>
    /// Localiza los bloques por la GRAMÁTICA que la hoja declara: una celda de la columna B es etiqueta de
    /// bloque si las cinco filas siguientes se declaran como ASEs + reportes + APLICADO + ANTICIPOS + TOTAL.
    /// Cero números de fila en el código: las filas salen de la referencia de la propia celda.
    /// </summary>
    private static IReadOnlyList<Bloque> LocalizarBloques(WorkbookPart workbookPart, IReadOnlyDictionary<uint, Row> filas)
    {
        var bloques = new List<Bloque>();
        foreach (var fila in filas.Values.OrderBy(f => f.RowIndex!.Value))
        {
            var etiqueta = TextoDeColumna(workbookPart, fila, ColumnaEtiquetaBloque).Trim();
            if (etiqueta.Length == 0
                || string.Equals(etiqueta, FilaAplicado, StringComparison.OrdinalIgnoreCase)
                || string.Equals(etiqueta, FilaAnticipos, StringComparison.OrdinalIgnoreCase)
                || string.Equals(etiqueta, FilaTotal, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var bloque = IntentarBloque(workbookPart, filas, fila.RowIndex!.Value, etiqueta);
            if (bloque is not null)
            {
                bloques.Add(bloque);
            }
        }

        if (bloques.Count == 0)
        {
            throw PlantillaInvalida(
                $"la hoja '{Hoja}' no declara ningún bloque con la gramática de ANT EXT-REV (etiqueta en la " +
                $"columna {ColumnaEtiquetaBloque} seguida de ASEs, reportes y las filas APLICADO/ANTICIPOS/TOTAL)");
        }

        var etiquetasRepetidas = bloques
            .GroupBy(bloque => Normalizar(bloque.Etiqueta), StringComparer.Ordinal)
            .Where(grupo => grupo.Count() > 1)
            .Select(grupo => grupo.Key)
            .ToArray();
        if (etiquetasRepetidas.Length > 0)
        {
            throw PlantillaInvalida(
                $"la hoja '{Hoja}' declara más de un bloque con la misma etiqueta: {string.Join(", ", etiquetasRepetidas)}");
        }

        foreach (var filaCatalogo in CatalogoFuentesAntExtRev.Filas)
        {
            var encontrados = bloques.Count(bloque =>
                CatalogoFuentesAntExtRev.BuscarPorEtiquetaBloque(bloque.Etiqueta) is { } filaDelBloque
                && CatalogoFuentesAntExtRev.EsMismaFuente(filaDelBloque, filaCatalogo));
            if (encontrados != 1)
            {
                throw PlantillaInvalida(
                    $"la hoja '{Hoja}' declara {encontrados} bloque(s) con la etiqueta esperada " +
                    $"'{filaCatalogo.EtiquetaBloque}' (se esperaba exactamente 1): el destino de esa fuente no se " +
                    "puede resolver sin adivinar la fila");
            }
        }

        return bloques;
    }

    /// <summary>
    /// Prueba la gramática del bloque cuya etiqueta vive en <paramref name="filaEtiqueta"/>; devuelve
    /// <c>null</c> si esa celda de la columna B no es una etiqueta de bloque (es otra cosa: un título o el
    /// rótulo de una sub-fila) y lanza si dice serlo pero declara una rejilla incompleta.
    /// </summary>
    private static Bloque? IntentarBloque(
        WorkbookPart workbookPart,
        IReadOnlyDictionary<uint, Row> filas,
        uint filaEtiqueta,
        string etiqueta)
    {
        var filaAses = filas.GetValueOrDefault(filaEtiqueta + 1);
        var filaReportes = filas.GetValueOrDefault(filaEtiqueta + 2);
        var filaAplicado = filas.GetValueOrDefault(filaEtiqueta + 3);
        var filaAnticipos = filas.GetValueOrDefault(filaEtiqueta + 4);
        var filaTotal = filas.GetValueOrDefault(filaEtiqueta + 5);
        if (filaAses is null || filaReportes is null || filaAplicado is null || filaAnticipos is null || filaTotal is null)
        {
            return null;
        }

        if (!EsRotulo(workbookPart, filaAplicado, ColumnaEtiquetaBloque, FilaAplicado)
            || !EsRotulo(workbookPart, filaAnticipos, ColumnaEtiquetaBloque, FilaAnticipos)
            || !EsRotulo(workbookPart, filaTotal, ColumnaEtiquetaBloque, FilaTotal))
        {
            return null;
        }

        var ases = new List<(int, string, string)>();
        string? columnaTotal = null;
        foreach (var celda in filaAses.Elements<Cell>())
        {
            var texto = LectorTextoHoja.Texto(workbookPart, celda).Trim();
            var columna = ColumnaDeReferencia(celda.CellReference?.Value);
            if (texto.Length == 0 || columna.Length == 0)
            {
                continue;
            }

            if (TryAseDeRotulo(texto, out var aseId))
            {
                var columnaSiguiente = ColumnaSiguiente(columna);
                var rotuloReporte3 = TextoDeColumna(workbookPart, filaReportes, columna).Trim();
                var rotuloReporte4 = TextoDeColumna(workbookPart, filaReportes, columnaSiguiente).Trim();
                if (!string.Equals(rotuloReporte3, Reporte3, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(rotuloReporte4, Reporte4, StringComparison.OrdinalIgnoreCase))
                {
                    throw PlantillaInvalida(
                        $"el bloque '{etiqueta}' de '{Hoja}' declara el ASE {aseId} en {columna} pero la fila de " +
                        $"reportes trae '{rotuloReporte3}'/'{rotuloReporte4}' en {columna}/{columnaSiguiente} " +
                        $"(se esperaba '{Reporte3}'/'{Reporte4}')");
                }

                ases.Add((aseId, columna, columnaSiguiente));
                continue;
            }

            if (string.Equals(texto, FilaTotal, StringComparison.OrdinalIgnoreCase))
            {
                columnaTotal = columna;
            }
        }

        if (ases.Count == 0)
        {
            return null;
        }

        if (columnaTotal is null)
        {
            throw PlantillaInvalida(
                $"el bloque '{etiqueta}' de '{Hoja}' no declara el rótulo '{FilaTotal}' en su fila de ASEs: no hay " +
                "columna de cierre del bloque");
        }

        var ids = ases.Select(ase => ase.Item1).ToArray();
        if (ids.Distinct().Count() != ids.Length)
        {
            throw PlantillaInvalida(
                $"el bloque '{etiqueta}' de '{Hoja}' declara el mismo ASE más de una vez: " +
                $"{string.Join(", ", ids)}. El destino de su celda sería ambiguo");
        }

        return new Bloque(etiqueta, filaEtiqueta, filaEtiqueta + 1, filaEtiqueta + 2, ases, columnaTotal);
    }

    /// <summary>
    /// Valida la rejilla de UN bloque y devuelve sus celdas de valor: por cada ASE y reporte declarados, la
    /// celda de <c>APLICADO</c> y la de <c>ANTICIPOS</c> existen, son numéricas y NO son fórmula; la celda
    /// homóloga de la fila <c>TOTAL</c> es fórmula (el cierre del bloque).
    /// </summary>
    private static RejillaBloque ResolverBloque(
        WorkbookPart workbookPart,
        IReadOnlyDictionary<uint, Row> filas,
        Bloque bloque)
    {
        var destinos = new List<Destino>();
        var filaAplicado = filas[bloque.FilaEtiqueta + 3];
        var filaAnticipos = filas[bloque.FilaEtiqueta + 4];
        var filaTotal = filas[bloque.FilaEtiqueta + 5];

        foreach (var (aseId, columna3, columna4) in bloque.Ases)
        {
            foreach (var (columna, reporte) in new[] { (columna3, Reporte3), (columna4, Reporte4) })
            {
                var celdaAplicado = $"{columna}{filaAplicado.RowIndex!.Value.ToString(CultureInfo.InvariantCulture)}";
                var celdaAnticipos = $"{columna}{filaAnticipos.RowIndex!.Value.ToString(CultureInfo.InvariantCulture)}";
                var celdaTotal = $"{columna}{filaTotal.RowIndex!.Value.ToString(CultureInfo.InvariantCulture)}";

                ExigirCeldaDeValor(filaAplicado, celdaAplicado, bloque.Etiqueta, aseId, FilaAplicado, reporte);
                ExigirCeldaDeValor(filaAnticipos, celdaAnticipos, bloque.Etiqueta, aseId, FilaAnticipos, reporte);
                ExigirCeldaDeCierre(workbookPart, filaTotal, celdaTotal, bloque.Etiqueta, aseId, columna);

                destinos.Add(new Destino(bloque.Etiqueta, aseId, FilaAplicado, reporte, columna, celdaAplicado));
                destinos.Add(new Destino(bloque.Etiqueta, aseId, FilaAnticipos, reporte, columna, celdaAnticipos));
            }
        }

        // Columnas M/N del bloque (cierre horizontal): fórmula, nunca destino de escritura.
        foreach (var columna in new[] { bloque.ColumnaTotalAses, ColumnaSiguiente(bloque.ColumnaTotalAses) })
        {
            foreach (var fila in new[] { filaAplicado, filaAnticipos })
            {
                var referencia = $"{columna}{fila.RowIndex!.Value.ToString(CultureInfo.InvariantCulture)}";
                if (fila.Elements<Cell>().Any(celda =>
                        string.Equals(celda.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase)))
                {
                    ExigirCeldaDeCierre(workbookPart, fila, referencia, bloque.Etiqueta, 0, columna);
                }
            }
        }

        return new RejillaBloque(destinos);
    }

    /// <summary>
    /// Banda superior de la hoja (encima del primer bloque): cuadro de totales y controles. Exige que las
    /// filas que declaran fórmulas no mezclen ningún valor fijo y que existan al menos dos controles
    /// booleanos con fórmula (los <c>P3</c>/<c>P4</c> de T0 §2.1, sin nombrarlos ni ubicarlos por celda).
    /// </summary>
    private static void ValidarBandaSuperior(
        WorkbookPart workbookPart,
        IReadOnlyDictionary<uint, Row> filas,
        uint primeraFilaBloque,
        IReadOnlyList<Destino> destinos)
    {
        var filasBanda = filas.Values
            .Where(fila => fila.RowIndex!.Value < primeraFilaBloque)
            .OrderBy(fila => fila.RowIndex!.Value)
            .ToArray();
        if (filasBanda.Length == 0)
        {
            throw PlantillaInvalida(
                $"la hoja '{Hoja}' no declara banda superior (cuadro de totales y controles) encima del primer bloque");
        }

        var columnas = new HashSet<string>(
            destinos.Select(destino => destino.Columna),
            StringComparer.OrdinalIgnoreCase);

        var filasDeFormula = new List<uint>();
        foreach (var fila in filasBanda)
        {
            var conFormula = 0;
            var conValorFijo = new List<string>();
            foreach (var celda in fila.Elements<Cell>())
            {
                var columna = ColumnaDeReferencia(celda.CellReference?.Value);
                if (columna.Length == 0 || !columnas.Contains(columna))
                {
                    continue;
                }

                if (celda.CellFormula is not null)
                {
                    conFormula++;
                }
                else if (!string.IsNullOrEmpty(celda.CellValue?.InnerText))
                {
                    conValorFijo.Add(celda.CellReference?.Value ?? columna);
                }
            }

            if (conFormula == 0)
            {
                continue;
            }

            if (conValorFijo.Count > 0)
            {
                throw PlantillaInvalida(
                    $"la banda superior de '{Hoja}' mezcla fórmulas y valores fijos en la fila " +
                    $"{fila.RowIndex!.Value}: {string.Join(", ", conValorFijo)}. El cuadro de totales debe ser " +
                    "íntegramente fórmula (nunca se escribe dinero ahí)");
            }

            filasDeFormula.Add(fila.RowIndex!.Value);
        }

        if (filasDeFormula.Count < 2)
        {
            throw PlantillaInvalida(
                $"la banda superior de '{Hoja}' declara {filasDeFormula.Count} fila(s) de fórmula sobre las columnas " +
                "de valor (se esperaban al menos 2: aplicados y anticipos del cuadro de totales)");
        }

        var controles = filasBanda
            .SelectMany(fila => fila.Elements<Cell>())
            .Count(celda => celda.CellFormula is not null && celda.DataType?.Value == CellValues.Boolean);
        if (controles < 2)
        {
            throw PlantillaInvalida(
                $"la banda superior de '{Hoja}' declara {controles} control(es) booleano(s) con fórmula (se " +
                "esperaban al menos 2: los chequeos de cuadre de la hoja)");
        }
    }

    /// <summary>La hoja por nombre; ausente = <c>ERR-PLANTILLA</c> con el nombre y la identidad buscada.</summary>
    private static Worksheet HojaPorNombre(WorkbookPart workbookPart)
    {
        var workbook = workbookPart.Workbook
            ?? throw PlantillaInvalida("el workbook de la plantilla no tiene metadata válida");
        var hoja = workbook.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, Hoja, StringComparison.OrdinalIgnoreCase));
        if (hoja?.Id?.Value is null)
        {
            throw PlantillaInvalida($"la plantilla no trae la hoja '{Hoja}'");
        }

        return (workbookPart.GetPartById(hoja.Id.Value) as WorksheetPart)?.Worksheet
            ?? throw PlantillaInvalida($"la hoja '{Hoja}' no tiene WorksheetPart válido");
    }

    /// <summary>
    /// Celda de valor: debe existir y NO ser fórmula (una celda de valor con fórmula es una plantilla
    /// incompatible con la escritura values-only).
    /// </summary>
    private static void ExigirCeldaDeValor(
        Row fila,
        string referencia,
        string etiqueta,
        int aseId,
        string rotuloFila,
        string reporte)
    {
        var celda = fila.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase))
            ?? throw PlantillaInvalida(
                $"el bloque '{etiqueta}' de '{Hoja}' no declara la celda de valor {referencia} " +
                $"(ASE {aseId} · {rotuloFila} · {reporte}): falta la rejilla");

        if (celda.CellFormula is not null)
        {
            throw PlantillaInvalida(
                $"la celda de valor {referencia} del bloque '{etiqueta}' de '{Hoja}' " +
                $"(ASE {aseId} · {rotuloFila} · {reporte}) trae una fórmula ('{celda.CellFormula.Text}'): ahí debe " +
                "vivir un valor y la escritura values-only nunca sobrescribe fórmulas");
        }
    }

    /// <summary>Celda de cierre del bloque (TOTAL por columna): debe ser fórmula.</summary>
    private static void ExigirCeldaDeCierre(
        WorkbookPart workbookPart,
        Row fila,
        string referencia,
        string etiqueta,
        int aseId,
        string columna)
    {
        var celda = fila.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        if (celda?.CellFormula is null)
        {
            var identidad = aseId > 0 ? $" (ASE {aseId})" : string.Empty;
            throw PlantillaInvalida(
                $"la celda de cierre {referencia} ('{columna}') del bloque '{etiqueta}' de '{Hoja}'{identidad} no es " +
                $"fórmula (encontrado: '{TextoDeCelda(workbookPart, celda)}'): el total del bloque debe calcularse solo");
        }
    }

    private static bool EsRotulo(WorkbookPart workbookPart, Row fila, string columna, string rotulo) =>
        string.Equals(TextoDeColumna(workbookPart, fila, columna).Trim(), rotulo, StringComparison.OrdinalIgnoreCase);

    private static string TextoDeColumna(WorkbookPart workbookPart, Row fila, string columna) =>
        LectorTextoHoja.TextoDeColumna(workbookPart, fila, columna);

    private static string TextoDeCelda(WorkbookPart workbookPart, Cell? celda) =>
        celda is null ? string.Empty : LectorTextoHoja.Texto(workbookPart, celda);

    /// <summary>Verdad si el rótulo declara un ASE (<c>ASE n</c>) y devuelve su identificador.</summary>
    private static bool TryAseDeRotulo(string texto, out int aseId)
    {
        aseId = 0;
        var partes = texto.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        return partes.Length == 2
            && string.Equals(partes[0], "ASE", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(partes[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out aseId)
            && aseId > 0;
    }

    /// <summary>Letras de columna de una referencia A1 (sin dígitos ni <c>$</c>).</summary>
    private static string ColumnaDeReferencia(string? referencia)
    {
        if (string.IsNullOrEmpty(referencia))
        {
            return string.Empty;
        }

        var fin = 0;
        while (fin < referencia.Length && (referencia[fin] == '$' || char.IsLetter(referencia[fin])))
        {
            fin++;
        }

        return referencia[..fin].Replace("$", string.Empty, StringComparison.Ordinal);
    }

    /// <summary>Columna inmediatamente a la derecha de <paramref name="columna"/> (A1: A→B, Z→AA).</summary>
    private static string ColumnaSiguiente(string columna)
    {
        var indice = 0;
        foreach (var caracter in columna)
        {
            indice = (indice * 26) + (char.ToUpperInvariant(caracter) - 'A' + 1);
        }

        indice++;
        var resultado = string.Empty;
        while (indice > 0)
        {
            var resto = (indice - 1) % 26;
            resultado = (char)('A' + resto) + resultado;
            indice = (indice - 1) / 26;
        }

        return resultado;
    }

    /// <summary>
    /// Columna donde la hoja declara las ETIQUETAS de bloque. Es la única gramática de la hoja que no se
    /// puede derivar de su contenido (todas las etiquetas viven en la misma columna): se declara una vez,
    /// con cita, como hace el resto de la solución con la columna de rótulos de un bloque.
    /// </summary>
    private const string ColumnaEtiquetaBloque = "B";

    private static string Normalizar(string? texto) => (texto ?? string.Empty).Trim().ToUpperInvariant();

    private static CalculoInvalidoException PlantillaInvalida(string detalle) =>
        new(
            CodigoError.Plantilla,
            $"ERR-PLANTILLA: {detalle}. La plantilla es la autoridad de los destinos: no se escribe nada hasta que " +
            "su estructura se pueda resolver por etiquetas (nunca por coordenada congelada).");
}
