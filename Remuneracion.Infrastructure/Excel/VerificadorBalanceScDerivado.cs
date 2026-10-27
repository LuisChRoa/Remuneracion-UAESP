using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 59 (H5, §0.1/§0.2): verificación de las protecciones de la hoja <c>BCE SC POR FACT.</c> por
/// IDENTIDAD + FORMA, derivada de lo que la propia hoja declara.
///
/// Sustituye la foto congelada <c>WorkbookLeafCellMapBalanceSc.Protegidas</c> (78 celdas literales de
/// la v1 que, además, codificaban el nombre de hoja de un período concreto en sus fragmentos). Esa
/// foto queda como material histórico de fixtures; el camino productivo no la consume.
///
/// Qué se deriva, sin un solo literal de fila/celda/período:
/// <list type="bullet">
/// <item><b>Las bandas de la hoja</b> por los rótulos que publican sus encabezados: la banda de
/// totales (<c>ASE · SUBSIDIO · CONTRIBUCION · TOTAL BSC · SISTEMA</c>) y la banda de validación
/// (<c>ASE · REMUN TOTAL · VALIDA 1 · VALIDA 2 · DIFERENCIA VALIDA 2</c>). Las columnas salen de esos
/// rótulos; las filas, del id de ASE que la propia hoja numera.</item>
/// <item><b>El vínculo con el detalle del período</b>: <c>SISTEMA</c> debe reflejar el balance del ASE
/// en la hoja <c>DetRetri</c> REAL del workbook (<see cref="HojasPeriodo"/>), resuelta —columna y
/// fila— por el rótulo <c>BALANCE DE SUBSIDIO O CONTRIBUCIONES</c> y el id del ASE. Es el punto donde
/// la foto vieja congelaba el nombre de una hoja de un período concreto.</item>
/// <item><b>El vínculo con el CONSOLIDADO</b>: <c>REMUN TOTAL</c> debe reflejar la celda que la tabla
/// «REPORTE CONSOLIDADO TOTAL DE RECAUDO» declara para ese ASE, resuelta por encabezado.</item>
/// <item><b>Las sumas y los cuadres</b>: la suma de cada columna de valor sobre su tramo de ASE, la
/// suma cruzada de subsidio+contribución, la diferencia de cada ASE y de sus agregados, y la
/// comparación de cuadre (<c>ROUND</c>) entre ambas sumas.</item>
/// </list>
///
/// Fail-fast accionable (<c>ERR-PLANTILLA</c>) por identidad (banda · rol · ASE · rótulo), con la celda
/// hallada solo como contexto. Cero ramas por empresa/ASE/período/quincena.
/// </summary>
internal static class VerificadorBalanceScDerivado
{
    /// <summary>Rótulo de la columna de identidad de fila de ambas bandas (vocabulario de la plantilla).</summary>
    private const string RotuloAse = "ASE";

    private const string RotuloSubsidio = "SUBSIDIO";
    private const string RotuloContribucion = "CONTRIBUCION";
    private const string RotuloTotalBsc = "TOTAL BSC";
    private const string RotuloSistema = "SISTEMA";
    private const string RotuloRemuneracionTotal = "REMUN TOTAL";
    private const string RotuloValida1 = "VALIDA 1";
    private const string RotuloValida2 = "VALIDA 2";
    private const string RotuloDiferenciaValida2 = "DIFERENCIA VALIDA 2";

    /// <summary>
    /// Rótulo del balance de subsidio o contribuciones: la MISMA identidad de columna en la hoja de
    /// detalle del período y en la tabla del CONSOLIDADO (la plantilla solo cambia mayúsculas y
    /// decoración; el verificador compara por rótulo, no por texto literal).
    /// </summary>
    private const string RotuloBalance = "BALANCE DE SUBSIDIO O CONTRIBUCIONES";

    /// <summary>
    /// Verifica la hoja de balance de subsidio/contribución y sus dos vínculos externos (detalle del
    /// período y CONSOLIDADO).
    /// </summary>
    internal static void Validar(WorkbookPart parte, Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(periodo);

        var hoja = WorkbookLeafCellMapBalanceSc.HojaBce;
        var worksheet = HojaFisica(parte, hoja);
        var filas = Filas(worksheet);

        var banda1 = LocalizarBanda(
            parte, filas, hoja,
            RotuloTotalBsc,
            [RotuloAse, RotuloSubsidio, RotuloContribucion, RotuloTotalBsc, RotuloSistema],
            "banda de totales (subsidio · contribución · total · sistema)");

        var banda2 = LocalizarBanda(
            parte,
            filas.Where(f => f.RowIndex!.Value > banda1.Encabezado.RowIndex!.Value).ToList(),
            hoja,
            RotuloRemuneracionTotal,
            [RotuloAse, RotuloRemuneracionTotal, RotuloValida1, RotuloValida2, RotuloDiferenciaValida2],
            "banda de validación (remuneración vs totales del balance)");

        var delPeriodo = filas
            .Where(f => f.RowIndex!.Value > banda1.Encabezado.RowIndex!.Value
                && f.RowIndex!.Value < banda2.Encabezado.RowIndex!.Value)
            .ToList();

        var aseFilas1 = FilasDeAse(parte, delPeriodo, banda1.ColumnaAse, hoja, "banda de totales");
        var aseFilas2 = FilasDeAse(
            parte,
            filas.Where(f => f.RowIndex!.Value > banda2.Encabezado.RowIndex!.Value).ToList(),
            banda2.ColumnaAse, hoja, "banda de validación");

        var colSubsidio = banda1.Columna(RotuloSubsidio);
        var colContribucion = banda1.Columna(RotuloContribucion);
        var colTotalBsc = banda1.Columna(RotuloTotalBsc);
        var colSistema = banda1.Columna(RotuloSistema);

        ValidarColumnaDiferencia(parte, hoja, aseFilas1, banda1.Columnas.Values, colTotalBsc, colSistema);

        var hojaDetalle = HojasPeriodo.DetRetri(parte, periodo);
        var detalle = HojaFisica(parte, hojaDetalle);
        var cierreDetalle = ResolvedorDestinoValor.ResolverTabla(
            parte, detalle,
            new ResolvedorDestinoValor.IdentidadTabla(hojaDetalle, Titulo: null, AseId: null, RotuloBalance));

        foreach (var (aseId, fila) in aseFilas1)
        {
            var r = TextoFila(fila);

            VerificadorDestinoProtegido.ExigirForma(
                parte, worksheet, hoja, $"{colTotalBsc}{r}",
                [(1, $"{colSubsidio}{r}"), (1, $"{colContribucion}{r}")],
                $"BCE · ASE {aseId} · '{RotuloTotalBsc}' = '{RotuloSubsidio}' + '{RotuloContribucion}'");

            var balanceDelAse = ResolvedorDestinoValor.ResolverTabla(
                parte, detalle,
                new ResolvedorDestinoValor.IdentidadTabla(hojaDetalle, Titulo: null, aseId, RotuloBalance));

            VerificadorDestinoProtegido.ExigirForma(
                parte, worksheet, hoja, $"{colSistema}{r}",
                [(1, $"{hojaDetalle}!{balanceDelAse.Celda}")],
                $"BCE · ASE {aseId} · '{RotuloSistema}' = balance del período en '{hojaDetalle}' (ASE {aseId})");
        }

        var totales = ValidarAgregados(parte, hoja, delPeriodo, aseFilas1, colSubsidio, colContribucion, colTotalBsc, colSistema);

        ValidarCuadre(parte, delPeriodo, totales.TotalBsc, totales.SubsidioMasContribucion);

        ValidarCierreContraDetalle(parte, delPeriodo, totales.Sistema, hojaDetalle, cierreDetalle.Celda);

        var colRemuneracion = banda2.Columna(RotuloRemuneracionTotal);
        var cierreBanda2 = LocalizarCierreBanda2(parte, filas, aseFilas2, colRemuneracion, hoja);

        var consolidado = HojaFisica(parte, DelimitadorBloquesHoja.NombreConsolidado);
        foreach (var (aseId, fila) in aseFilas2)
        {
            var r = TextoFila(fila);
            var fila1 = aseFilas1.FirstOrDefault(par => par.AseId == aseId);
            if (fila1.Fila is null)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"ERR-PLANTILLA: BCE · la banda de validación declara el ASE {aseId} y la banda de totales no: la hoja no describe el mismo catálogo de ASE en sus dos bandas.");
            }

            var balanceConsolidado = ResolvedorDestinoValor.ResolverTabla(
                parte, consolidado,
                new ResolvedorDestinoValor.IdentidadTabla(
                    DelimitadorBloquesHoja.NombreConsolidado,
                    VerificadorDestinoProtegido.Tablas.TotalConsolidado.Titulo,
                    aseId,
                    RotuloBalance));

            VerificadorDestinoProtegido.ExigirForma(
                parte, worksheet, hoja, $"{colRemuneracion}{r}",
                [(1, $"{DelimitadorBloquesHoja.NombreConsolidado}!{balanceConsolidado.Celda}")],
                $"BCE · ASE {aseId} · '{RotuloRemuneracionTotal}' = balance del ASE en el CONSOLIDADO");

            var r1 = TextoFila(fila1.Fila);

            ExigirComparacion(
                parte, worksheet, hoja, $"{banda2.Columna(RotuloValida1)}{r}",
                $"{colRemuneracion}{r}", $"{colTotalBsc}{r1}",
                $"BCE · ASE {aseId} · '{RotuloValida1}' (remuneración del ASE = '{RotuloTotalBsc}' del ASE)");

            ExigirComparacion(
                parte, worksheet, hoja, $"{banda2.Columna(RotuloValida2)}{r}",
                $"{colRemuneracion}{r}", $"{colSistema}{r1}",
                $"BCE · ASE {aseId} · '{RotuloValida2}' (remuneración del ASE = '{RotuloSistema}' del ASE)");

            VerificadorDestinoProtegido.ExigirForma(
                parte, worksheet, hoja, $"{banda2.Columna(RotuloDiferenciaValida2)}{r}",
                [(1, $"{colRemuneracion}{r}"), (-1, $"{colSistema}{r1}")],
                $"BCE · ASE {aseId} · '{RotuloDiferenciaValida2}' = remuneración del ASE − '{RotuloSistema}' del ASE");
        }

        var rCierre = TextoFila(cierreBanda2);

        ExigirComparacion(
            parte, worksheet, hoja, $"{banda2.Columna(RotuloValida1)}{rCierre}",
            $"{colRemuneracion}{rCierre}", totales.TotalBsc,
            $"BCE · cierre de la banda de validación · '{RotuloValida1}' (remuneración total = suma de '{RotuloTotalBsc}')");

        ExigirComparacion(
            parte, worksheet, hoja, $"{banda2.Columna(RotuloValida2)}{rCierre}",
            $"{colRemuneracion}{rCierre}", totales.Sistema,
            $"BCE · cierre de la banda de validación · '{RotuloValida2}' (remuneración total = suma de '{RotuloSistema}')");

        VerificadorDestinoProtegido.ExigirForma(
            parte, worksheet, hoja, $"{banda2.Columna(RotuloDiferenciaValida2)}{rCierre}",
            [(1, $"{colRemuneracion}{rCierre}"), (-1, totales.Sistema)],
            $"BCE · cierre de la banda de validación · '{RotuloDiferenciaValida2}' = remuneración total − suma de '{RotuloSistema}'");
    }

    // ── Banda 1: agregados y cuadre ──────────────────────────────────────────────────────────────

    /// <summary>Agregados derivados de la banda de totales (referencias ya resueltas por rol).</summary>
    private sealed record Agregados(string Subsidio, string Contribucion, string TotalBsc, string Sistema, string SubsidioMasContribucion);

    private static Agregados ValidarAgregados(
        WorkbookPart parte,
        string hoja,
        IReadOnlyList<Row> region,
        IReadOnlyList<(int AseId, Row Fila)> aseFilas,
        string colSubsidio,
        string colContribucion,
        string colTotalBsc,
        string colSistema)
    {
        var primera = TextoFila(aseFilas[0].Fila);
        var ultima = TextoFila(aseFilas[^1].Fila);

        var subsidio = ExigirSumaDeTramo(parte, hoja, region, colSubsidio, colSubsidio, colSubsidio, primera, ultima, RotuloSubsidio);
        var contribucion = ExigirSumaDeTramo(parte, hoja, region, colContribucion, colContribucion, colContribucion, primera, ultima, RotuloContribucion);
        var total = ExigirSumaDeTramo(parte, hoja, region, colTotalBsc, colTotalBsc, colTotalBsc, primera, ultima, RotuloTotalBsc);
        var sistema = ExigirSumaDeTramo(parte, hoja, region, colSistema, colSistema, colSistema, primera, ultima, RotuloSistema);

        // Suma cruzada: el total del balance se cuadra contra la suma de subsidio+contribución.
        var cruzada = ExigirSumaDeTramo(
            parte, hoja, region, colTotalBsc, colSubsidio, colContribucion, primera, ultima,
            $"'{RotuloTotalBsc}' sobre '{RotuloSubsidio}'+'{RotuloContribucion}'");

        return new Agregados(subsidio, contribucion, total, sistema, cruzada);
    }

    /// <summary>
    /// Exige, dentro de la región de totales de la banda, exactamente una celda que sume el tramo de
    /// ASE indicado. La suma es el ROL (Σ del tramo que la propia hoja numera), no una coordenada.
    /// </summary>
    private static string ExigirSumaDeTramo(
        WorkbookPart parte,
        string hoja,
        IReadOnlyList<Row> region,
        string columna,
        string columnaInicio,
        string columnaFin,
        string primeraFila,
        string ultimaFila,
        string identidad)
    {
        var esperado = $"SUM({columnaInicio}{primeraFila}:{columnaFin}{ultimaFila})";
        var halladas = Formulas(region, columna)
            .Where(f => string.Equals(
                VerificadorDestinoProtegido.NormalizarFormula(f.Formula),
                VerificadorDestinoProtegido.NormalizarFormula(esperado),
                StringComparison.Ordinal))
            .ToList();

        if (halladas.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: BCE · la banda de totales declara {halladas.Count} sumas de '{identidad}' sobre su tramo de ASE y se exige exactamente una (forma '{esperado}'). " +
                $"Hoja '{hoja}'. La suma es el rol de la banda; no se reacomoda por posición.");
        }

        _ = parte;
        return $"{columna}{TextoFila(halladas[0].Fila)}";
    }

    /// <summary>
    /// Deriva la columna del rol «diferencia» (la plantilla no la rotula) por la FORMA que declara en
    /// cada fila de ASE: <c>TOTAL BSC − SISTEMA</c>. Debe ser una sola para toda la banda.
    /// </summary>
    private static void ValidarColumnaDiferencia(
        WorkbookPart parte,
        string hoja,
        IReadOnlyList<(int AseId, Row Fila)> aseFilas,
        IEnumerable<string> columnasConocidas,
        string colTotalBsc,
        string colSistema)
    {
        var conocidas = columnasConocidas.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var candidatas = ColumnasConValor(parte, aseFilas.Select(par => par.Fila).ToList())
            .Where(columna => !conocidas.Contains(columna))
            .ToList();

        var coincidentes = candidatas
            .Where(columna => aseFilas.All(par => CoincideDiferencia(par.Fila, columna, colTotalBsc, colSistema)))
            .ToList();

        if (coincidentes.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: BCE · la banda de totales de '{hoja}' declara {coincidentes.Count} columnas con el rol de diferencia ('{RotuloTotalBsc}' − '{RotuloSistema}') por ASE y se exige exactamente una. " +
                $"Columnas candidatas: {(candidatas.Count == 0 ? "(ninguna)" : string.Join(", ", candidatas))}. El rol se deriva de la forma que la fila declara, nunca de una letra fija.");
        }
    }

    private static bool CoincideDiferencia(Row fila, string columna, string colTotalBsc, string colSistema)
    {
        var r = TextoFila(fila);
        var celda = BuscarCelda(fila, $"{columna}{r}");
        if (celda?.CellFormula is null)
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(celda.CellFormula.Text) && celda.CellFormula.SharedIndex is not null)
        {
            return true;
        }

        var terminos = VerificadorDestinoProtegido.Terminos(celda.CellFormula.Text, WorkbookLeafCellMapBalanceSc.HojaBce);
        return terminos is not null
            && VerificadorDestinoProtegido.MismaForma(
                terminos,
                [(1, $"{colTotalBsc}{r}"), (-1, $"{colSistema}{r}")],
                WorkbookLeafCellMapBalanceSc.HojaBce);
    }

    /// <summary>
    /// Exige la comparación de cuadre entre la suma del total del balance y la suma cruzada
    /// subsidio+contribución (la plantilla la publica con <c>ROUND(…,2)</c>). Las referencias se
    /// derivan; nunca se compara una lista de operandos congelada.
    /// </summary>
    private static void ValidarCuadre(
        WorkbookPart parte,
        IReadOnlyList<Row> region,
        string totalBsc,
        string subsidioMasContribucion)
    {
        var esperadoTotal = VerificadorDestinoProtegido.NormalizarFormula(totalBsc);
        var esperadoCruzada = VerificadorDestinoProtegido.NormalizarFormula(subsidioMasContribucion);

        var halladas = region.Count(fila => fila.Elements<Cell>().Any(celda =>
        {
            var formula = celda.CellFormula?.Text;
            if (string.IsNullOrWhiteSpace(formula))
            {
                return false;
            }

            var texto = VerificadorDestinoProtegido.NormalizarFormula(formula);
            return texto.Contains("ROUND(", StringComparison.Ordinal)
                && texto.Contains('=')
                && texto.Contains(esperadoTotal, StringComparison.Ordinal)
                && texto.Contains(esperadoCruzada, StringComparison.Ordinal);
        }));

        if (halladas == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: BCE · la banda de totales perdió el cuadre con ROUND entre la suma de '{RotuloTotalBsc}' ('{totalBsc}') y la suma cruzada de '{RotuloSubsidio}'+'{RotuloContribucion}' ('{subsidioMasContribucion}'). " +
                "El rol de cuadre se verifica por las referencias que la banda declara, no por la celda que ocupaba en una foto anterior.");
        }

        _ = parte;
    }

    /// <summary>
    /// El agregado de <c>SISTEMA</c> debe contrastarse contra el CIERRE del balance en la hoja de
    /// detalle del período (la fila que la tabla declara como total). El nombre de la hoja es el REAL
    /// del workbook (<see cref="HojasPeriodo"/>) y la celda se resuelve por rótulo.
    /// </summary>
    private static void ValidarCierreContraDetalle(
        WorkbookPart parte,
        IReadOnlyList<Row> region,
        string sistema,
        string hojaDetalle,
        string celdaCierreDetalle)
    {
        var esperadoDerecha = $"{hojaDetalle}!{celdaCierreDetalle}";
        var halladas = region.Count(fila => fila.Elements<Cell>().Any(celda =>
        {
            var (izquierda, derecha) = PartirComparacion(celda.CellFormula?.Text);
            return izquierda is not null
                && VerificadorDestinoProtegido.MismaReferencia(izquierda, sistema, WorkbookLeafCellMapBalanceSc.HojaBce)
                && VerificadorDestinoProtegido.MismaReferencia(derecha, esperadoDerecha, WorkbookLeafCellMapBalanceSc.HojaBce);
        }));

        if (halladas != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: BCE · la banda de totales declara {halladas} comparaciones del agregado '{RotuloSistema}' ('{sistema}') contra el cierre del balance del período ('{esperadoDerecha}') y se exige exactamente una. " +
                "El cierre se resuelve por rótulo en la hoja real del período, no por el nombre de una hoja de otro período.");
        }

        _ = parte;
    }

    // ── Banda 2: cierre de la validación ────────────────────────────────────────────────────────

    private static Row LocalizarCierreBanda2(
        WorkbookPart parte,
        IReadOnlyList<Row> filas,
        IReadOnlyList<(int AseId, Row Fila)> aseFilas,
        string colRemuneracion,
        string hoja)
    {
        var primera = TextoFila(aseFilas[0].Fila);
        var ultima = TextoFila(aseFilas[^1].Fila);
        var esperado = VerificadorDestinoProtegido.NormalizarFormula($"SUM({colRemuneracion}{primera}:{colRemuneracion}{ultima})");

        var candidatas = filas
            .Where(f => f.RowIndex!.Value > aseFilas[^1].Fila.RowIndex!.Value)
            .Select(f => (Fila: f, Celda: BuscarCelda(f, $"{colRemuneracion}{TextoFila(f)}")))
            .Where(par => par.Celda?.CellFormula is not null
                && string.Equals(
                    VerificadorDestinoProtegido.NormalizarFormula(par.Celda.CellFormula.Text),
                    esperado,
                    StringComparison.Ordinal))
            .ToList();

        if (candidatas.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: BCE · la banda de validación declara {candidatas.Count} filas de cierre (suma de '{RotuloRemuneracionTotal}' sobre su tramo de ASE, forma '{esperado}') y se exige exactamente una. Hoja '{hoja}'.");
        }

        _ = parte;
        return candidatas[0].Fila;
    }

    // ── Detección estructural de las bandas ─────────────────────────────────────────────────────

    /// <summary>Banda localizada: fila de encabezado + columna de identidad + rótulo → letra.</summary>
    private sealed record Banda(Row Encabezado, string ColumnaAse, IReadOnlyDictionary<string, string> Columnas)
    {
        internal string Columna(string rotulo) => Columnas[rotulo];
    }

    /// <summary>
    /// Localiza la fila de encabezado de una banda por sus rótulos (todos en la MISMA fila) y devuelve
    /// la fila más el mapa rótulo → letra de columna declarada. Falla si no hay exactamente una.
    /// </summary>
    private static Banda LocalizarBanda(
        WorkbookPart parte,
        IReadOnlyList<Row> filas,
        string hoja,
        string rotuloDistintivo,
        IReadOnlyList<string> rotulos,
        string identidadBanda)
    {
        var candidatas = filas
            .Where(fila => TextosDeFila(parte, fila)
                .Any(texto => string.Equals(texto.Valor.Trim(), rotuloDistintivo, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        if (candidatas.Count != 1)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{hoja}' declara {candidatas.Count} encabezados con el rótulo '{rotuloDistintivo}' y la {identidadBanda} se identifica por él: se exige exactamente uno.");
        }

        var encabezado = candidatas[0];
        var textos = TextosDeFila(parte, encabezado)
            .Where(par => par.Valor.Trim().Length > 0)
            .ToList();
        var columnas = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var rotulo in rotulos)
        {
            var coincidencias = textos
                .Where(par => string.Equals(par.Valor.Trim(), rotulo, StringComparison.OrdinalIgnoreCase))
                .Select(par => par.Columna)
                .ToList();

            if (coincidencias.Count != 1)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"ERR-PLANTILLA: la {identidadBanda} de '{hoja}' declara {coincidencias.Count} columnas '{rotulo}' y se exige exactamente una. " +
                    $"Rótulos declarados en el encabezado: {string.Join(", ", textos.Select(par => par.Valor.Trim()))}.");
            }

            columnas[rotulo] = coincidencias[0];
        }

        return new Banda(encabezado, columnas[RotuloAse], columnas);
    }

    /// <summary>Filas de ASE de una banda, identificadas por el id numérico que la propia hoja numera.</summary>
    private static IReadOnlyList<(int AseId, Row Fila)> FilasDeAse(
        WorkbookPart parte,
        IReadOnlyList<Row> filas,
        string columnaAse,
        string hoja,
        string identidadBanda)
    {
        var total = Remuneracion.Core.Constants.CarpetasAse.Prefijos.Count;
        var halladas = new List<(int AseId, Row Fila)>();

        foreach (var fila in filas)
        {
            var texto = LectorTextoHoja.TextoDeColumna(parte, fila, columnaAse).Trim();
            if (!int.TryParse(texto, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id)
                || id < 1
                || id > total)
            {
                continue;
            }

            if (halladas.Exists(par => par.AseId == id))
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"ERR-PLANTILLA: la {identidadBanda} de '{hoja}' declara dos filas para el ASE {id}: la identidad de fila es el id del catálogo y debe ser única.");
            }

            halladas.Add((id, fila));
        }

        if (halladas.Count != total)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la {identidadBanda} de '{hoja}' declara {halladas.Count} filas de ASE y el catálogo de dominio tiene {total}. " +
                $"ASE hallados: {(halladas.Count == 0 ? "(ninguno)" : string.Join(", ", halladas.Select(par => par.AseId)))}.");
        }

        return halladas.OrderBy(par => par.AseId).ToList();
    }

    // ── Comparaciones (rol «valida») ─────────────────────────────────────────────────────────────

    private static void ExigirComparacion(
        WorkbookPart parte,
        Worksheet worksheet,
        string hoja,
        string celda,
        string izquierdaEsperada,
        string derechaEsperada,
        string identidad)
    {
        var hallada = ResolvedorDestinoValor.BuscarCelda(worksheet, celda)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidad}: la plantilla no declara la celda del rol (celda hallada {celda} de '{hoja}').");

        if (hallada.CellFormula is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidad} debería ser la comparación que el template declara (celda hallada {celda} de '{hoja}') y la plantilla trae un valor fijo.");
        }

        if (string.IsNullOrWhiteSpace(hallada.CellFormula.Text) && hallada.CellFormula.SharedIndex is not null)
        {
            return; // follower de fórmula compartida: la forma se verifica en el maestro
        }

        var (izquierda, derecha) = PartirComparacion(hallada.CellFormula.Text);
        if (izquierda is null
            || !VerificadorDestinoProtegido.MismaReferencia(izquierda, izquierdaEsperada, hoja)
            || !VerificadorDestinoProtegido.MismaReferencia(derecha, derechaEsperada, hoja))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidad}: la comparación no es la que su geometría exige (celda hallada {celda} de '{hoja}'). " +
                $"Esperado: '{izquierdaEsperada}' = '{derechaEsperada}'; obtenido: '{hallada.CellFormula.Text}'. " +
                "Los operandos se derivan por rol y se comparan por identidad, nunca por texto congelado.");
        }

        _ = parte;
    }

    /// <summary>Parte una comparación plana <c>LHS = RHS</c>; <c>null</c> si no tiene esa forma.</summary>
    private static (string? Izquierda, string? Derecha) PartirComparacion(string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return (null, null);
        }

        var texto = formula.Replace(" ", string.Empty, StringComparison.Ordinal);
        var indice = texto.IndexOf('=', StringComparison.Ordinal);
        if (indice <= 0 || indice != texto.LastIndexOf('='))
        {
            return (null, null);
        }

        return (texto[..indice], texto[(indice + 1)..]);
    }

    // ── Utilidades ──────────────────────────────────────────────────────────────────────────────

    private static List<Row> Filas(Worksheet worksheet) =>
        worksheet.Elements<SheetData>().FirstOrDefault()?.Elements<Row>()
            .Where(fila => fila.RowIndex is not null)
            .OrderBy(fila => fila.RowIndex!.Value)
            .ToList()
        ?? [];

    private static IEnumerable<(string Columna, string Valor)> TextosDeFila(WorkbookPart parte, Row fila) =>
        fila.Elements<Cell>()
            .Select(celda => (
                Columna: ResolvedorDestinoValor.LetraDeColumna(celda.CellReference?.Value),
                Valor: LectorTextoHoja.Texto(parte, celda)))
            .Where(par => par.Columna.Length > 0);

    private static IEnumerable<(Row Fila, Cell Celda, string Formula)> Formulas(
        IReadOnlyList<Row> filas, string columna)
    {
        foreach (var fila in filas)
        {
            var celda = BuscarCelda(fila, $"{columna}{TextoFila(fila)}");
            if (celda?.CellFormula is not null && !string.IsNullOrWhiteSpace(celda.CellFormula.Text))
            {
                yield return (fila, celda, celda.CellFormula.Text);
            }
        }
    }

    private static IReadOnlyCollection<string> ColumnasConValor(WorkbookPart parte, IReadOnlyList<Row> filas) =>
        filas
            .SelectMany(fila => TextosDeFila(parte, fila))
            .Where(par => par.Valor.Trim().Length > 0)
            .Select(par => par.Columna)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static Cell? BuscarCelda(Row fila, string referencia) =>
        fila.Elements<Cell>().FirstOrDefault(celda =>
            string.Equals(celda.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));

    private static string TextoFila(Row fila) =>
        (fila.RowIndex ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "Fila sin índice en la hoja de balance de subsidio/contribución."))
            .Value.ToString(CultureInfo.InvariantCulture);

    private static Worksheet HojaFisica(WorkbookPart parte, string nombre)
    {
        ArgumentNullException.ThrowIfNull(parte);
        var hoja = parte.Workbook?.Sheets?.Elements<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: el workbook no trae la hoja '{nombre}' que exige la verificación del balance de subsidio/contribución.");
        return ((WorksheetPart)parte.GetPartById(hoja.Id!)).Worksheet
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{nombre}' no tiene WorksheetPart válido.");
    }
}
