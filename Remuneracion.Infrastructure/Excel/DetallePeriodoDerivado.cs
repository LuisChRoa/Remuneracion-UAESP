using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 59 (H2-bis, §2-D2.3-bis): verificación de las PROTECCIONES de las hojas de detalle del
/// período (<c>DetRetri{AAAAMMQ}</c> y <c>DetValiRetri{AAAAMMQ}</c>) por ROL + FORMA.
///
/// Sustituye la lista congelada <c>WorkbookLeafCellMapQ2.DetRetriProtected(periodo)</c> — 23 celdas
/// literales (D23:D28, D32:D36, D16:D21, D24:D29) con sus listas de operandos de la v2. La v3 movió
/// la geometría y cambió los operandos: exigir <c>['F75','E47','D73','F77']</c> en D32 era exigirle
/// a la plantilla nueva la foto de la vieja (en la v3 esa celda es
/// <c>+D23+'Reporte Componentes R1'!F89+…</c>).
///
/// Gramática MEDIDA por bytes en la v3 (T0 de esta tanda; ver
/// <c>Docs/T0-Plan59-H2bis-Detalle-Derivado.md</c>). Cada hoja declara TRES familias de filas por
/// ASE, en este orden físico, y cada familia se identifica por sus ETIQUETAS, nunca por filas:
/// <list type="number">
/// <item><b>Principal</b>: encabezado con la columna de identidad <c>B='ASE'</c>; una fila por ASE
/// (<c>B</c> = id del catálogo de dominio) y la fila de cierre (<c>C='Total'</c>).</item>
/// <item><b>Diferencia vs CONSOLIDADO</b>: filas por ASE con la misma identidad de fila. En DetRetri
/// declara encabezado propio (mismos rótulos de columna que la principal); en DetValiRetri NO lo
/// declara y hereda los rótulos de columna de la tabla principal (la hoja declara un único juego de
/// columnas de valor).</item>
/// <item><b>Verificación de composición</b>: filas por ASE con encabezado propio; en DetValiRetri
/// declara además su fila de cierre (<c>C='TOTAL'</c>).</item>
/// </list>
///
/// FORMA derivada, sin un solo operando declarado:
/// <list type="bullet">
/// <item><b>Diferencia</b> (identidad ASE k o cierre): <c>+&lt;celda principal de la misma
/// fila/columna&gt; − &lt;CONSOLIDADO · tabla TOTAL · misma identidad · misma columna&gt;</c>.</item>
/// <item><b>Verificación</b>: comparación <c>LHS = RHS</c>, <c>RHS</c> = celda principal de la misma
/// fila/columna y <c>LHS</c> = celda de diferencia de esa identidad MÁS las hojas de origen de las
/// cinco tablas de recaudo del CONSOLIDADO para esa identidad y columna. Las hojas se obtienen
/// EXPANDIENDO la fórmula que el CONSOLIDADO declara (referencia directa, suma de referencias o
/// <c>SUM(rango)</c>): un cambio de geometría o de cardinalidad de subsecciones cambia las hojas y la
/// verificación sigue siendo cierta sin tocar código.</item>
/// </list>
///
/// Fail-fast accionable por IDENTIDAD (hoja · familia · ASE · rótulo de columna), con la celda
/// hallada solo como contexto: un error que se explica con un número de celda es un defecto
/// (§2-D2.3-bis).
/// </summary>
internal static class DetallePeriodoDerivado
{
    /// <summary>Familia de filas por ASE dentro de una hoja de detalle (orden físico documentado).</summary>
    internal enum RolFamilia
    {
        /// <summary>Tabla principal: encabezado con <c>B='ASE'</c>.</summary>
        Principal,

        /// <summary>Diferencias contra el <c>CONSOLIDADO_TOTAL RECAUDO</c>.</summary>
        Diferencia,

        /// <summary>Verificación de composición (comparación contra la tabla principal).</summary>
        Verificacion,
    }

    /// <summary>Rótulo de la columna de identidad de fila (gramática del template).</summary>
    private const string RotuloIdentidad = "ASE";

    /// <summary>Rótulo de la columna de nombre del ASE (no es columna de valor).</summary>
    private const string RotuloNombre = "Concesionarios";

    /// <summary>Profundidad máxima de expansión de una fórmula del CONSOLIDADO (defecto de plantilla si se supera).</summary>
    private const int ProfundidadMaxima = 4;

    /// <summary>Familias exigidas por la gramática medida de las hojas de detalle del período.</summary>
    private static readonly RolFamilia[] FamiliasExigidas =
        [RolFamilia.Principal, RolFamilia.Diferencia, RolFamilia.Verificacion];

    // ── Puerta pública ──────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Verifica las protecciones de las dos hojas de detalle del período: por cada ASE del catálogo
    /// (y por la fila de cierre de la familia, cuando la declara) y por cada columna de valor que el
    /// encabezado principal declara, exige la forma del rol.
    /// </summary>
    internal static void Validar(WorkbookPart parte, Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(parte);
        ArgumentNullException.ThrowIfNull(periodo);

        ValidarHoja(parte, HojasPeriodo.DetRetri(parte, periodo));
        ValidarHoja(parte, HojasPeriodo.DetValiRetri(parte, periodo));
    }

    // ── Verificación por hoja ───────────────────────────────────────────────────────────────────

    private static void ValidarHoja(WorkbookPart parte, string nombreHoja)
    {
        var worksheet = HojaFisica(parte, nombreHoja);
        var consolidado = HojaFisica(parte, DelimitadorBloquesHoja.NombreConsolidado);
        var filas = Filas(worksheet);
        var familias = DelimitarFamilias(parte, filas, nombreHoja);

        var encabezadoPrincipal = EncabezadoConIdentidad(parte, filas, familias[RolFamilia.Principal], nombreHoja);

        var etiquetas = EtiquetasDeValor(parte, encabezadoPrincipal, nombreHoja);

        // Columna de valor PRINCIPAL de la hoja: la primera que la tabla declara. Es la columna que la
        // plantilla publica con una hoja de origen por tabla de recaudo y la que la protección congelada
        // cubría en las dos hojas de detalle (D de DetRetri y D de DetValiRetri). Las columnas derivadas
        // —BCE, INTERVENTORIA, sumas anidadas— tienen composición heterogénea (medido en el T0 de esta
        // tanda) y su verificación no es expresable por la regla de una hoja por tabla de recaudo.
        var columnaProtegida = etiquetas[0].Columna;
        var rotuloProtegido = etiquetas[0].Rotulo;

        // Contraparte del CONSOLIDADO: el bloque que la hoja refleja se identifica por el rótulo de su
        // primera columna de valor, y su primera columna es la del concepto protegido.
        var bloqueConsolidado = BloqueDeColumnas(parte, consolidado, rotuloProtegido);
        if (bloqueConsolidado.Count == 0
            || NormalizarRotulo(bloqueConsolidado[0].Rotulo) != NormalizarRotulo(rotuloProtegido))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la primera columna de valor de '{nombreHoja}' declara '{rotuloProtegido}' y el bloque del CONSOLIDADO que la hoja refleja empieza en '{(bloqueConsolidado.Count == 0 ? "(vacío)" : bloqueConsolidado[0].Rotulo)}'. La correspondencia es por rótulo: no se remapea en silencio.");
        }

        var consolidadoColumna = bloqueConsolidado[0].Columna;

        foreach (var identidad in Identidades(parte, familias[RolFamilia.Principal]))
        {
            var filaPrincipal = FilaDeIdentidad(parte, familias[RolFamilia.Principal], identidad)
                ?? throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"ERR-SIN-SLOT: la tabla principal de '{nombreHoja}' no declara la fila del ASE {identidad.Descripcion} (columna B = id del catálogo, o C = 'Total' para el cierre).");

            var mainRef = $"{columnaProtegida}{filaPrincipal.ToString(CultureInfo.InvariantCulture)}";

            if (FilaDeIdentidad(parte, familias[RolFamilia.Diferencia], identidad) is not { } filaDiferencia)
            {
                continue;
            }

            var difRef = $"{columnaProtegida}{filaDiferencia.ToString(CultureInfo.InvariantCulture)}";
            ValidarDiferencia(parte, worksheet, consolidado, nombreHoja, identidad, rotuloProtegido, difRef, mainRef, consolidadoColumna);

            if (FilaDeIdentidad(parte, familias[RolFamilia.Verificacion], identidad) is not { } filaVerificacion)
            {
                continue;
            }

            var verRef = $"{columnaProtegida}{filaVerificacion.ToString(CultureInfo.InvariantCulture)}";
            ValidarVerificacion(parte, worksheet, consolidado, nombreHoja, identidad, rotuloProtegido, verRef, difRef, mainRef, consolidadoColumna);
        }
    }

    // ── Rol Diferencia ──────────────────────────────────────────────────────────────────────────

    private static void ValidarDiferencia(
        WorkbookPart parte,
        Worksheet worksheet,
        Worksheet consolidado,
        string nombreHoja,
        Identidad identidad,
        string rotulo,
        string referencia,
        string mainRef,
        string consolidadoColumna)
    {
        var identidadTexto = IdentidadTexto(nombreHoja, identidad, rotulo, RolFamilia.Diferencia);
        var celda = Celda(worksheet, referencia, identidadTexto);
        var filaConsolidado = FilaEnTabla(parte, consolidado, DelimitadorBloquesHoja.NombreConsolidado,
            VerificadorDestinoProtegido.Tablas.TotalConsolidado.Titulo, identidad, identidadTexto);

        var esperados = new List<Termino>
        {
            new(1, nombreHoja, mainRef),
            new(-1, DelimitadorBloquesHoja.NombreConsolidado, $"{consolidadoColumna}{filaConsolidado.ToString(CultureInfo.InvariantCulture)}"),
        };

        ExigirForma(celda.CellFormula, nombreHoja, identidadTexto, referencia,
            $"{identidadTexto} (principal − CONSOLIDADO de la misma identidad y columna)", esperados, celda.CellFormula?.Text);
    }

    // ── Rol Verificación ────────────────────────────────────────────────────────────────────────

    private static void ValidarVerificacion(
        WorkbookPart parte,
        Worksheet worksheet,
        Worksheet consolidado,
        string nombreHoja,
        Identidad identidad,
        string rotulo,
        string referencia,
        string difRef,
        string mainRef,
        string consolidadoColumna)
    {
        var identidadTexto = IdentidadTexto(nombreHoja, identidad, rotulo, RolFamilia.Verificacion);
        var celda = Celda(worksheet, referencia, identidadTexto);

        if (celda.CellFormula is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidadTexto}: la plantilla trae un valor fijo donde el rol exige la comparación de composición (celda hallada {referencia}).");
        }

        var texto = celda.CellFormula.Text;
        if (string.IsNullOrWhiteSpace(texto) && celda.CellFormula.SharedIndex is not null)
        {
            return; // follower de fórmula compartida: la forma se verifica en el maestro
        }

        var comparacion = PartirComparacion(texto, identidadTexto, referencia);
        var esperados = new List<Termino> { new(1, nombreHoja, difRef) };
        esperados.AddRange(HojasDeOrigen(parte, consolidado, identidad, rotulo, consolidadoColumna, identidadTexto));

        ExigirForma(celda.CellFormula, nombreHoja, identidadTexto, referencia,
            $"{identidadTexto} (diferencia + hojas de origen de las cinco tablas de recaudo)", esperados, comparacion.Izquierda);

        var derecha = TerminosFormula(comparacion.Derecha, nombreHoja);
        if (derecha.Count != 1 || derecha[0].Signo != 1
            || !string.Equals(derecha[0].Hoja, nombreHoja, StringComparison.OrdinalIgnoreCase)
            || !string.Equals(derecha[0].Referencia, mainRef, StringComparison.OrdinalIgnoreCase))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidadTexto}: el lado derecho de la comparación debe ser la celda principal de la misma identidad y columna ('{mainRef}') y la fórmula trae '{comparacion.Derecha}' (celda hallada {referencia}).");
        }
    }

    /// <summary>
    /// Hojas de origen que alimentan una columna del CONSOLIDADO para una identidad, derivadas
    /// EXPANDIENDO la fórmula que la plantilla declara en cada una de las cinco tablas de recaudo.
    /// </summary>
    private static IEnumerable<Termino> HojasDeOrigen(
        WorkbookPart parte,
        Worksheet consolidado,
        Identidad identidad,
        string rotulo,
        string consolidadoColumna,
        string identidadTexto)
    {
        foreach (var tabla in VerificadorDestinoProtegido.Tablas.TodasLasDeRecaudo)
        {
            if (ReferenceEquals(tabla, VerificadorDestinoProtegido.Tablas.SaldosR5) && identidad.EsTotal)
            {
                // El cierre del CONSOLIDADO de la tabla R5 suma el rango de sus ASE; la verificación de
                // cierre usa el cierre propio de la hoja de origen (su fila C='TOTAL'), que es el valor
                // que la plantilla declara como total. Se resuelve por rol, nunca por fila.
                var hojaOrigen = HojaFisica(parte, DelimitadorBloquesHoja.NombreAjustesSfT);
                var filaOrigen = FilaEnTabla(parte, hojaOrigen, DelimitadorBloquesHoja.NombreAjustesSfT,
                    VerificadorDestinoProtegido.TituloAjustesSfT, identidad, identidadTexto);
                var columnaOrigen = ColumnaEnTabla(parte, DelimitadorBloquesHoja.NombreAjustesSfT,
                    VerificadorDestinoProtegido.TituloAjustesSfT, rotulo);
                yield return new Termino(1, DelimitadorBloquesHoja.NombreAjustesSfT,
                    $"{columnaOrigen}{filaOrigen.ToString(CultureInfo.InvariantCulture)}");
                continue;
            }

            var fila = FilaEnTabla(parte, consolidado, DelimitadorBloquesHoja.NombreConsolidado, tabla.Titulo, identidad, identidadTexto);
            var celdaConsolidado = $"{consolidadoColumna}{fila.ToString(CultureInfo.InvariantCulture)}";

            foreach (var termino in Expandir(parte, consolidado, celdaConsolidado, identidadTexto, 0))
            {
                yield return termino;
            }
        }
    }

    /// <summary>
    /// Expande una celda del CONSOLIDADO hasta las hojas de ORIGEN: una referencia a otra hoja es una
    /// hoja terminal; una referencia interna, una suma de referencias o un <c>SUM(rango)</c> se
    /// expanden. No se declaran operandos: se sigue la fórmula que la plantilla trae.
    /// </summary>
    private static IEnumerable<Termino> Expandir(
        WorkbookPart parte,
        Worksheet consolidado,
        string referencia,
        string identidadTexto,
        int profundidad)
    {
        if (profundidad > ProfundidadMaxima)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidadTexto}: la fórmula del CONSOLIDADO en '{referencia}' encadena más de {ProfundidadMaxima} expansiones y sus hojas de origen no se pueden derivar (posible ciclo en la plantilla).");
        }

        var celda = ResolvedorDestinoValor.BuscarCelda(consolidado, referencia)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidadTexto}: la celda '{referencia}' del CONSOLIDADO que la verificación necesita no existe en la plantilla.");

        if (celda.CellFormula is null || string.IsNullOrWhiteSpace(celda.CellFormula.Text))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidadTexto}: la celda '{referencia}' del CONSOLIDADO debería ser la fórmula que compone el recaudo y la plantilla trae un valor fijo. La verificación no adivina hojas de origen.");
        }

        var terminos = TerminosFormula(celda.CellFormula.Text, DelimitadorBloquesHoja.NombreConsolidado);
        if (terminos.Count == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidadTexto}: la fórmula de '{referencia}' ('{celda.CellFormula.Text}') no es una composición de referencias ni una suma de rango y sus hojas de origen no se pueden derivar.");
        }

        foreach (var termino in terminos)
        {
            if (!string.Equals(termino.Hoja, DelimitadorBloquesHoja.NombreConsolidado, StringComparison.OrdinalIgnoreCase))
            {
                yield return termino;
                continue;
            }

            foreach (var hoja in Expandir(parte, consolidado, termino.Referencia, identidadTexto, profundidad + 1))
            {
                yield return hoja with { Signo = termino.Signo > 0 ? hoja.Signo : -hoja.Signo };
            }
        }
    }

    // ── Familias, identidad y encabezados ───────────────────────────────────────────────────────

    /// <summary>Identidad de fila: un ASE del catálogo o la fila de cierre de la familia.</summary>
    internal sealed record Identidad(int? AseId)
    {
        internal bool EsTotal => AseId is null;

        internal string Descripcion => AseId is { } ase
            ? ase.ToString(CultureInfo.InvariantCulture)
            : "cierre";
    }

    private static IReadOnlyList<Identidad> Identidades(WorkbookPart parte, IReadOnlyList<Row> principal)
    {
        var lista = new List<Identidad>();
        for (var id = 1; id <= CarpetasAse.Prefijos.Count; id++)
        {
            lista.Add(new Identidad(id));
        }

        if (principal.Any(f => EsCierre(parte, f)))
        {
            lista.Add(new Identidad(null));
        }

        return lista;
    }

    /// <summary>
    /// Familias de filas por ASE: tramos consecutivos de filas cuya identidad de fila es un ASE del
    /// catálogo (<c>B = id</c>) o el cierre de la familia (<c>C = 'Total'</c>). La hoja debe declarar
    /// exactamente las tres familias documentadas, en orden.
    /// </summary>
    private static IReadOnlyDictionary<RolFamilia, IReadOnlyList<Row>> DelimitarFamilias(
        WorkbookPart parte,
        IReadOnlyList<Row> filas,
        string nombreHoja)
    {
        var tramos = new List<List<Row>>();
        List<Row>? actual = null;
        var filaAnterior = int.MinValue;

        foreach (var fila in filas)
        {
            var numero = (int)fila.RowIndex!.Value;
            if (IdentidadDeFila(parte, fila) is null)
            {
                actual = null;
                continue;
            }

            if (actual is null || numero != filaAnterior + 1)
            {
                actual = [];
                tramos.Add(actual);
            }

            actual.Add(fila);
            filaAnterior = numero;
        }

        if (tramos.Count != FamiliasExigidas.Length)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{nombreHoja}' declara {tramos.Count} familias de filas por ASE ({DescribirTramos(tramos)}) y la gramática de las hojas de detalle exige {FamiliasExigidas.Length}: {string.Join(", ", FamiliasExigidas.Select(Describir))}. La identidad de cada familia es su rol, nunca su posición.");
        }

        var resultado = new Dictionary<RolFamilia, IReadOnlyList<Row>>();
        for (var i = 0; i < FamiliasExigidas.Length; i++)
        {
            resultado[FamiliasExigidas[i]] = tramos[i];
        }

        return resultado;
    }

    private static string DescribirTramos(IReadOnlyList<List<Row>> tramos) =>
        tramos.Count == 0
            ? "ninguna"
            : string.Join(", ", tramos.Select(t => $"{t[0].RowIndex!.Value}-{t[^1].RowIndex!.Value} ({t.Count} filas)"));

    private static bool EsCierre(WorkbookPart parte, Row fila) =>
        NormalizarRotulo(LectorTextoHoja.TextoDeColumna(parte, fila, "C")) == NormalizarRotulo(DelimitadorBloquesHoja.MarcaTotal);

    /// <summary>
    /// Identidad de fila por etiqueta: <c>B</c> = id de un ASE del catálogo de dominio, o <c>C</c> =
    /// marca de cierre con <c>B</c> vacía. Cualquier otra fila no pertenece a una familia.
    /// </summary>
    private static Identidad? IdentidadDeFila(WorkbookPart parte, Row fila)
    {
        var textoB = LectorTextoHoja.TextoDeColumna(parte, fila, "B").Trim();
        if (int.TryParse(textoB, NumberStyles.Integer, CultureInfo.InvariantCulture, out var aseId)
            && aseId >= 1 && aseId <= CarpetasAse.Prefijos.Count)
        {
            return new Identidad(aseId);
        }

        if (textoB.Length == 0 && EsCierre(parte, fila))
        {
            return new Identidad(null);
        }

        return null;
    }

    private static int? FilaDeIdentidad(WorkbookPart parte, IReadOnlyList<Row> familia, Identidad identidad)
    {
        foreach (var fila in familia)
        {
            var actual = IdentidadDeFila(parte, fila);
            if (actual is null)
            {
                continue;
            }

            if (identidad.EsTotal ? actual.EsTotal : actual.AseId == identidad.AseId)
            {
                return (int)fila.RowIndex!.Value;
            }
        }

        return null;
    }

    /// <summary>
    /// Encabezado de una familia: la fila de rótulos de columna más cercana por ENCIMA del tramo (sin
    /// cruzar el tramo anterior). Si no la hay, la familia hereda el único juego de columnas de la
    /// hoja y usa el encabezado de la tabla principal.
    /// </summary>
    private static Row EncabezadoDeFamilia(
        WorkbookPart parte,
        IReadOnlyList<Row> filas,
        IReadOnlyDictionary<RolFamilia, IReadOnlyList<Row>> familias,
        RolFamilia rol,
        Row encabezadoPrincipal)
    {
        var familia = familias[rol];
        var primera = (int)familia[0].RowIndex!.Value;
        var limite = int.MinValue;

        foreach (var otra in familias)
        {
            if (otra.Key == rol)
            {
                continue;
            }

            var fin = (int)otra.Value[^1].RowIndex!.Value;
            if (fin < primera && fin > limite)
            {
                limite = fin;
            }
        }

        for (var i = filas.Count - 1; i >= 0; i--)
        {
            var numero = (int)filas[i].RowIndex!.Value;
            if (numero >= primera || numero <= limite)
            {
                continue;
            }

            if (EsFilaDeRotulos(parte, filas[i]))
            {
                return filas[i];
            }
        }

        return encabezadoPrincipal;
    }

    /// <summary>
    /// Encabezado de la tabla principal: la fila de rótulos más cercana por encima del tramo que
    /// declara la columna de identidad <c>ASE</c>.
    /// </summary>
    private static Row EncabezadoConIdentidad(WorkbookPart parte, IReadOnlyList<Row> filas, IReadOnlyList<Row> familia, string nombreHoja)
    {
        var primera = (int)familia[0].RowIndex!.Value;
        for (var i = filas.Count - 1; i >= 0; i--)
        {
            var numero = (int)filas[i].RowIndex!.Value;
            if (numero >= primera)
            {
                continue;
            }

            if (EsEncabezadoDeTabla(parte, filas[i]))
            {
                return filas[i];
            }
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"ERR-PLANTILLA: la hoja '{nombreHoja}' no declara el encabezado de columnas de su tabla principal (la fila de rótulos con la columna de identidad '{RotuloIdentidad}'). Sin rótulos no hay resolución por concepto.");
    }

    /// <summary>
    /// Fila de rótulos: declara al menos dos textos en las columnas de valor y ningún valor numérico
    /// (misma regla estructural que el delimitador de bloques de recaudo). Es el encabezado de las
    /// familias derivadas de las hojas de detalle, que no repiten la columna de identidad.
    /// </summary>
    private static bool EsFilaDeRotulos(WorkbookPart parte, Row fila)
    {
        var textos = 0;
        foreach (var celda in fila.Elements<Cell>())
        {
            var columna = ResolvedorDestinoValor.LetraDeColumna(celda.CellReference?.Value);
            if (columna.Length == 0 || columna is "A" or "B" or "C")
            {
                continue;
            }

            if (EsCeldaDeTexto(celda) && !string.IsNullOrWhiteSpace(LectorTextoHoja.Texto(parte, celda)))
            {
                textos++;
                continue;
            }

            if (!string.IsNullOrWhiteSpace(celda.CellValue?.InnerText) && !EsCeldaDeTexto(celda))
            {
                return false;
            }
        }

        return textos >= 2;
    }

    /// <summary>Verdad si la celda guarda TEXTO (shared string o inline string) y no un valor calculado.</summary>
    private static bool EsCeldaDeTexto(Cell celda) =>
        celda.DataType?.Value == CellValues.SharedString || celda.DataType?.Value == CellValues.InlineString;

    /// <summary>Verdad si la fila declara un rótulo de texto en la columna indicada (nunca un valor con caché).</summary>
    private static bool TieneRotulo(WorkbookPart parte, Row fila, string columna) =>
        fila.Elements<Cell>()
            .Where(c => string.Equals(ResolvedorDestinoValor.LetraDeColumna(c.CellReference?.Value), columna, StringComparison.OrdinalIgnoreCase))
            .Any(c => EsCeldaDeTexto(c) && !string.IsNullOrWhiteSpace(LectorTextoHoja.Texto(parte, c)));

    /// <summary>Encabezado de una tabla identificada por su título: además declara la columna de identidad <c>ASE</c>.</summary>
    private static bool EsEncabezadoDeTabla(WorkbookPart parte, Row fila) =>
        EsFilaDeRotulos(parte, fila) && ColumnaEnFila(parte, fila, RotuloIdentidad) is not null;

    /// <summary>Columnas de valor del encabezado principal: las que declaran rótulo menos identidad y nombre.</summary>
    private static IReadOnlyList<(string Columna, string Rotulo)> EtiquetasDeValor(
        WorkbookPart parte,
        Row encabezado,
        string nombreHoja)
    {
        var etiquetas = new List<(string, string)>();
        foreach (var celda in encabezado.Elements<Cell>())
        {
            var columna = ResolvedorDestinoValor.LetraDeColumna(celda.CellReference?.Value);
            var rotulo = LectorTextoHoja.Texto(parte, celda).Trim();
            if (columna.Length == 0 || rotulo.Length == 0)
            {
                continue;
            }

            var normalizado = NormalizarRotulo(rotulo);
            if (normalizado == NormalizarRotulo(RotuloIdentidad) || normalizado == NormalizarRotulo(RotuloNombre))
            {
                continue;
            }

            etiquetas.Add((columna, rotulo));
        }

        if (etiquetas.Count == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: el encabezado de la tabla principal de '{nombreHoja}' no declara columnas de valor además de '{RotuloIdentidad}'/'{RotuloNombre}'. Sin rótulo no hay verificación por concepto.");
        }

        return etiquetas;
    }

    // ── Localización por etiqueta ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Bloque de columnas de valor de una tabla, identificado por el rótulo de su PRIMERA columna: se
    /// recorre la fila de encabezado desde esa columna hasta la columna de identidad del bloque
    /// siguiente (<c>'ASE'</c>). Devuelve las columnas en el orden que la plantilla declara.
    /// </summary>
    private static IReadOnlyList<(string Columna, string Rotulo)> BloqueDeColumnas(
        WorkbookPart parte,
        Worksheet worksheet,
        string primerRotulo)
    {
        var buscado = NormalizarRotulo(primerRotulo);
        var identidad = NormalizarRotulo(RotuloIdentidad);
        var nombre = NormalizarRotulo(RotuloNombre);

        foreach (var fila in Filas(worksheet))
        {
            if (!EsEncabezadoDeTabla(parte, fila))
            {
                continue;
            }

            var celdas = fila.Elements<Cell>().ToList();
            var inicio = -1;
            for (var i = 0; i < celdas.Count; i++)
            {
                if (NormalizarRotulo(LectorTextoHoja.Texto(parte, celdas[i])) == buscado)
                {
                    inicio = i;
                    break;
                }
            }

            if (inicio < 0)
            {
                continue;
            }

            var bloque = new List<(string, string)>();
            for (var i = inicio; i < celdas.Count; i++)
            {
                var columna = ResolvedorDestinoValor.LetraDeColumna(celdas[i].CellReference?.Value);
                var rotulo = LectorTextoHoja.Texto(parte, celdas[i]).Trim();
                var normalizado = NormalizarRotulo(rotulo);
                if (normalizado == identidad)
                {
                    break;
                }

                if (columna.Length == 0 || normalizado.Length == 0 || normalizado == nombre)
                {
                    continue;
                }

                bloque.Add((columna, rotulo));
            }

            return bloque;
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"ERR-SIN-SLOT: ninguna tabla del CONSOLIDADO declara el bloque de columnas que empieza en el rótulo '{primerRotulo}'.");
    }

    private static string? ColumnaEnFila(WorkbookPart parte, Row fila, string rotulo)
    {
        var buscado = NormalizarRotulo(rotulo);
        foreach (var celda in fila.Elements<Cell>())
        {
            var columna = ResolvedorDestinoValor.LetraDeColumna(celda.CellReference?.Value);
            if (columna.Length == 0)
            {
                continue;
            }

            if (NormalizarRotulo(LectorTextoHoja.Texto(parte, celda)) == buscado)
            {
                return columna;
            }
        }

        return null;
    }

    /// <summary>
    /// Columna de un rótulo dentro de una tabla identificada por su título (columna <c>F</c>): el
    /// encabezado de la tabla es la primera fila de identidad por debajo del título.
    /// </summary>
    private static string ColumnaEnTabla(WorkbookPart parte, string nombreHoja, string? titulo, string rotulo)
    {
        var worksheet = HojaFisica(parte, nombreHoja);
        var filas = Filas(worksheet);
        var desde = 0;

        if (titulo is not null)
        {
            desde = IndiceDeTitulo(parte, filas, nombreHoja, titulo) + 1;
        }

        for (var i = desde; i < filas.Count; i++)
        {
            if (!EsEncabezadoDeTabla(parte, filas[i]))
            {
                continue;
            }

            var columna = ColumnaEnFila(parte, filas[i], rotulo);
            if (columna is not null)
            {
                return columna;
            }
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"ERR-SIN-SLOT: la tabla '{(titulo is null ? nombreHoja : titulo)}' de '{nombreHoja}' no declara la columna '{rotulo}'.");
    }

    private static int IndiceDeTitulo(WorkbookPart parte, IReadOnlyList<Row> filas, string nombreHoja, string titulo)
    {
        for (var i = 0; i < filas.Count; i++)
        {
            if (TieneRotulo(parte, filas[i], "F")
                && NormalizarRotulo(LectorTextoHoja.TextoDeColumna(parte, filas[i], "F")) == NormalizarRotulo(titulo))
            {
                return i;
            }
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"ERR-PLANTILLA: la hoja '{nombreHoja}' no declara la tabla '{titulo}' (se esperaba su título en la columna F).");
    }

    /// <summary>
    /// Fila de una identidad dentro de la tabla identificada por <paramref name="titulo"/>: se busca
    /// entre el encabezado de esa tabla y el título de la siguiente (la región que la tabla declara).
    /// </summary>
    private static int FilaEnTabla(
        WorkbookPart parte,
        Worksheet worksheet,
        string nombreHoja,
        string titulo,
        Identidad identidad,
        string identidadTexto)
    {
        var filas = Filas(worksheet);
        var indiceTitulo = IndiceDeTitulo(parte, filas, nombreHoja, titulo);

        var indiceEncabezado = -1;
        for (var i = indiceTitulo + 1; i < filas.Count; i++)
        {
            if (EsEncabezadoDeTabla(parte, filas[i]))
            {
                indiceEncabezado = i;
                break;
            }
        }

        if (indiceEncabezado < 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidadTexto}: la tabla '{titulo}' de '{nombreHoja}' no declara encabezado de columnas debajo de su título.");
        }

        for (var i = indiceEncabezado + 1; i < filas.Count; i++)
        {
            var fila = filas[i];
            if (EsEncabezadoDeTabla(parte, fila))
            {
                break;
            }

            if (TieneRotulo(parte, fila, "F"))
            {
                break;
            }

            var actual = IdentidadDeFila(parte, fila);
            if (actual is null)
            {
                continue;
            }

            if (identidad.EsTotal ? actual.EsTotal : actual.AseId == identidad.AseId)
            {
                return (int)fila.RowIndex!.Value;
            }
        }

        throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"ERR-SIN-SLOT: {identidadTexto}: la tabla '{titulo}' de '{nombreHoja}' no declara la fila de identidad '{identidad.Descripcion}'. La fila se resuelve por etiqueta, nunca por posición.");
    }

    // ── Forma de fórmulas ────────────────────────────────────────────────────────────────────────

    private sealed record Termino(int Signo, string Hoja, string Referencia);

    private sealed record Comparacion(string Izquierda, string Derecha);

    private static string IdentidadTexto(string nombreHoja, Identidad identidad, string rotulo, RolFamilia rol) =>
        $"hoja '{nombreHoja}' · ASE {identidad.Descripcion} · columna '{rotulo}' · familia '{Describir(rol)}'";

    private static string Describir(RolFamilia rol) => rol switch
    {
        RolFamilia.Principal => "principal",
        RolFamilia.Diferencia => "diferencia vs CONSOLIDADO",
        _ => "verificación de composición",
    };

    private static Comparacion PartirComparacion(string formula, string identidadTexto, string referencia)
    {
        var texto = Compactar(formula);
        var partes = texto.Split('=');
        if (partes.Length != 2 || partes[0].Length == 0 || partes[1].Length == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidadTexto}: se esperaba una comparación 'composición = valor' y la fórmula es '{formula}' (celda hallada {referencia}).");
        }

        return new Comparacion(partes[0], partes[1]);
    }

    /// <summary>
    /// Descompone una fórmula de suma/resta de referencias (con soporte de <c>SUM(rango)</c> como
    /// término) en hojas terminales <c>(signo, hoja, celda)</c>. Las referencias sin hoja se atribuyen
    /// a <paramref name="hojaPorDefecto"/>.
    /// </summary>
    private static IReadOnlyList<Termino> TerminosFormula(string formula, string? hojaPorDefecto)
    {
        var texto = Compactar(formula);
        var terminos = new List<Termino>();
        var i = 0;
        var signo = 1;

        if (texto.StartsWith('-'))
        {
            signo = -1;
            i = 1;
        }
        else if (texto.StartsWith('+'))
        {
            i = 1;
        }

        while (i < texto.Length)
        {
            var termino = LeerTermino(texto, ref i, formula);
            if (termino.StartsWith("SUM(", StringComparison.OrdinalIgnoreCase))
            {
                foreach (var expandido in ExpandirRango(termino, formula, hojaPorDefecto))
                {
                    terminos.Add(expandido with { Signo = signo * expandido.Signo });
                }
            }
            else
            {
                var (hoja, referencia) = SepararHoja(termino, hojaPorDefecto, formula);
                terminos.Add(new Termino(signo, hoja, referencia));
            }

            signo = 1;
            if (i < texto.Length)
            {
                if (texto[i] == '-')
                {
                    signo = -1;
                }
                else if (texto[i] != '+')
                {
                    throw new CalculoInvalidoException(
                        CodigoError.Plantilla,
                        $"ERR-PLANTILLA: la fórmula '{formula}' no es una suma/resta de referencias (carácter inesperado '{texto[i]}' en la posición {i + 1}).");
                }

                i++;
            }
        }

        return terminos;
    }

    private static IEnumerable<Termino> ExpandirRango(string termino, string formula, string? hojaPorDefecto)
    {
        var rango = termino[4..].TrimEnd(')');
        var partes = rango.Split(':');
        if (partes.Length != 2)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la fórmula '{formula}' declara 'SUM({rango})' y solo se admite la suma de un rango de una columna.");
        }

        var (hojaInicio, desde) = SepararHoja(partes[0], hojaPorDefecto, formula);
        var (hojaFin, hasta) = SepararHoja(partes[1], hojaPorDefecto, formula);
        if (!string.Equals(hojaInicio, hojaFin, StringComparison.OrdinalIgnoreCase))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la fórmula '{formula}' suma un rango entre hojas distintas ('{rango}'); no se puede derivar su forma.");
        }

        var columna = ResolvedorDestinoValor.LetraDeColumna(desde);
        var filaDesde = int.Parse(new string(desde.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);
        var filaHasta = int.Parse(new string(hasta.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);
        for (var fila = filaDesde; fila <= filaHasta; fila++)
        {
            yield return new Termino(1, hojaInicio, $"{columna}{fila.ToString(CultureInfo.InvariantCulture)}");
        }
    }

    /// <summary>
    /// Quita los espacios y saltos de línea FUERA de los nombres de hoja entrecomillados: la plantilla
    /// puede escribir 'AJUSTES - SF-T' y ese espacio interior es parte de la identidad de la hoja.
    /// </summary>
    private static string Compactar(string formula)
    {
        var sb = new StringBuilder(formula.Length);
        var entreComillas = false;
        foreach (var caracter in formula)
        {
            if (caracter == '\'')
            {
                entreComillas = !entreComillas;
                sb.Append(caracter);
                continue;
            }

            if (!entreComillas && char.IsWhiteSpace(caracter))
            {
                continue;
            }

            sb.Append(caracter);
        }

        return sb.ToString();
    }

    private static string LeerTermino(string texto, ref int i, string formula)
    {
        var inicio = i;
        var profundidad = 0;
        while (i < texto.Length)
        {
            var caracter = texto[i];
            if (caracter == '\'')
            {
                var cierre = texto.IndexOf('\'', i + 1);
                if (cierre < 0)
                {
                    throw new CalculoInvalidoException(
                        CodigoError.Plantilla,
                        $"ERR-PLANTILLA: la fórmula '{formula}' tiene una comilla de hoja sin cerrar.");
                }

                i = cierre + 1;
                continue;
            }

            if (caracter == '(')
            {
                profundidad++;
            }
            else if (caracter == ')')
            {
                profundidad--;
            }
            else if (profundidad == 0 && caracter is '+' or '-')
            {
                break;
            }

            i++;
        }

        if (i == inicio)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la fórmula '{formula}' declara un término vacío en la posición {inicio + 1}.");
        }

        return texto[inicio..i];
    }

    private static (string Hoja, string Referencia) SepararHoja(string termino, string? hojaPorDefecto, string formula)
    {
        var separador = termino.LastIndexOf('!');
        if (separador < 0)
        {
            if (string.IsNullOrWhiteSpace(hojaPorDefecto))
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"ERR-PLANTILLA: el término '{termino}' de la fórmula '{formula}' no declara la hoja y la verificación no puede atribuirlo a ninguna.");
            }

            return (hojaPorDefecto, termino.Replace("$", string.Empty, StringComparison.Ordinal));
        }

        var hoja = termino[..separador].Trim('\'').Trim();
        var referencia = termino[(separador + 1)..].Replace("$", string.Empty, StringComparison.Ordinal);
        return (hoja, referencia);
    }

    private static void ExigirForma(
        CellFormula? formula,
        string hojaPorDefecto,
        string identidadTexto,
        string referencia,
        string rol,
        IReadOnlyList<Termino> esperados,
        string? textoExplicito)
    {
        if (formula is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidadTexto}: la plantilla trae un valor fijo donde el rol exige fórmula (celda hallada {referencia}).");
        }

        if (string.IsNullOrWhiteSpace(formula.Text) && formula.SharedIndex is not null)
        {
            return;
        }

        var texto = textoExplicito ?? formula.Text;
        var obtenidos = TerminosFormula(texto, hojaPorDefecto).ToList();

        var esperado = esperados
            .OrderBy(t => t.Hoja, StringComparer.OrdinalIgnoreCase)
            .ThenBy(t => t.Referencia, StringComparer.OrdinalIgnoreCase)
            .ThenBy(t => t.Signo)
            .ToList();
        var actual = obtenidos
            .OrderBy(t => t.Hoja, StringComparer.OrdinalIgnoreCase)
            .ThenBy(t => t.Referencia, StringComparer.OrdinalIgnoreCase)
            .ThenBy(t => t.Signo)
            .ToList();

        if (!esperado.SequenceEqual(actual))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: {identidadTexto}: la fórmula no tiene la forma que su geometría exige ({rol}; celda hallada {referencia}). " +
                $"Esperado por forma: {Describir(esperado)}; obtenido: {Describir(actual)}. " +
                "La forma se deriva de las etiquetas y de las fórmulas que la plantilla declara; no se compara una lista de operandos congelada (§2-D2.3-bis).");
        }
    }

    private static string Describir(IReadOnlyList<Termino> terminos) =>
        terminos.Count == 0
            ? "(ninguno)"
            : string.Join(" ", terminos.Select(t => $"{(t.Signo > 0 ? "+" : "−")}{t.Hoja}!{t.Referencia}"));

    // ── Utilidades ──────────────────────────────────────────────────────────────────────────────

    private static Cell Celda(Worksheet hoja, string referencia, string identidadTexto) =>
        ResolvedorDestinoValor.BuscarCelda(hoja, referencia)
        ?? throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"ERR-SIN-SLOT: {identidadTexto}: la celda que el rol resuelve ({referencia}) no existe en la plantilla. No se inventa: la plantilla debe declararla.");

    private static IReadOnlyList<Row> Filas(Worksheet worksheet) =>
        worksheet.Elements<SheetData>().FirstOrDefault()?.Elements<Row>()
            .Where(r => r.RowIndex is not null)
            .OrderBy(r => r.RowIndex!.Value)
            .ToList()
        ?? [];

    private static Worksheet HojaFisica(WorkbookPart parte, string nombre)
    {
        ArgumentNullException.ThrowIfNull(parte);
        var hoja = parte.Workbook?.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: el workbook no trae la hoja '{nombre}' que exige la verificación de las hojas de detalle del período.");
        return ((WorksheetPart)parte.GetPartById(hoja.Id!)).Worksheet
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la hoja '{nombre}' no tiene WorksheetPart válido.");
    }

    /// <summary>
    /// Normaliza un rótulo para compararlo: sin acentos, sin espacios redundantes, sin saltos de línea
    /// y en mayúsculas. Es normalización de TEXTO: la plantilla escribe el mismo rótulo con distinto
    /// espaciado entre hojas (p. ej. 'Tarifa de Comercialización del Servicio.  ( TCS)'), y eso no
    /// cambia el concepto al que se refiere.
    /// </summary>
    internal static string NormalizarRotulo(string? valor)
    {
        if (string.IsNullOrWhiteSpace(valor))
        {
            return string.Empty;
        }

        var descompuesto = valor.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder(descompuesto.Length);
        foreach (var caracter in descompuesto)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(caracter) == UnicodeCategory.NonSpacingMark)
            {
                continue;
            }

            // Solo letras y dígitos: la plantilla decora el mismo concepto con '(*)' en una hoja y sin
            // él en otra, o con doble espacio; la identidad del concepto es el texto alfanumérico.
            if (!char.IsLetterOrDigit(caracter))
            {
                continue;
            }

            sb.Append(char.ToUpperInvariant(caracter));
        }

        return sb.ToString();
    }
}
