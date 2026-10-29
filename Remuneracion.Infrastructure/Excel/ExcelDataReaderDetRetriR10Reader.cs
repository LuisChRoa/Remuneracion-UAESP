using System.Globalization;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// HU-20 (G3) + Plan 62 (H1): lector read-only del insumo <c>R10_Remuneracion_{AAAAMMQ}.xlsx</c>
/// (hojas <c>DetRetri{AAAAMMQ}</c> y <c>DetValiRetri{AAAAMMQ}</c>).
///
/// <para>
/// T0c (evidencia en disco): la hoja trae el encabezado de la tabla en la fila 8
/// (<c>B8=ASE</c>, <c>C8=CONCESIONARIOS</c>); las filas 9..13 son ASE 1..5 y la 14 es
/// <c>Total</c>. Los montos están almacenados como TEXTO con formato monetario es-CO (p. ej.
/// <c>"$17.450.250.543"</c>, negativo <c>"-$5.093.259.797"</c>), por eso se convierten con
/// <see cref="ParseoMonetarioR10"/> (fuente única, T1.2). El rango del período vive en <c>G7</c>
/// ("Fecha Desde") y <c>J7</c> ("Fecha Hasta"); el sello de proceso en <c>D6</c>/<c>D7</c>.
/// </para>
///
/// <para>
/// La fila de encabezado se localiza DINÁMICAMENTE por sus etiquetas (nunca por número fijo) y las
/// hojas se resuelven con <see cref="NombresHojaPeriodo"/> (nunca enumerando el workbook).
/// Fail-fast: si la hoja, el encabezado, un ASE del 1..5, la fila Total o —en el bloque— una celda
/// del contrato faltan o son ilegibles, lanza la excepción del catálogo que nombra archivo, hoja,
/// período y celda — nunca 0 silencioso ni bloque parcial. El reader NUNCA escribe.
/// </para>
///
/// <para>
/// <b>Predicado de encabezado:</b> <c>B = ASE</c> y <c>C = CONCESIONARIOS</c>. El predicado
/// histórico exigía además <c>D = "RECAUDO TOTAL…"</c>, pero esa etiqueta solo existe en
/// <c>DetRetri</c>: en <c>DetValiRetri</c> la columna D es <c>TARIFA DE COMERCIALIZACIÓN…</c>
/// (medido en <c>Docs/T0-Plan62-Pegado-R10.md</c> §1, T0-04). Un predicado común a AMBAS hojas es
/// requisito del bloque; para <c>DetRetri</c> es el mismo hecho (fila 8), así que el oráculo de julio
/// queda byte-idéntico.
/// </para>
/// </summary>
public sealed class ExcelDataReaderDetRetriR10Reader : IDetRetriR10Reader
{
    /// <summary>Primera fila del bloque del contrato T0 (A3:O19 / A3:O14).</summary>
    private const int FilaInicioBloque = 3;

    /// <summary>Fila del encabezado de la tabla (T0-02/T0-04: fila 8 en ambas hojas).</summary>
    private const int FilaEncabezadoTabla = 8;

    /// <summary>Última fila del bloque de <c>DetRetri</c> (incluye el pie 15–19).</summary>
    private const int FilaFinBloqueDetRetri = 19;

    /// <summary>Última fila del bloque de <c>DetValiRetri</c> (cierra en la fila del Total).</summary>
    private const int FilaFinBloqueDetValiRetri = 14;

    /// <summary>Primer índice de columna del bloque (A).</summary>
    private const int ColumnaInicioBloque = 0;

    /// <summary>Último índice de columna del bloque (O).</summary>
    private const int ColumnaFinBloque = 14;

    /// <summary>Filas de ASE dentro del bloque (B9:B13).</summary>
    private const int FilasAse = 5;

    /// <summary>Fila del cierre <c>Total</c> (B14/C14 = D9:D13 Σ).</summary>
    private const int FilaTotal = FilaEncabezadoTabla + FilasAse + 1;

    /// <summary>Fila del pie del bloque de <c>DetRetri</c> (T0-04: <c>C16</c> "* Costos de Interventoría…").</summary>
    private const int FilaPieInterventoria = 16;

    /// <inheritdoc />
    public DetRetriInputs LeerDetRetri(Periodo periodo, string rutaR10)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(rutaR10);

        var nombreArchivo = Path.GetFileName(rutaR10);
        var hoja = NombresHojaPeriodo.DetRetri(periodo);
        var filas = ExcelWorksheetNavigator.LeerFilas(rutaR10, hoja);
        if (filas.Count == 0)
        {
            throw FormatoInvalido(
                $"R10 {nombreArchivo}: la hoja '{hoja}' está vacía; no se puede leer el DetRetri del período.");
        }

        var indiceHeader = filas.FindIndex(EsFilaEncabezadoDeTabla);
        if (indiceHeader < 0)
        {
            throw FormatoInvalido(
                $"R10 {nombreArchivo}: no se encontró el encabezado 'ASE / CONCESIONARIOS' en la hoja '{hoja}' (T0c: fila {FilaEncabezadoTabla}).");
        }

        var porAse = new Dictionary<int, decimal>();
        decimal? total = null;
        for (var i = indiceHeader + 1; i < filas.Count; i++)
        {
            var fila = filas[i];
            var aseTexto = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(1));
            var concepto = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(2));

            if (int.TryParse(aseTexto, NumberStyles.Integer, CultureInfo.InvariantCulture, out var aseId)
                && aseId is >= 1 and <= 5)
            {
                porAse[aseId] = ParseoMonetarioR10.Convertir(fila.ElementAtOrDefault(3));
                continue;
            }

            if (concepto.Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                total = ParseoMonetarioR10.Convertir(fila.ElementAtOrDefault(3));
                break;
            }
        }

        var faltantes = Enumerable.Range(1, 5).Where(id => !porAse.ContainsKey(id)).ToArray();
        if (faltantes.Length > 0)
        {
            throw FormatoInvalido(
                $"R10 {nombreArchivo}: la hoja '{hoja}' no trae el DetRetri de los ASE {string.Join(", ", faltantes)} en la columna D (T0c: D9:D13).");
        }

        if (total is null)
        {
            throw FormatoInvalido(
                $"R10 {nombreArchivo}: la hoja '{hoja}' no trae la fila 'Total' del DetRetri (T0c: D14).");
        }

        return new DetRetriInputs
        {
            CodigoRemuneracion = LeerCodigoRemuneracion(filas, periodo),
            FechaDesde = BuscarFecha(filas, "Fecha Desde", columnaEtiqueta: 5, columnaValor: 6) ?? default,
            FechaHasta = BuscarFecha(filas, "Fecha Hasta", columnaEtiqueta: 8, columnaValor: 9) ?? default,
            FechaProceso = BuscarFecha(filas, "Fecha de Proceso", columnaEtiqueta: 2, columnaValor: 3),
            HoraProceso = BuscarHora(filas, "Hora", columnaEtiqueta: 2, columnaValor: 3),
            DetRetriPorAse = porAse,
            Total = total.Value
        };
    }

    /// <inheritdoc />
    public BloqueHojaR10 LeerBloqueDetRetri(Periodo periodo, string rutaR10) =>
        LeerBloque(periodo, rutaR10, NombresHojaPeriodo.DetRetri(periodo), FilaFinBloqueDetRetri);

    /// <inheritdoc />
    public BloqueHojaR10 LeerBloqueDetValiRetri(Periodo periodo, string rutaR10) =>
        LeerBloque(periodo, rutaR10, NombresHojaPeriodo.DetValiRetri(periodo), FilaFinBloqueDetValiRetri);

    /// <summary>
    /// Lee el bloque completo del rango del contrato (<c>A3:O{filaFin}</c>) de la hoja indicada.
    ///
    /// <para>
    /// Todo lo que puede impedir el pegado se valida ANTES de devolver: el encabezado de la tabla en
    /// su fila del contrato, los 5 ASE en orden, el cierre <c>Total</c> (case-insensitive: el rótulo
    /// real de los cierres del negocio alterna mayúsculas — Plan 62 §11 R-3), el pie del bloque
    /// cuando el rango lo incluye y CADA monto del bloque legible. Si algo falla, no hay bloque
    /// parcial: el error nombra archivo, hoja, período y celda.
    /// </para>
    /// </summary>
    private static BloqueHojaR10 LeerBloque(Periodo periodo, string rutaR10, string hoja, int filaFin)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(rutaR10);

        var nombreArchivo = Path.GetFileName(rutaR10);
        var periodoTexto = periodo.CodigoCompleto;
        var filas = ExcelWorksheetNavigator.LeerFilas(rutaR10, hoja);
        if (filas.Count == 0)
        {
            throw FormatoInvalido(
                $"R10 {nombreArchivo}: la hoja '{hoja}' está vacía; no se puede copiar el bloque del período {periodoTexto}.");
        }

        var indiceEncabezado = filas.FindIndex(EsFilaEncabezadoDeTabla);
        if (indiceEncabezado < 0)
        {
            throw FormatoInvalido(
                $"R10 {nombreArchivo}: la hoja '{hoja}' no trae el encabezado 'ASE / CONCESIONARIOS' del período {periodoTexto}; no se copia el bloque y no se genera salida.");
        }

        if (indiceEncabezado + 1 != FilaEncabezadoTabla)
        {
            throw FormatoInvalido(
                $"R10 {nombreArchivo}: en la hoja '{hoja}' el encabezado 'ASE / CONCESIONARIOS' está en la fila {indiceEncabezado + 1} y el bloque del período {periodoTexto} lo exige en la fila {FilaEncabezadoTabla}; no se copia el bloque y no se genera salida.");
        }

        if (filas.Count < filaFin)
        {
            throw FormatoInvalido(
                $"R10 {nombreArchivo}: la hoja '{hoja}' termina en la fila {filas.Count} y el bloque del período {periodoTexto} llega hasta la fila {filaFin}; no se copia el bloque y no se genera salida.");
        }

        var encabezado = filas[FilaEncabezadoTabla - 1];
        var columnasMonto = ClasificarColumnas(encabezado, periodoTexto, hoja, nombreArchivo, out var columnaAse, out var columnaRotulo);

        var ids = new HashSet<int>();
        for (var offset = 1; offset <= FilasAse; offset++)
        {
            var fila = FilaEncabezadoTabla + offset;
            var celda = $"{LetraColumna(columnaAse)}{fila}";
            var aseTexto = ExcelWorksheetNavigator.CeldaTexto(filas[fila - 1].ElementAtOrDefault(columnaAse));
            if (!int.TryParse(aseTexto, NumberStyles.Integer, CultureInfo.InvariantCulture, out var aseId)
                || aseId is < 1 or > 5
                || !ids.Add(aseId))
            {
                throw FormatoInvalido(
                    $"R10 {nombreArchivo}: la celda {celda} de la hoja '{hoja}' dice '{aseTexto}' y el bloque del período {periodoTexto} exige el ASE {offset} de 1 a 5, una sola vez; no se copia nada y no se genera salida.");
            }
        }

        var rotuloTotal = ExcelWorksheetNavigator.CeldaTexto(filas[FilaTotal - 1].ElementAtOrDefault(columnaRotulo));
        if (!rotuloTotal.Equals("Total", StringComparison.OrdinalIgnoreCase))
        {
            throw FormatoInvalido(
                $"R10 {nombreArchivo}: la celda {LetraColumna(columnaRotulo)}{FilaTotal} de la hoja '{hoja}' dice '{rotuloTotal}' y el bloque del período {periodoTexto} exige el cierre 'Total'; no se copia nada y no se genera salida.");
        }

        if (filaFin >= FilaPieInterventoria)
        {
            var pie = ExcelWorksheetNavigator.CeldaTexto(filas[FilaPieInterventoria - 1].ElementAtOrDefault(columnaRotulo));
            if (string.IsNullOrWhiteSpace(pie))
            {
                throw FormatoInvalido(
                    $"R10 {nombreArchivo}: la celda {LetraColumna(columnaRotulo)}{FilaPieInterventoria} de la hoja '{hoja}' no trae el pie del bloque ('* Costos de Interventoría…') del período {periodoTexto}; no se copia el bloque y no se genera salida.");
            }
        }

        var celdas = new List<CeldaBloqueR10>(
            (filaFin - FilaInicioBloque + 1) * (ColumnaFinBloque - ColumnaInicioBloque + 1));

        for (var fila = FilaInicioBloque; fila <= filaFin; fila++)
        {
            var valores = filas[fila - 1];
            var esFilaDeMontos = fila > FilaEncabezadoTabla && fila <= FilaTotal;

            for (var indice = ColumnaInicioBloque; indice <= ColumnaFinBloque; indice++)
            {
                celdas.Add(CrearCelda(
                    fila,
                    indice,
                    valores.ElementAtOrDefault(indice),
                    esFilaDeMontos && columnasMonto.Contains(indice),
                    periodoTexto,
                    hoja,
                    nombreArchivo));
            }
        }

        return new BloqueHojaR10(hoja, periodo, celdas);
    }

    /// <summary>
    /// Clasifica las columnas del bloque por la ETIQUETA del encabezado (nunca por letra física):
    /// la columna <c>ASE</c> (id del ASE), la columna <c>CONCESIONARIOS</c> (rótulo) y las columnas
    /// de monto (toda etiqueta restante del rango). La columna sin etiqueta del contrato (la primera)
    /// no participa: el bloque la transporta como celda libre (T0-04: vacía en R10 y plantilla).
    /// </summary>
    private static HashSet<int> ClasificarColumnas(
        object?[] encabezado,
        string periodoTexto,
        string hoja,
        string nombreArchivo,
        out int columnaAse,
        out int columnaRotulo)
    {
        columnaAse = -1;
        columnaRotulo = -1;
        var columnasMonto = new HashSet<int>();

        for (var indice = ColumnaInicioBloque; indice <= ColumnaFinBloque; indice++)
        {
            var etiqueta = ExcelWorksheetNavigator.CeldaTexto(encabezado.ElementAtOrDefault(indice));

            if (indice == ColumnaInicioBloque)
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(etiqueta))
            {
                throw FormatoInvalido(
                    $"R10 {nombreArchivo}: la hoja '{hoja}' no trae la etiqueta de la columna {LetraColumna(indice)} en la fila {FilaEncabezadoTabla} del bloque del período {periodoTexto}; no se copia el bloque y no se genera salida.");
            }

            var normalizada = Normalizar(etiqueta);
            if (normalizada.Equals("ase", StringComparison.Ordinal))
            {
                columnaAse = indice;
            }
            else if (normalizada.Equals("concesionarios", StringComparison.Ordinal))
            {
                columnaRotulo = indice;
            }
            else
            {
                columnasMonto.Add(indice);
            }
        }

        if (columnaAse < 0 || columnaRotulo < 0 || columnasMonto.Count == 0)
        {
            throw FormatoInvalido(
                $"R10 {nombreArchivo}: la hoja '{hoja}' no declara las columnas 'ASE', 'CONCESIONARIOS' y los montos del bloque del período {periodoTexto}; no se copia el bloque y no se genera salida.");
        }

        return columnasMonto;
    }

    /// <summary>
    /// Representa UNA celda del rango con su forma exacta: texto (rótulos, sello, fechas y horas),
    /// numérico (montos y celdas numéricas, SIN redondear) o vacía (presente y explícita). Un monto
    /// vacío o ilegible en el bloque es fail-fast: el pegado no puede inventar un valor.
    /// </summary>
    private static CeldaBloqueR10 CrearCelda(
        int fila,
        int indiceColumna,
        object? valor,
        bool esMonto,
        string periodoTexto,
        string hoja,
        string nombreArchivo)
    {
        var letra = LetraColumna(indiceColumna);
        var referencia = $"{letra}{fila}";
        var esVacia = valor is null
            || (valor is string textoCrudo && string.IsNullOrWhiteSpace(textoCrudo));

        if (esVacia)
        {
            if (esMonto)
            {
                throw FormatoInvalido(
                    $"R10 {nombreArchivo}: la celda {referencia} de la hoja '{hoja}' está vacía y el bloque del período {periodoTexto} exige el monto; no se copia nada y no se genera salida.");
            }

            return new CeldaBloqueR10(fila, letra, null, null);
        }

        if (esMonto)
        {
            if (!ParseoMonetarioR10.TryConvertir(valor, out var monto))
            {
                throw FormatoInvalido(
                    $"R10 {nombreArchivo}: la celda {referencia} de la hoja '{hoja}' trae '{ExcelWorksheetNavigator.CeldaTexto(valor)}', que no es un monto legible del período {periodoTexto}; no se copia nada y no se genera salida.");
            }

            return new CeldaBloqueR10(fila, letra, null, monto);
        }

        if (valor is string texto)
        {
            return new CeldaBloqueR10(fila, letra, texto, null);
        }

        var esNumero = valor is decimal or double or float or int or long or short or byte;
        if (esNumero && ParseoMonetarioR10.TryConvertir(valor, out var numero))
        {
            return new CeldaBloqueR10(fila, letra, null, numero);
        }

        throw FormatoInvalido(
            $"R10 {nombreArchivo}: la celda {referencia} de la hoja '{hoja}' trae un valor que no se puede copiar tal cual (fecha u hora con formato de Excel) y el bloque del período {periodoTexto} lo exige como texto; no se copia nada y no se genera salida.");
    }

    /// <summary>Letra de la columna a partir de su índice 0-based (0 → A, 14 → O).</summary>
    private static string LetraColumna(int indice)
    {
        var letras = string.Empty;
        var numero = indice + 1;
        while (numero > 0)
        {
            var resto = (numero - 1) % 26;
            letras = (char)('A' + resto) + letras;
            numero = (numero - 1) / 26;
        }

        return letras;
    }

    /// <summary>
    /// Predicado ÚNICO del encabezado de la tabla (lo usan el oráculo y el lector de bloque):
    /// <c>ASE</c> en la segunda columna y <c>CONCESIONARIOS</c> en la tercera, case-insensitive.
    /// </summary>
    private static bool EsFilaEncabezadoDeTabla(object?[] fila) =>
        ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(1)).Equals("ASE", StringComparison.OrdinalIgnoreCase)
        && ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(2)).Equals("CONCESIONARIOS", StringComparison.OrdinalIgnoreCase);

    private static CalculoInvalidoException FormatoInvalido(string mensaje) =>
        new(CodigoError.FormatoFuente, mensaje);

    /// <summary>
    /// Código de remuneración declarado en la hoja (M3 = "2026071"); si no se puede leer, cae al
    /// código del período (nunca falla por un dato informativo).
    /// </summary>
    private static string LeerCodigoRemuneracion(List<object?[]> filas, Periodo periodo)
    {
        foreach (var fila in filas)
        {
            var etiqueta = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(11));
            if (Normalizar(etiqueta).StartsWith("codigoremuneracion", StringComparison.Ordinal))
            {
                var valor = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(12));
                if (!string.IsNullOrWhiteSpace(valor))
                {
                    return valor;
                }
            }
        }

        return periodo.CodigoCompleto;
    }

    /// <summary>
    /// Busca la celda de fecha cuya etiqueta (columna <paramref name="columnaEtiqueta"/>) sea igual
    /// al texto normalizado esperado y devuelve el valor de la columna contigua
    /// (<paramref name="columnaValor"/>). Devuelve <c>null</c> si no existe/no parsea.
    /// </summary>
    private static DateTime? BuscarFecha(List<object?[]> filas, string etiquetaEsperada, int columnaEtiqueta, int columnaValor)
    {
        var esperado = Normalizar(etiquetaEsperada);
        foreach (var fila in filas)
        {
            if (!Normalizar(ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(columnaEtiqueta)))
                .Equals(esperado, StringComparison.Ordinal))
            {
                continue;
            }

            return ParsearFecha(fila.ElementAtOrDefault(columnaValor));
        }

        return null;
    }

    private static DateTime? ParsearFecha(object? valor)
    {
        if (valor is DateTime fecha)
        {
            return fecha;
        }

        if (valor is double oa)
        {
            return DateTime.FromOADate(oa);
        }

        var texto = ExcelWorksheetNavigator.CeldaTexto(valor);
        if (string.IsNullOrWhiteSpace(texto))
        {
            return null;
        }

        var formatos = new[] { "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy" };
        return DateTime.TryParseExact(texto, formatos, CultureInfo.GetCultureInfo("es-CO"), DateTimeStyles.None, out var exacta)
            ? exacta
            : DateTime.TryParse(texto, CultureInfo.GetCultureInfo("es-CO"), DateTimeStyles.None, out var general) ? general : null;
    }

    /// <summary>
    /// Plan 31 (T3, R-S-1): busca la celda de hora cuya etiqueta sea igual al texto normalizado
    /// esperado ("Hora") y devuelve la hora-del-día de la columna contigua. El R10 la almacena como
    /// TEXTO en formato 12h (<c>"07:42 AM"</c>). Devuelve <c>null</c> si no existe/no parsea.
    /// </summary>
    private static TimeSpan? BuscarHora(List<object?[]> filas, string etiquetaEsperada, int columnaEtiqueta, int columnaValor)
    {
        var esperado = Normalizar(etiquetaEsperada);
        foreach (var fila in filas)
        {
            if (!Normalizar(ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(columnaEtiqueta)))
                .Equals(esperado, StringComparison.Ordinal))
            {
                continue;
            }

            return ParsearHora(fila.ElementAtOrDefault(columnaValor));
        }

        return null;
    }

    private static TimeSpan? ParsearHora(object? valor)
    {
        if (valor is TimeSpan hora)
        {
            return hora;
        }

        if (valor is DateTime fecha)
        {
            return fecha.TimeOfDay;
        }

        if (valor is double oa)
        {
            // Fracción de día de Excel (0 = 00:00, 0.5 = 12:00).
            return TimeSpan.FromDays(oa);
        }

        var texto = ExcelWorksheetNavigator.CeldaTexto(valor)?.Trim();
        if (string.IsNullOrWhiteSpace(texto))
        {
            return null;
        }

        var formatos = new[] { "hh:mm tt", "h:mm tt", "hh:mm:ss tt", "HH:mm", "H:mm", "HH:mm:ss" };
        if (DateTime.TryParseExact(texto, formatos, CultureInfo.InvariantCulture, DateTimeStyles.None, out var exacta))
        {
            return exacta.TimeOfDay;
        }

        return DateTime.TryParse(texto, CultureInfo.InvariantCulture, DateTimeStyles.None, out var general)
            ? general.TimeOfDay
            : null;
    }

    /// <summary>Normaliza una etiqueta (minúsculas, solo letras/dígitos) para el match por título.</summary>
    private static string Normalizar(string texto)
    {
        var sb = new System.Text.StringBuilder(texto.Length);
        foreach (var ch in texto)
        {
            if (char.IsLetterOrDigit(ch))
            {
                sb.Append(char.ToLowerInvariant(ch));
            }
        }

        return sb.ToString();
    }
}
