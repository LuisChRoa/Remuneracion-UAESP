using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 61 (D9 ratificada — H1.6): RESOLUCIÓN CONCEPTUAL del destino de las columnas del
/// desglose-detalle de <c>Rem. Anticipos R2</c> y <c>Reversion Pagos R4</c>.
///
/// <para><b>El defecto que cierra.</b> La escritura del detalle resolvía la columna destino contra
/// un mapa congelado de 13 pares <c>encabezado → letra</c> (<c>ColumnasR4</c>). La plantilla declara
/// 15 columnas de valor en ese bloque: <c>SERVICIO ESPECIALES</c> y <c>Ajuste Recargo Aseo</c> no
/// figuraban en el mapa, así que un componente con monto y sin entrada <b>se descartaba en
/// silencio</b> (6 de 30 fuentes R4 del corpus traen la 14.ª/15.ª columna). El verificador
/// (<c>VerificarDetalleR2R4</c>) sí derivaba las columnas de la hoja y por eso lo cazaba: el writer y
/// el gate no compartían la misma autoridad.</para>
///
/// <para><b>La regla.</b> La plantilla no es una rejilla de letras: es una jerarquía declarada. La
/// identidad de un valor es (hoja · bloque de ASE · sección · subsección · fuente/canal · concepto ·
/// componente) y la columna destino es la que la PROPIA fila de encabezados del bloque declara para
/// ese componente. Se resuelve por rótulo normalizado (<c>Trim</c> + <c>OrdinalIgnoreCase</c>) con
/// UNA sola tabla de alias (<see cref="ResolvedorConciliacionEmpresa.EncabezadosEspeciales"/>), nunca
/// por posición ni por conteo. Una columna nueva/movida en la plantilla se absorbe sin código nuevo;
/// un encabezado ambiguo (dos columnas con la misma identidad) es fail-fast.</para>
///
/// <para><b>Quién lo consume.</b> El escritor (<c>OpenXmlPlantillaWriter.EscribirDetalleR2</c> /
/// <c>EscribirDetalleR4</c>) y el gate de valores (<c>VerificarDetalleR2R4</c>): los tres derivan de
/// aquí, de modo que la autoridad del destino es única y una pérdida deja de ser invisible.</para>
/// </summary>
internal static class ResolvedorDestinoDetalleR2R4
{
    /// <summary>Encabezado de columna de valor tal como la fila de encabezados lo declara, con su letra.</summary>
    internal sealed record EncabezadoDestino(string Etiqueta, string Columna);

    /// <summary>
    /// Encabezados de las columnas de VALOR del bloque, leídos de la fila de encabezados que la
    /// plantilla declara (la que <see cref="DelimitadorBloquesHoja"/> ubica por forma: rótulos sin
    /// valores numéricos). Cero letras de columna escritas en el código: la letra sale de la
    /// referencia A1 de cada celda.
    /// </summary>
    /// <exception cref="CalculoInvalidoException"><c>ERR-PLANTILLA</c>: el bloque no declara su fila
    /// de encabezados o esa fila no trae rótulos.</exception>
    internal static IReadOnlyList<EncabezadoDestino> ColumnasDeclaradas(
        WorkbookPart workbookPart,
        Worksheet worksheet,
        int filaEncabezados,
        string identidad)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);
        ArgumentNullException.ThrowIfNull(worksheet);

        var fila = worksheet.Elements<SheetData>()
            .SelectMany(tabla => tabla.Elements<Row>())
            .FirstOrDefault(f => f.RowIndex?.Value == filaEncabezados)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La {identidad} no trae la fila de encabezados ({filaEncabezados}): sin los rótulos que la " +
                "plantilla declara no se puede resolver el destino de ningún componente.");

        var columnas = new List<EncabezadoDestino>();
        foreach (var celda in fila.Elements<Cell>())
        {
            var etiqueta = LectorTextoHoja.Texto(workbookPart, celda).Trim();
            if (etiqueta.Length == 0)
            {
                continue;
            }

            var columna = ColumnaDeReferencia(celda.CellReference?.Value);
            if (columna.Length > 0)
            {
                columnas.Add(new EncabezadoDestino(etiqueta, columna));
            }
        }

        if (columnas.Count == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La fila de encabezados ({filaEncabezados}) de la {identidad} no declara rótulos de columna: " +
                "no hay destino conceptual que resolver para el desglose-detalle por componente.");
        }

        return columnas;
    }

    /// <summary>
    /// Índice <c>encabezado declarado → columna</c> con la normalización del proyecto
    /// (<c>Trim</c> + <c>OrdinalIgnoreCase</c>). Dos columnas que declaren la MISMA identidad —o dos
    /// miembros de la misma clase de alias, p. ej. <c>Especiales</c> y <c>SERVICIO ESPECIALES</c> en
    /// el mismo bloque— son una plantilla ambigua y se fallan: nunca se elige una por orden.
    /// </summary>
    /// <exception cref="CalculoInvalidoException"><c>ERR-PLANTILLA</c> si la identidad no es única.</exception>
    internal static IReadOnlyDictionary<string, string> IndicePorEtiqueta(
        IReadOnlyList<EncabezadoDestino> columnas,
        string identidad)
    {
        ArgumentNullException.ThrowIfNull(columnas);

        var indice = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (etiqueta, columna) in columnas)
        {
            var clave = etiqueta.Trim();
            if (indice.TryGetValue(clave, out var previa)
                && !string.Equals(previa, columna, StringComparison.OrdinalIgnoreCase))
            {
                throw Ambigua(identidad, clave, previa, columna);
            }

            indice[clave] = columna;
        }

        // La clase de alias de Especiales resuelve a UNA sola columna: si el bloque declarara dos
        // miembros de la clase en columnas distintas, el destino del componente sería ambiguo.
        var deLaClase = new List<(string Etiqueta, string Columna)>();
        foreach (var miembro in ResolvedorConciliacionEmpresa.EncabezadosEspeciales)
        {
            if (indice.TryGetValue(miembro.Trim(), out var columna)
                && !deLaClase.Any(c => string.Equals(c.Columna, columna, StringComparison.OrdinalIgnoreCase)))
            {
                deLaClase.Add((miembro, columna));
            }
        }

        if (deLaClase.Count > 1)
        {
            throw Ambigua(identidad, deLaClase[0].Etiqueta, deLaClase[0].Columna, deLaClase[1].Columna);
        }

        return indice;
    }

    /// <summary>
    /// Columna destino de un componente de la FUENTE: coincidencia por rótulo normalizado y, si el
    /// encabezado pertenece a la clase de alias de Especiales, resolución contra esa misma clase.
    /// <c>null</c> = la plantilla no declara destino para ese componente (el consumidor decide: fallo
    /// duro si el valor no es cero, cero documentado si lo es).
    /// </summary>
    internal static string? ColumnaDe(IReadOnlyDictionary<string, string> indice, string? encabezadoFuente)
    {
        ArgumentNullException.ThrowIfNull(indice);

        var etiqueta = encabezadoFuente?.Trim() ?? string.Empty;
        if (etiqueta.Length == 0)
        {
            return null;
        }

        if (indice.TryGetValue(etiqueta, out var columna))
        {
            return columna;
        }

        if (!PerteneceALaClaseDeAlias(etiqueta))
        {
            return null;
        }

        foreach (var alias in ResolvedorConciliacionEmpresa.EncabezadosEspeciales)
        {
            if (indice.TryGetValue(alias.Trim(), out var deLaClase))
            {
                return deLaClase;
            }
        }

        return null;
    }

    /// <summary>Verdad si el rótulo es un miembro de la (única) clase de alias del proyecto.</summary>
    internal static bool PerteneceALaClaseDeAlias(string? encabezado)
    {
        var etiqueta = encabezado?.Trim() ?? string.Empty;
        return etiqueta.Length > 0
            && ResolvedorConciliacionEmpresa.EncabezadosEspeciales.Any(
                alias => string.Equals(alias.Trim(), etiqueta, StringComparison.OrdinalIgnoreCase));
    }

    private static CalculoInvalidoException Ambigua(string identidad, string etiqueta, string primera, string segunda) =>
        new(
            CodigoError.Plantilla,
            $"ERR-PLANTILLA: la {identidad} declara dos columnas de valor para la misma identidad " +
            $"('{etiqueta}'): {primera} y {segunda}. El destino del componente sería ambiguo y no se elige " +
            "por posición: revise los encabezados de la plantilla.");

    /// <summary>Letras de columna de una referencia A1 (sin dígitos ni <c>$</c>).</summary>
    private static string ColumnaDeReferencia(string? referencia)
    {
        if (string.IsNullOrEmpty(referencia))
        {
            return string.Empty;
        }

        var fin = 0;
        while (fin < referencia.Length && (char.IsLetter(referencia[fin]) || referencia[fin] == '$'))
        {
            fin++;
        }

        return referencia[..fin].Replace("$", string.Empty, StringComparison.Ordinal);
    }
}
