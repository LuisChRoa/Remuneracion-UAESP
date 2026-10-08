using System.Globalization;
using System.Text;

namespace Remuneracion.Core.Services;

/// <summary>
/// Plan 32 (T2-interior, RONDA A — R-B-1/D-B): compositor PURO (sin I/O, sin OpenXML, sin Excel) del
/// INTERIOR de <c>Reporte Componentes R1</c> (sub-visibles por empresa).
///
/// Dado el mapa <paramref name="filasPorFirma"/> (etiquetas de firma del sub-bloque -> número de fila
/// real del período), reproduce el TEXTO de <c>&lt;f&gt;</c> que el manual del administrativo trae para
/// la celda destino. Devuelve <c>null</c> cuando la celda va como LITERAL (clase L-1: el manual no trae
/// <c>&lt;f&gt;</c>) o cuando falta un ancla del sub-bloque (el compositor NUNCA inventa filas); en ese
/// segundo caso <see cref="AnclasFaltantes(string, IReadOnlyDictionary{string, int}, IReadOnlySet{string}?)"/>
/// nombra las etiquetas ausentes para que el caller decida el fail-fast.
///
/// Etiquetas reconocidas en <c>filasPorFirma</c>:
///   - <c>Mes0..MesN</c>: filas-dato Mes/empresa del sub-bloque (orden de emisión de los términos F).
///   - <c>Aplic0..AplicM</c>: filas-dato Aplicacion del EXTEMP-interior (orden de emisión).
///   - <c>Ancla</c>: fila del single-ref (SUBS/AFASEO).
///   - <c>QuirkDobleSignoL</c>: etiqueta reservada (valor ignorado) que reproduce el quirk textual
///     <c>--L</c> del template base (H-DOBLE-SIGNO, p. ej. <c>F204</c>); el disco manda, no se normaliza.
///
/// Reglas de composición (el manual es la fuente de verdad; formas verificadas contra el fixture
/// <c>InterioresR1Esperados</c>):
///   - SUB_EMP (col F, Mes): <c>ΣF(Mes) − ΣL(Mes excepto la de MAYOR número de fila)</c>, F en orden de
///     etiqueta y L ascendente por fila.
///   - SUB_TDF (col G/H, Mes): G/H de las Mes excepto la de mayor fila, emitidas descendente.
///   - SUB_L (col L, Mes): igual que SUB_TDF pero en columna L.
///   - EXT_INT (col F, Aplic): <c>ΣF(Aplic) − L(última Aplic por orden de etiqueta)</c>; 0 Aplic -> null.
///   - EXT_INT (col G/H, Aplic): single-ref G/H de la última Aplic (espejo del ancla restada).
///   - SUBS / AFASEO (Ancla): single-ref <c>col + Ancla</c>.
/// </summary>
public static class CompositorInteriorR1
{
    /// <summary>Prefijo de etiqueta de fila-dato Mes/empresa del sub-bloque.</summary>
    public const string EtiquetaMes = "Mes";

    /// <summary>Prefijo de etiqueta de fila-dato Aplicacion (EXTEMP-interior).</summary>
    public const string EtiquetaAplic = "Aplic";

    /// <summary>Etiqueta del ancla single-ref (SUBS/AFASEO).</summary>
    public const string EtiquetaAncla = "Ancla";

    /// <summary>
    /// Etiqueta reservada que reproduce el quirk textual <c>--L</c> del template base (H-DOBLE-SIGNO).
    /// Su valor en el mapa se ignora.
    /// </summary>
    public const string EtiquetaQuirkDobleSignoL = "QuirkDobleSignoL";

    private enum Rol
    {
        Ninguno,
        Single,
        ExtInt,
        SubEmp,
        SubTdf,
        SubL
    }

    /// <summary>
    /// Compone el texto de <c>&lt;f&gt;</c> de la celda interior <paramref name="celdaDestino"/> a partir
    /// de las filas reales del sub-bloque. Devuelve <c>null</c> si el rol no se puede inferir (p. ej.
    /// mapa vacío = clase L-1 literal) o si falta un ancla requerida (nunca inventa filas).
    /// </summary>
    /// <param name="aseId">ASE (1..5); solo trazabilidad/validación.</param>
    /// <param name="celdaDestino">Referencia destino (p. ej. <c>F60</c>, <c>G60</c>, <c>L569</c>).</param>
    /// <param name="filasPorFirma">Etiquetas de firma -> número de fila real.</param>
    /// <param name="filasAplicInexistentes">
    /// Anclas Aplic conocidas-inexistentes del período (0-Aplic -> la celda va literal). Opcional.
    /// </param>
    public static string? ComponerInterior(
        int aseId,
        string celdaDestino,
        IReadOnlyDictionary<string, int> filasPorFirma,
        IReadOnlySet<string>? filasAplicInexistentes = null)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(aseId, 1);
        ArgumentNullException.ThrowIfNull(filasPorFirma);
        ArgumentException.ThrowIfNullOrWhiteSpace(celdaDestino);

        var columna = Columna(celdaDestino);
        var rol = InferirRol(columna, filasPorFirma);
        if (rol == Rol.Ninguno)
        {
            return null;
        }

        if (AnclasFaltantes(rol, filasPorFirma, filasAplicInexistentes).Count > 0)
        {
            return null;
        }

        return rol switch
        {
            Rol.Single => RefSimple(columna, Requerido(filasPorFirma, EtiquetaAncla)),
            Rol.ExtInt => ExtInt(columna, Aplic(filasPorFirma)),
            Rol.SubEmp => SubEmp(filasPorFirma),
            Rol.SubTdf or Rol.SubL => EspejoMes(columna, Mes(filasPorFirma)),
            _ => null
        };
    }

    /// <summary>
    /// Reporta las anclas requeridas por la etiqueta de la celda que NO están en el mapa (o están
    /// marcadas como inexistentes). Lista vacía = el compositor puede componer.
    /// </summary>
    /// <param name="etiqueta">SUB_EMP | SUB_TDF | SUB_L | EXT_INT | SUBS | AFASEO.</param>
    /// <param name="filasPorFirma">Etiquetas de firma -> número de fila real.</param>
    /// <param name="filasAplicInexistentes">Anclas Aplic conocidas-inexistentes del período. Opcional.</param>
    public static IReadOnlyList<string> AnclasFaltantes(
        string etiqueta,
        IReadOnlyDictionary<string, int> filasPorFirma,
        IReadOnlySet<string>? filasAplicInexistentes = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(etiqueta);
        ArgumentNullException.ThrowIfNull(filasPorFirma);

        return AnclasFaltantes(RolDeEtiqueta(etiqueta), filasPorFirma, filasAplicInexistentes);
    }

    private static IReadOnlyList<string> AnclasFaltantes(
        Rol rol,
        IReadOnlyDictionary<string, int> filasPorFirma,
        IReadOnlySet<string>? filasAplicInexistentes)
    {
        var faltantes = new List<string>();
        switch (rol)
        {
            case Rol.Single:
                if (!filasPorFirma.ContainsKey(EtiquetaAncla))
                {
                    faltantes.Add(EtiquetaAncla);
                }

                break;
            case Rol.ExtInt:
                if (!filasPorFirma.ContainsKey(EtiquetaAplic + "0")
                    || (filasAplicInexistentes?.Contains(EtiquetaAplic + "0") ?? false))
                {
                    faltantes.Add(EtiquetaAplic + "0");
                }

                break;
            case Rol.SubEmp:
            case Rol.SubTdf:
            case Rol.SubL:
                if (!filasPorFirma.ContainsKey(EtiquetaMes + "0"))
                {
                    faltantes.Add(EtiquetaMes + "0");
                }

                break;
        }

        return faltantes;
    }

    private static Rol InferirRol(string columna, IReadOnlyDictionary<string, int> filasPorFirma)
    {
        if (filasPorFirma.Keys.Any(k => k.StartsWith(EtiquetaAplic, StringComparison.Ordinal)))
        {
            return Rol.ExtInt;
        }

        if (filasPorFirma.ContainsKey(EtiquetaAncla))
        {
            return Rol.Single;
        }

        if (!filasPorFirma.Keys.Any(k => k.StartsWith(EtiquetaMes, StringComparison.Ordinal)))
        {
            return Rol.Ninguno;
        }

        return columna switch
        {
            "G" or "H" => Rol.SubTdf,
            "L" => Rol.SubL,
            _ => Rol.SubEmp
        };
    }

    private static Rol RolDeEtiqueta(string etiqueta) => etiqueta.ToUpperInvariant() switch
    {
        "SUBS" or "AFASEO" => Rol.Single,
        "EXT_INT" => Rol.ExtInt,
        "SUB_TDF" => Rol.SubTdf,
        "SUB_L" => Rol.SubL,
        _ => Rol.SubEmp
    };

    private static string SubEmp(IReadOnlyDictionary<string, int> filasPorFirma)
    {
        var mes = Mes(filasPorFirma);
        var mayor = mes.Max();
        var lFilas = mes.Where(f => f != mayor).OrderBy(f => f).ToList();
        var quirk = filasPorFirma.ContainsKey(EtiquetaQuirkDobleSignoL);

        var partes = new List<string>(mes.Count + lFilas.Count);
        foreach (var fila in mes)
        {
            partes.Add("+F" + Formatear(fila));
        }

        for (var i = 0; i < lFilas.Count; i++)
        {
            partes.Add((i == 0 && quirk ? "--L" : "-L") + Formatear(lFilas[i]));
        }

        return FormatearFormula(partes);
    }

    private static string EspejoMes(string columna, IReadOnlyList<int> mes)
    {
        // Mes menos la de mayor número de fila (n==1: la única), emitidas descendente (forma del manual).
        IEnumerable<int> filas = mes.Count > 1
            ? mes.Where(f => f != mes.Max()).OrderByDescending(f => f)
            : mes;

        return FormatearFormula(filas.Select(f => "+" + columna + Formatear(f)).ToList());
    }

    private static string ExtInt(string columna, IReadOnlyList<int> aplic)
    {
        // El espejo G/H del EXTEMP-interior es single-ref de la última Aplic (la que se resta en F).
        if (columna is "G" or "H")
        {
            return RefSimple(columna, aplic[^1]);
        }

        var partes = new List<string>(aplic.Count + 1);
        foreach (var fila in aplic)
        {
            partes.Add("+F" + Formatear(fila));
        }

        partes.Add("-L" + Formatear(aplic[^1]));
        return FormatearFormula(partes);
    }

    private static IReadOnlyList<int> Mes(IReadOnlyDictionary<string, int> filasPorFirma) =>
        OrdenadasPorIndice(filasPorFirma, EtiquetaMes);

    private static IReadOnlyList<int> Aplic(IReadOnlyDictionary<string, int> filasPorFirma) =>
        OrdenadasPorIndice(filasPorFirma, EtiquetaAplic);

    private static IReadOnlyList<int> OrdenadasPorIndice(
        IReadOnlyDictionary<string, int> filasPorFirma,
        string prefijo) =>
        filasPorFirma
            .Where(p => p.Key.StartsWith(prefijo, StringComparison.Ordinal)
                && int.TryParse(p.Key.AsSpan(prefijo.Length), NumberStyles.None, CultureInfo.InvariantCulture, out _))
            .Select(p => (
                Indice: int.Parse(p.Key.AsSpan(prefijo.Length), NumberStyles.None, CultureInfo.InvariantCulture),
                Fila: p.Value))
            .OrderBy(x => x.Indice)
            .Select(x => x.Fila)
            .ToList();

    private static int Requerido(IReadOnlyDictionary<string, int> filasPorFirma, string etiqueta) =>
        filasPorFirma.TryGetValue(etiqueta, out var fila)
            ? fila
            : throw new InvalidOperationException($"El compositor interior R1 requiere el ancla '{etiqueta}'.");

    private static string RefSimple(string columna, int fila) =>
        columna + fila.ToString(CultureInfo.InvariantCulture);

    private static string FormatearFormula(IReadOnlyList<string> partes)
    {
        var texto = new StringBuilder();
        foreach (var parte in partes)
        {
            texto.Append(parte);
        }

        var resultado = texto.ToString();
        return resultado.StartsWith('+') ? resultado[1..] : resultado;
    }

    private static string Columna(string celdaDestino)
    {
        var texto = new StringBuilder();
        foreach (var caracter in celdaDestino)
        {
            if (!char.IsLetter(caracter))
            {
                break;
            }

            texto.Append(char.ToUpperInvariant(caracter));
        }

        if (texto.Length == 0)
        {
            throw new ArgumentException($"Referencia interior inválida: '{celdaDestino}'.", nameof(celdaDestino));
        }

        return texto.ToString();
    }

    private static string Formatear(int fila) => fila.ToString(CultureInfo.InvariantCulture);
}
