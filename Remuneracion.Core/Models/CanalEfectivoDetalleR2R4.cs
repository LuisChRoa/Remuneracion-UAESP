namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 47 (H1, S2.1 — §2.2 y autorización H1.1): CANAL EFECTIVO del detalle de las hojas
/// <c>Rem. Anticipos R2</c> y <c>Reversion Pagos R4</c> — análogo de la regla R46-1 del frente R1
/// (<see cref="ContratoSeccionPeriodFreeR1.CanalEfectivo"/>).
///
/// <para>
/// Una fila SIN etiqueta propia de grupo (R2: <c>A|B|C</c> vacíos; R4: <c>A</c> vacío — las filas
/// <c>Vlr Servicio</c>, <c>Vlr Intereses</c>, <c>Subsidio(-)/Contribucion(+)</c> y, en R4, los
/// sub-cierres <c>Componente</c>/<c>Subs/Cont</c>) NO se ubica por orden físico: se ubica en el slot
/// que comparte el canal/empresa de la fila <c>Total</c> que <b>CIERRA</b> su grupo (la primera hacia
/// abajo que trae el literal estructural <c>Total</c> y una etiqueta de grupo no vacía).
/// </para>
///
/// <para>
/// El MISMO cálculo se aplica a las filas de la FUENTE y a los slots del OFICIAL, de modo que la
/// clave de colocación es <c>(firma + canal efectivo)</c>. Predicado PURO: sin I/O, sin período y sin
/// empresa (D-H); las etiquetas de grupo son DATOS (qué columna porta el canal/empresa lo declara el
/// mapa geométrico de la plantilla, no este predicado).
/// </para>
/// </summary>
public static class CanalEfectivoDetalleR2R4
{
    /// <summary>Rótulo crudo de una fila del detalle (columnas A–D tal como están en la hoja).</summary>
    public sealed record Rotulo(string A, string B, string C, string D)
    {
        /// <summary>Rótulo vacío (fila sin ninguna etiqueta).</summary>
        public static Rotulo Vacio { get; } = new(string.Empty, string.Empty, string.Empty, string.Empty);

        /// <summary>Texto de la columna indicada (0=A, 1=B, 2=C, 3=D).</summary>
        public string Texto(int columna) => columna switch
        {
            0 => A,
            1 => B,
            2 => C,
            3 => D,
            _ => throw new ArgumentOutOfRangeException(nameof(columna), columna, "El rótulo del detalle solo tiene columnas 0..3 (A..D)."),
        };
    }

    /// <summary>
    /// Etiqueta propia de grupo de la fila: el primer texto no vacío de las columnas de grupo
    /// declaradas por el mapa (R2: canal <c>C</c> y empresa <c>B</c>; R4: empresa <c>A</c>), excluyendo
    /// el literal estructural <c>Total</c>. <c>null</c> = la fila no trae etiqueta de grupo propia.
    /// </summary>
    public static string? EtiquetaDeGrupo(Rotulo fila, IReadOnlyList<int> columnasGrupo)
    {
        ArgumentNullException.ThrowIfNull(fila);
        ArgumentNullException.ThrowIfNull(columnasGrupo);

        foreach (var columna in columnasGrupo)
        {
            var texto = fila.Texto(columna).Trim();
            if (texto.Length > 0 && !EsLiteralTotal(texto))
            {
                return texto;
            }
        }

        return null;
    }

    /// <summary>
    /// Verdad si la fila trae una etiqueta PROPIA que la identifica (se localiza por firma, sin canal
    /// efectivo): cualquier texto no vacío, distinto del literal <c>Total</c>, en las columnas
    /// declaradas por el mapa — R2: <c>A|B|C</c> (roles de sección <c>Componente</c>/<c>Subs/Cont</c>,
    /// empresa y canal; el concepto vive en <c>D</c>); R4: <c>A</c> (empresa; el rol de sub-bloque
    /// <c>Componente</c>/<c>Subs/Cont</c> vive en <c>B</c> DENTRO del grupo y el concepto en <c>C</c>).
    ///
    /// Distinguirla de <see cref="EtiquetaDeGrupo"/> es lo que evita dar "canal" a las filas de sección
    /// (<c>A=Componente B=Total</c> en R2), que no pertenecen a ningún grupo.
    /// </summary>
    public static bool TieneEtiquetaPropia(Rotulo fila, IReadOnlyList<int> columnasEtiquetaPropia)
    {
        ArgumentNullException.ThrowIfNull(fila);
        ArgumentNullException.ThrowIfNull(columnasEtiquetaPropia);

        foreach (var columna in columnasEtiquetaPropia)
        {
            var texto = fila.Texto(columna).Trim();
            if (texto.Length > 0 && !EsLiteralTotal(texto))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// CANAL (rótulo de grupo) EFECTIVO de la fila <paramref name="indice"/>, con **segundo cierre**
    /// (Plan 47 W-3):
    /// <list type="number">
    /// <item>si la fila se identifica por su PROPIA etiqueta fuera del grupo (roles de sección
    /// <c>Componente</c>/<c>Subs/Cont</c>, empresa, <c>Total</c> del bloque) → <c>null</c> (se localiza
    /// por firma);</item>
    /// <item>si no, su rótulo de grupo es el propio si la fila ES el cierre de grupo
    /// (<c>C=E D=Total</c> en R2) o el de la fila <c>Total</c> que CIERRA su grupo;</item>
    /// <item>y se COMPONE <c>&lt;grupo&gt;/&lt;empresa&gt;</c> con la empresa del PRIMER cierre de
    /// empresa por debajo (<c>B=ENEL C=Total</c> en R2), atravesando filas de detalle y otros cierres
    /// de grupo. Así los dos grupos del MISMO canal quedan desambiguados por geometría
    /// (<c>E/ENEL</c> vs <c>E/ENERBIT</c> en R2 ASE5), sin recurrir al orden. Si el mapa no declara
    /// columnas de segundo cierre (R4) o no hay cierre de empresa debajo, devuelve el rótulo simple
    /// (declarado).</item>
    /// </list>
    /// </summary>
    public static string? CanalEfectivo(
        IReadOnlyList<Rotulo> filas,
        int indice,
        IReadOnlyList<int> columnasEtiquetaPropia,
        IReadOnlyList<int> columnasGrupo,
        IReadOnlyList<int>? columnasSegundoCierre = null)
    {
        ArgumentNullException.ThrowIfNull(filas);
        ArgumentNullException.ThrowIfNull(columnasEtiquetaPropia);
        ArgumentNullException.ThrowIfNull(columnasGrupo);
        ArgumentOutOfRangeException.ThrowIfNegative(indice);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(indice, filas.Count);

        var fila = filas[indice];

        // (1) Identidad por etiqueta propia FUERA del grupo (roles de sección, empresa, Total) → firma.
        if (TieneEtiquetaPropia(fila, columnasEtiquetaPropia) && EtiquetaDeGrupo(fila, columnasGrupo) is null)
        {
            return null;
        }

        // (2) Rótulo de grupo: el propio (fila de cierre de grupo) o el del cierre que lo cierra.
        var canal = EtiquetaDeGrupo(fila, columnasGrupo);
        var desde = indice + 1;
        if (canal is null)
        {
            var cierre = -1;
            for (var j = indice + 1; j < filas.Count && cierre < 0; j++)
            {
                if (EsFilaDeCierre(filas[j]) && EtiquetaDeGrupo(filas[j], columnasGrupo) is not null)
                {
                    cierre = j;
                }
            }

            if (cierre < 0)
            {
                return null;
            }

            canal = EtiquetaDeGrupo(filas[cierre], columnasGrupo);
            desde = cierre + 1;
        }

        // (3) Segundo cierre (empresa del bloque): compone el rótulo del grupo.
        var empresa = columnasSegundoCierre is { Count: > 0 }
            ? PrimerCierreConEtiqueta(filas, desde, columnasSegundoCierre)
            : null;
        return empresa is null ? canal : $"{canal}/{empresa}";
    }

    /// <summary>Etiqueta del primer cierre (fila con <c>Total</c>) por debajo de <paramref name="desde"/>
    /// que trae rótulo propio en las columnas del segundo cierre (la empresa del bloque).</summary>
    private static string? PrimerCierreConEtiqueta(IReadOnlyList<Rotulo> filas, int desde, IReadOnlyList<int> columnas)
    {
        for (var j = desde; j < filas.Count; j++)
        {
            if (!EsFilaDeCierre(filas[j]))
            {
                continue;
            }

            var etiqueta = EtiquetaDeGrupo(filas[j], columnas);
            if (etiqueta is not null)
            {
                return etiqueta;
            }
        }

        return null;
    }

    /// <summary>Fila de cierre: alguna de sus cuatro columnas trae el literal estructural <c>Total</c>.</summary>
    private static bool EsFilaDeCierre(Rotulo fila) =>
        EsLiteralTotal(fila.A) || EsLiteralTotal(fila.B) || EsLiteralTotal(fila.C) || EsLiteralTotal(fila.D);

    private static bool EsLiteralTotal(string texto) =>
        string.Equals(texto.Trim(), "Total", StringComparison.OrdinalIgnoreCase);
}
