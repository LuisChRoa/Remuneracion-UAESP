namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 29 (T3, Unidad R — ESCRITURA): mapa de destino del desglose-detalle de las hojas
/// <c>Rem. Anticipos R2</c> y <c>Reversion Pagos R4</c>. Declara —por ASE— el bloque de filas de
/// captura (que existen en la plantilla como literales-0, T0b) y la resolución de COLUMNAS por
/// ENCABEZADO de componente (nunca por índice fijo, T0b §2.1).
///
/// Plan 47 (H1, S2.1 — §2.2 + autorización H1.1/H1.3): la fila destino NO se resuelve por orden
/// físico. La clave de colocación es <c>(firma + CANAL EFECTIVO)</c>:
/// <list type="bullet">
/// <item><see cref="CanalEfectivoDetalleR2R4"/> (Core, puro, R46-1 análogo) calcula el canal/empresa
/// de la fila <c>Total</c> que CIERRA el grupo, tanto en la fuente como en el oficial;</item>
/// <item>una fila-fuente sin slot es <b>cero documentado</b> sólo si está en
/// <see cref="CerosDocumentadosDetalleR2R4"/> (evidencia citada en el T0 del Plan 47); fuera de esa
/// lista es <c>ERR-SIN-SLOT</c> accionable (hoja + ASE + firma + canal) — nunca remap/insert ni
/// colocación por posición (Plan 41 E1-1).</item>
/// </list>
///
/// Cero geometría (D-C): este mapa NO inserta ni borra filas, solo nombra filas y columnas que ya
/// existen. La escritura la ejecuta <see cref="OpenXmlPlantillaWriter"/> con el guard anti-fórmula
/// existente (celda-destino fórmula ⇒ <c>ERR-PLANTILLA</c>, nunca sobrescritura).
/// </summary>
public static class WorkbookLeafCellMapDetalleR2R4
{
    /// <summary>Hoja de detalle R2 (Saldos a Favor / Anticipos).</summary>
    public const string HojaR2 = "Rem. Anticipos R2";

    /// <summary>Hoja de detalle R4 (Reversión por componente).</summary>
    public const string HojaR4 = "Reversion Pagos R4";

    /// <summary>
    /// Columna (0=A) que porta el RÓTULO DE GRUPO del detalle R2 — el canal en <c>C</c> (bytes:
    /// `C5=A D=Total`, `C9=E D=Total`, `C77=1 D=Total`, `C79=2 D=Total`, `C82=5 D=Total`). El canal
    /// efectivo (R46-1) se calcula con esta columna + la fila de cierre y se COMPONE con la empresa del
    /// segundo cierre (R2 ASE5: `E/ENEL` vs `E/ENERBIT`); el predicado vive en Core.
    /// </summary>
    public static readonly IReadOnlyList<int> ColumnasEtiquetaGrupoR2 = [2];

    /// <summary>
    /// Columnas (0=A, 1=B, 2=C, 3=D) que portan la ETIQUETA PROPIA del detalle R2 (identifican la fila →
    /// se localiza por firma, sin canal efectivo): <c>A</c> (roles de sección `Componente`/`Subs/Cont`/
    /// `Total`), <c>B</c> (empresa) y <c>C</c> (canal). El concepto vive en <c>D</c>. Sin esta
    /// distinción, la fila de sección `A=Componente B=Total` tomaría un "canal" que no le corresponde.
    /// </summary>
    public static readonly IReadOnlyList<int> ColumnasEtiquetaPropiaR2 = [0, 1, 2];

    /// <summary>
    /// Columnas (0=A, 1=B, 2=C, 3=D) que portan la ETIQUETA DE GRUPO del detalle R4 — empresa en
    /// <c>A</c> (bytes: `A8=ENEL B=Total`, `A14=OCCIDENTE B=Total`).
    /// </summary>
    public static readonly IReadOnlyList<int> ColumnasEtiquetaGrupoR4 = [0];

    /// <summary>
    /// Columnas (0=A, 1=B, 2=C, 3=D) que portan la ETIQUETA PROPIA del detalle R4: solo la empresa
    /// <c>A</c>. El rol de sub-bloque (<c>Componente</c>/<c>Subs/Cont</c>) vive en <c>B</c> DENTRO de
    /// cada grupo (bytes: `B95=Componente C=Total`, `B97=Subs/Cont C=Total`), por lo que esas filas SÍ
    /// participan del canal efectivo (cada grupo tiene su juego Componente/Subsidio/Subs-Cont).
    /// </summary>
    public static readonly IReadOnlyList<int> ColumnasEtiquetaPropiaR4 = [0];

    /// <summary>
    /// Columnas (0=A, 1=B, 2=C, 3=D) del SEGUNDO cierre (Plan 47 W-3): la empresa del bloque en R2 —
    /// <c>B</c>, género `B=ENEL C=Total` / `B=ENERBIT C=Total` (bytes: `B10=ENEL C10=Total`,
    /// `B389=ENEL C389=Total` y `B393=ENERBIT C393=Total`). El canal efectivo se compone
    /// <c>&lt;canal&gt;/&lt;empresa&gt;</c> para desambiguar dos grupos del MISMO canal (R2 ASE5:
    /// <c>E/ENEL</c> vs <c>E/ENERBIT</c>) sin recurrir al orden.
    /// </summary>
    public static readonly IReadOnlyList<int> ColumnasSegundoCierreR2 = [1];

    /// <summary>
    /// R4 no tiene segundo cierre: su etiqueta de grupo (columna <c>A</c>) YA es la empresa
    /// (`A8=ENEL B=Total`, `A14=OCCIDENTE B=Total`) → canal simple, sin composición.
    /// </summary>
    public static readonly IReadOnlyList<int> ColumnasSegundoCierreR4 = [];

    /// <summary>
    /// Bloque de detalle por ASE en <c>Rem. Anticipos R2</c>: rango INCLUSIVO de filas con labels
    /// y valores literales-0 (T0b §2.2). El cierre <c>Total</c> es la última fila del rango.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (int FilaInicio, int FilaFin)> BloquesR2 =
        new Dictionary<int, (int, int)>
        {
            [1] = (3, 33),
            [2] = (67, 124),
            [3] = (175, 205),
            [4] = (293, 350),
            [5] = (404, 441)
        };

    /// <summary>
    /// Bloque de detalle por ASE en <c>Reversion Pagos R4</c>: rango INCLUSIVO de filas con labels
    /// y valores literales-0 (T0b §2.3). El cierre <c>Total</c> es la última fila del rango.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (int FilaInicio, int FilaFin)> BloquesR4 =
        new Dictionary<int, (int, int)>
        {
            [1] = (3, 15),
            [2] = (93, 111),
            [3] = (194, 206),
            [4] = (231, 249),
            [5] = (346, 358)
        };

    /// <summary>
    /// Columnas destino del detalle R2 en orden: (encabezado-fuente, columna-destino). El
    /// encabezado <c>null</c> marca una columna destino SIN contraparte en la fuente (se escribe 0
    /// explícito). <c>Especiales K</c> tiene contraparte solo en agosto; en julio la fuente no trae
    /// la columna y el modelo emite 0 explícito (invariante).
    ///
    /// <para><b>Plan 60 (H2.2):</b> las dos últimas columnas son los puertos del par
    /// <c>Deb/Cred</c> / <c>Deb/Cred-otros</c> y su reparto NO es invariante. La fuente R2 solo trae
    /// <c>Deb/Cred</c>; <c>Deb/Cred-otros</c> arranca en 0 y pasa a ser DESTINO de la reclasificación
    /// por canal: la regla <see cref="ReclasificadorDebCredR2"/> mueve a <c>R</c> el bloque completo
    /// del canal cuyo cierre supera el umbral y re-propaga los agregados literales, de modo que ambos
    /// valores se escriben POR VALOR MOVIDO (el par <c>Q + R</c> se conserva). Antes de esta regla el
    /// doc declaraba <c>R</c> como «invariante en 0»; ya no lo es.</para>
    /// </summary>
    public static readonly IReadOnlyList<(string? Encabezado, string Columna)> ColumnasR2 =
    [
        ("Total", "E"),
        ("Componente TDF", "F"),
        ("Componente TTL", "G"),
        ("Componente TVIAT", "H"),
        ("Aprovechamiento", "I"),
        ("CCSA Prest.Aprov.", "J"),
        ("Especiales", "K"),
        ("Componente TCS", "L"),
        ("Componente TLU", "M"),
        ("Componente TBL", "N"),
        ("Componente TRT", "O"),
        ("CCSA Prest. No Aprov.", "P"),
        ("Deb/Cred", "Q"),
        ("Deb/Cred-otros", "R")
    ];

    /// <summary>
    /// Columnas destino del detalle R4 en orden: (encabezado-fuente, columna-destino). <c>D..O</c>
    /// tienen contraparte en la fuente en los 5 ASE; <c>Otros → R</c> tiene contraparte en la fuente
    /// de ASE4 (Plan 47 H1.3 — falso hueco de columna: el oficial tiene `R2=Otros` como VALOR y la
    /// fuente de ASE4 trae el encabezado `Otros` con valores, p. ej. `P6=-0,29`). Las columnas
    /// <c>SERVICIO ESPECIALES P</c> y <c>Ajuste Recargo Aseo Q</c> NO tienen contraparte en ninguna
    /// fuente julio-2 (ausencia probada por bytes, T0 §6.3): quedan en el 0 de la plantilla.
    /// </summary>
    public static readonly IReadOnlyList<(string? Encabezado, string Columna)> ColumnasR4 =
    [
        ("Total", "D"),
        ("Componente TDF", "E"),
        ("Componente TTL", "F"),
        ("Componente TVIAT", "G"),
        ("Aprovechamiento", "H"),
        ("CCSA Prest.Aprov.", "I"),
        ("Componente TCS", "J"),
        ("Componente TLU", "K"),
        ("Componente TBL", "L"),
        ("Componente TRT", "M"),
        ("CCSA Prest. No Aprov.", "N"),
        ("Ajuste Decena", "O"),
        ("Otros", "R")
    ];
}
