namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 29 (T3, Unidad R — ESCRITURA): mapa de destino del desglose-detalle de las hojas
/// <c>Rem. Anticipos R2</c> y <c>Reversion Pagos R4</c>. Declara —por ASE— el bloque de filas de
/// captura (que existen en la plantilla como literales-0, T0b) y la resolución de COLUMNAS por
/// ENCABEZADO de componente (nunca por índice fijo, T0b §2.1).
///
/// La fila destino se resuelve por LABEL (firma A|B|C|D para R2, A|B|C para R4) contra la propia
/// plantilla, de modo que la malla derivada de agosto (sub-bloques intercalados en LIME/BOGOTA
/// para R2; filas reducidas en PROMO/CIUDAD/BOGOTA para R4) no rompe el mapeo: las filas-fuente
/// sin contraparte se omiten y las filas-destino sin fuente quedan en su 0 de plantilla.
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
    /// Bloque de detalle por ASE en <c>Rem. Anticipos R2</c>: rango INCLUSIVO de filas con labels
    /// y valores literales-0 (T0b §2.2). El cierre <c>Total</c> es la última fila del rango.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (int FilaInicio, int FilaFin)> BloquesR2 =
        new Dictionary<int, (int, int)>
        {
            [1] = (3, 29),
            [2] = (63, 108),
            [3] = (159, 188),
            [4] = (276, 324),
            [5] = (378, 411)
        };

    /// <summary>
    /// Bloque de detalle por ASE en <c>Reversion Pagos R4</c>: rango INCLUSIVO de filas con labels
    /// y valores literales-0 (T0b §2.3). El cierre <c>Total</c> es la última fila del rango.
    /// </summary>
    public static readonly IReadOnlyDictionary<int, (int FilaInicio, int FilaFin)> BloquesR4 =
        new Dictionary<int, (int, int)>
        {
            [1] = (3, 15),
            [2] = (93, 105),
            [3] = (188, 200),
            [4] = (225, 243),
            [5] = (340, 352)
        };

    /// <summary>
    /// Columnas destino del detalle R2 en orden: (encabezado-fuente, columna-destino). El
    /// encabezado <c>null</c> marca una columna destino SIN contraparte en la fuente (se escribe 0
    /// explícito; invariante <c>Deb/Cred-otros R</c>). <c>Especiales K</c> tiene contraparte solo
    /// en agosto; en julio la fuente no trae la columna y el modelo emite 0 explícito (invariante).
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
    /// Columnas destino del detalle R4 en orden: (encabezado-fuente, columna-destino). Solo D..O
    /// tienen contraparte en la fuente; <c>SERVICIO ESPECIALES P</c>, <c>Ajuste Recargo Aseo Q</c> y
    /// <c>Otros R</c> quedan en el 0 de la plantilla (T0b §2.3 — no se escriben).
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
        ("Ajuste Decena", "O")
    ];
}
