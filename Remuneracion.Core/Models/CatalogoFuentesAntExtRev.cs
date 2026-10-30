namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 64 (H1.1, §1.2 / §2.2 / D-03): catálogo CERRADO de fuentes de <c>ANT EXT-REV</c> como DATOS
/// citados — etiqueta de bloque de la plantilla ← etiquetas con las que el insumo
/// <c>RecaudosReversados_*.xlsx</c> rotula esa fuente en la fila de total de su grupo (columna A del
/// bloque 2).
///
/// <para><b>Doctrina (misma que <see cref="SinonimosEmpresaR1"/>, Plan 34 D-B).</b> El sistema no
/// conoce fuentes por ramas de código: una fuente nueva se absorbe añadiendo una FILA con su cita de
/// negocio. El matcher es PURO y reutiliza <see cref="SinonimosEmpresaR1.SonMismaEmpresa"/>, de modo
/// que la sinonimia citada del Plan 34 (<c>RECIPROCIDAD ↔ NUEVO ESQUEMA</c>) vale aquí sin duplicar
/// la tabla. Sin cita no hay fila (mitigación R-SINONIMO-ABUSO).</para>
///
/// <para><b>HUE-1 (declarado, no resuelto).</b> <c>EAAB CIUDAD LIMPIA</c> y <c>ENERBIT</c> NO tienen
/// fila: aparecen 0/30 en los insumos y están en 0 en los 4 manuales (T0 §5.1). Si algún período los
/// trajera, el lector falla con <c>ERR-FORMATO-FUENTE</c> (nunca un 0 silencioso, nunca un mapeo
/// inventado).</para>
/// </summary>
public static class CatalogoFuentesAntExtRev
{
    /// <summary>
    /// Fila del catálogo: etiqueta de BLOQUE del template (columna B), conjunto de etiquetas con las
    /// que el insumo puede rotular la fuente y la cita de negocio que prueba la equivalencia.
    /// </summary>
    /// <param name="EtiquetaBloque">Etiqueta del bloque en la plantilla (dato leído, nunca coordenada).</param>
    /// <param name="EtiquetasFuente">Etiquetas de fuente del insumo que designan esa misma fuente.</param>
    /// <param name="Cita">Fuente de negocio que prueba la equivalencia (sin cita no hay fila).</param>
    public sealed record FilaFuente(
        string EtiquetaBloque,
        IReadOnlyList<string> EtiquetasFuente,
        string Cita);

    /// <summary>
    /// Tabla citada. Filas iniciales (T0 §5.1/§5.3): identidad para <c>OCCIDENTE</c> y <c>ENEL</c>,
    /// sinonimia citada del Plan 34 para <c>NUEVO ESQUEMA</c> (EAAB Reciprocidad).
    /// </summary>
    public static IReadOnlyList<FilaFuente> Filas { get; } =
    [
        new(
            "BANCO DE OCCIDENTE",
            ["OCCIDENTE"],
            "Docs/T0-Plan64-Ant-Ext-Rev.md §5.1/§5.3 + §1.2 del plan 64: el insumo rotula la fuente " +
            "\"OCCIDENTE\" en la fila de total de su grupo y el manual del administrativo la escribe en el " +
            "bloque \"BANCO DE OCCIDENTE\" (identidad por rótulo normalizado, misma regla que " +
            "SinonimosEmpresaR1.Normalizar)."),
        new(
            "EAAB RECIPROCIDAD",
            ["NUEVO ESQUEMA"],
            "SinonimosEmpresaR1.Clases[0] = {RECIPROCIDAD, NUEVO ESQUEMA} (cita triple: Detalle de plantilla " +
            "«NUEVO ESQUEMA: Corresponde al recaudo \"EAAB Reciprocidad\"»; Proceso de Recaudo (ASE2 = ENEL + " +
            "EAAB); oráculo Reporte Componentes R1!F199) + T0 §5.3 (paridad fuente×etiqueta en 20 archivos)."),
        new(
            "ENEL",
            ["ENEL"],
            "Docs/T0-Plan64-Ant-Ext-Rev.md §5.1/§5.3 + §1.2 del plan 64: identidad por rótulo normalizado; el " +
            "manual del administrativo escribe \"ENEL\" del insumo en el bloque \"ENEL\" de la plantilla.")
    ];

    /// <summary>
    /// Fila del catálogo cuya fuente designa <paramref name="etiquetaFuente"/>; <c>null</c> = la
    /// etiqueta no está tabulada (evidencia, o fail-fast si además hay una fuente del catálogo ausente:
    /// §2.4 paso 6). La comparación es pura (sin períodos, sin I/O) y pasa por
    /// <see cref="SinonimosEmpresaR1.SonMismaEmpresa"/>.
    /// </summary>
    /// <param name="etiquetaFuente">Etiqueta observada en el insumo (se normaliza dentro del matcher).</param>
    public static FilaFuente? BuscarPorEtiquetaFuente(string? etiquetaFuente)
    {
        foreach (var fila in Filas)
        {
            foreach (var etiqueta in fila.EtiquetasFuente)
            {
                if (SinonimosEmpresaR1.SonMismaEmpresa(etiquetaFuente, etiqueta))
                {
                    return fila;
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Fila del catálogo cuya etiqueta de BLOQUE designa <paramref name="etiquetaBloque"/>;
    /// <c>null</c> = el bloque no tiene fuente mapeada (HUE-1: CIUDAD LIMPIA / ENERBIT quedan en el
    /// cero del template). Usa la MISMA regla pura de identidad del proyecto
    /// (<see cref="SinonimosEmpresaR1.SonMismaEmpresa"/>): una implementación única de la comparación de
    /// rótulos, sin reglas paralelas que puedan divergir.
    /// </summary>
    /// <param name="etiquetaBloque">Etiqueta de bloque leída de la plantilla.</param>
    public static FilaFuente? BuscarPorEtiquetaBloque(string? etiquetaBloque)
    {
        foreach (var fila in Filas)
        {
            if (SinonimosEmpresaR1.SonMismaEmpresa(etiquetaBloque, fila.EtiquetaBloque))
            {
                return fila;
            }
        }

        return null;
    }

    /// <summary>
    /// Verdad si dos filas del catálogo designan la MISMA fuente (misma etiqueta de bloque). Se usa
    /// para detectar que el insumo rotule dos veces la misma fuente con etiquetas distintas
    /// (sinónimos): el destino sería ambiguo ⇒ fail-fast (§2.4 paso 6).
    /// </summary>
    /// <param name="primera">Primera fila.</param>
    /// <param name="segunda">Segunda fila.</param>
    public static bool EsMismaFuente(FilaFuente primera, FilaFuente segunda)
    {
        ArgumentNullException.ThrowIfNull(primera);
        ArgumentNullException.ThrowIfNull(segunda);
        return string.Equals(primera.EtiquetaBloque, segunda.EtiquetaBloque, StringComparison.Ordinal);
    }
}

/// <summary>
/// Plan 64 (H1.1, §1.2): catálogo citado CONCEPTO del bloque 2 → (fila de valor, reporte) del bloque
/// de la plantilla. T0 lo congeló con bytes y el manual del administrativo lo confirma en 20 archivos
/// (§5.3). El rótulo del concepto es DATO del insumo (se compara normalizado); la fila y el reporte son
/// rótulos de la plantilla (se resuelven por etiqueta, nunca por índice).
/// </summary>
public static class CatalogoConceptosAntExtRev
{
    /// <summary>
    /// Concepto del insumo con su destino declarado en la plantilla.
    /// </summary>
    /// <param name="Etiqueta">Rótulo del concepto tal como viene en el insumo (columna de concepto).</param>
    /// <param name="Fila">Rótulo de la fila de valor del bloque (<c>APLICADO</c> / <c>ANTICIPOS</c>).</param>
    /// <param name="Reporte">Rótulo de la columna de reporte (<c>REPORTE 3</c> / <c>REPORTE 4</c>).</param>
    /// <param name="Indice">Orden congelado por T0 para el registro de la fuente.</param>
    public sealed record Concepto(string Etiqueta, string Fila, string Reporte, int Indice);

    /// <summary>
    /// Tabla congelada por T0 §1.2 (4 conceptos; la ausencia de uno en una fuente presente vale 0
    /// explícito, D-04).
    /// </summary>
    public static IReadOnlyList<Concepto> Conceptos { get; } =
    [
        new("1-APLICADOS A FACTURACION", "APLICADO", "REPORTE 3", 0),
        new("2-SALDOS A FAVOR GENERADOS", "ANTICIPOS", "REPORTE 3", 1),
        new("9-REVERSION PAGOS A COMPONENTES", "APLICADO", "REPORTE 4", 2),
        new("9-REVERSION SALDOS A FAVOR", "ANTICIPOS", "REPORTE 4", 3)
    ];

    /// <summary>
    /// Concepto cuya etiqueta (ya normalizada por el llamador con la regla del §2.4: mayúsculas, sin
    /// diacríticos, espacios colapsados) corresponde a <paramref name="etiquetaNormalizada"/>;
    /// <c>null</c> = el rótulo no es un concepto del catálogo.
    /// </summary>
    /// <param name="etiquetaNormalizada">Rótulo normalizado observado en el insumo.</param>
    public static Concepto? Buscar(string? etiquetaNormalizada)
    {
        if (string.IsNullOrEmpty(etiquetaNormalizada))
        {
            return null;
        }

        foreach (var concepto in Conceptos)
        {
            if (string.Equals(concepto.Etiqueta, etiquetaNormalizada, StringComparison.Ordinal))
            {
                return concepto;
            }
        }

        return null;
    }
}
