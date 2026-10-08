namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 32 (T1, pieza b — R-B-1/D-B): firmas puras del INTERIOR de <c>Reporte Componentes R1</c>
/// (sub-visibles por empresa). Declaraciones puras — el consumidor (la recomposición por firma del
/// pase final del mutador) llega en T2.
///
/// Causa (T0c, evidencia en disco 2026-10-08): los anclajes del interior son MIXTOS. Parte resuelve
/// con las firmas vigentes de <see cref="FilaEspejoR1"/> (<c>EsMesTotal</c> =
/// <c>B='Mes' ∧ C='Total'</c>; <c>EsAplicacionTotal</c> = <c>B</c> contiene
/// <c>Aplicacion nuevos x reversion</c> ∧ <c>C='Total'</c>), pero los términos de empresa resuelven a
/// filas con firma <b><c>C=&lt;empresa&gt; ∧ D='Total'</c></b> (p. ej. <c>[ENEL/Total]</c>,
/// <c>[OCCIDENTE/Total]</c>, <c>[NUEVO ESQUEMA/Total]</c>) — firma que NINGÚN predicado vigente cubre
/// (<c>EsMesTotal</c> exige <c>B='Mes'</c>; <c>EsCodigoD</c> mira la columna D de códigos, no C).
///
/// El <b>catálogo de empresas de la columna C es ABIERTO</b> (el disco trae <c>NUEVO ESQUEMA</c> además
/// de ENEL/OCCIDENTE/RECIPROCIDAD/ENERBIT/CIUDAD LIMPIA-ACUEDUCTO): estas firmas NO enumeran nombres,
/// solo la forma <c>(B,C,D,E)</c>. La recomposición alinea el sub-bloque fuente ↔ filas finales por
/// ORDEN (misma cardinalidad por construcción del delta), nunca por enésima ocurrencia (D-B).
///
/// Todas las comparaciones de etiqueta usan la misma normalización <c>OrdinalIgnoreCase</c> de las
/// firmas de <see cref="FilaEspejoR1"/> (la secuencia del espejo ya viene normalizada por
/// <c>LeerEspejoR1</c>).
/// </summary>
public static class R1FirmaInterior
{
    /// <summary>
    /// Plan 32 (T0c, D-B): firma de una FILA-DATO de empresa — ancla de los términos del interior
    /// (<c>C=&lt;empresa&gt; ∧ D='Total'</c>, con <c>B</c> vacía). Catálogo de empresas ABIERTO:
    /// no se enumera ningún nombre; <c>NUEVO ESQUEMA</c> y cualquier nombre por-fuente caben.
    /// </summary>
    public static bool EsDatoEmpresa(FilaEspejoR1 fila)
    {
        ArgumentNullException.ThrowIfNull(fila);
        return !string.IsNullOrWhiteSpace(fila.C)
            && string.IsNullOrWhiteSpace(fila.B)
            && EsTotal(fila.D);
    }

    /// <summary>
    /// Igual que <see cref="EsDatoEmpresa(FilaEspejoR1)"/> pero exigiendo una empresa concreta de la
    /// columna C. La comparación aprende la equivalencia de sinónimos legado↔vigente
    /// (<see cref="SinonimosEmpresaR1.SonMismaEmpresa"/>): el rótulo del template <c>RECIPROCIDAD</c>
    /// y el nombre de la fuente <c>NUEVO ESQUEMA</c> son la misma empresa. El catálogo sigue siendo
    /// abierto por construcción: un rótulo no tabulado solo se iguala a sí mismo (Plan 34 D-A/D-B).
    /// </summary>
    public static bool EsDatoEmpresa(FilaEspejoR1 fila, string empresa)
    {
        ArgumentNullException.ThrowIfNull(fila);
        return EsDatoEmpresa(fila) && SinonimosEmpresaR1.SonMismaEmpresa(fila.C, empresa);
    }

    /// <summary>
    /// Plan 32 (T0c, D-B): firma de la FILA-SUBTOTAL de empresa del interior
    /// (<c>C=&lt;empresa&gt; ∧ D='OPORTUNO' ∧ E='Total'</c>). Excluye la fila visible TOT_OPT del
    /// bloque (<c>C='TOTAL'</c>), que NO es un subtotal de empresa. Catálogo de empresas ABIERTO.
    /// </summary>
    public static bool EsSubtotalEmpresa(FilaEspejoR1 fila)
    {
        ArgumentNullException.ThrowIfNull(fila);
        return !string.IsNullOrWhiteSpace(fila.C)
            && !EsTotal(fila.C)
            && string.Equals(fila.D, "OPORTUNO", StringComparison.OrdinalIgnoreCase)
            && EsTotal(fila.E);
    }

    /// <summary>
    /// Igual que <see cref="EsSubtotalEmpresa(FilaEspejoR1)"/> pero exigiendo una empresa concreta de
    /// la columna C (localizador de fila-subtotal por <c>(C,D,E)</c>; D-B).
    /// </summary>
    public static bool EsSubtotalEmpresa(FilaEspejoR1 fila, string empresa)
    {
        ArgumentNullException.ThrowIfNull(fila);
        return EsSubtotalEmpresa(fila) && string.Equals(fila.C, empresa, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Plan 32 (T0c): firma de la fila EXTEMP-interior (<c>D='EXTEMPORANEO' ∧ E='Total'</c>). Aplica
    /// tanto al EXTEMP visible del bloque como a los EXTEMP por empresa del interior (la posición en
    /// la secuencia decide cuál es cuál).
    /// </summary>
    public static bool EsExtemporaneoInterior(FilaEspejoR1 fila)
    {
        ArgumentNullException.ThrowIfNull(fila);
        return string.Equals(fila.D, "EXTEMPORANEO", StringComparison.OrdinalIgnoreCase)
            && EsTotal(fila.E);
    }

    /// <summary>
    /// Plan 32 (T0c): firma de la fila Subsidio del interior
    /// (<c>E='Subsidio(-)/Contribucion(+)'</c>). Sus celdas de columna F son single-ref del ancla del
    /// sub-bloque (Subs single-ref, R-1) en el manual.
    /// </summary>
    public static bool EsSubsidioInterior(FilaEspejoR1 fila)
    {
        ArgumentNullException.ThrowIfNull(fila);
        return fila.E.Contains("Subsidio(-)/Contribucion(+)", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Plan 32 (T0c): firma de la fila AFaseo del interior
    /// (<c>D='AFASEO' ∧ E='Total'</c>). Su columna F es single-ref del ancla del sub-bloque (R-1).
    /// </summary>
    public static bool EsAFaseoInterior(FilaEspejoR1 fila)
    {
        ArgumentNullException.ThrowIfNull(fila);
        return string.Equals(fila.D, "AFASEO", StringComparison.OrdinalIgnoreCase)
            && EsTotal(fila.E);
    }

    private static bool EsTotal(string? texto) =>
        string.Equals(texto, "TOTAL", StringComparison.OrdinalIgnoreCase);
}
