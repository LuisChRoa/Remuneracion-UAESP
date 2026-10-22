namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 53 (D1, A1): SECCIÓN ESPERADA de una celda destino del contrato por operando/empresa de
/// <c>Reporte Componentes R1</c>. Cada celda del mapa
/// (<c>WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa</c>) declara <b>qué sección de la fuente
/// debe alimentarla</b>, además del tipo de valor. El contrato se deriva de la plantilla/manual (T0
/// del Plan 53: <c>Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md</c>) y <b>no</b> se infiere por
/// ocurrencia/índice ni por magnitud (fin del camino ordinal del Plan 52 §13).
///
/// Los roles son <b>relativos</b> a las secciones presentes del bloque de la empresa (Totales o Subs,
/// según el operando), en orden de aparición: cero literales de período/empresa (D-H).
/// </summary>
public enum SeccionEsperadaR1
{
    /// <summary>
    /// La sección de recaudo única del bloque (primera == última; caso ASE5-oct: las 3 empresas traen
    /// una sola fila en <c>Oportuno#1</c>). Es la MISMA resolución que
    /// <see cref="UltimaOportuna"/> (la última sección presente del bloque); se declara donde el T0
    /// probó la geometría de sección única y su valor contractual es <b>blindar la celda</b> contra el
    /// rol <see cref="Primera"/> (prohíbe la lectura "única primera ⇒ Menor" de la regla D3-del-52,
    /// derogada). Nunca es un fallback posicional: la sección sale de la lista de secciones del bloque.
    /// </summary>
    UnicaRecaudo,

    /// <summary>
    /// Primera sección del bloque <b>distinta de la última</b> (el "otro" bloque: Aplicación/primera
    /// sección). Si el bloque tiene UNA sola sección no hay primera distinta → la celda queda en
    /// <b>0 explícito</b> (la fuente no trae ese bloque; regla 6 del project-context).
    /// </summary>
    Primera,

    /// <summary>
    /// Última sección presente del bloque (la que cierra el bloque: el recaudo oportuno en las formas
    /// congeladas por el T0).
    /// </summary>
    UltimaOportuna,

    /// <summary>
    /// Sección intermedia del bloque (solo si hay ≥ 3 secciones presentes; con menos, la celda queda
    /// en <b>0 explícito</b>). Ver D4: <c>TotalMedio = Intermedia si ≥3, si no 0</c>.
    /// </summary>
    Intermedia,
}
