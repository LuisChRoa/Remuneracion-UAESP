namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 49 (A1-interior, §2.3): mapa de anclas EXTEMPORANEO de <c>Reporte Componentes R1</c>
/// como DATO citado del T0 (<c>Docs/T0-EXTEMP-Ase5-Plan49.md</c> §3.4), no como lógica:
/// ASE5 es el único ASE cuya ancla EXTEMP es un literal en la plantilla v2 (ASE1-4 traen
/// <c>&lt;f&gt;</c> y el resolutor las evalúa solo). El mapa se congela desde los bytes de la v2
/// (verificado 2026-10-19: <c>F660/F629/L629/F659/F628/L628</c> existen como literales;
/// <c>F682/F702</c> son literales <c>0</c> sin <c>&lt;f&gt;</c>) y NO contiene ramas por período ni
/// por empresa (D-H): el ASE es la clave y la celda/expresión son datos.
///
/// <para>
/// El <b>núcleo</b> del gate absoluto no usa este mapa: resuelve
/// <c>CONSOLIDADO_TOTAL RECAUDO!D{46+aseId}</c> (D47..D51) desde lo escrito. Este mapa solo
/// aporta la EVIDENCIA del <b>interior</b> (dónde se perdió el importe: hoja, celda ancla y la
/// suma de sus secciones) para que el diagnóstico no exija un T0 de bytes a mano.
/// </para>
/// </summary>
public static class AnclasExtempR1
{
    /// <summary>Hoja del interior R1.</summary>
    public const string Hoja = "Reporte Componentes R1";

    /// <summary>Hoja del consolidado de salida.</summary>
    public const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";

    /// <summary>
    /// Primera fila del bloque EXTEMP del CONSOLIDADO (D47..D51 = ASE1..ASE5). El plan lo cita
    /// como <c>D{47+k}</c> con <c>k</c> = índice del ASE (0..4).
    /// </summary>
    public const int PrimeraFilaConsolidadoExtemp = 47;

    /// <summary>
    /// Ancla EXTEMP de un ASE: celda destino + expresión de la suma de sus secciones (misma
    /// forma que la fórmula del manual v2, T0 §3.4) + las celdas de sección que la componen
    /// (para distinguir una geometría completa de una forma parcial).
    /// </summary>
    /// <param name="Celda">Celda ancla en <see cref="Hoja"/> (p. ej. <c>F682</c>).</param>
    /// <param name="ExpresionSecciones">
    /// Suma de secciones en notación del resolutor (p. ej. <c>F660+F629-L629</c>).
    /// </param>
    /// <param name="CeldasSeccion">Celdas que la expresión referencia (sin signo ni duplicados).</param>
    public sealed record Ancla(string Celda, string ExpresionSecciones, IReadOnlyList<string> CeldasSeccion);

    private static readonly IReadOnlyDictionary<int, IReadOnlyList<Ancla>> Mapa =
        new Dictionary<int, IReadOnlyList<Ancla>>
        {
            // T0 §3.4 (manual v2 julio): R1!F682 = F660+F629-L629 y R1!F702 = F659+F628-L628.
            // F682 consume CONSOLIDADO!D51; F702 consume REMUNERACION_OCCIDENTE!D51 (T0 §5.2).
            [5] =
            [
                new Ancla("F682", "F660+F629-L629", ["F660", "F629", "L629"]),
                new Ancla("F702", "F659+F628-L628", ["F659", "F628", "L628"])
            ]
        };

    /// <summary>Anclas EXTEMP conocidas del ASE; lista vacía = el ASE no se gatea en el interior.</summary>
    public static IReadOnlyList<Ancla> PorAse(int aseId) =>
        Mapa.TryGetValue(aseId, out var anclas) ? anclas : [];

    /// <summary>
    /// Celda del CONSOLIDADO que recibe el EXTEMP del ASE (<c>D47</c> para ASE1 … <c>D51</c>
    /// para ASE5).
    /// </summary>
    public static string CeldaConsolidadoExtemp(int aseId) =>
        $"D{PrimeraFilaConsolidadoExtemp + aseId - 1}";
}
