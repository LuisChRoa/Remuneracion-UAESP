namespace Remuneracion.Core.Models;

/// <summary>
/// Geometría de PUBLICACIÓN del EXTEMPORANEO en <c>CONSOLIDADO_TOTAL RECAUDO</c> (D47..D51 = ASE1..ASE5):
/// la celda del consolidado que publica el EXTEMP de cada bloque. Es la dirección del NÚCLEO del gate
/// (Plan 49 A1), no un mapa de anclas: el ancla del interior (a qué celda de
/// <c>Reporte Componentes R1</c> apunta) se deriva en runtime siguiendo la referencia que la propia
/// plantilla declara en esta celda (Plan 57 §2-D3.1).
///
/// <para>
/// <b>Plan 57 — retiro citado.</b> La clase <c>AnclasExtempR1</c> del Plan 49 congelaba como DATO la
/// celda-ancla y su expresión de secciones (<c>ASE5: F682 = F660+F629-L629</c>,
/// <c>F702 = F659+F628-L628</c>) y declaraba su naturaleza («literales»). La v2 vigente
/// (SHA <c>0F192AD2…7F2DAB</c>) trae <c>F682 = &lt;f&gt;F660+F629-L629&lt;/f&gt;</c>: la foto quedó
/// desactualizada y el oráculo abortaba con <c>ERR-PLANTILLA</c>. Evidencia por bytes:
/// <c>Docs/T0-Plan57-Anclas-Derivadas.md</c> §1-§2 (fixture
/// <c>AnclasDerivadasT0Plan57Tests</c>). Ese mapa —y su tipo <c>Ancla</c>— <b>se retira</b>: la
/// posición y la naturaleza del ancla se leen del workbook en el momento (Plan 57 §2-D2), y la
/// expresión de secciones sale del propio <c>&lt;f&gt;</c> leído. Lo único que sobrevive es lo que no
/// es foto: la fila de publicación del consolidado (geometría del núcleo, idéntica a la que congela
/// <c>GeometriaPlantillaEstandarizadaT0Tests</c>).
/// </para>
/// </summary>
public static class PublicacionExtempConsolidado
{
    /// <summary>Hoja del consolidado de salida.</summary>
    public const string HojaConsolidado = "CONSOLIDADO_TOTAL RECAUDO";

    /// <summary>
    /// Primera fila del bloque EXTEMP del CONSOLIDADO (D47..D51 = ASE1..ASE5). El plan la cita como
    /// <c>D{47+k}</c> con <c>k</c> = índice del ASE (0..4).
    /// </summary>
    public const int PrimeraFilaConsolidadoExtemp = 47;

    /// <summary>
    /// Fila del CONSOLIDADO que publica el EXTEMP del ASE. En la v2 las hojas publicadoras
    /// (<c>REMUNERACION_*</c>) usan la MISMA fila para el EXTEMP del mismo bloque (T0 §2), de ahí que
    /// la derivación de la familia de anclas se apoye en esta fila.
    /// </summary>
    public static int FilaConsolidadoExtemp(int aseId) =>
        PrimeraFilaConsolidadoExtemp + aseId - 1;

    /// <summary>
    /// Celda del CONSOLIDADO que recibe el EXTEMP del ASE (<c>D47</c> para ASE1 … <c>D51</c> para ASE5).
    /// </summary>
    public static string CeldaConsolidadoExtemp(int aseId) =>
        $"D{FilaConsolidadoExtemp(aseId)}";
}
