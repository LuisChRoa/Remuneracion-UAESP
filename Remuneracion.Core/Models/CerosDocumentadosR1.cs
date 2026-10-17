namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 45 (H1.1, S2.3 — §2.2 del plan; Enmienda E-44.1): lista documentada de CEROS de la hoja
/// <c>Reporte Componentes R1</c>. Una fila fuente cuya (firma, sección-contrato) NO tiene slot en el
/// contrato T0 y que pertenece a esta lista queda en el <b>cero documentado</b> del template (no se
/// escribe, no es error, no se simula): el desglose se difiere y se reporta al dueño; los Totales ya
/// llevan el dinero.
///
/// Caso probado en bytes (T0 del Plan 45 §E3/E4): la 1.ª sección del OFICIAL NO declara la fila
/// <c>Vlr Intereses</c> en ninguno de los 5 ASE; la fuente de ASE5-julio-2 SÍ la trae (R8 <c>O</c>
/// 402,40 y R11 <c>T</c> 7.697,23, con desglose G..U) y el manual legacy ASE5 también la registraba.
/// Es un HUECO DE PLANTILLA (no se toca el template; se reporta al dueño).
///
/// Identidad por <c>(firma, sección-contrato)</c> — period-free y company-free (D-H): la consume el
/// escritor values-only por slot (<c>EscritorValoresSlotR1</c>) SOLO cuando la fila no encontró slot.
/// Cualquier fila sin slot FUERA de esta lista sigue fallando en duro con <c>ERR-SIN-SLOT</c>.
/// </summary>
public static class CerosDocumentadosR1
{
    /// <summary>
    /// Verdad si la fila es un cero documentado en la sección-contrato indicada. Hoy: la fila
    /// <c>Vlr Intereses</c> de nivel canal (A/B/C/D vacíos) en la 1.ª sección — el hueco de plantilla
    /// del <c>Vlr Intereses</c> sin slot. Otras filas sin slot NO son cero documentado.
    /// </summary>
    public static bool EsCeroDocumentado(FilaEspejoR1 fila, SeccionFronteraArriba? seccionContrato)
    {
        ArgumentNullException.ThrowIfNull(fila);

        return seccionContrato is { Tipo: TipoSeccionEspejoR1.Primera }
            && string.IsNullOrWhiteSpace(fila.A)
            && string.IsNullOrWhiteSpace(fila.B)
            && string.IsNullOrWhiteSpace(fila.C)
            && string.IsNullOrWhiteSpace(fila.D)
            && string.Equals(fila.E?.Trim(), "Vlr Intereses", StringComparison.OrdinalIgnoreCase);
    }
}
