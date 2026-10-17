namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 45 (H1.1, S2.3 — §2.2; Enmienda E-44.1) + Plan 46 (H1.3, S2.4): lista documentada de CEROS de
/// la hoja <c>Reporte Componentes R1</c>. Una fila fuente cuya (firma, sección-contrato) NO tiene slot
/// en el contrato T0 y que pertenece a esta lista queda en el <b>cero documentado</b> del template (no
/// se escribe, no es error, no se simula); cualquier fila sin slot FUERA de la lista falla en duro con
/// <c>ERR-SIN-SLOT</c>.
///
/// Plan 46 (FP-3, WITHDRAW): la única entrada que tuvo esta lista —la fila <c>Vlr Intereses</c> de la
/// 1.ª sección de ASE5-julio2— queda <b>RETIRADA</b>. Con el tipado de sección por CIERRE
/// (<see cref="ContratoSeccionPeriodFreeR1.SeccionPorCierre"/>) esas dos filas caen en su slot real
/// (<c>Aplic#1</c> oficial, <c>F623</c>/<c>F626</c>), no en un hueco de plantilla. El MECANISMO de cero
/// documentado se conserva (lista <see cref="Patrones"/> vacía) por si otra forma lo exige; ver
/// <c>plans/46</c> §0.3 FP-3.
///
/// Identidad por <c>(firma, sección-contrato)</c> — period-free y company-free (D-H): la consume el
/// escritor values-only por slot (<c>EscritorValoresSlotR1</c>) SOLO cuando la fila no encontró slot.
/// </summary>
public static class CerosDocumentadosR1
{
    /// <summary>
    /// Patrones de cero documentado VIGENTES. Plan 46 (FP-3): VACÍA — las dos entradas del Plan 45
    /// (Vlr Intereses de la 1.ª de ASE5-julio2) se retiraron al tiparse la sección por su frontera de
    /// CIERRE (esas filas ya tienen slot). El mecanismo queda disponible para una forma futura.
    /// </summary>
    private static readonly (bool NivelCanal, string Concepto, TipoSeccionEspejoR1? Tipo)[] Patrones = [];

    /// <summary>
    /// Verdad si la fila es un cero documentado en la sección-contrato indicada. Plan 46 (FP-3): hoy la
    /// lista está vacía → siempre <c>false</c> (no hay ceros documentados activos). El mecanismo se
    /// conserva para una forma futura.
    /// </summary>
    public static bool EsCeroDocumentado(FilaEspejoR1 fila, SeccionFronteraArriba? seccionContrato)
    {
        ArgumentNullException.ThrowIfNull(fila);

        foreach (var (nivelCanal, concepto, tipo) in Patrones)
        {
            if (tipo is not null && seccionContrato is not { } seccion)
            {
                continue;
            }

            if (tipo is not null && seccionContrato!.Value.Tipo != tipo)
            {
                continue;
            }

            if (nivelCanal
                && (!string.IsNullOrWhiteSpace(fila.A)
                    || !string.IsNullOrWhiteSpace(fila.B)
                    || !string.IsNullOrWhiteSpace(fila.C)
                    || !string.IsNullOrWhiteSpace(fila.D)))
            {
                continue;
            }

            if (string.Equals(fila.E?.Trim(), concepto, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }
}
