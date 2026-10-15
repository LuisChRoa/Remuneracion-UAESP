using System.Globalization;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;

namespace Remuneracion.WinForms;

// Plan 42 (T3-UI, R-CLS-3/R-RE-RUN-UI, D-D/D-F): presentador PURO de la superficie de
// reconocimiento explícito de divergencias declaradas. No decide nada (la clasificación ya
// viene hecha en Core) y no toca ningún control: solo agrupa, formatea y redacta el copy
// administrativo que el diálogo muestra. Sin usings de WinForms para poder compilarse también
// en tests (archivo vinculado). Cero literales de período o empresa (D-H): todo llega por
// parámetro; la única lectura del catálogo es la búsqueda por id para el copy/evidencia,
// explícitamente en contrato.
public static class ReconocimientoDivergencia
{
    private static readonly NumberFormatInfo FormatoNegocio = new()
    {
        NumberGroupSeparator = ".",
        NumberDecimalSeparator = ",",
        NumberGroupSizes = [3],
        NumberDecimalDigits = 2
    };

    // R-RE-RUN-UI / D-F: la superficie SOLO se ofrece cuando lo único que impide cerrar es una
    // divergencia declarada aún no reconocida en esta corrida. Con no-declaradas coexistentes
    // (o sin clasificación) devuelve falso y el flujo sigue al FAIL duro habitual.
    public static bool OfreceReconocimiento(CalculoInvalidoException? excepcion) =>
        excepcion?.Clasificacion?.SoloDeclaradas == true;

    // La clasificación trae un veredicto por ASE·celda (hoy 15 = 5 ASE × 3 celdas del mismo par):
    // la superficie muestra UNA fila por divergencia, deduplicando por (EntradaId, Celda).
    public static IReadOnlyList<VeredictoDivergenciaValidacion> AgruparPorDivergencia(
        ClasificacionDivergenciasDeclaradas clasificacion)
    {
        ArgumentNullException.ThrowIfNull(clasificacion);

        return clasificacion.Declaradas
            .GroupBy(v => (v.EntradaId, v.Celda))
            .Select(g => g.First())
            .OrderBy(v => v.EntradaId, StringComparer.Ordinal)
            .ThenBy(v => v.Celda, StringComparer.Ordinal)
            .ToList();
    }

    // Ids reconocidos en ESTA corrida: exactamente las entradas declaradas de la clasificación
    // (acto consciente sobre lo presentado; nada persistente, nada pre-marcado).
    public static IReadOnlyList<string> IdsAceptados(ClasificacionDivergenciasDeclaradas clasificacion)
    {
        ArgumentNullException.ThrowIfNull(clasificacion);

        return clasificacion.EntradasDeclaradas;
    }

    // Formato de negocio (miles con punto, decimales con coma: 417.519,62), estable e
    // independiente de la cultura del equipo — el mismo del reporte durable.
    public static string FormatearMontoNegocio(decimal valor) =>
        valor.ToString("#,##0.00", FormatoNegocio);

    // §2.3 (QUÉ, primero): hoja/celda + observado vs esperado + Δ, en lenguaje de negocio.
    public static string TextoQue(VeredictoDivergenciaValidacion veredicto)
    {
        ArgumentNullException.ThrowIfNull(veredicto);

        return $"Dónde: hoja VALIDACION_TOTAL, celda {veredicto.Celda}. "
            + $"Qué muestra el archivo: {FormatearMontoNegocio(veredicto.MontoObservado)}. "
            + $"Qué esperaba la validación: {FormatearMontoNegocio(veredicto.MontoEsperado)}. "
            + $"Diferencia: {FormatearMontoNegocio(veredicto.Delta)}.";
    }

    // §2.3 (DE QUIÉN ES, después): dueño pendiente + cita de evidencia + fecha + retiro.
    public static string TextoDeQuienEs(DivergenciasDeclaradasCatalogo.Entrada? entrada)
    {
        if (entrada is null)
        {
            return "De quién es: responsable externo por confirmar.";
        }

        return $"De quién es: pendiente de {entrada.Dueno}."
            + $"{Environment.NewLine}Evidencia: {entrada.CitaEvidencia}"
            + $"{Environment.NewLine}Fecha: {entrada.Fecha}"
            + $"{Environment.NewLine}Se retira cuando: {entrada.CondicionRetiro}";
    }

    // §2.3 (QUÉ SE PIDE): el acto de aceptación nombra QUÉ se acepta. El id es el del catálogo
    // (el mismo que pide el flag del CLI); el dueño sale de la entrada citada.
    public static string TextoAceptacion(string entradaId, decimal delta, string? dueno)
    {
        var responsable = string.IsNullOrWhiteSpace(dueno) ? "el responsable externo" : dueno;
        return $"Acepto cerrar la corrida con la divergencia {entradaId}, "
            + $"Δ = {FormatearMontoNegocio(delta)}, pendiente de {responsable}.";
    }

    // Búsqueda por id SOLO para copy/evidencia (en contrato). No clasifica: si el id no está en
    // el catálogo vigente, devuelve nulo y el copy usa el fallback administrativo.
    public static DivergenciasDeclaradasCatalogo.Entrada? BuscarEntrada(string? id)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            return null;
        }

        foreach (var entrada in DivergenciasDeclaradasCatalogo.Entradas)
        {
            if (string.Equals(entrada.Id, id, StringComparison.OrdinalIgnoreCase))
            {
                return entrada;
            }
        }

        return null;
    }
}
