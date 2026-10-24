namespace Remuneracion.Core.Errors;

/// <summary>
/// HU-14 (3.1, D1): catálogo de códigos estables de error. Strings estables para no romper
/// la serialización de mensajes; el código viaja en la excepción (propiedad <c>Codigo</c>),
/// no en el tipo. Cada fail-fast del path período porta un código del catálogo en su mensaje
/// y en la propiedad, y <see cref="CatalogoErrores"/> lo traduce a UX accionable + código de
/// salida (contrato HU-15).
/// </summary>
public static class CodigoError
{
    /// <summary>Carpeta ASE / R1/R2/R4 / banco / balance / notas / retribución ausente.</summary>
    public const string FuenteNoEncontrada = "ERR-FUENTE-NO-ENCONTRADA";

    /// <summary>Plantilla ausente, salida == plantilla, salida sin WorkbookPart válido.</summary>
    public const string Plantilla = "ERR-PLANTILLA";

    /// <summary>Estructura inesperada de fuente (header ausente, celda no numérica, hoja faltante).</summary>
    public const string FormatoFuente = "ERR-FORMATO-FUENTE";

    /// <summary>Gate de dominio/coherencia que no cierra (nombra ASE + validación + celda).</summary>
    public const string Validacion = "ERR-VALIDACION";

    /// <summary>Fallo durante escritura (atomicidad: salida parcial borrada, patrón V5).</summary>
    public const string Escritura = "ERR-ESCRITURA";

    /// <summary>Cualquier otra excepción (catch-all de UI).</summary>
    public const string Inesperado = "ERR-INESPERADO";

    /// <summary>Decisión del usuario que cancela (no es un fallo; nivel Warning).</summary>
    public const string CanceladoPorUsuario = "WARN-CANCELADO";

    /// <summary>
    /// Plan 36 (T1, R-B-2/D-A): la fuente R1 trae una fila (empresa × tipo) sin slot en el
    /// contrato T0 de la plantilla estandarizada. Fail-fast con ASE + reporte + celda + slot
    /// esperado; nunca remap silencioso, nunca insert, nunca 0 silencioso.
    /// </summary>
    public const string SinSlot = "ERR-SIN-SLOT";

    /// <summary>
    /// Plan 58 (H1.2): el paquete generado no es integro por bytes (grupo de formula compartida sin
    /// maestro, formula alterada/creada/perdida, celda duplicada, fila desordenada, indice de
    /// sharedStrings fuera de rango o referencia a una hoja inexistente). Excel repararia el archivo
    /// borrando formulas, asi que el archivo NO se entrega. El mensaje nombra concepto + hoja + celda.
    /// </summary>
    public const string Integridad = "ERR-INTEGRIDAD";
}
