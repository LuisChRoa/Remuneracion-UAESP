namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 42 (T1, R-CLS-1/R-CLS-3, D-B): veredicto de UN sub-bloque FALSE de <c>VALIDACION_TOTAL</c>
/// (por ASE y celda) tras la clasificación del validador. <see cref="Declarada"/> = la divergencia
/// corresponde a una entrada vigente de <see cref="DivergenciasDeclaradasCatalogo"/> (externa, con
/// dueño fuera de la app); <c>false</c> = divergencia NO declarada (defecto interno) que falla en
/// duro con el nombre accionable de siempre. Viaja en el snapshot/resultado (aditivo) para que UI
/// y CLI la presenten SIN re-evaluar (R-DOBLE-PUERTA).
/// </summary>
public sealed class VeredictoDivergenciaValidacion
{
    /// <summary>Id del ASE del snapshot que reporta el sub-bloque (1..5).</summary>
    public int AseId { get; init; }

    /// <summary>Celda del sub-bloque (p. ej. "O25").</summary>
    public string Celda { get; init; } = string.Empty;

    /// <summary><c>true</c> si la divergencia está declarada (entrada del catálogo aplicable).</summary>
    public bool Declarada { get; init; }

    /// <summary>Id de la entrada del catálogo que aplica; <c>null</c> si no está declarada.</summary>
    public string? EntradaId { get; init; }

    /// <summary>Monto observado (lado izquierdo de la igualdad evaluada desde lo escrito).</summary>
    public decimal MontoObservado { get; init; }

    /// <summary>Monto esperado (lado derecho de la igualdad evaluada desde lo escrito).</summary>
    public decimal MontoEsperado { get; init; }

    /// <summary>Diferencia material |observado − esperado|.</summary>
    public decimal Delta => Math.Abs(MontoObservado - MontoEsperado);
}

/// <summary>
/// Plan 42 (T1, R-CLS-1/R-CLS-2/R-CLS-3, D-F): resultado de clasificar los sub-bloques FALSE de
/// <c>VALIDACION_TOTAL</c> en <see cref="VeredictoDivergenciaValidacion.Declarada"/> vs no declarada.
/// Es la ÚNICA decisión (Core); UI y CLI solo la presentan.
///
/// D-F (la declarada es la ÚLTIMA puerta): si hay CUALQUIER divergencia no declarada
/// (<see cref="HayNoDeclaradas"/>), el run falla en duro aunque la declarada esté reconocida.
/// </summary>
public sealed class ClasificacionDivergenciasDeclaradas
{
    /// <summary>Veredictos de los sub-bloques FALSE (uno por ASE·celda), clasificados.</summary>
    public IReadOnlyList<VeredictoDivergenciaValidacion> Veredictos { get; init; } = [];

    /// <summary>
    /// Todos los errores de la validación cruzada que NO son sub-bloques declarados: el resto de
    /// gates internos (por empresa, DetValiRetri, O9/P9, multi-ASE) y los sub-bloques FALSE no
    /// declarados. Si está vacío, lo único que queda son divergencias declaradas.
    /// </summary>
    public IReadOnlyList<string> ErroresNoDeclarados { get; init; } = [];

    /// <summary>Veredictos declarados (entrada del catálogo aplicable).</summary>
    public IReadOnlyList<VeredictoDivergenciaValidacion> Declaradas =>
        Veredictos.Where(v => v.Declarada).ToList();

    /// <summary>Veredictos NO declarados (defecto interno).</summary>
    public IReadOnlyList<VeredictoDivergenciaValidacion> NoDeclaradas =>
        Veredictos.Where(v => !v.Declarada).ToList();

    /// <summary>Ids de entrada declarados (distintos), para el reconocimiento y el reporte.</summary>
    public IReadOnlyList<string> EntradasDeclaradas =>
        Declaradas.Select(v => v.EntradaId).Where(id => !string.IsNullOrWhiteSpace(id)).Distinct(StringComparer.Ordinal).ToList()!;

    /// <summary>
    /// Verdad si hay algo que NO es una divergencia declarada (gates internos o sub-bloques no
    /// declarados). Con esto verdadero, el run FALLA en duro (D-F).
    /// </summary>
    public bool HayNoDeclaradas => NoDeclaradas.Count > 0 || ErroresNoDeclarados.Count > 0;

    /// <summary>Verdad si lo único que queda son divergencias declaradas (D-F: última puerta).</summary>
    public bool SoloDeclaradas => !HayNoDeclaradas && Declaradas.Count > 0;

    /// <summary>Verdad si no hay ninguna divergencia (flujo actual intacto, cero fricción).</summary>
    public bool SinDivergencias => Veredictos.Count == 0 && ErroresNoDeclarados.Count == 0;
}
