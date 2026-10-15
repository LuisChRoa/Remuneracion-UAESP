using Remuneracion.Core.Errors;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Exceptions;

/// <summary>
/// Excepción lanzada cuando un cálculo de remuneración es inválido.
/// HU-14 (3.1, D1): porta <see cref="Codigo"/> del catálogo (default
/// <see cref="CodigoError.Validacion"/>); constructores previos intactos (compat 160 tests).
/// </summary>
public class CalculoInvalidoException : Exception
{
    /// <summary>
    /// Inicializa la excepción con un mensaje.
    /// </summary>
    /// <param name="message">Mensaje de error.</param>
    public CalculoInvalidoException(string message) : base(message)
    {
    }

    /// <summary>
    /// Inicializa la excepción con un mensaje y una excepción interna.
    /// </summary>
    /// <param name="message">Mensaje de error.</param>
    /// <param name="inner">Excepción interna.</param>
    public CalculoInvalidoException(string message, Exception inner) : base(message, inner)
    {
    }

    /// <summary>
    /// Inicializa la excepción con un código explícito del catálogo y un mensaje.
    /// </summary>
    /// <param name="codigo">Código del catálogo (p. ej. <see cref="CodigoError.Validacion"/>).</param>
    /// <param name="message">Mensaje de error.</param>
    public CalculoInvalidoException(string codigo, string message) : base(message)
    {
        Codigo = codigo;
    }

    /// <summary>
    /// Inicializa la excepción con un código explícito del catálogo, un mensaje y una excepción interna.
    /// </summary>
    /// <param name="codigo">Código del catálogo.</param>
    /// <param name="message">Mensaje de error.</param>
    /// <param name="inner">Excepción interna.</param>
    public CalculoInvalidoException(string codigo, string message, Exception inner) : base(message, inner)
    {
        Codigo = codigo;
    }

    /// <summary>
    /// Plan 42 (T1/T3, D-D/D-F/D-G): inicializa la excepción portando además la clasificación de
    /// divergencias declaradas que originó el bloqueo. Se usa cuando lo ÚNICO que impide cerrar es
    /// una divergencia DECLARADA aún no reconocida en esta corrida: la UI/CLI leen
    /// <see cref="Clasificacion"/> para presentar la superficie de reconocimiento SIN re-evaluar.
    /// El tipo sigue siendo <see cref="CalculoInvalidoException"/> (mismo contrato y código).
    /// </summary>
    /// <param name="codigo">Código del catálogo (p. ej. <see cref="CodigoError.Validacion"/>).</param>
    /// <param name="message">Mensaje de error (nombra las celdas y las entradas declaradas).</param>
    /// <param name="clasificacion">Clasificación que originó el bloqueo.</param>
    public CalculoInvalidoException(string codigo, string message, ClasificacionDivergenciasDeclaradas clasificacion)
        : base(message)
    {
        Codigo = codigo;
        Clasificacion = clasificacion;
    }

    /// <summary>
    /// Código del catálogo de errores (HU-14). Default: <see cref="CodigoError.Validacion"/>.
    /// </summary>
    public string Codigo { get; } = CodigoError.Validacion;

    /// <summary>
    /// Plan 42 (T1/T3): clasificación de divergencias declaradas asociada al bloqueo; <c>null</c>
    /// cuando la excepción no proviene de un cierre bloqueado por una divergencia declarada.
    /// </summary>
    public ClasificacionDivergenciasDeclaradas? Clasificacion { get; }
}
