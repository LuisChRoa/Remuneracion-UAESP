using Remuneracion.Core.Errors;

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
    /// Código del catálogo de errores (HU-14). Default: <see cref="CodigoError.Validacion"/>.
    /// </summary>
    public string Codigo { get; } = CodigoError.Validacion;
}
