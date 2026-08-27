namespace Remuneracion.Core.Exceptions;

/// <summary>
/// Excepción lanzada cuando un cálculo de remuneración es inválido.
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
}
