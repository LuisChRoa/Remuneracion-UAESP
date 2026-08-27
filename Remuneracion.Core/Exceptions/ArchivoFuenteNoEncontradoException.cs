namespace Remuneracion.Core.Exceptions;

/// <summary>
/// Excepción lanzada cuando no se encuentra un archivo fuente requerido.
/// </summary>
public class ArchivoFuenteNoEncontradoException : Exception
{
    /// <summary>
    /// Inicializa la excepción con un mensaje.
    /// </summary>
    /// <param name="message">Mensaje de error.</param>
    public ArchivoFuenteNoEncontradoException(string message) : base(message)
    {
    }

    /// <summary>
    /// Inicializa la excepción con un mensaje y una excepción interna.
    /// </summary>
    /// <param name="message">Mensaje de error.</param>
    /// <param name="inner">Excepción interna.</param>
    public ArchivoFuenteNoEncontradoException(string message, Exception inner) : base(message, inner)
    {
    }
}
