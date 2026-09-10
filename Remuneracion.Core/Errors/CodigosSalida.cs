namespace Remuneracion.Core.Errors;

/// <summary>
/// HU-14 (3.1, D3): contrato de códigos de salida para HU-15 (modo CLI). El contrato vive en
/// Core sin dependencias; <c>Form1</c> lo REGISTRA (status + log, <c>UltimoCodigoSalida</c>)
/// pero NUNCA lo emite con <c>Environment.Exit</c> (V6: WinForms es GUI); HU-15 lo emitirá.
/// </summary>
public static class CodigosSalida
{
    /// <summary>Proceso completado correctamente.</summary>
    public const int Ok = 0;

    /// <summary>Fallo de validación (ERR-VALIDACION).</summary>
    public const int Validacion = 1;

    /// <summary>Fallo de fuente o plantilla (ERR-FUENTE-NO-ENCONTRADA / ERR-PLANTILLA / ERR-FORMATO-FUENTE).</summary>
    public const int FuenteOPlantilla = 2;

    /// <summary>Fallo de escritura (ERR-ESCRITURA).</summary>
    public const int Escritura = 3;

    /// <summary>Error inesperado (ERR-INESPERADO).</summary>
    public const int Inesperado = 4;

    /// <summary>Cancelado por el usuario (WARN-CANCELADO).</summary>
    public const int CanceladoPorUsuario = 5;
}
