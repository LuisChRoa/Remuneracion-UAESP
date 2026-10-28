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

    /// <summary>
    /// Plan 42 (T1/T3, D-G): la corrida CERRÓ con una divergencia DECLARADA reconocida
    /// explícitamente por el usuario (UI: acto consciente; CLI: <c>--aceptar-divergencia-declarada</c>).
    /// DISTINTO de <see cref="Ok"/> por diseño: el orquestador del lote debe poder distinguir un
    /// cierre limpio de uno con deuda externa reconocida (R-CLI-CIEGO). El reporte durable y el log
    /// quedan emitidos. Sin el reconocimiento, la declarada bloquea con <see cref="Validacion"/>.
    /// </summary>
    public const int DivergenciaDeclaradaReconocida = 6;

    /// <summary>
    /// Plan 61 (H4.2 / S3.4, D-CODIGO): la corrida ENTREGÓ el Excel marcado <c>(sin validar)</c> más
    /// el reporte de novedades, con al menos una novedad de DATOS sin resolver. Es un estado de
    /// entrega, NO un error: el orquestador del lote lo trata como "revisar y seguir". Distinto de
    /// <see cref="Ok"/> (0) y de <see cref="DivergenciaDeclaradaReconocida"/> (6) por diseño; el
    /// contrato 0..6 queda intacto (este valor se suma, no se renumera).
    /// </summary>
    public const int NovedadesEntrega = 7;
}
