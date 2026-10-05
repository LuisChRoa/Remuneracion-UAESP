using System.Text;

namespace Remuneracion.Core.Services;

/// <summary>
/// Plan 27 (T2, §2.1/D-C): veredicto de la firma binaria de un archivo insumo, antes de abrir
/// cualquier workbook. La regla es deliberadamente mínima (R-FALSO-POSITIVO): solo lo que el
/// runtime no puede leer de ninguna forma se reporta como inválido.
/// </summary>
public enum FirmaExcel
{
    /// <summary>
    /// Firma de xlsx (ZIP/PK <c>50 4B 03 04</c>) o xls legacy (OLE <c>D0 CF 11 E0</c>): el
    /// <c>ExcelDataReader</c> del proyecto abre ambos (D-D), por lo que se aceptan.
    /// </summary>
    Valida,

    /// <summary>
    /// El prefijo es texto con marcadores de página web (MHTML/HTML renombrado a <c>.xlsx</c>).
    /// Es el caso típico de una captura/descarga web que no es el archivo real.
    /// </summary>
    NoExcelPareceCopiaWeb,

    /// <summary>Cualquier otra firma: no es un Excel reconocible (archivo corrupto u otro formato).</summary>
    NoExcelFormatoDesconocido,

    /// <summary>
    /// No se pudo leer el archivo (error de I/O). Se trata como formato desconocido: el inspector
    /// nunca lanza su propia excepción (el runtime ya tiene su fail-fast).
    /// </summary>
    Ilegible
}

/// <summary>
/// Plan 27 (T2, D-C): inspector de firma binaria con BCL puro (<see cref="FileStream"/>), sin
/// <c>ExcelDataReader</c>, sin <c>OpenXml</c> ni referencia a Infrastructure. <see cref="Clasificar(byte[])"/>
/// es una función pura (testeable sin archivos); <see cref="Inspeccionar(string)"/> lee a lo sumo
/// <see cref="BytesALeer"/> bytes por archivo (nunca el archivo entero) y lo cierra de inmediato.
/// </summary>
public static class InspectorFirmaExcel
{
    /// <summary>
    /// Tope de lectura por archivo (magia binaria + sniff de texto acotado). Fijado en
    /// implementación y traído a revisión en el PR (Plan 27 §5): 512 B alcanzan para la firma y
    /// para los encabezados de un MHTML real.
    /// </summary>
    public const int BytesALeer = 512;

    // xlsx = contenedor ZIP: "PK\x03\x04".
    private static readonly byte[] FirmaXlsx = [0x50, 0x4B, 0x03, 0x04];

    // xls legacy = OLE Compound File: D0 CF 11 E0 (D-D: se acepta como válido).
    private static readonly byte[] FirmaXlsOle = [0xD0, 0xCF, 0x11, 0xE0];

    // Marcadores de texto web en minúsculas (el sniff se hace en minúsculas). El marcador "<"
    // suelto se evalúa aparte sobre el prefijo recortado (ver ContieneMarcadorWeb).
    private static readonly string[] MarcadoresTextoWeb =
        ["<!doctype", "<html", "mime-version", "saved by", "from:"];

    /// <summary>Clasificador puro: bytes → veredicto (sin I/O).</summary>
    /// <param name="bytes">Bytes del archivo (al menos el prefijo a inspeccionar).</param>
    public static FirmaExcel Clasificar(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        return Clasificar(bytes.AsSpan());
    }

    /// <summary>Clasificador puro sobre un span de bytes: PK/OLE → válido; texto web → copia web; resto → desconocido.</summary>
    /// <param name="bytes">Prefijo de bytes del archivo.</param>
    public static FirmaExcel Clasificar(ReadOnlySpan<byte> bytes)
    {
        if (EmpiezaCon(bytes, FirmaXlsx) || EmpiezaCon(bytes, FirmaXlsOle))
        {
            return FirmaExcel.Valida;
        }

        return ContieneMarcadorWeb(bytes)
            ? FirmaExcel.NoExcelPareceCopiaWeb
            : FirmaExcel.NoExcelFormatoDesconocido;
    }

    /// <summary>
    /// Lee a lo sumo <see cref="BytesALeer"/> bytes del archivo y clasifica su firma. Nunca lanza
    /// por error de I/O: devuelve <see cref="FirmaExcel.Ilegible"/> y los bytes leídos (vacíos si
    /// no se pudo leer) para el log técnico.
    /// </summary>
    /// <param name="rutaArchivo">Ruta del archivo a inspeccionar.</param>
    public static ResultadoFirmaExcel Inspeccionar(string rutaArchivo)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rutaArchivo);

        try
        {
            using var stream = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var buffer = new byte[BytesALeer];
            var leidos = stream.Read(buffer, 0, buffer.Length);
            if (leidos < buffer.Length)
            {
                Array.Resize(ref buffer, leidos);
            }

            return new ResultadoFirmaExcel(Clasificar(buffer), buffer);
        }
        catch (IOException)
        {
            return new ResultadoFirmaExcel(FirmaExcel.Ilegible, []);
        }
        catch (UnauthorizedAccessException)
        {
            return new ResultadoFirmaExcel(FirmaExcel.Ilegible, []);
        }
    }

    private static bool EmpiezaCon(ReadOnlySpan<byte> bytes, ReadOnlySpan<byte> firma) =>
        bytes.Length >= firma.Length && bytes[..firma.Length].SequenceEqual(firma);

    private static bool ContieneMarcadorWeb(ReadOnlySpan<byte> bytes)
    {
        if (bytes.IsEmpty)
        {
            return false;
        }

        // Latin1: 1 byte = 1 char; los marcadores son ASCII, así que el decode es fiel y no
        // depende del encoding real del archivo. El BOM UTF-8 (EF BB BF) se ve como "ï»¿" y se
        // recorta junto con espacios para tolerar un texto que "empieza con <" tras el BOM.
        var texto = Encoding.Latin1.GetString(bytes).ToLowerInvariant();
        var recortado = texto.TrimStart('\uFEFF', '\u00EF', '\u00BB', '\u00BF', ' ', '\t', '\r', '\n');
        if (recortado.StartsWith('<'))
        {
            return true;
        }

        return MarcadoresTextoWeb.Any(marcador => texto.Contains(marcador, StringComparison.Ordinal));
    }
}

/// <summary>
/// Resultado de <see cref="InspectorFirmaExcel.Inspeccionar(string)"/>: el veredicto y los bytes
/// efectivamente leídos (para el log técnico; nunca van al mensaje administrativo).
/// </summary>
/// <param name="Firma">Veredicto de la firma.</param>
/// <param name="PrimerosBytes">Bytes leídos (hasta <see cref="InspectorFirmaExcel.BytesALeer"/>).</param>
public readonly record struct ResultadoFirmaExcel(FirmaExcel Firma, byte[] PrimerosBytes);
