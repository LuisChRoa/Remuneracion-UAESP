using System.Globalization;
using System.Text;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 42 (T2, R-REP-1, D-E): emite el REPORTE DURABLE de divergencias declaradas como un archivo
/// de texto junto a la salida (mismo directorio, nombre derivado del archivo generado). El texto es
/// de negocio — legible por el administrativo, sin jerga de código (prohibido <c>ERR-*</c>,
/// <c>&lt;f&gt;</c>, nombres de clase o códigos internos). Es la primera de las DOS huellas
/// obligatorias del cierre con declarada (la segunda es el log Serilog); sin reporte NO hay ÉXITO.
/// Plan 43 (O-3): registra además QUIÉN aceptó (usuario del sistema) y POR QUÉ CANAL (UI o CLI).
/// </summary>
public sealed class ReporteDivergenciasDeclaradasWriter : IReporteDivergenciasDeclaradas
{
    /// <inheritdoc />
    public string Emitir(
        string rutaSalida,
        Periodo periodo,
        Guid runId,
        ClasificacionDivergenciasDeclaradas clasificacion,
        string? usuario = null,
        string? canal = null)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(clasificacion);
        if (string.IsNullOrWhiteSpace(rutaSalida))
        {
            throw new ArgumentException("Debe indicarse la ruta de salida para emitir el reporte de divergencias declaradas.", nameof(rutaSalida));
        }

        var rutaAbsoluta = Path.GetFullPath(rutaSalida);
        var directorio = Path.GetDirectoryName(rutaAbsoluta) ?? ".";
        var nombreBase = Path.GetFileNameWithoutExtension(rutaAbsoluta);
        var rutaReporte = Path.Combine(directorio, $"{nombreBase} - Divergencias declaradas.txt");

        var texto = ConstruirTexto(rutaAbsoluta, periodo, runId, clasificacion, usuario, canal);
        File.WriteAllText(rutaReporte, texto, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        return rutaReporte;
    }

    /// <summary>
    /// Construye el texto del reporte (puro, testeable sin I/O). Contenido mandatorio (R-REP-1):
    /// celda, observado vs esperado, Δ, dueño, cita de evidencia, fecha y qué se reconoció.
    /// Plan 43 (O-3): además quién aceptó (<paramref name="usuario"/>) y por qué canal
    /// (<paramref name="canal"/>), resueltos de forma honesta cuando no llegan.
    /// </summary>
    public static string ConstruirTexto(
        string rutaSalida,
        Periodo periodo,
        Guid runId,
        ClasificacionDivergenciasDeclaradas clasificacion,
        string? usuario = null,
        string? canal = null)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(clasificacion);

        var sb = new StringBuilder();
        var nombreArchivo = Path.GetFileName(rutaSalida);
        var quincenaTexto = periodo.NumeroQuincena == 1 ? "primera" : "segunda";
        var fecha = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        var usuarioTexto = ResolverUsuario(usuario);
        var canalTexto = string.IsNullOrWhiteSpace(canal) ? HuellaAceptacion.CanalIndefinido : canal;

        sb.AppendLine("REPORTE DE DIVERGENCIAS DECLARADAS");
        sb.AppendLine("==================================");
        sb.AppendLine();
        sb.AppendLine($"Período: {periodo.CodigoCompleto} ({quincenaTexto} quincena)");
        sb.AppendLine($"Archivo generado: {nombreArchivo}");
        sb.AppendLine($"Fecha del reporte: {fecha}");
        sb.AppendLine($"Identificador de la corrida: {runId}");
        sb.AppendLine();
        sb.AppendLine("Qué significa este reporte");
        sb.AppendLine("--------------------------");
        sb.AppendLine("La corrida cerró con una diferencia conocida y reconocida por el usuario. No es un error");
        sb.AppendLine("del cálculo: es una discrepancia entre el reporte oficial y el manual del administrativo,");
        sb.AppendLine("cuya decisión no está en manos de la aplicación. Queda registrada aquí y en el log de la");
        sb.AppendLine("corrida, y se retira cuando el responsable indicado la resuelva.");
        sb.AppendLine();
        sb.AppendLine("Diferencias reconocidas");
        sb.AppendLine("-----------------------");

        var grupos = clasificacion.Declaradas
            .GroupBy(v => (v.EntradaId, v.Celda))
            .Select(g => g.First())
            .ToList();

        var numero = 1;
        foreach (var veredicto in grupos)
        {
            var entrada = BuscarEntrada(veredicto.EntradaId);
            sb.AppendLine($"{numero}) Dónde: hoja VALIDACION_TOTAL, celda {veredicto.Celda}.");
            sb.AppendLine($"   Qué muestra el archivo: {FormatearMonto(veredicto.MontoObservado)}.");
            sb.AppendLine($"   Qué esperaba la validación: {FormatearMonto(veredicto.MontoEsperado)}.");
            sb.AppendLine($"   Diferencia: {FormatearMonto(veredicto.Delta)}.");
            sb.AppendLine($"   De quién es: {entrada?.Dueno ?? "Responsable externo por confirmar"}.");
            sb.AppendLine($"   Evidencia: {entrada?.CitaEvidencia ?? "Documentada en el catálogo de divergencias declaradas."}");
            sb.AppendLine($"   Fecha: {entrada?.Fecha ?? fecha}.");
            sb.AppendLine($"   Se retira cuando: {entrada?.CondicionRetiro ?? "El responsable externo resuelva la discrepancia."}");
            sb.AppendLine();
            numero++;
        }

        sb.AppendLine("Aceptación");
        sb.AppendLine("----------");
        sb.AppendLine($"Aceptado por: {usuarioTexto} (canal: {canalTexto}).");
        sb.AppendLine("El usuario reconoció explícitamente estas diferencias para ESTA corrida. La aceptación no");
        sb.AppendLine("se guarda ni se repite automáticamente: cada corrida vuelve a presentar la diferencia.");
        sb.AppendLine();

        return sb.ToString();
    }

    /// <summary>
    /// Plan 43 (O-3): resuelve el usuario de la aceptación de forma honesta. Usa el valor declarado
    /// si llega; si no, el usuario del sistema (<see cref="Environment.UserName"/>) y, si el entorno
    /// no lo expone, <see cref="HuellaAceptacion.UsuarioNoDisponible"/> (nunca un nombre inventado).
    /// </summary>
    private static string ResolverUsuario(string? usuario) =>
        !string.IsNullOrWhiteSpace(usuario)
            ? usuario
            : string.IsNullOrWhiteSpace(Environment.UserName) ? HuellaAceptacion.UsuarioNoDisponible : Environment.UserName;

    private static DivergenciasDeclaradasCatalogo.Entrada? BuscarEntrada(string? id)
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

    /// <summary>
    /// Formato de monto de negocio (miles con punto, decimales con coma), estable e independiente de
    /// la cultura del equipo (p. ej. <c>417.519,62</c>).
    /// </summary>
    private static string FormatearMonto(decimal valor) =>
        valor.ToString("#,##0.00", FormatoNegocio);

    private static readonly NumberFormatInfo FormatoNegocio = new()
    {
        NumberGroupSeparator = ".",
        NumberDecimalSeparator = ",",
        NumberGroupSizes = [3],
        NumberDecimalDigits = 2
    };
}
