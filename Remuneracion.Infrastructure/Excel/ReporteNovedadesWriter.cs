using System.Globalization;
using System.Text;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 61 (H4.1 / S3.3, §2-D-H4): emite el REPORTE DE NOVEDADES como archivo de texto junto a la
/// salida. Es el artefacto que el administrativo lee para saber, sin abrir nada más, que la corrida
/// <b>no está certificada</b>, qué hay que revisar y cómo hacerlo.
///
/// <para><b>Lenguaje de negocio (G-TRAZA).</b> Prohibido <c>ERR-*</c>, <c>&lt;f&gt;</c>, nombres de
/// clase y códigos internos: aquí solo aparecen períodos, hojas reales de la plantilla, etiquetas de
/// ASE/empresa/concepto, montos y qué debe hacer la persona que revisa.</para>
///
/// <para><b>Cinco secciones.</b> (1) veredicto inequívoco; (2) datos de la corrida; (3) una ficha por
/// novedad; (4) guía de revisión por etiqueta contra el administrativo y el R10; (5) divergencias
/// declaradas, si las hay (un solo reporte con secciones: prohibido tener dos TXT que se
/// contradigan).</para>
///
/// <para><b>Determinismo.</b> La ruta se deriva del nombre definitivo de la salida, así que la misma
/// corrida reescribe el mismo archivo; UTF-8 sin BOM (precedente Plan 42). La fecha y hora son
/// inyectables para que los tests puedan fijarlas.</para>
/// </summary>
public sealed class ReporteNovedadesWriter : IReporteNovedades
{
    /// <summary>Nombre (sufijo) del reporte cuando hay novedades de DATOS (D-TXT ratificada).</summary>
    public const string SufijoNovedades = " - Novedades de la corrida.txt";

    /// <inheritdoc />
    public string Emitir(
        string rutaSalida,
        string rutaExcelEntregado,
        Periodo periodo,
        Guid runId,
        IReadOnlyList<NovedadEntrega> novedades,
        ClasificacionDivergenciasDeclaradas? declaradas = null,
        string? usuario = null,
        string? canal = null,
        DateTime? fechaHora = null)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(novedades);
        if (string.IsNullOrWhiteSpace(rutaSalida))
        {
            throw new ArgumentException("Debe indicarse la ruta de salida para emitir el reporte de novedades.", nameof(rutaSalida));
        }

        if (novedades.Count == 0)
        {
            throw new ArgumentException("El reporte de novedades exige al menos una novedad; una corrida sin novedades se certifica y no emite este reporte.", nameof(novedades));
        }

        var rutaAbsoluta = Path.GetFullPath(rutaSalida);
        var directorio = Path.GetDirectoryName(rutaAbsoluta) ?? ".";
        var nombreBase = Path.GetFileNameWithoutExtension(rutaAbsoluta);
        var rutaReporte = Path.Combine(directorio, $"{nombreBase}{SufijoNovedades}");

        var texto = ConstruirTexto(rutaExcelEntregado, periodo, runId, novedades, declaradas, usuario, canal, fechaHora);
        File.WriteAllText(rutaReporte, texto, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        return rutaReporte;
    }

    /// <summary>
    /// Construye el texto del reporte (puro, testeable sin I/O).
    /// </summary>
    public static string ConstruirTexto(
        string rutaExcelEntregado,
        Periodo periodo,
        Guid runId,
        IReadOnlyList<NovedadEntrega> novedades,
        ClasificacionDivergenciasDeclaradas? declaradas = null,
        string? usuario = null,
        string? canal = null,
        DateTime? fechaHora = null)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(novedades);

        var sb = new StringBuilder();

        // ── 1. Veredicto inequívoco ──────────────────────────────────────────────────────────────
        sb.AppendLine("ENTREGADO CON NOVEDADES — NO CERTIFICADO");
        sb.AppendLine("========================================");
        sb.AppendLine();
        // ── 2. Una ficha por novedad ─────────────────────────────────────────────────────────────
        sb.AppendLine($"Detalle de las novedades ({novedades.Count})");
        sb.AppendLine("---------------------------------");
        var numero = 1;
        foreach (var novedad in novedades)
        {
            sb.AppendLine($"{numero}) {Ubicacion(novedad)}");
            sb.AppendLine($"   Concepto: {novedad.ConceptoOCelda}");
            if (novedad.ValorObservado is not null)
            {
                sb.AppendLine($"   Qué muestra el Excel entregado: {FormatearMonto(novedad.ValorObservado.Value)}");
            }

            if (novedad.ValorEsperado is not null)
            {
                sb.AppendLine($"   Qué esperaba la revisión: {FormatearMonto(novedad.ValorEsperado.Value)}");
            }

            if (novedad.Delta is not null)
            {
                sb.AppendLine($"   Diferencia (lo entregado menos lo esperado): {FormatearMonto(novedad.Delta.Value)}");
            }

            if (!string.IsNullOrWhiteSpace(novedad.Explicacion))
            {
                sb.AppendLine($"   Qué pasó: {novedad.Explicacion}");
            }

            sb.AppendLine($"   De quién es: {novedad.Dueno}");
            sb.AppendLine($"   Qué revisar: {novedad.QueRevisar}");
            sb.AppendLine();
            numero++;
        }

        // ── 3. Guía de revisión ─────────────────────────────────────────────────────────────────
        sb.AppendLine("Cómo revisar");
        sb.AppendLine("------------");
        sb.AppendLine("1. Abra el Excel entregado (el que lleva \"(sin validar)\" en el nombre).");
        sb.AppendLine("2. Para cada novedad, vaya a la hoja que se indica arriba y busque el ASE, la empresa o el");
        sb.AppendLine("   concepto POR SU NOMBRE (su etiqueta).");
        sb.AppendLine("3. Compare esa cifra contra el reporte oficial del período (R10) y contra el archivo del");
        sb.AppendLine("   administrativo, también por etiqueta.");
        sb.AppendLine("4. Anote la diferencia y decida: si es una diferencia real de datos, corríjala en su fuente;");
        sb.AppendLine("   si es una diferencia de transcripción del reporte oficial, corríjala en el reporte.");
        sb.AppendLine("5. Vuelva a ejecutar la aplicación. Cuando no queden diferencias, el archivo saldrá con su");
        sb.AppendLine("   nombre definitivo y sin este documento.");
        sb.AppendLine();

        // ── 4. Divergencias declaradas (si aplica) ───────────────────────────────────────────────
        var declaradasDistintas = declaradas?.Declaradas
            .GroupBy(v => (v.EntradaId, v.Celda))
            .Select(g => g.First())
            .ToList() ?? [];
        if (declaradasDistintas.Count > 0)
        {
            sb.AppendLine("Divergencias declaradas presentes en esta corrida");
            sb.AppendLine("-----------------------------------------------");
            sb.AppendLine("Son diferencias conocidas, con responsable asignado fuera de la aplicación. Se registran aquí");
            sb.AppendLine("junto con las novedades para que no queden dos documentos que digan cosas distintas.");
            sb.AppendLine();
            foreach (var veredicto in declaradasDistintas)
            {
                var entrada = BuscarEntrada(veredicto.EntradaId);
                sb.AppendLine($"· Código {veredicto.EntradaId ?? "sin código"} · ASE {veredicto.AseId} · hoja VALIDACION_TOTAL, bloque {veredicto.Celda}");
                sb.AppendLine($"  Qué muestra el archivo: {FormatearMonto(veredicto.MontoObservado)}; qué esperaba la validación: {FormatearMonto(veredicto.MontoEsperado)}; diferencia: {FormatearMonto(veredicto.MontoObservado - veredicto.MontoEsperado)}.");
                sb.AppendLine($"  De quién es: {entrada?.Dueno ?? "Responsable externo por confirmar"}.");
                sb.AppendLine($"  Evidencia: {entrada?.CitaEvidencia ?? "Documentada en el registro de divergencias conocidas."}");
                sb.AppendLine($"  Se retira cuando: {entrada?.CondicionRetiro ?? "El responsable externo resuelva la discrepancia."}");
                sb.AppendLine();
            }
        }


        return sb.ToString();
    }

    private static string Ubicacion(NovedadEntrega novedad)
    {
        var ase = novedad.AseId == 0 ? "Período completo (los 5 ASE)" : $"ASE {novedad.AseId}";
        var hoja = string.IsNullOrWhiteSpace(novedad.Hoja) ? "sin hoja asociada" : $"hoja {novedad.Hoja}";
        var reporte = string.IsNullOrWhiteSpace(novedad.Reporte) ? "revisión de la corrida" : novedad.Reporte;
        return $"{ase} · {reporte} · {hoja}";
    }

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
    /// Monto de negocio: miles con punto, decimales con coma, sin ceros de relleno (739 y no 739,00;
    /// −178.619.709,53 con sus centavos cuando existen).
    /// </summary>
    private static string FormatearMonto(decimal valor) => valor.ToString("#,##0.####", FormatoNegocio);

    private static readonly NumberFormatInfo FormatoNegocio = new()
    {
        NumberGroupSeparator = ".",
        NumberDecimalSeparator = ",",
        NumberGroupSizes = [3],
        NumberDecimalDigits = 4,
    };
}
