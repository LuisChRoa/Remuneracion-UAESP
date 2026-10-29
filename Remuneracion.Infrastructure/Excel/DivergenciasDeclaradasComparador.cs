namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 62 (H5 — T5.1, D-J9 + R-Aclareo): catálogo VERSIONADO de divergencias DECLARADAS del
/// comparador salida-vs-manual (<see cref="ComparadorSalidaVsManual"/>).
///
/// <para>
/// Una DIVERGENCIA DECLARADA es una diferencia real entre la salida de la app y el archivo manual
/// del administrativo que NO es defecto de la app: la app reprodujo fielmente su fuente (el R10 del
/// período, R-Aclareo) y el manual conserva un dato heredado. No es una exclusión de forma (esas
/// viven en <see cref="ComparadorSalidaVsManual.EsExclusionVersionada"/>): se declara la CELDA y el
/// PAR DE VALORES que la constituyen, de modo que cualquier otro valor en esa misma celda vuelve a
/// ser divergencia inesperada (R-FALSO-POSITIVO: prohibido tapar una regresión con una entrada
/// ancha).
/// </para>
///
/// <para>
/// <b>Contrato de cada entrada:</b> Id estable con el prefijo <c>P62-</c>, hoja (normalizada sin el
/// token de período), celdas EXCLUSIVAS, predicado de coincidencia sobre el par app/manual, motivo
/// en lenguaje de negocio, cita de la evidencia (plan + T0/§), fecha de alta y condición de retiro.
/// Sin los seis datos no hay entrada. Añadir una divergencia futura = añadir una FILA citada.
/// </para>
///
/// <para>
/// <b>MECANISMO DE RETIRO (Plan 63 §7.1, ratificación del Ingeniero 2026-10-29):</b> una entrada
/// puede quedar <see cref="Entrada.Retirada"/> con su estado de retiro citado
/// (<see cref="Entrada.FechaRetiro"/>, <see cref="Entrada.MotivoRetiro"/>,
/// <see cref="Entrada.EvidenciaRetiro"/>) — mismo patrón que
/// <c>Remuneracion.Core.Models.DivergenciasDeclaradasCatalogo</c>. Una entrada retirada YA NO
/// aplica en <see cref="Buscar"/> (no hay verde silencioso ni divergencia fantasma) pero se
/// CONSERVA en <see cref="Entradas"/> como historia auditable: el mecanismo queda intacto y solo
/// cambia el dato. El retiro exige fecha + motivo + evidencia; <see cref="Entrada.RetiroDocumentado"/>
/// lo verifica.
/// </para>
/// </summary>
public static class DivergenciasDeclaradasComparador
{
    /// <summary>Prefijo obligatorio de los Id de este plan (formato verificado por fixture).</summary>
    public const string PrefijoId = "P62-";

    /// <summary>Hoja <c>DetRetri{AAAAMMQ}</c> normalizada (sin token de período).</summary>
    public const string HojaDetRetri = "DETRETRI";

    /// <summary>Hoja <c>DetValiRetri{AAAAMMQ}</c> normalizada (sin token de período).</summary>
    public const string HojaDetValiRetri = "DETVALIRETRI";

    /// <summary>
    /// Una divergencia declarada: celda(s) + par de valores que la constituyen + cita y condición de
    /// retiro. <see cref="Coincide"/> recibe los valores TAL COMO los publica el comparador
    /// (<c>Divergencia.ValorApp</c> / <c>ValorManual</c>) y decide si ESA diferencia es la declarada.
    /// Una entrada puede estar RETIRADA (el dato se corrigió en la fuente): se conserva como historia
    /// auditable y no aplica.
    /// </summary>
    public sealed record Entrada(
        string Id,
        string Hoja,
        IReadOnlyList<string> Celdas,
        string Motivo,
        Func<string?, string?, bool> Coincide,
        string Cita,
        string Fecha,
        string CondicionRetiro,
        bool Retirada = false,
        string? FechaRetiro = null,
        string? MotivoRetiro = null,
        string? EvidenciaRetiro = null)
    {
        /// <summary>Una entrada incompleta NO aplica (no se admite divergencia sin cita ni retiro).</summary>
        public bool Completa =>
            !string.IsNullOrWhiteSpace(Id)
            && Id.StartsWith(PrefijoId, StringComparison.Ordinal)
            && !string.IsNullOrWhiteSpace(Hoja)
            && Celdas.Count > 0
            && Celdas.All(c => !string.IsNullOrWhiteSpace(c))
            && !string.IsNullOrWhiteSpace(Motivo)
            && Coincide is not null
            && !string.IsNullOrWhiteSpace(Cita)
            && !string.IsNullOrWhiteSpace(Fecha)
            && !string.IsNullOrWhiteSpace(CondicionRetiro);

        /// <summary>
        /// Verdad si el estado de retiro está DOCUMENTADO (fecha + motivo + evidencia citados). Una
        /// entrada vigente siempre lo está; una retirada sin documentar es un dato incompleto que el
        /// fixture del catálogo rechaza (Plan 63 §7.1).
        /// </summary>
        public bool RetiroDocumentado =>
            !Retirada
            || (!string.IsNullOrWhiteSpace(FechaRetiro)
                && !string.IsNullOrWhiteSpace(MotivoRetiro)
                && !string.IsNullOrWhiteSpace(EvidenciaRetiro));

        /// <summary>
        /// Verdad si la entrada APLICA hoy: completa, no retirada y con el retiro (si lo hay)
        /// documentado. El comparador solo considera entradas vigentes.
        /// </summary>
        public bool Vigente => Completa && !Retirada && RetiroDocumentado;
    }

    /// <summary>
    /// Catálogo (DATOS citados; sin ramas por período ni por ASE). Una sola entrada VIGENTE: el pie
    /// de interventoría heredado de <c>C16</c>. Las dos entradas de errata del peso (<c>D11</c> y
    /// <c>D14</c> de septiembre-1) quedan RETIRADAS desde 2026-10-29 (Plan 63 §7.1, ratificación del
    /// Ingeniero): el administrativo corrigió el dato en su manual y la divergencia ya no existe.
    /// </summary>
    public static readonly IReadOnlyList<Entrada> Entradas =
    [
        new Entrada(
            Id: "P62-DETRETRI-C16-PIE-INTERVENTORIA",
            Hoja: HojaDetRetri,
            Celdas: ["C16"],
            Motivo:
                "Pie de interventoría: el R10 del período que se ejecuta trae la REMUNERACIÓN del "
                + "período vigente y el manual del administrativo conserva la del período heredado. "
                + "El pegado reproduce el R10 (R-Aclareo); no es defecto de la app.",
            Coincide: CoincidePieInterventoria,
            Cita:
                "Plan 62 §3 T5.1 + §11 R-Aclareo (los encabezados/pies viajan pegados desde el R10) + "
                + "§Doctrina D-G3/D-61R10; valores medidos del caso en "
                + "Docs/T5-Plan62-Reclasificacion-Rojos.md (§2, fila 1) y Docs/T0-Plan62-Pegado-R10.md. "
                + "Valor manual re-verificado el 2026-10-29 sobre el manual corregido por el negocio: "
                + "C16 conserva el pie de interventoría con el período heredado (valor y cita medidos en "
                + "Docs/Plan63-Cierre.md §7.1).",
            Fecha: "2026-10-29",
            CondicionRetiro:
                "Se retira cuando el administrativo actualice el pie de interventoría de su plantilla "
                + "manual al período que ejecuta (o cuando el pie deje de viajar en el bloque pegado)."),

        new Entrada(
            Id: "P62-DETRETRI-D11-ERRATA-PESO-ASE3",
            Hoja: HojaDetRetri,
            Celdas: ["D11"],
            Motivo:
                "Errata humana del peso en el ASE 3: el R10 vigente ya traía el valor redondeado por "
                + "la regla única y el manual del administrativo conservaba el entero superior "
                + "histórico (Δ = 1 peso exacto). RETIRADA: el administrativo corrigió su manual.",
            Coincide: CoincideUnPeso,
            Cita:
                "Plan 62 §3 T5.1 + §11 R-Aclareo + Plan 61 §13.2-C (el caso del peso humano); valores "
                + "medidos del caso en Docs/T5-Plan62-Reclasificacion-Rojos.md (§2, fila 24).",
            Fecha: "2026-10-29",
            CondicionRetiro:
                "Se retira cuando el administrativo corrija su DetRetri manual al entero de la regla "
                + "única de redondeo o reemita el período desde la salida de la app. CUMPLIDA.",
            Retirada: true,
            FechaRetiro: "2026-10-29",
            MotivoRetiro:
                "El dato fue CORREGIDO en la fuente: el manual del administrativo de septiembre-1 ya "
                + "trae DetRetri!D11 = 10.891.338.739 (idéntico al R10 y a la salida de la app), así "
                + "que la divergencia de 1 peso YA NO EXISTE y declararla taparía una regresión real. "
                + "El Ingeniero ratificó el 2026-10-29 mantener el manual corregido.",
            EvidenciaRetiro:
                "Docs/Plan63-Cierre.md §7.1 (SHA del manual 9EF8F023…87383AC6, D11 medido "
                + "10.891.338.739 y D14 63.145.073.622 el 2026-10-29, ratificación del Ingeniero); "
                + "fixture ComparadorPegadoH5Plan62Tests.T5_1_Septiembre1_LaSuperficieDelPegadoSoloTraeLasDeclaradas "
                + "(observa 1 sola declarada aplicada: C16)."),

        new Entrada(
            Id: "P62-DETRETRI-D14-TOTAL-ERRATA-PESO",
            Hoja: HojaDetRetri,
            Celdas: ["D14"],
            Motivo:
                "Arrastre al TOTAL de la fila de cierre de la errata del peso del ASE 3: mismo Δ = 1 "
                + "peso entre el R10 (suma de sus filas) y el manual heredado. RETIRADA: el "
                + "administrativo corrigió su manual.",
            Coincide: CoincideUnPeso,
            Cita:
                "Plan 62 §3 T5.1 + §11 R-Aclareo; medición en "
                + "Docs/T5-Plan62-Reclasificacion-Rojos.md (§2, fila 24, mismo peso de la errata).",
            Fecha: "2026-10-29",
            CondicionRetiro:
                "Se retira junto con P62-DETRETRI-D11-ERRATA-PESO-ASE3 (misma causa, mismo retiro). "
                + "CUMPLIDA.",
            Retirada: true,
            FechaRetiro: "2026-10-29",
            MotivoRetiro:
                "El dato fue CORREGIDO en la fuente: el manual del administrativo de septiembre-1 ya "
                + "trae DetRetri!D14 = 63.145.073.622 (idéntico al R10 y a la salida de la app); el "
                + "arrastre de la errata al TOTAL desapareció con su causa. Retiro conjunto ratificado "
                + "por el Ingeniero el 2026-10-29.",
            EvidenciaRetiro:
                "Docs/Plan63-Cierre.md §7.1 (D14 medido 63.145.073.622 el 2026-10-29); fixture "
                + "ComparadorPegadoH5Plan62Tests.T5_1_Septiembre1_LaSuperficieDelPegadoSoloTraeLasDeclaradas "
                + "(1 sola declarada aplicada: C16).")
    ];

    /// <summary>Entradas que APLICAN hoy (completas, no retiradas); es lo que consulta el comparador.</summary>
    public static IReadOnlyList<Entrada> Vigentes => Entradas.Where(e => e.Vigente).ToList();

    /// <summary>Entradas retiradas (historia auditable; nunca aplican al comparador).</summary>
    public static IReadOnlyList<Entrada> Retiradas => Entradas.Where(e => e.Retirada).ToList();

    /// <summary>
    /// Devuelve la entrada que declara EXACTAMENTE esta diferencia, o <c>null</c> si es inesperada.
    /// La hoja se compara normalizada (sin token de período) y la celda de forma exacta. Las entradas
    /// retiradas se ignoran (Plan 63 §7.1: una retirada no aplica).
    /// </summary>
    public static Entrada? Buscar(string hoja, string celda, string? valorApp, string? valorManual)
    {
        var hojaNormalizada = ComparadorSalidaVsManual.NormalizarHoja(hoja);
        foreach (var entrada in Entradas)
        {
            if (!entrada.Vigente
                || !string.Equals(entrada.Hoja, hojaNormalizada, StringComparison.OrdinalIgnoreCase)
                || !entrada.Celdas.Contains(celda, StringComparer.OrdinalIgnoreCase))
            {
                continue;
            }

            if (entrada.Coincide(valorApp, valorManual))
            {
                return entrada;
            }
        }

        return null;
    }

    /// <summary>
    /// Pie de interventoría: mismo texto salvo el token de período de la REMUNERACIÓN. Cualquier otra
    /// diferencia de texto en C16 NO es la declarada.
    /// </summary>
    private static bool CoincidePieInterventoria(string? app, string? manual)
    {
        const string prefijo = "* COSTOS DE INTERVENTORIA: REMUNERACION ";
        var a = ComparadorSalidaVsManual.NormalizarTexto(app ?? string.Empty);
        var m = ComparadorSalidaVsManual.NormalizarTexto(manual ?? string.Empty);
        if (!a.StartsWith(prefijo, StringComparison.Ordinal) || !m.StartsWith(prefijo, StringComparison.Ordinal))
        {
            return false;
        }

        var tokenApp = a[prefijo.Length..].Trim();
        var tokenManual = m[prefijo.Length..].Trim();
        return tokenApp.Length > 0
            && tokenManual.Length > 0
            && !string.Equals(tokenApp, tokenManual, StringComparison.Ordinal);
    }

    /// <summary>
    /// Δ = 1 peso EXACTO y entero: el par de la errata humana documentada. Cualquier otro Δ (0, 2+,
    /// fraccionario) NO se declara — una entrada ancha taparía una regresión real.
    /// </summary>
    private static bool CoincideUnPeso(string? app, string? manual)
    {
        var a = Importe(app);
        var m = Importe(manual);
        return a is not null
            && m is not null
            && a.Value == decimal.Truncate(a.Value)
            && m.Value == decimal.Truncate(m.Value)
            && m.Value - a.Value == 1m;
    }

    /// <summary>Parsea un importe es-CO ('$17.450.250.543', '-$5.093.259.797', '739').</summary>
    private static decimal? Importe(string? valor)
    {
        if (string.IsNullOrWhiteSpace(valor))
        {
            return null;
        }

        var limpio = ComparadorSalidaVsManual.NormalizarTexto(valor)
            .Replace("$", string.Empty, StringComparison.Ordinal)
            .Replace(" ", string.Empty, StringComparison.Ordinal);
        if (limpio.Length == 0)
        {
            return null;
        }

        var negativo = limpio.StartsWith('-');
        limpio = limpio.Trim('-');
        limpio = limpio.Replace(".", string.Empty, StringComparison.Ordinal).Replace(',', '.');
        return decimal.TryParse(limpio, System.Globalization.NumberStyles.Number,
            System.Globalization.CultureInfo.InvariantCulture, out var importe)
            ? negativo ? -importe : importe
            : null;
    }
}
