namespace Remuneracion.Core.Models;

/// <summary>Hoja de detalle del contrato R2/R4 (discriminador puro; el nombre de hoja vive en el mapa).</summary>
public enum HojaDetalleR2R4
{
    /// <summary><c>Rem. Anticipos R2</c>.</summary>
    R2,

    /// <summary><c>Reversion Pagos R4</c>.</summary>
    R4,
}

/// <summary>
/// Plan 47 (H1.4 — §1.4 de la autorización; doctrina <b>E-44.1</b> ya ratificada en el frente R1):
/// lista documentada de las filas del detalle R2/R4 que quedan probadas <b>sin contraparte</b> en el
/// bloque de la plantilla oficial.
///
/// <para>
/// Una fila-fuente cuya clave <c>(firma + canal efectivo)</c> no encuentra slot:
/// </para>
/// <list type="bullet">
/// <item>si pertenece a esta lista → <b>cero documentado</b> (no se escribe, no es error, no se
/// simula) y queda reportada aquí con la evidencia citada (bytes del T0 del Plan 47);</item>
/// <item>si NO pertenece a esta lista → <c>ERR-SIN-SLOT</c> accionable (hoja + ASE + firma + canal),
/// sin remap ni colocación por posición (Plan 41 E1-1).</item>
/// </list>
///
/// La identidad es <c>(hoja, ASE, firma, canal efectivo)</c> con una <see cref="Entrada.Cantidad"/>
/// de ocurrencias: period-free (no hay literales de período) y company-free en el predicado (el ASE es
/// el mismo índice de bloque que ya declara el mapa geométrico; el canal es DATO de la fila).
/// </summary>
public static class CerosDocumentadosDetalleR2R4
{
    /// <summary>Entrada documentada: clave + ocurrencias + motivo + evidencia citada (bytes).</summary>
    public sealed record Entrada(
        HojaDetalleR2R4 Hoja,
        int Ase,
        string Firma,
        string? Canal,
        int Cantidad,
        string Motivo,
        string Evidencia);

    private const string MotivoCanalA =
        "El bloque OFICIAL no enumera el canal `A`/empresa `EAB` en la mitad Subs/Cont (bytes: ASE1 22..31 arranca en `Subsidio` cerrado por `C=E`): la fuente de agosto trae ese grupo y el contrato no lo tiene.";

    /// <summary>
    /// Entradas VIGENTES (**3 filas fuente = 3 claves**, todas de agosto-2). El Plan 48 (H0/T0 §2.2)
    /// probó por bytes que la geometría v2 **sí tiene slot** para el canal 9 y el `Vlr Intereses` del
    /// canal 2 de ASE2/ASE4: las 9 filas de julio-2 que el Plan 47 documentaba eran un **falso hueco**
    /// y quedaron RETIRADAS (no sobrevive clave sin prueba de ausencia en bytes contra la v2). Las 3
    /// claves A/EAB de agosto-2 siguen probadas ausentes en la v2 (la mitad Subs/Cont de ASE4 no trae
    /// el grupo `A`/`EAB`). Evidencia: <c>Docs/T0-Plan48-Adopcion-Plantilla-v2.md</c> §2.2/§2.4.
    /// </summary>
    public static IReadOnlyList<Entrada> Entradas { get; } =
    [
        // R2 ASE4 (BOGOTA LIMPIA) — 3 filas de agosto-2 sin slot en la geometría v2.
        new(HojaDetalleR2R4.R2, 4, "|||Subsidio(-)/Contribucion(+)", "A/EAB", 1, MotivoCanalA,
            "T0-Plan48 §2.2: la mitad Subs/Cont del bloque v2 de ASE4 no trae el grupo A/EAB; fuente agosto-2 `src#28` Total = -12.462,26."),
        new(HojaDetalleR2R4.R2, 4, "||A|Total", "A/EAB", 1, MotivoCanalA,
            "T0-Plan48 §2.2: la mitad Subs/Cont del bloque v2 de ASE4 no trae `||A|Total`; fuente agosto-2 `src#29` Total = -12.462,26."),
        new(HojaDetalleR2R4.R2, 4, "|EAB|Total|", null, 1, MotivoCanalA,
            "T0-Plan48 §2.2: la mitad Subs/Cont del bloque v2 de ASE4 no trae `|EAB|Total|`; fuente agosto-2 `src#30` Total = -12.462,26."),
    ];

    /// <summary>Total de filas-fuente documentadas (suma de <see cref="Entrada.Cantidad"/>).</summary>
    public static int TotalFilasDocumentadas => Entradas.Sum(e => e.Cantidad);

    /// <summary>
    /// Ocurrencias documentadas para la clave exacta <c>(hoja, ASE, firma, canal)</c>. <c>0</c> = fuera
    /// de la lista → el consumidor debe fallar con <c>ERR-SIN-SLOT</c>.
    /// </summary>
    public static int CantidadDocumentada(HojaDetalleR2R4 hoja, int ase, string firma, string? canal)
    {
        ArgumentNullException.ThrowIfNull(firma);

        var total = 0;
        foreach (var entrada in Entradas)
        {
            if (entrada.Hoja != hoja || entrada.Ase != ase)
            {
                continue;
            }

            if (!string.Equals(entrada.Firma, firma.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!string.Equals(entrada.Canal?.Trim(), canal?.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            total += entrada.Cantidad;
        }

        return total;
    }
}
