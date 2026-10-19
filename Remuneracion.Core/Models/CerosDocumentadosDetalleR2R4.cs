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

    private const string MotivoCanal9 =
        "El bloque OFICIAL no tiene fila de cierre `C=9 D=Total` (ni `Subsidio` del canal 9): la fuente trae el canal 9 y el contrato de la plantilla no lo enumera.";

    private const string MotivoCanalA =
        "El bloque OFICIAL no enumera el canal `A`/empresa `EAB` en la mitad Subs/Cont (bytes: ASE1 22..31 arranca en `Subsidio` cerrado por `C=E`): la fuente de agosto trae ese grupo y el contrato no lo tiene.";

    private const string MotivoInteresesCanal2 =
        "El grupo del canal 2 del bloque OFICIAL tiene UNA sola fila (`Vlr Servicio`) entre los cierres `1` y `2`: no existe slot para `Vlr Intereses`.";

    /// <summary>
    /// Entradas VIGENTES (14 filas fuente = 11 claves): julio-2 (9 filas) + agosto-2 (5 filas). Las
    /// claves con 2 ocurrencias (<c>||9|Total</c> de ASE2, <c>||A|Total</c>/<c>|EAB|Total|</c> de ASE4)
    /// corresponden a los dos mitades del bloque (Anticipos y Subs/Cont).
    /// Evidencia completa en <c>Docs/T0-Plan47-Revision-Pestanas-LenteR1.md</c> §6.1/§6.4.
    /// </summary>
    public static IReadOnlyList<Entrada> Entradas { get; } =
    [
        // R2 ASE2 (LIME) — 4 filas: julio-2 (3) + agosto-2 (1).
        new(HojaDetalleR2R4.R2, 2, "|||Vlr Intereses", "2/NUEVO ESQUEMA", 1, MotivoInteresesCanal2,
            "T0-Plan47 §6.1: oficial R2 67..113 con una sola fila (78) entre `||1|Total` y `||2|Total`; fuente julio-2 `src#8` Total = 406,42."),
        new(HojaDetalleR2R4.R2, 2, "|||Vlr Servicio", "9/NUEVO ESQUEMA", 1, MotivoCanal9,
            "T0-Plan47 §6.1: no existe fila `C=9 D=Total` en R2 ASE2; fuente julio-2 `src#13` Total = 831.919,29 y agosto-2 `src#12` Total = 81.263,29 (ambos incluidos en `NUEVO ESQUEMA`)."),
        new(HojaDetalleR2R4.R2, 2, "||9|Total", "9/NUEVO ESQUEMA", 2, MotivoCanal9,
            "T0-Plan47 §6.1/§6.4: no existe fila `C=9 D=Total` en R2 ASE2; fuente julio-2 `src#14` Total = 831.919,29 y agosto-2 `src#13` Total = 81.263,29 (mitades Anticipos y Subs/Cont)."),
        new(HojaDetalleR2R4.R2, 2, "|||Subsidio(-)/Contribucion(+)", "9/NUEVO ESQUEMA", 1, MotivoCanal9,
            "T0-Plan47 §6.4: la mitad Subs/Cont del bloque oficial de ASE2 enumera E/1/2/5/H/O/T (sin `9`); fuente agosto-2 `src#35` Total = 6.719,84."),

        // R2 ASE4 (BOGOTA LIMPIA) — 10 filas: julio-2 (6) + agosto-2 (4).
        new(HojaDetalleR2R4.R2, 4, "|||Vlr Intereses", "2/NUEVO ESQUEMA", 1, MotivoInteresesCanal2,
            "T0-Plan47 §6.1: oficial R2 282..328 con una sola fila (293) entre `||1|Total` y `||2|Total`; fuente julio-2 `src#8` Total = -8.785,59 y agosto-2 `src#11` Total = -1.647,69."),
        new(HojaDetalleR2R4.R2, 4, "|||Vlr Servicio", "9/NUEVO ESQUEMA", 1, MotivoCanal9,
            "T0-Plan47 §6.1: no existe fila `C=9 D=Total` en R2 ASE4; fuente julio-2 `src#13` Total = 20.152.773,11."),
        new(HojaDetalleR2R4.R2, 4, "|||Vlr Intereses", "9/NUEVO ESQUEMA", 1, MotivoCanal9,
            "T0-Plan47 §6.1: no existe fila `C=9 D=Total` en R2 ASE4; fuente julio-2 `src#14` Total = -2.377,86."),
        new(HojaDetalleR2R4.R2, 4, "||9|Total", "9/NUEVO ESQUEMA", 2, MotivoCanal9,
            "T0-Plan47 §6.1: no existe fila `C=9 D=Total` en R2 ASE4; fuente julio-2 `src#15` Total = 20.150.395,25 (mitad Anticipos) y `src#38` Total = 5.663.564,75 (mitad Subs/Cont)."),
        new(HojaDetalleR2R4.R2, 4, "|||Subsidio(-)/Contribucion(+)", "9/NUEVO ESQUEMA", 1, MotivoCanal9,
            "T0-Plan47 §6.1: no existe fila `C=9 D=Total` en la mitad Subs/Cont de R2 ASE4; fuente julio-2 `src#37` Total = 5.663.564,75."),
        new(HojaDetalleR2R4.R2, 4, "|||Subsidio(-)/Contribucion(+)", "A/EAB", 1, MotivoCanalA,
            "T0-Plan47 §6.4: la mitad Subs/Cont del bloque oficial de ASE4 no trae el grupo A/EAB; fuente agosto-2 `src#28` Total = -12.462,26."),
        new(HojaDetalleR2R4.R2, 4, "||A|Total", "A/EAB", 1, MotivoCanalA,
            "T0-Plan47 §6.4: la mitad Subs/Cont del bloque oficial de ASE4 no trae `||A|Total`; fuente agosto-2 `src#29` Total = -12.462,26."),
        new(HojaDetalleR2R4.R2, 4, "|EAB|Total|", null, 1, MotivoCanalA,
            "T0-Plan47 §6.4: la mitad Subs/Cont del bloque oficial de ASE4 no trae `|EAB|Total|`; fuente agosto-2 `src#30` Total = -12.462,26."),
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
