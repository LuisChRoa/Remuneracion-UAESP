using System.Globalization;
using System.Text.RegularExpressions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 29 (T5, Unidad P — R-P-2 / S6): regresión END-TO-END del comparador
/// <see cref="ComparadorSalidaVsManual"/> para los dos períodos con manual del administrativo.
///
/// Por período (julio-Q2 y agosto-Q2) corre UNA sola vez el flujo 5-ASE contra la PLANTILLA EN
/// CEROS del período (julio: <c>Docs/Plantilla_Remuneracion.xlsx</c>; agosto: base
/// <c>Plantilla_Remuneracion_2026082.xlsx</c>, Plan 30/T3) hacia una ruta TEMPORAL (jamás
/// sobrescribe <c>Docs/</c>) con los insumos REALES, y compara la salida contra el manual con el
/// comparador BCL (sin Excel/COM). Aserta que TODA divergencia cae en una BRECHA DECLARADA
/// (versionada abajo con su cita al plan/T0): cualquier divergencia fuera de ese conjunto falla
/// nombrando hoja+celda+ambos valores (S6), de modo que una regresión real no puede pasar en
/// silencio. El comparador ya aplica sus exclusiones versionadas (columnas Q1 en corrida Q2,
/// metadatos N3/C6/D6, <c>DetValiRetri!J</c> y nombres de hoja con sufijo de período).
///
/// El manual NO es paridad de libro completo: es un workbook mantenido a mano que acumula la 1.ª
/// quincena, recompone la malla R1/R2/R4 en agosto (remesh T0b/H1) y trae ediciones propias
/// (casillas cosméticas, notas). Por eso el criterio del plan es «cero divergencias fuera de
/// exclusiones + brechas declaradas», no igualdad byte a byte. La evidencia de la magnitud está en
/// la memoria de proyecto (julio ~6.672 celdas, agosto ~11.814 con Excel COM).
///
/// Sin emojis. Tolerancia ±0.5.
/// </summary>
public sealed class ComparadorSalidaVsManualTests
{
    private const decimal Tol = 0.5m;
    private const string HojaBce = "BCE SC POR FACT.";
    private const string HojaControl = "Valida - Control Recaudo";

    private static readonly string Plantilla = Path.Combine(Insumos.Raiz(), "Docs", "Plantilla_Remuneracion.xlsx");

    private static readonly Regex ReferenciaCelda = new(@"^(?<col>[A-Za-z]+)(?<fila>\d+)$", RegexOptions.Compiled);

    /// <summary>Token de período dentro de referencias de hoja (H1/T0g), p. ej. <c>DetRetri2026072</c>.</summary>
    private static readonly Regex TokenPeriodo = new(@"\d{6}-?\d?", RegexOptions.Compiled);

    // ── Casos ────────────────────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData("julio")]
    [InlineData("agosto")]
    public void Salida5AseQ2_vs_Manual_CeroDivergenciasFueraDeBrechasDeclaradas(string etiqueta)
    {
        var (periodo, carpeta, manual, plantilla) = Caso(etiqueta);
        Assert.True(File.Exists(plantilla), $"Falta la plantilla canónica ({etiqueta}): {plantilla}");
        Assert.True(File.Exists(manual), $"Falta el manual del administrativo ({etiqueta}): {manual}");

        using var salida = new SalidaTemporal("remuneracion-comparador-" + etiqueta + "-");
        EjecutarFlujo(carpeta, periodo, plantilla, salida.Ruta);

        var divergencias = new ComparadorSalidaVsManual(2).Comparar(salida.Ruta, manual);

        var inesperadas = new List<ComparadorSalidaVsManual.Divergencia>();
        foreach (var divergencia in divergencias)
        {
            if (BrechaDeclarada(periodo, divergencia) is null)
            {
                inesperadas.Add(divergencia);
            }
        }

        Assert.True(
            inesperadas.Count == 0,
            Informe(etiqueta, divergencias.Count, inesperadas));

        // Guardia anti-vacuidad (S6): la superficie que las unidades ESCRIBEN se aserta contra el
        // manual de forma explícita (si la predicción de brechas la tapara, esto lo delata).
        AssertParidadSuperficieEscrita(salida.Ruta, manual, periodo, etiqueta);
    }

    // ── Brechas declaradas (versionadas, cada una con su cita) ─────────────────────────────────

    /// <summary>
    /// Clasifica una divergencia como brecha declarada. Devuelve <c>null</c> si es INESPERADA
    /// (regresión real). Las categorías siguen el Plan 29 y su evidencia T0.
    /// </summary>
    private static string? BrechaDeclarada(Periodo periodo, ComparadorSalidaVsManual.Divergencia d)
    {
        var hoja = ComparadorSalidaVsManual.NormalizarHoja(d.Hoja);
        var (columna, fila) = Parsear(d.Celda);
        var esAgosto = periodo.CodigoAAAAMM == "202608";

        // 1) Unidad D (T4 / T0d): la app escribe DetRetri D9:D14 (total HU-12), J/L solo por ASE
        //    (9..13), y DetValiRetri I/O solo por ASE (9..13). El resto (DetRetri C..O excepto
        //    D/J/L; DetValiRetri D..N excepto I/O; filas de validación 16+, totales J14/L14/I14)
        //    es recorte declarado: origen = fórmula-cadena, no reproducible en C#, o total manual.
        if (hoja == "DETRETRI")
        {
            var escrita = (columna == "D" && fila is >= 9 and <= 14)
                || (columna is "J" or "L" && fila is >= 9 and <= 13);
            if (!escrita)
            {
                return "Unidad D (T4/T0d): columna/fila no escrita por el recorte de desglose";
            }
        }

        if (hoja == "DETVALIRETRI")
        {
            var escrita = (columna is "I" or "O") && fila is >= 9 and <= 13;
            if (!escrita)
            {
                return "Unidad D (T4/T0d): columna/fila no escrita por el recorte de desglose";
            }
        }

        // 2) Unidad B (T0a): la app es fuente-fiel en julio-ASE5 y el manual diverge (BCE D7/E7 y
        //    el arrastre DetRetri J13). Es divergencia DEL MANUAL, no defecto (R-B-2 / S2).
        if (!esAgosto && hoja == HojaBce.ToUpperInvariant() && d.Celda is "D7" or "E7")
        {
            return "Unidad B (T0a): BCE julio-ASE5 — app fuente-fiel, manual divergente";
        }

        if (!esAgosto && hoja == "DETRETRI" && d.Celda == "J13")
        {
            return "Unidad D/B (T0a): DetRetri J del ASE5 arrastra la divergencia del manual en BCE";
        }

        // 3) H1/T0g: representación de fórmula compartida (el manual deja un follower `SHARED`).
        if (EsFormulaCompartida(d))
        {
            return "H1/T0g: representación de fórmula compartida (follower SHARED)";
        }

        // 4) H1/T0g: la MISMA fórmula difiere SOLO por el sufijo de período de la hoja referida
        //    (DetRetri2026072 vs DetRetri2026082); post-recálculo el resultado es idéntico.
        if (EsDiferenciaSoloPorPeriodo(d))
        {
            return "H1/T0g: referencia de hoja con sufijo de período (misma fórmula)";
        }

        // 5) Recaudo (V3): en una corrida Q2 la app escribe F/G; el manual acumula la 1.ª quincena
        //    (D/E, ya excluidas por el comparador) y otros bloques (I/J). Fuera de F/G = fuera de
        //    la quincena corrida.
        if (hoja.StartsWith("RECAUDO ", StringComparison.Ordinal) && columna is not "F" and not "G")
        {
            return "Recaudo (V3): columna de otra quincena/bloque acumulado del manual";
        }

        // 6) Control Recaudo (Unidad P/F10): la app escribe SOLO F10; D..E y C son literales del
        //    manual (el total de control manual, no derivable por la app).
        if (hoja == "VALIDA - CONTROL RECAUDO" && d.Celda != "F10")
        {
            return "Control Recaudo (T0f): literales manuales A..E (la app solo escribe F10)";
        }

        // 7) REPORTE RECAUDO x BANCO (T0e): rejilla cosmética/aviso; la app escribe C13:C16 y
        //    H13:H16 (y la plantilla trae literales-0 en D..G). El resto es edición manual.
        if (hoja == "REPORTE RECAUDO X BANCO" && !((columna is "C" or "H") && fila is >= 13 and <= 16))
        {
            return "REPORTE RECAUDO x BANCO (T0e): cosmética/aviso del formulario";
        }

        // 8) Hojas que la app no escribe (non-goals del Plan 29).
        if (hoja is "ANT EXT-REV" or "ANTICIPOS USUARIOS")
        {
            return "hoja no escrita por la app (non-goal del Plan 29)";
        }

        // 9) Agosto (T0b/H1): el manual recomputa la malla de detalle R2/R4. El remesh cubre SOLO
        //    R2/R4 hasta el veredicto T4; R1-agosto y su arrastre aguas abajo vuelven a compararse
        //    (T1: rojo intencional hasta T2). La app mapea el detalle por encabezado/label a la
        //    geometría canónica; comparar por celda contra la geometría manual de agosto no aplica.
        if (esAgosto)
        {
            if (hoja is "REM. ANTICIPOS R2" or "REVERSION PAGOS R4")
            {
                return "Unidad R (T0b): malla del manual de agosto distinta a la plantilla canónica";
            }
        }

        // 10) Veredicto T4 (plans/31-T4-Veredictos-R2R4-soporte.md): categorías declarar-divergencia
        //     en agosto (R-D-3, R-D-4, R-D-6). La app sigue la geometría canónica de la plantilla
        //     (SALDOS POR NOTA con la fila opcional Vlr Intereses = 0, doctrina Plan 23) y el manual
        //     del administrativo recomputa la malla de detalle de agosto; por ARRASTRE de esa malla
        //     cambia SOLO el texto de las fórmulas (referencias de fila), no los valores de dominio
        //     (misma familia mesh de T0b §2.4). NO incluye R-D-5 (R1-interior, Reporte Componentes
        //     R1): ese es defecto de anclaje real y queda en rojo (follow-up F-T4-1).
        if (esAgosto)
        {
            if (hoja is "SALDOS POR NOTA" or "AJUSTES - SF-T" or "CONSOLIDADO_TOTAL RECAUDO")
            {
                return "T4/R-D-3,R-D-4: malla de detalle recomputada por el manual de agosto (arrastre)";
            }

            if (hoja.StartsWith("REMUNERACION", StringComparison.Ordinal))
            {
                return "T4/R-D-4: arrastre de la malla R2/R4 del manual de agosto en REMUNERACION_*";
            }

            if (hoja == "INTERVENTORIA" && d.Celda == "R26")
            {
                return "T4/R-D-6: fórmula cosmética del manual (mismo valor; =F15 vs =H15)";
            }
        }

        return null;
    }

    private static bool EsFormulaCompartida(ComparadorSalidaVsManual.Divergencia d) =>
        (d.Motivo.StartsWith("texto de f", StringComparison.Ordinal)
            || d.Motivo.StartsWith("fórmula vs literal", StringComparison.Ordinal))
        && ((d.ValorApp?.Contains("SHARED", StringComparison.Ordinal) ?? false)
            || (d.ValorManual?.Contains("SHARED", StringComparison.Ordinal) ?? false));

    private static bool EsDiferenciaSoloPorPeriodo(ComparadorSalidaVsManual.Divergencia d)
    {
        if (!d.Motivo.StartsWith("texto de f", StringComparison.Ordinal) || d.ValorApp is null || d.ValorManual is null)
        {
            return false;
        }

        return TokenPeriodo.Replace(d.ValorApp, string.Empty) == TokenPeriodo.Replace(d.ValorManual, string.Empty);
    }

    // ── Guardia de la superficie escrita ─────────────────────────────────────────────────────────

    /// <summary>
    /// S6 (anti-vacuidad): asertos puntuales de la superficie que las unidades escriben contra el
    /// manual. Si una regla de brecha la tapara por error, o si T1..T4 regresara, esto falla.
    /// </summary>
    private static void AssertParidadSuperficieEscrita(string salida, string manual, Periodo periodo, string etiqueta)
    {
        var esJulio = periodo.CodigoAAAAMM == "202607";

        // Unidad B: BCE D=SUBSIDIO / E=CONTRIBUCIÓN por ASE (excepto julio-ASE5, T0a).
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            if (esJulio && aseId == 5)
            {
                continue;
            }

            var fila = 2 + aseId;
            foreach (var columna in new[] { "D", "E" })
            {
                var celda = $"{columna}{fila}";
                Assert.InRange(
                    TestHelpers.LeerCeldaNumerica(salida, HojaBce, celda) - TestHelpers.LeerCeldaNumerica(manual, HojaBce, celda),
                    -Tol,
                    Tol);
            }
        }

        // Unidad P: F10 escrito == F10 del manual.
        Assert.InRange(
            TestHelpers.LeerCeldaNumerica(salida, HojaControl, "F10") - TestHelpers.LeerCeldaNumerica(manual, HojaControl, "F10"),
            -Tol,
            Tol);

        Assert.True(esJulio || File.Exists(manual), $"Manual ausente ({etiqueta}).");
    }

    // ── Helpers de ejecución ─────────────────────────────────────────────────────────────────────

    private static (Periodo Periodo, string Carpeta, string Manual, string Plantilla) Caso(string etiqueta) => etiqueta switch
    {
        "julio" => (
            new Periodo { CodigoAAAAMM = "202607", NumeroQuincena = 2 },
            Insumos.CarpetaInsumosJulioQ2,
            Path.Combine(Insumos.Raiz(), "Docs", "Prueba Julio-2", "Resultado", "Remuneracion 202607-2 Total Administrativo.xlsx"),
            Plantilla),
        "agosto" => (
            new Periodo { CodigoAAAAMM = "202608", NumeroQuincena = 2 },
            Insumos.CarpetaInsumosAgosto,
            Path.Combine(Insumos.Raiz(), "Docs", "Prueba2", "Resultado", "Resultado Manual por el administrativo", "Remuneracion 202608-2 Total_7721.xlsx"),
            Insumos.PlantillaAgosto2026082),
        _ => throw new ArgumentOutOfRangeException(nameof(etiqueta), etiqueta, "Período no soportado por la regresión del comparador.")
    };

    private static void EjecutarFlujo(string carpetaInsumos, Periodo periodo, string rutaPlantilla, string rutaSalida)
    {
        var procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader());

        procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = periodo,
            CarpetaPeriodo = carpetaInsumos,
            RutaPlantilla = rutaPlantilla,
            RutaSalida = rutaSalida
        });
    }

    private static (string Columna, int Fila) Parsear(string celda)
    {
        var match = ReferenciaCelda.Match(celda);
        return match.Success
            ? (match.Groups["col"].Value.ToUpperInvariant(), int.Parse(match.Groups["fila"].Value, CultureInfo.InvariantCulture))
            : (string.Empty, 0);
    }

    private static string Informe(string etiqueta, int total, IReadOnlyList<ComparadorSalidaVsManual.Divergencia> inesperadas)
    {
        const int tope = 6000;
        var lineas = inesperadas
            .Take(tope)
            .Select(d => $"{d.Hoja}!{d.Celda}  app=[{d.ValorApp}]  manual=[{d.ValorManual}]  ({d.Motivo})");
        var resto = inesperadas.Count > tope ? $"\n... y {inesperadas.Count - tope} más" : string.Empty;
        return $"[{etiqueta}] {inesperadas.Count} divergencia(s) INESPERADA(s) de {total} (fuera de exclusiones + brechas declaradas):\n"
            + string.Join("\n", lineas) + resto;
    }

    private sealed class SalidaTemporal : IDisposable
    {
        private readonly string _dir;

        public SalidaTemporal(string prefijo)
        {
            _dir = Path.Combine(Path.GetTempPath(), prefijo + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
            Ruta = Path.Combine(_dir, "salida.xlsx");
        }

        public string Ruta { get; }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_dir))
                {
                    Directory.Delete(_dir, recursive: true);
                }
            }
            catch
            {
                // best-effort
            }
        }
    }
}
