using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Fixtures compartidos de los insumos reales Q1 (<c>Docs/Prueba Julio-1/Insumos</c>). Nombres
/// reales de carpeta (cuidado con la tilde de "5-Área Limpia") y archivos prefix-based (mismo
/// patrón del locator).
/// </summary>
internal static class Insumos
{
    public const decimal Tolerancia = 0.5m;

    public static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }

    /// <summary>
    /// HISTÓRICA — base <b>v2</b> (<c>Docs/Plantilla estandarizada/v2/Formato Remuneracion 2026xx-x Total vf.xlsx</c>).
    ///
    /// <para>
    /// Plan 59 (§15, 2026-10-27): el nombre anterior (<c>PlantillaOficial</c>) mentía — devolvía la v2
    /// mientras la plantilla VIGENTE es la <b>v4</b> (<see cref="PlantillaOficialV4"/>). Se renombró sin
    /// alias: si algo pide "la oficial" debe pedir la v4. La v2 solo la consumen las caracterizaciones
    /// cuyo OBJETO es la geometría v2 (mapas LEGACY) y la evidencia congelada que cita sus celdas; la v1
    /// (<c>Docs/Plantilla estandarizada/Formato Remuneracion 2026xx-x Total.xlsx</c>) también queda
    /// histórica intacta y NO la usa el pipeline. Ruta EXPLÍCITA (nunca la raíz <c>Docs/Insumos</c>
    /// como fuente de plantilla).
    /// </para>
    /// </summary>
    public static string PlantillaHistoricaV2() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "v2", "Formato Remuneracion 2026xx-x Total vf.xlsx");

    /// <summary>
    /// Plantilla <b>v4</b> VIGENTE desde 2026-10-27
    /// (<c>Docs/Plantilla estandarizada/v4/Formato Remuneracion 2026xx-x Total v4.xlsx</c>).
    /// Es la v3 más exactamente dos celdas de fórmula corregidas en <c>Reporte Componentes R1</c>
    /// (<c>F328</c> y <c>F333</c>: <c>-M213</c>/<c>--L126</c> → <c>-L213</c>/<c>-L126</c>); el resto del
    /// contenido es idéntico (mismas 40 hojas, mismo orden, mismas dimensiones, mismas 42.134 celdas).
    /// Con la v4 julio-1, julio-2 y octubre-1 cierran. Declara, como la v3, el grupo
    /// <c>canal A / empresa EAB</c> del SEGUNDO cuadro de <c>Rem. Anticipos R2</c> en los 5 ASE.
    /// La v3 queda en disco como <b>histórica</b> y ningún camino de producción ni fixture de
    /// procesamiento la consume.
    /// </summary>
    public static string PlantillaOficialV4() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "v4", "Formato Remuneracion 2026xx-x Total v4.xlsx");

    private static readonly object CandadoBases = new();
    private static readonly Dictionary<string, string> BasesPorPeriodo = new();

    /// <summary>
    /// Plan 39 (T4): base del período producida por Phase 0 (<see cref="PreparadorBasePipeline"/>)
    /// desde la plantilla <b>histórica v2</b> (<see cref="PlantillaHistoricaV2"/>) hacia una ruta temp
    /// FRESCA (una sola vez por proceso; los fixtures comparan hashes entre llamadas). Solo la usan
    /// las caracterizaciones ancladas a la geometría v2 (vía <see cref="BaseV2DelPeriodo"/>); la base
    /// vigente es <see cref="BaseDelPeriodoV4"/>. Sustituye a la carpeta de bases preparadas offline
    /// (retirada): el fixture ya no apunta a una carpeta versionada y el flujo es el de producción
    /// (plantilla → Phase 0 → copia de trabajo).
    /// </summary>
    public static string BaseDelPeriodo(string periodo)
    {
        lock (CandadoBases)
        {
            if (BasesPorPeriodo.TryGetValue(periodo, out var existente))
            {
                return existente;
            }

            var oficial = PlantillaHistoricaV2();
            Assert.True(File.Exists(oficial), $"Falta la plantilla oficial: {oficial}");

            var dir = Path.Combine(
                Path.GetTempPath(), $"p39-fixtures-{Environment.ProcessId}-{Guid.NewGuid():N}");
            Directory.CreateDirectory(dir);
            var salida = Path.Combine(dir, $"Plantilla_Remuneracion_{periodo}.xlsx");
            PreparadorBasePipeline.Ejecutar(oficial, Remuneracion.Core.Models.Periodo.Parse(periodo), salida);
            BasesPorPeriodo[periodo] = salida;
            return salida;
        }
    }

    /// <summary>
    /// Base canónica Q1 (período <c>2026071</c>) preparada por Phase 0.
    ///
    /// Plan 59 (H5, rebaseado 2026-10-27): se prepara desde la plantilla <b>v4</b> — única base
    /// vigente. La expectativa de los fixtures que la consumen se conserva porque la v4 es la v3 con
    /// dos fórmulas corregidas (misma geometría, mismas etiquetas) y la v3 era superconjunto semántico
    /// de la v2 (§0.4-a: 0 faltantes de etiqueta en las tres hojas).
    /// </summary>
    public static string Plantilla => BaseDelPeriodoV4("2026071");

    private static readonly object CandadoBasesV4 = new();
    private static readonly Dictionary<string, string> BasesV4PorPeriodo = new();

    /// <summary>
    /// Plan 59 (H5, rebaseado 2026-10-27): base del período producida por Phase 0 desde la plantilla
    /// <b>v4</b> hacia una ruta temp FRESCA (una sola vez por proceso). Mismo flujo que
    /// <see cref="BaseDelPeriodo"/>; la diferencia es la plantilla de origen vigente.
    /// </summary>
    public static string BaseDelPeriodoV4(string periodo)
    {
        lock (CandadoBasesV4)
        {
            if (BasesV4PorPeriodo.TryGetValue(periodo, out var existente))
            {
                return existente;
            }

            var oficial = PlantillaOficialV4();
            Assert.True(File.Exists(oficial), $"Falta la plantilla v4 vigente: {oficial}");

            var dir = Path.Combine(
                Path.GetTempPath(), $"p59-fixtures-v4-{Environment.ProcessId}-{Guid.NewGuid():N}");
            Directory.CreateDirectory(dir);
            var salida = Path.Combine(dir, $"Plantilla_Remuneracion_v4_{periodo}.xlsx");
            PreparadorBasePipeline.Ejecutar(oficial, Remuneracion.Core.Models.Periodo.Parse(periodo), salida);
            BasesV4PorPeriodo[periodo] = salida;
            return salida;
        }
    }


    /// <summary>
    /// Plan 59 (H5): base del período desde la plantilla <b>v2</b> (histórica). Existe SOLO para las
    /// caracterizaciones cuyo OBJETO es la geometría v2 —la foto del mapa Q2 y las filas que cita la
    /// evidencia T0b—. NO es la base de producción: la app consume la v4 (`PlantillaOficialV4`).
    /// </summary>
    public static string BaseV2DelPeriodo(string periodo) => BaseDelPeriodo(periodo);

    public static string CarpetaPeriodo => Path.Combine(Raiz(), "Docs", "Prueba Julio-1", "Insumos");

    public static string[] CarpetasAse => new[]
    {
        Path.Combine(CarpetaPeriodo, "1-Promoambiental"),
        Path.Combine(CarpetaPeriodo, "2-Lime"),
        Path.Combine(CarpetaPeriodo, "3-Ciudad Limpia"),
        Path.Combine(CarpetaPeriodo, "4-Bogota Limpia"),
        Path.Combine(CarpetaPeriodo, "5-Área Limpia")
    };

    public static string R1(int aseId) => Buscar(CarpetasAse[aseId - 1], "Recaudoporcomponente", $"R1 ASE{aseId}");

    public static string R2(int aseId) => Buscar(CarpetasAse[aseId - 1], "RerpoteDetalleSaldosaFavor", $"R2 ASE{aseId}");

    public static string R4(int aseId)
    {
        var carpeta = CarpetasAse[aseId - 1];
        var r4 = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("ReversiónPorComponente", StringComparison.OrdinalIgnoreCase))
            ?? Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
                .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("ReversionPorComponente", StringComparison.OrdinalIgnoreCase));

        return r4 ?? throw new FileNotFoundException($"No se encontró R4 ASE{aseId} en {carpeta}.");
    }

    /// <summary>
    /// HU-09 (2.3, V13): ruta del <c>ReportePagosxBanco_*.xlsx</c> del ASE (naming verificado por T0).
    /// </summary>
    public static string ReporteBanco(int aseId) => Buscar(CarpetasAse[aseId - 1], "ReportePagosxBanco", $"ReportePagosxBanco ASE{aseId}");

    /// <summary>
    /// HU-10 (2.4, V8): ruta del <c>R4-BalanceSubsidioyContribuciones_*.xlsx</c> del ASE.
    /// Mismo criterio del locator: prefijo base (con guion bajo, excluye la variante
    /// <c>-Optimizado_</c>) primero y variante <c>R4-BalanceSubsidioyContribuciones-Optimizado_</c>
    /// como fallback. El archivo <c>Reca_BalanceSubsidiosyContribuciones_*</c> de ASE5-Q1 es otro
    /// reporte y NO matchea ninguno de los dos prefijos.
    /// </summary>
    public static string Balance(int aseId)
    {
        var carpeta = CarpetasAse[aseId - 1];
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("R4-BalanceSubsidioyContribuciones_", StringComparison.OrdinalIgnoreCase))
            ?? Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
                .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("R4-BalanceSubsidioyContribuciones-Optimizado_", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Balance ASE{aseId} en {carpeta}");
        return archivo!;
    }

    /// <summary>
    /// HU-14 (S-3): helper migrado a la factoría única <see cref="AseFactory.DesdeId"/>
    /// (fuente única desde <see cref="Remuneracion.Core.Constants.CarpetasAse.Prefijos"/>).
    /// </summary>
    public static Ase Ase(int id) => AseFactory.DesdeId(id);

    public static Periodo Periodo() => new() { CodigoAAAAMM = "202607", NumeroQuincena = 1 };

    /// <summary>
    /// HU-20 (G1): carpeta de conciliaciones de la nueva organización
    /// (<c>{periodo}/Conciliaciones/</c>; <c>Consolidado/</c> eliminada).
    /// </summary>
    public static string CarpetaConciliaciones => Path.Combine(CarpetaPeriodo, "Conciliaciones");

    /// <summary>
    /// HU-20 (G3): R10 del período Q1 (<c>R10_Remuneracion_{AAAAMMQ}.xlsx</c>).
    /// </summary>
    public static string R10 => Path.Combine(CarpetaPeriodo, "R10_Remuneracion_2026071.xlsx");

    // ── HU-11 (2.5): insumos Q2 ──────────────────────────────────────────────────────────────

    /// <summary>
    /// HU-11 (2.5, T0-0.1): plantilla canónica en ceros Q2 (fórmulas sin caché),
    /// <c>Docs/Prueba Julio-2/Plantilla_Remuneracion.xlsx</c>.
    /// </summary>
    /// <summary>
    /// Plan 36 (T4): base canónica Q2 del período <c>2026072</c>, preparada por Phase 0.
    ///
    /// Plan 59 (H5, rebaseado 2026-10-27): se prepara desde la plantilla <b>v4</b> (única base
    /// vigente). La expectativa se conserva: el período Q2 es el mismo, la fuente es la misma y la v4
    /// declara todo lo que la v2 declaraba más los grupos de canal que faltaban (§0.4-a/c) y las dos
    /// fórmulas que la v3 traía con typo. La base deja de ser la foto de la v2.
    /// </summary>
    public static string PlantillaQ2 => BaseDelPeriodoV4("2026072");


    /// <summary>
    /// HU-11 (2.5, T0-0.1): golden de VALORES Q2 (referencia completada) usado como caché en la
    /// Capa A (D85:D89, D104:D109, AJUSTES-SF-T). Es el manual del administrativo
    /// <c>Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx</c>,
    /// byte-idéntico al golden Q2 previo a la reorganización de insumos (mismo SHA256 58431010…,
    /// verificado en disco).
    /// </summary>
    public static string GoldenQ2 => ManualJulioQ2;

    /// <summary>
    /// Manual del administrativo de julio-Q2
    /// (<c>Docs/Prueba Julio-2/Resultado/Remuneracion 202607-2 Total Administrativo.xlsx</c>):
    /// golden de valores y referencia de brechas del comparador.
    /// </summary>
    public static string ManualJulioQ2 => Path.Combine(
        Raiz(), "Docs", "Prueba Julio-2", "Resultado", "Remuneracion 202607-2 Total Administrativo.xlsx");

    /// <summary>
    /// Plan 29 (T5): salida de la app del período Q2 de julio
    /// (<c>Docs/Prueba Julio-2/Resultado/Resultado{numero}/Remuneración 202607-2 Total.xlsx</c>).
    /// </summary>
    public static string ResultadoJulioQ2(int numero) => Path.Combine(
        Raiz(), "Docs", "Prueba Julio-2", "Resultado", $"Resultado{numero}", "Remuneración 202607-2 Total.xlsx");

    /// <summary>
    /// Plan 29 (T5): salida de la app del período Q2 de agosto
    /// (<c>Docs/Prueba Agosto-2/Resultado/Resultado{numero}/Remuneración 202608-2 Total.xlsx</c>).
    /// </summary>
    public static string ResultadoAgostoQ2(int numero) => Path.Combine(
        Raiz(), "Docs", "Prueba Agosto-2", "Resultado", $"Resultado{numero}", "Remuneración 202608-2 Total.xlsx");

    public static string CarpetaPeriodoQ2 => Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Insumos");

    public static string[] CarpetasAseQ2 => new[]
    {
        Path.Combine(CarpetaPeriodoQ2, "1-Promoambiental"),
        Path.Combine(CarpetaPeriodoQ2, "2-Lime"),
        Path.Combine(CarpetaPeriodoQ2, "3-Ciudad Limpia"),
        Path.Combine(CarpetaPeriodoQ2, "4-Bogota Limpia"),
        Path.Combine(CarpetaPeriodoQ2, "5-Área Limpia")
    };

    public static Periodo PeriodoQ2() => new() { CodigoAAAAMM = "202607", NumeroQuincena = 2 };

    /// <summary>
    /// HU-20 (G1): carpeta de conciliaciones Q2 (nueva organización).
    /// </summary>
    public static string CarpetaConciliacionesQ2 => Path.Combine(CarpetaPeriodoQ2, "Conciliaciones");

    /// <summary>
    /// HU-20 (G3): R10 del período Q2 (<c>R10_Remuneracion_{AAAAMMQ}.xlsx</c>).
    /// </summary>
    public static string R10Q2 => Path.Combine(CarpetaPeriodoQ2, "R10_Remuneracion_2026072.xlsx");

    /// <summary>
    /// Plan 27 (T3): resuelve la conciliación por prefijo dentro de
    /// <c>{carpetaPeriodo}/Conciliaciones/</c> (mismo criterio del locator: prefijo case-insensitive).
    /// </summary>
    public static string Conciliacion(string carpetaPeriodo, string prefijo)
    {
        var carpeta = Path.Combine(carpetaPeriodo, "Conciliaciones");
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta de conciliaciones: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta la conciliación '{prefijo}' en {carpeta}");
        return archivo!;
    }

    /// <summary>
    /// Plan 27 (T3): las 15 conciliaciones reales en disco (5 Q1 + 5 Q2 + 5 agosto).
    /// </summary>
    public static IReadOnlyList<string> ConciliacionesReales() =>
    [
        .. Directory.EnumerateFiles(CarpetaConciliaciones, "*.xlsx", SearchOption.TopDirectoryOnly),
        .. Directory.EnumerateFiles(CarpetaConciliacionesQ2, "*.xlsx", SearchOption.TopDirectoryOnly),
        .. Directory.EnumerateFiles(Path.Combine(CarpetaInsumosAgosto, "Conciliaciones"), "*.xlsx", SearchOption.TopDirectoryOnly)
    ];

    /// <summary>
    /// HU-11 (2.5): ruta del <c>Recaudoporcomponente_*.xlsx</c> del ASE Q2 (layout ASE1-4
    /// paritario con Q1; ASE5 diverge — recorte T0-0.6).
    /// </summary>
    public static string R1Q2(int aseId) => Buscar(CarpetasAseQ2[aseId - 1], "Recaudoporcomponente", $"R1 Q2 ASE{aseId}");

    /// <summary>
    /// HU-12 (2.6 ampliada, V0.3): ruta del R1-Q2 de ASE5 (variante 2 filas Mes/Total — el
    /// dispatch Q2 del reader la soporta). Alias explícito para los tests de la variante.
    /// </summary>
    public static string R1Q2Ase5() => R1Q2(5);

    /// <summary>
    /// HU-11 (2.5): ruta del <c>RerpoteDetalleSaldosaFavor_*.xlsx</c> del ASE Q2.
    /// </summary>
    public static string R2Q2(int aseId) => Buscar(CarpetasAseQ2[aseId - 1], "RerpoteDetalleSaldosaFavor", $"R2 Q2 ASE{aseId}");

    /// <summary>
    /// HU-11 (2.5): ruta del <c>ReversiónPorComponente_*.xlsx</c> del ASE Q2 (diacríticos en disco).
    /// </summary>
    public static string R4Q2(int aseId)
    {
        var carpeta = CarpetasAseQ2[aseId - 1];
        var r4 = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Normalizar(Path.GetFileNameWithoutExtension(f)).StartsWith("reversionporcomponente", StringComparison.Ordinal));
        Assert.True(r4 is not null, $"Falta R4 Q2 ASE{aseId} en {carpeta}");
        return r4!;
    }

    /// <summary>
    /// HU-11 (2.5): ruta del <c>R4-BalanceSubsidioyContribuciones_*</c> del ASE Q2 (ASE5 solo
    /// variante -Optimizado; fallback obligatorio, V4).
    /// </summary>
    public static string BalanceQ2(int aseId)
    {
        var carpeta = CarpetasAseQ2[aseId - 1];
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("R4-BalanceSubsidioyContribuciones_", StringComparison.OrdinalIgnoreCase))
            ?? Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
                .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("R4-BalanceSubsidioyContribuciones-Optimizado_", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Balance Q2 ASE{aseId} en {carpeta}");
        return archivo!;
    }

    /// <summary>
    /// HU-11 (2.5): ruta del <c>ReportePagosxBanco_*.xlsx</c> del ASE Q2.
    /// </summary>
    public static string ReporteBancoQ2(int aseId) => Buscar(CarpetasAseQ2[aseId - 1], "ReportePagosxBanco", $"ReportePagosxBanco Q2 ASE{aseId}");

    /// <summary>
    /// HU-11 (2.5, T0-0.5): ruta del <c>SaldosaFavorAplicadosPorNotas_*.xlsx</c> del ASE Q2.
    /// </summary>
    public static string SaldosNotas(int aseId) => Buscar(CarpetasAseQ2[aseId - 1], "SaldosaFavorAplicadosPorNotas", $"SaldosNotas ASE{aseId}");

    /// <summary>
    /// HU-11 (2.5, T0-0.4/0.5): ruta del <c>RetribuciónNegativa_*.xlsx</c> del ASE Q2. El nombre
    /// en disco lleva diacríticos (V9): el match usa el stem normalizado sin acentos.
    /// </summary>
    public static string RetribucionNegativa(int aseId)
    {
        var carpeta = CarpetasAseQ2[aseId - 1];
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Normalizar(Path.GetFileNameWithoutExtension(f)).StartsWith("retribucionnegativa", StringComparison.Ordinal));
        Assert.True(archivo is not null, $"Falta RetribuciónNegativa ASE{aseId} en {carpeta}");
        return archivo!;
    }

    // ── HU-22 (Plan 23): insumos de agosto (Docs/Prueba Agosto-2/Insumos) ────────────────────

    /// <summary>
    /// HU-22 (Plan 23): raíz de los insumos de agosto (<c>Docs/Prueba Agosto-2/Insumos</c>),
    /// usados para ejercitar la opcionalidad de <c>Vlr Intereses</c> (ASE2 sin la fila, T0d §5.1).
    /// </summary>
    public static string CarpetaInsumosAgosto => Path.Combine(Raiz(), "Docs", "Prueba Agosto-2", "Insumos");

    /// <summary>Plan 27 (T3): período de agosto Q2 (<c>2026082</c>).</summary>
    public static Periodo PeriodoAgosto() => new() { CodigoAAAAMM = "202608", NumeroQuincena = 2 };

    /// <summary>
    /// Plan 30 (T3, R-B-1): base canónica de agosto <c>2026082</c>, producida UNA vez de forma
    /// offline por <c>Herramientas/PreparadorBasePeriodo</c> (rename de hojas + reescritura del
    /// token de período + metadatos/Q1 del manual). Su consistencia la fija
    /// <see cref="BaseConsistente2026082Tests"/>.
    ///
    /// Plan 59 (§11-bis, rebaseado 2026-10-27): se prepara desde la plantilla <b>v4</b>
    /// (§<see cref="BaseDelPeriodoV4"/>) porque la fuente de agosto trae, en el SEGUNDO cuadro de
    /// <c>Rem. Anticipos R2</c>, filas del grupo <c>canal A / empresa EAB</c> (ASE4: −12.462,26) que la
    /// v2 no declara: escribirlas sobre la v2 dejaba el bloque incompleto (el detalle sumaba 12.462,26
    /// de más frente a la fuente).
    /// </summary>
    public static string PlantillaAgosto2026082 => BaseDelPeriodoV4("2026082");

    /// <summary>
    /// Plan 30 (T3, D-E): manual del administrativo de agosto (<c>202608-2</c>), fuente de los
    /// metadatos de período y de las columnas Q1 (D:E) que la base hereda celda por celda.
    /// </summary>
    public static string ManualAgosto2026082 => Path.Combine(
        Raiz(), "Docs", "Prueba Agosto-2", "Resultado", "Resultado Manual por el administrativo",
        "Remuneracion 202608-2 Total_7721.xlsx");

    public static string[] CarpetasAseAgosto => new[]
    {
        Path.Combine(CarpetaInsumosAgosto, "1-Promoambiental"),
        Path.Combine(CarpetaInsumosAgosto, "2-Lime"),
        Path.Combine(CarpetaInsumosAgosto, "3-Ciudad Limpia"),
        Path.Combine(CarpetaInsumosAgosto, "4-Bogota Limpia"),
        Path.Combine(CarpetaInsumosAgosto, "5-Área Limpia")
    };

    /// <summary>
    /// HU-22 (Plan 23): ruta del <c>SaldosaFavorAplicadosPorNotas_*.xlsx</c> de agosto del ASE.
    /// </summary>
    public static string SaldosNotasAgosto(int aseId) =>
        Buscar(CarpetasAseAgosto[aseId - 1], "SaldosaFavorAplicadosPorNotas", $"SaldosNotas agosto ASE{aseId}");

    /// <summary>
    /// Plan 25 (WU-1): ruta del <c>Recaudoporcomponente_*.xlsx</c> de agosto (R1) del ASE.
    /// </summary>
    public static string R1Agosto(int aseId) =>
        Buscar(CarpetasAseAgosto[aseId - 1], "Recaudoporcomponente", $"R1 agosto ASE{aseId}");

    /// <summary>
    /// Plan 25 (WU-1): ruta del <c>RerpoteDetalleSaldosaFavor_*.xlsx</c> de agosto (R2) del ASE.
    /// </summary>
    public static string R2Agosto(int aseId) =>
        Buscar(CarpetasAseAgosto[aseId - 1], "RerpoteDetalleSaldosaFavor", $"R2 agosto ASE{aseId}");

    /// <summary>
    /// Plan 25 (WU-1): ruta del <c>ReversiónPorComponente_*.xlsx</c> de agosto (R4) del ASE
    /// (diacríticos en disco; match por stem normalizado).
    /// </summary>
    public static string R4Agosto(int aseId)
    {
        var carpeta = CarpetasAseAgosto[aseId - 1];
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta: {carpeta}");
        var r4 = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Normalizar(Path.GetFileNameWithoutExtension(f)).StartsWith("reversionporcomponente", StringComparison.Ordinal));
        Assert.True(r4 is not null, $"Falta R4 agosto ASE{aseId} en {carpeta}");
        return r4!;
    }

    // ── Plan 29 (T2, Unidad R): insumos julio-Q2 (Docs/Prueba Julio-2/Insumos) ──────────────
    // Es la misma carpeta que el período Q2 (CarpetaPeriodoQ2), expuesta con la ruta literal que
    // cita la evidencia T0b (`plans/29-T0-Evidencia-B.md`).

    /// <summary>Plan 29 (T2): raíz de los insumos de julio-Q2 (<c>Docs/Prueba Julio-2/Insumos</c>).</summary>
    public static string CarpetaInsumosJulioQ2 => Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Insumos");

    public static string[] CarpetasAseJulioQ2 => new[]
    {
        Path.Combine(CarpetaInsumosJulioQ2, "1-Promoambiental"),
        Path.Combine(CarpetaInsumosJulioQ2, "2-Lime"),
        Path.Combine(CarpetaInsumosJulioQ2, "3-Ciudad Limpia"),
        Path.Combine(CarpetaInsumosJulioQ2, "4-Bogota Limpia"),
        Path.Combine(CarpetaInsumosJulioQ2, "5-Área Limpia")
    };

    /// <summary>Plan 29 (T2): <c>RerpoteDetalleSaldosaFavor_*</c> de julio-Q2 (R2) del ASE.</summary>
    public static string R2JulioQ2(int aseId) =>
        Buscar(CarpetasAseJulioQ2[aseId - 1], "RerpoteDetalleSaldosaFavor", $"R2 julio-Q2 ASE{aseId}");

    /// <summary>Plan 29 (T2): <c>Recaudoporcomponente_*</c> de julio-Q2 (R1) del ASE.</summary>
    public static string R1JulioQ2(int aseId) =>
        Buscar(CarpetasAseJulioQ2[aseId - 1], "Recaudoporcomponente", $"R1 julio-Q2 ASE{aseId}");

    /// <summary>Plan 29 (T2): <c>ReversiónPorComponente_*</c> de julio-Q2 (R4) del ASE (diacríticos en disco).</summary>
    public static string R4JulioQ2(int aseId) =>
        BuscarR4(CarpetasAseJulioQ2[aseId - 1], $"R4 julio-Q2 ASE{aseId}");

    /// <summary>
    /// Plan 29 (T2): <c>RerpoteDetalleSaldosaFavor_*</c> de agosto-Q2 (R2) del ASE
    /// (<c>Docs/Prueba Agosto-2/Insumos</c>).
    /// </summary>
    public static string R2AgostoQ2(int aseId) =>
        Buscar(CarpetasAseAgosto[aseId - 1], "RerpoteDetalleSaldosaFavor", $"R2 agosto-Q2 ASE{aseId}");

    /// <summary>
    /// Plan 29 (T2): <c>ReversiónPorComponente_*</c> de agosto-Q2 (R4) del ASE (alias explícito de
    /// <see cref="R4Agosto"/>; diacríticos en disco).
    /// </summary>
    public static string R4AgostoQ2(int aseId) => R4Agosto(aseId);

    /// <summary>
    /// Localiza el <c>ReversiónPorComponente_*</c> (R4) en una carpeta por prefijo normalizado
    /// (sin diacríticos, case-insensitive). Fail-fast si falta.
    /// </summary>
    private static string BuscarR4(string carpeta, string etiqueta)
    {
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Normalizar(Path.GetFileNameWithoutExtension(f)).StartsWith("reversionporcomponente", StringComparison.Ordinal));
        Assert.True(archivo is not null, $"Falta {etiqueta} en {carpeta}");
        return archivo!;
    }

    /// <summary>
    /// HU-11 (2.5, T0-0.4): normaliza un nombre (minúsculas, sin diacríticos) para el match por prefijo.
    /// </summary>
    private static string Normalizar(string texto)
    {
        var normalizado = texto.Normalize(System.Text.NormalizationForm.FormD);
        var sb = new System.Text.StringBuilder(normalizado.Length);
        foreach (var ch in normalizado)
        {
            if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch) != System.Globalization.UnicodeCategory.NonSpacingMark)
            {
                sb.Append(char.ToLowerInvariant(ch));
            }
        }

        return sb.ToString();
    }

    private static string Buscar(string carpeta, string prefijo, string etiqueta)
    {
        Assert.True(Directory.Exists(carpeta), $"Falta carpeta: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta {etiqueta} en {carpeta}");
        return archivo!;
    }
}
