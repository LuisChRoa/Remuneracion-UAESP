using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 64 (§2.9 — requisito pendiente tras H0..H5): GATES DE COHERENCIA de <c>ANT EXT-REV</c> por ASE
/// en <see cref="ValidadorBasico"/> y su HERENCIA en
/// <see cref="ValidadorBasico.ClasificarDivergenciasDeclaradas"/>.
///
/// <para><b>Qué se juzga (aditivo, SRP: el juicio R1/R2/R4 no se toca).</b>
/// <list type="number">
/// <item><b>Σ por fuente:</b> la suma de los cuatro conceptos de una fuente PRESENTE debe coincidir
/// ±0,5 con el total con el que esa fuente cierra su grupo en el insumo (dos lecturas independientes
/// del mismo archivo: las cuatro filas de concepto y la fila de total).</item>
/// <item><b>Σ por bloque (ASE):</b> el dinero LEÍDO en los bloques de fuente del ASE debe llegar
/// completo a las celdas que la plantilla escribe. El bloque de la plantilla tiene UN destino por ASE,
/// de modo que dos grupos del insumo que el catálogo identifica como la MISMA fuente dejan dinero sin
/// celda: eso solo lo ve el cierre por bloque, no el cierre por fuente.</item>
/// </list></para>
///
/// <para><b>Efecto (Plan 61 §2-D-H3-ii).</b> Son gates de DATOS: la divergencia se entrega como
/// NOVEDAD (el Excel ya escrito no se retiene) y, al no existir entrada declarable para esta hoja,
/// <see cref="ValidadorBasico.ClasificarDivergenciasDeclaradas"/> la conserva como NO DECLARADA
/// (Plan 42 D-F). Los gates de ESTRUCTURA (insumo roto, plantilla sin destino) siguen siendo fallo
/// duro aguas arriba.</para>
///
/// <para><b>Oráculos.</b> El caso VERDE se ancla al corpus real medido por H0 (los 30 insumos de
/// <c>Docs/T0-Plan64-Ant-Ext-Rev.md</c>). Los casos negativos usan lecturas SINTÉTICAS construidas con
/// el contrato del modelo: solo forma, nunca un monto real del negocio.</para>
/// </summary>
public sealed class AntExtRevGatesCoherenciaTests
{
    /// <summary>Marca con la que los gates nuevos nombran la hoja (G-TRAZA).</summary>
    private const string MarcaHoja = "ANT EXT-REV";

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Rutas explícitas (prohibido Insumos.Raiz() en lo nuevo, §2.1)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static string RaizRepositorio()
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

        throw new InvalidOperationException($"No se encontró AGENTS.md partiendo de '{AppContext.BaseDirectory}'.");
    }

    private static string Ruta(params string[] partes) => Path.Combine([RaizRepositorio(), .. partes]);

    private static readonly (string Periodo, string Carpeta)[] Periodos =
    [
        ("2026071", "Prueba Julio-1"),
        ("2026072", "Prueba Julio-2"),
        ("2026082", "Prueba Agosto-2"),
        ("2026091", "Prueba Septiembre 1"),
        ("2026092", "Prueba Septiembre-2"),
        ("2026101", "Prueba Octubre-1")
    ];

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia"
    ];

    private static string Insumo(string periodo, int aseId)
    {
        var carpeta = Path.Combine(
            Ruta("Docs", Periodos.Single(p => p.Periodo == periodo).Carpeta, "Insumos"),
            CarpetasAse[aseId - 1]);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta del ASE{aseId} en {periodo}: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .SingleOrDefault(f => Path.GetFileName(f).StartsWith("RecaudosReversados_", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta un único RecaudosReversados_*.xlsx en {carpeta}");
        return archivo!;
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Caso VERDE — corpus real (H0 midió los 30 insumos; aquí solo se juzga lo ya leído)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Ninguno de los 30 insumos reales dispara los gates: la coherencia que exige §2.9 ya se cumple en
    /// el corpus medido por H0. Se comprueba además que la lectura NO venga vacía (un verde por ausencia
    /// de datos no probaría nada).
    /// </summary>
    [Fact]
    public void Gate_CorpusReal30Insumos_NoReportaDivergenciasAntExtRev()
    {
        var evaluados = 0;
        foreach (var (periodo, _) in Periodos)
        {
            var leafs = new List<WorkbookLeafInputs>();
            for (var aseId = 1; aseId <= 5; aseId++)
            {
                var leido = Leer(periodo, aseId);
                Assert.NotEmpty(leido.Fuentes);
                leafs.Add(new WorkbookLeafInputs
                {
                    Ase = AseFactory.DesdeId(aseId),
                    Periodo = PeriodoDe(periodo),
                    AntExtRev = leido
                });
            }

            var errores = new ValidadorBasico().Validar(ResultadoDe(periodo), leafs);
            Assert.DoesNotContain(errores, EsDelGate);
            evaluados += leafs.Count;
        }

        Assert.Equal(30, evaluados);
    }

    /// <summary>
    /// Compatibilidad por construcción (D-03): sin lectura del insumo el gate es NO-OP, igual que el
    /// resto del frente ANT EXT-REV. No se inventa una divergencia donde no hay hoja que juzgar.
    /// </summary>
    [Fact]
    public void Gate_SinLecturaDelInsumo_EsNoOp()
    {
        var leafs = new List<WorkbookLeafInputs>();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            leafs.Add(new WorkbookLeafInputs
            {
                Ase = AseFactory.DesdeId(aseId),
                Periodo = PeriodoDe("2026091")
            });
        }

        var errores = new ValidadorBasico().Validar(ResultadoDe("2026091"), leafs);

        Assert.DoesNotContain(errores, EsDelGate);
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Caso ROJO 1 — Σ de los cuatro conceptos ≠ total de la fuente
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Una fuente PRESENTE cuyos cuatro conceptos no suman el total con el que cierra su grupo rompe la
    /// atribución del dinero: el gate lo reporta con ASE + fuente + observado/esperado + Δ, y solo para
    /// el ASE afectado.
    /// </summary>
    [Fact]
    public void Gate_SigmaConceptosVsTotalFuente_Diverge_ReportaAseFuenteYDelta()
    {
        // Insumo SINTÉTICO (solo forma, nunca un monto real del negocio): los cuatro conceptos suman
        // 90 y la fuente cierra su grupo en 91 (Δ = 1, fuera de la tolerancia).
        var leafs = LeafsBase();
        leafs[1].AntExtRev = LecturaSintetica(2, Fuente("ENEL", 100m, 20m, -30m, 0m, totalFuente: 91m));

        var errores = new ValidadorBasico().Validar(ResultadoDe("2026091"), leafs);

        var porFuente = Assert.Single(
            errores,
            e => EsDelGate(e) && e.Contains("fuente 'ENEL'", StringComparison.Ordinal));
        Assert.Contains("ASE 2", porFuente, StringComparison.Ordinal);
        Assert.Contains("(90)", porFuente, StringComparison.Ordinal);
        Assert.Contains("(91)", porFuente, StringComparison.Ordinal);
        Assert.Contains("Diferencia=1", porFuente, StringComparison.Ordinal);

        // La divergencia de ANT EXT-REV no se disfraza de sub-bloque declarable (Plan 42 D-F).
        Assert.DoesNotContain("VALIDACION_TOTAL", porFuente, StringComparison.Ordinal);

        // Fuera del ASE afectado no hay hallazgo: el gate es por ASE, no global.
        Assert.DoesNotContain(errores, e => EsDelGate(e) && e.Contains("ASE 1", StringComparison.Ordinal));
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Caso ROJO 2 — dinero leído que no llega a ninguna celda (cierre por bloque)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// El insumo trae DOS grupos que el catálogo identifica como la MISMA fuente (sinonimia citada del
    /// Plan 34: <c>NUEVO ESQUEMA</c> ↔ <c>RECIPROCIDAD</c>). Cada grupo cierra solo —el gate por fuente
    /// no tiene nada que decir—, pero el bloque de la plantilla tiene un único destino por ASE: el
    /// segundo grupo no llega a ninguna celda. El cierre por bloque es el ÚNICO gate que ve ese dinero
    /// leído y no escrito (por eso no es una tautología del gate por fuente).
    /// </summary>
    [Fact]
    public void Gate_SigmaBloquesLeidosVsLoEscrito_DineroLeidoSinDestino_ReportaAseYDelta()
    {
        // Insumo SINTÉTICO (solo forma, nunca un monto real del negocio): ambos grupos cierran y el
        // segundo (55) se queda sin bloque del catálogo donde escribirse.
        var leafs = LeafsBase();
        leafs[0].AntExtRev = LecturaSintetica(
            1,
            Fuente("NUEVO ESQUEMA", 200m, 20m, -7m, 0m, totalFuente: 213m),
            Fuente("RECIPROCIDAD", 50m, 5m, 0m, 0m, totalFuente: 55m));

        var errores = new ValidadorBasico().Validar(ResultadoDe("2026091"), leafs);

        // El gate por fuente está en silencio: los dos grupos suman su propio total.
        Assert.DoesNotContain(
            errores, e => EsDelGate(e) && e.Contains("fuente '", StringComparison.Ordinal));

        var porBloque = Assert.Single(
            errores,
            e => EsDelGate(e) && e.Contains("por bloque", StringComparison.Ordinal));
        Assert.Contains("ASE 1", porBloque, StringComparison.Ordinal);
        Assert.Contains("(268)", porBloque, StringComparison.Ordinal);
        Assert.Contains("(213)", porBloque, StringComparison.Ordinal);
        Assert.Contains("Diferencia=55", porBloque, StringComparison.Ordinal);
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Herencia en la clasificación (Plan 42 D-F) y ficha de novedad (G-TRAZA)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// <see cref="ValidadorBasico.ClasificarDivergenciasDeclaradas"/> reutiliza el validador completo: la
    /// divergencia de ANT EXT-REV llega como NO DECLARADA (no hay entrada declarable para esta hoja y no
    /// se agrega ninguna), de modo que la clasificación no la puede hacer pasar por conocida.
    /// </summary>
    [Fact]
    public void Clasificacion_HeredaLaDivergenciaAntExtRev_ComoNoDeclarada()
    {
        var leafs = LeafsBase();
        leafs[1].AntExtRev = LecturaSintetica(2, Fuente("ENEL", 100m, 20m, -30m, 0m, totalFuente: 91m));

        var clasificacion = new ValidadorBasico().ClasificarDivergenciasDeclaradas(
            ResultadoDe("2026091"), leafs, SnapshotsVacios());

        Assert.True(clasificacion.HayNoDeclaradas);
        Assert.Empty(clasificacion.Veredictos);
        Assert.Contains(
            clasificacion.ErroresNoDeclarados,
            e => EsDelGate(e) && e.Contains("fuente 'ENEL'", StringComparison.Ordinal));
    }

    /// <summary>
    /// G-TRAZA: la novedad que el administrativo ve NOMBRA la hoja real (<c>ANT EXT-REV</c>), el ASE y
    /// los dos montos comparados. Una ficha que apunte a otra hoja manda a revisar donde no está el
    /// problema.
    /// </summary>
    [Fact]
    public void Novedad_DesdeCoherencia_NombraHojaAseYMontosComparados()
    {
        var leafs = LeafsBase();
        leafs[1].AntExtRev = LecturaSintetica(2, Fuente("ENEL", 100m, 20m, -30m, 0m, totalFuente: 91m));

        var errores = new ValidadorBasico().Validar(ResultadoDe("2026091"), leafs);
        var novedades = ClasificadorNovedadesEntrega.DesdeCoherencia(errores, PeriodoDe("2026091"));

        var fichas = novedades.Where(n => string.Equals(n.Hoja, MarcaHoja, StringComparison.Ordinal)).ToList();
        Assert.NotEmpty(fichas);
        Assert.All(fichas, ficha => Assert.Equal(2, ficha.AseId));

        var deLaFuente = Assert.Single(fichas, f => f.ValorObservado == 90m && f.ValorEsperado == 91m);
        Assert.Equal(-1m, deLaFuente.Delta);
        Assert.Contains(MarcaHoja, deLaFuente.Reporte, StringComparison.Ordinal);
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Apoyo: lectura real, contrato sintético, resultado y snapshots
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static bool EsDelGate(string error) => error.Contains(MarcaHoja, StringComparison.Ordinal);

    private static Periodo PeriodoDe(string codigo) => Periodo.Parse(codigo);

    private static AntExtRevAseInputs Leer(string periodo, int aseId) =>
        new ExcelDataReaderWorkbookLeafInputReader().LeerAntExtRev(
            AseFactory.DesdeId(aseId),
            PeriodoDe(periodo),
            Insumo(periodo, aseId));

    private static AntExtRevFuenteInputs Fuente(
        string etiqueta,
        decimal aplicadoR3,
        decimal anticiposR3,
        decimal aplicadoR4,
        decimal anticiposR4,
        decimal totalFuente) =>
        new(etiqueta, aplicadoR3, anticiposR3, aplicadoR4, anticiposR4, totalFuente);

    private static AntExtRevAseInputs LecturaSintetica(int aseId, params AntExtRevFuenteInputs[] fuentes) =>
        new() { Ase = AseFactory.DesdeId(aseId), Fuentes = fuentes };

    /// <summary>
    /// Cinco lecturas que CIERRAN (una fuente presente por ASE, con Σ de conceptos == total de fuente):
    /// cualquier divergencia que reporten los gates viene del caso que el test inyecta, no del apoyo.
    /// </summary>
    private static List<WorkbookLeafInputs> LeafsBase()
    {
        var fuentes = new[]
        {
            Fuente("OCCIDENTE", 100m, 10m, -5m, 0m, totalFuente: 105m),
            Fuente("ENEL", 100m, 20m, -30m, 0m, totalFuente: 90m),
            Fuente("NUEVO ESQUEMA", 200m, 20m, -7m, 0m, totalFuente: 213m),
            Fuente("ENEL", 300m, 30m, -9m, 0m, totalFuente: 321m),
            Fuente("OCCIDENTE", 40m, 0m, 0m, 0m, totalFuente: 40m)
        };

        var leafs = new List<WorkbookLeafInputs>();
        for (var aseId = 1; aseId <= 5; aseId++)
        {
            leafs.Add(new WorkbookLeafInputs
            {
                Ase = AseFactory.DesdeId(aseId),
                Periodo = PeriodoDe("2026091"),
                AntExtRev = LecturaSintetica(aseId, fuentes[aseId - 1])
            });
        }

        return leafs;
    }

    /// <summary>
    /// Resultado mínimo del período: cinco consolidados en cero (los gates R1/R2/R4 comparan contra los
    /// leafs en cero de <see cref="LeafsBase"/>, así que el único hallazgo posible es el que se inyecta).
    /// </summary>
    private static ResultadoRemuneracion ResultadoDe(string codigo)
    {
        var periodo = PeriodoDe(codigo);
        return new ResultadoRemuneracion
        {
            Periodo = periodo,
            Exitoso = true,
            Consolidados = Enumerable.Range(1, 5)
                .Select(aseId => new ConsolidadoAse
                {
                    Ase = AseFactory.DesdeId(aseId),
                    TotOpt = 0m,
                    R2TotalOportuno = 0m,
                    Extemp = 0m,
                    ReversionR4 = 0m,
                    AjustesSfT = 0m
                })
                .ToList()
        };
    }

    /// <summary>Snapshots vacíos (5 ASE): sin evidencia de validaciones cruzadas, sin veredictos.</summary>
    private static IReadOnlyList<ValidacionCruzadaSnapshot> SnapshotsVacios() =>
        Enumerable.Range(1, 5)
            .Select(aseId => new ValidacionCruzadaSnapshot
            {
                Ase = AseFactory.DesdeId(aseId),
                ValidacionTotal = 0m,
                ValidacionTotalOkP = true
            })
            .ToList();
}
