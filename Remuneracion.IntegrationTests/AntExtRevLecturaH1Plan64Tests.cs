using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 64 (H1.5, REQ-04/05/06 — CA-02): lectura del insumo <c>RecaudosReversados_*</c> y resolución del
/// destino de <c>ANT EXT-REV</c> por etiquetas.
///
/// <para><b>Oráculos (no se inventa ningún número).</b>
/// <list type="bullet">
/// <item>Los 30 insumos se contrastan contra la tabla CONGELADA por H0 (<c>Docs/T0-Plan64-Ant-Ext-Rev.md</c>
/// §4.3: total por fuente de cada archivo) y contra la invariante Σ-conceptos ≡ total-fuente.</item>
/// <item>Los 20 archivos con manual se contrastan CELDA A CELDA contra el manual del administrativo,
/// resolviendo la celda por la MISMA autoridad de producción
/// (<see cref="ResolvedorDestinoAntExtRev"/>), no por una tabla de letras del fixture.</item>
/// <item>Los casos negativos usan insumos SINTÉTICOS con la gramática medida: solo se manipula la forma
/// (marcadores, etiqueta, terna de valores, cierre aritmético), nunca un monto real.</item>
/// </list></para>
/// </summary>
public sealed class AntExtRevLecturaH1Plan64Tests
{
    private const decimal Tolerancia = 0.5m;

    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Prns = "http://schemas.openxmlformats.org/package/2006/relationships";

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Rutas explícitas (prohibido Insumos.Raiz() en lo nuevo)
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

    private static string PlantillaV4() =>
        Ruta("Docs", "Plantilla estandarizada", "v4", "Formato Remuneracion 2026xx-x Total v4.xlsx");

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

    private static readonly (string Periodo, string[] Ruta)[] Manuales =
    [
        ("2026072", ["Docs", "Prueba Julio-2", "Resultado", "Resultado Administrativo con plantilla estandar v2", "Formato Remuneracion 202607-2 Total.xlsx"]),
        ("2026091", ["Docs", "Prueba Septiembre 1", "Resultados", "Administrativo", "Remuneracion 202609-1 Total.xlsx"]),
        ("2026092", ["Docs", "Prueba Septiembre-2", "Resultados", "Administrativo", "Remuneracion 202609-2 Total.xlsx"]),
        ("2026101", ["Docs", "Prueba Octubre-1", "Resultados", "Resultado administrativo", "Remuneracion 202610-1 Total.xlsx"])
    ];

    private static string Manual(string periodo) => Ruta(Manuales.Single(m => m.Periodo == periodo).Ruta);

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Llamada a la unidad de producción
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static AntExtRevAseInputs Leer(string periodo, int aseId) =>
        new ExcelDataReaderWorkbookLeafInputReader().LeerAntExtRev(
            new Ase { Id = aseId, NombreCorto = $"ASE{aseId}" },
            Periodo.Parse(periodo),
            Insumo(periodo, aseId));

    private static decimal[] ValoresDe(AntExtRevFuenteInputs fuente) =>
        [fuente.AplicadoR3, fuente.AnticiposR3, fuente.AplicadoR4, fuente.AnticiposR4];

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // T0 §4.3 — tabla CONGELADA por H0 (evidencia, no contrato)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static readonly (string Periodo, int Ase, string Fuentes, decimal TotalBloque2)[] EstructuraT0 =
    [
        ("2026071", 1, "OCCIDENTE=-7402760", -7402760m),
        ("2026071", 2, "ENEL=-9917803|NUEVO ESQUEMA=-620493", -10538296m),
        ("2026071", 3, "ENEL=-5627802|OCCIDENTE=-10051780", -15679582m),
        ("2026071", 4, "ENEL=-3475400|NUEVO ESQUEMA=0|OCCIDENTE=-522720", -3998120m),
        ("2026071", 5, "ENEL=-5470988|OCCIDENTE=-1386390", -6857378m),
        ("2026072", 1, "ENEL=-15501271|OCCIDENTE=-1086060", -16587331m),
        ("2026072", 2, "ENEL=-9472909|OCCIDENTE=-178170", -9651079m),
        ("2026072", 3, "ENEL=-30776490|OCCIDENTE=-2650410", -33426900m),
        ("2026072", 4, "ENEL=-14910059|NUEVO ESQUEMA=0|OCCIDENTE=-14926610", -29836669m),
        ("2026072", 5, "ENEL=-2114880|OCCIDENTE=-663150", -2778030m),
        ("2026082", 1, "ENEL=0", 0m),
        ("2026082", 2, "ENEL=-10736662|OCCIDENTE=0", -10736662m),
        ("2026082", 3, "ENEL=-4617354", -4617354m),
        ("2026082", 4, "ENEL=-3442420|NUEVO ESQUEMA=0|OCCIDENTE=0", -3442420m),
        ("2026082", 5, "ENEL=-1454550|OCCIDENTE=-66010", -1520560m),
        ("2026091", 1, "ENEL=-3462950|OCCIDENTE=0", -3462950m),
        ("2026091", 2, "ENEL=-8739913", -8739913m),
        ("2026091", 3, "ENEL=-6939290|OCCIDENTE=0", -6939290m),
        ("2026091", 4, "ENEL=-3196930|NUEVO ESQUEMA=0|OCCIDENTE=-64000", -3260930m),
        ("2026091", 5, "ENEL=-8838800|OCCIDENTE=0", -8838800m),
        ("2026092", 1, "ENEL=-690790|OCCIDENTE=0", -690790m),
        ("2026092", 2, "ENEL=-15212817|NUEVO ESQUEMA=0|OCCIDENTE=0", -15212817m),
        ("2026092", 3, "ENEL=-6595914", -6595914m),
        ("2026092", 4, "ENEL=-4393374|NUEVO ESQUEMA=0|OCCIDENTE=0", -4393374m),
        ("2026092", 5, "ENEL=-4790099", -4790099m),
        ("2026101", 1, "ENEL=-3184910|OCCIDENTE=0", -3184910m),
        ("2026101", 2, "ENEL=-6297940|OCCIDENTE=0", -6297940m),
        ("2026101", 3, "ENEL=-2259272", -2259272m),
        ("2026101", 4, "ENEL=-1581684|NUEVO ESQUEMA=0|OCCIDENTE=0", -1581684m),
        ("2026101", 5, "ENEL=-4005397", -4005397m)
    ];

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Contrato y catálogo (REQ-05)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// REQ-05/D-03: el contrato crece de forma ADITIVA y la compatibilidad es por construcción —
    /// <c>WorkbookLeafInputs.AntExtRev</c> nace <c>null</c>, de modo que lo que no lee el insumo se
    /// comporta exactamente como antes (rollback de H2).
    /// </summary>
    [Fact]
    public void Contrato_AntExtRevPorDefectoEsNull_NoCambiaElComportamientoAnterior()
    {
        var leaf = new WorkbookLeafInputs();

        Assert.Null(leaf.AntExtRev);
        Assert.Empty(leaf.AntExtRev?.Fuentes ?? []);
        Assert.Empty(new AntExtRevAseInputs().Fuentes);
    }

    /// <summary>
    /// REQ-05: el catálogo es DATO citado — identidad (<c>OCCIDENTE</c>, <c>ENEL</c>) y sinonimia citada
    /// (<c>NUEVO ESQUEMA</c> ↔ EAAB Reciprocidad, Plan 34). HUE-1: CIUDAD LIMPIA y ENERBIT no tienen fila y
    /// una etiqueta no tabulada NO se mapea a nada (nunca un sinónimo inventado).
    /// </summary>
    [Fact]
    public void Catalogo_IdentidadSinonimiaYHueco1_SinFilasInventadas()
    {
        Assert.Equal("BANCO DE OCCIDENTE", CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente("OCCIDENTE")?.EtiquetaBloque);
        Assert.Equal("BANCO DE OCCIDENTE", CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente("occidente")?.EtiquetaBloque);
        Assert.Equal("ENEL", CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente("ENEL")?.EtiquetaBloque);
        Assert.Equal("EAAB RECIPROCIDAD", CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente("NUEVO ESQUEMA")?.EtiquetaBloque);

        // La sinonimia citada del Plan 34 se reutiliza (no se duplica la tabla): RECIPROCIDAD ≡ NUEVO ESQUEMA.
        Assert.Equal("EAAB RECIPROCIDAD", CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente("RECIPROCIDAD")?.EtiquetaBloque);

        // HUE-1 declarado: sin fila ⇒ null (el lector lo convierte en evidencia o fail-fast, nunca en 0).
        Assert.Null(CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente("EAAB CIUDAD LIMPIA"));
        Assert.Null(CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente("ENERBIT"));
        Assert.Null(CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente("EMPRESA FANTASMA"));

        // Catálogo ABIERTO del matcher (Plan 34): un rótulo no tabulado SOLO se iguala a sí mismo; jamás a
        // una fuente del catálogo. Sin cita no hay fila y sin fila no hay mapeo.
        Assert.True(SinonimosEmpresaR1.SonMismaEmpresa("EMPRESA FANTASMA", "empresa fantasma"));
        Assert.False(SinonimosEmpresaR1.SonMismaEmpresa("EMPRESA FANTASMA", "ENEL"));
        Assert.False(SinonimosEmpresaR1.SonMismaEmpresa("OCCIDENTE", "BANCO DE OCCIDENTE"));

        // El bloque sin fuente mapeada tampoco se resuelve por parecido (nunca un match difuso).
        Assert.Null(CatalogoFuentesAntExtRev.BuscarPorEtiquetaBloque("EAAB CIUDAD LIMPIA"));
        Assert.Null(CatalogoFuentesAntExtRev.BuscarPorEtiquetaBloque("ENERBIT"));
        Assert.Equal(3, CatalogoFuentesAntExtRev.Filas.Count);
        Assert.All(CatalogoFuentesAntExtRev.Filas, fila => Assert.False(string.IsNullOrWhiteSpace(fila.Cita)));

        // Conceptos: el orden congelado por T0 es el del registro de cada fuente.
        Assert.Equal(
            ["1-APLICADOS A FACTURACION", "2-SALDOS A FAVOR GENERADOS", "9-REVERSION PAGOS A COMPONENTES", "9-REVERSION SALDOS A FAVOR"],
            CatalogoConceptosAntExtRev.Conceptos.Select(concepto => concepto.Etiqueta).ToArray());
        Assert.Equal([0, 1, 2, 3], CatalogoConceptosAntExtRev.Conceptos.Select(concepto => concepto.Indice).ToArray());
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // CA-02 / REQ-04 — los 30 insumos ≡ T0 ±0,5 (lectura pura)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// CA-02: los 30 insumos (6 períodos × 5 ASE) se leen en puro y su total por fuente coincide con la
    /// tabla congelada por H0 (§4.3) ±0,5; el cierre Σ-conceptos ≡ total-fuente se cumple en todos y no hay
    /// etiquetas no tabuladas ni fuentes del catálogo sin traza de ausencia.
    /// </summary>
    [Fact]
    public void Lectura_Los30Insumos_TotalPorFuenteIgualAlT0()
    {
        Assert.Equal(30, EstructuraT0.Length);
        var vistas = 0;

        foreach (var esperado in EstructuraT0)
        {
            var leido = Leer(esperado.Periodo, esperado.Ase);
            var contexto = $"{esperado.Periodo} ASE{esperado.Ase}";

            Assert.Equal(esperado.Ase, leido.Ase.Id);
            Assert.NotEmpty(leido.Fuentes);
            Assert.Empty(leido.EtiquetasNoReconocidas);

            // Total por fuente, en el orden del archivo, contra la tabla congelada.
            Assert.Equal(esperado.Fuentes, string.Join(
                "|",
                leido.Fuentes.Select(fuente =>
                    $"{fuente.FuenteEtiqueta}={fuente.TotalFuente.ToString(CultureInfo.InvariantCulture)}")));

            // Invariante §1.4/T0 §4.4: Σ de los cuatro conceptos ≡ total de la fuente.
            foreach (var fuente in leido.Fuentes)
            {
                Assert.InRange(fuente.SumaConceptos - fuente.TotalFuente, -Tolerancia, Tolerancia);
                vistas++;
            }

            Assert.InRange(leido.TotalBloque2 - esperado.TotalBloque2, -Tolerancia, Tolerancia);

            // D-04: toda fuente del catálogo que no vino deja su traza; ninguna desaparece en silencio.
            var esperadasAusentes = CatalogoFuentesAntExtRev.Filas
                .Where(fila => !leido.Fuentes.Any(fuente =>
                    CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente(fuente.FuenteEtiqueta) is { } deLaFuente
                    && CatalogoFuentesAntExtRev.EsMismaFuente(deLaFuente, fila)))
                .Select(fila => $"ANTEXTREV:{fila.EtiquetasFuente[0]}")
                .Order(StringComparer.Ordinal)
                .ToArray();
            Assert.Equal(esperadasAusentes, leido.FuentesAusentes.Order(StringComparer.Ordinal).ToArray());
            Assert.Equal(CatalogoFuentesAntExtRev.Filas.Count - leido.Fuentes.Count, leido.FuentesAusentes.Count);

            // El control del bloque 3 cierra contra el bloque 2 en los 30 archivos (T0 §4.3): no hay hallazgo.
            Assert.NotNull(leido.TotalControlBloque3);
            Assert.InRange(leido.TotalControlBloque3!.Value - esperado.TotalBloque2, -Tolerancia, Tolerancia);
            Assert.Null(leido.HallazgoControlBloque3);
        }

        Assert.Equal(59, vistas); // 30 archivos: 22 con OCCIDENTE + 29 ENEL + 8 NUEVO ESQUEMA (T0 §5.1).
    }

    /// <summary>
    /// REQ-04 / M8: los cinco casos citados por el §1.4 del plan, re-derivados del insumo por la unidad de
    /// producción (no por un parser del fixture). Incluye los dos de ENEL de Sept-1.
    /// </summary>
    [Fact]
    public void Lectura_CasosCitadosDelPlan_CoincidenExactamente()
    {
        var casos = new (string Periodo, int Ase, string Fuente, decimal AplicadoR3, decimal AnticiposR3, decimal AplicadoR4, decimal AnticiposR4)[]
        {
            ("2026091", 1, "OCCIDENTE", 42680m, 40630m, -83310m, 0m),
            ("2026091", 1, "ENEL", 0m, 0m, -3461197.29m, -1752.71m),
            ("2026091", 4, "OCCIDENTE", 239834m, 1438701m, -1742535m, 0m),
            ("2026091", 4, "NUEVO ESQUEMA", 208970m, 53030m, -262000m, 0m),
            ("2026091", 4, "ENEL", 8186910m, 9321147m, -20245524.78m, -459462.22m)
        };

        foreach (var caso in casos)
        {
            var fuente = Leer(caso.Periodo, caso.Ase).Fuentes.Single(f => f.FuenteEtiqueta == caso.Fuente);
            var esperado = new[] { caso.AplicadoR3, caso.AnticiposR3, caso.AplicadoR4, caso.AnticiposR4 };
            var observado = ValoresDe(fuente);

            for (var i = 0; i < esperado.Length; i++)
            {
                Assert.InRange(
                    observado[i] - esperado[i],
                    -Tolerancia,
                    Tolerancia);
            }
        }
    }

    /// <summary>
    /// CA-02 / T0 §5.3: paridad completa <c>fuente × concepto</c> contra el manual del administrativo en los
    /// 4 períodos con manual (20 archivos). La celda destino se resuelve con la MISMA autoridad de producción
    /// (<see cref="ResolvedorDestinoAntExtRev"/>). Única celda vacía declarada del universo:
    /// <c>2026092!J20</c>, que el manual deja en blanco donde D-04 manda 0 (vacío ≡ 0, nunca se inventa).
    /// </summary>
    [Theory]
    [InlineData("2026072")]
    [InlineData("2026091")]
    [InlineData("2026092")]
    [InlineData("2026101")]
    public void Lectura_ParidadCeldaACelda_ContraElManual(string periodo)
    {
        var destinos = DestinosDelManual(periodo, out var celdas);
        var vacias = new List<string>();

        for (var ase = 1; ase <= 5; ase++)
        {
            var leido = Leer(periodo, ase);
            foreach (var fuente in leido.Fuentes)
            {
                var filaCatalogo = CatalogoFuentesAntExtRev.BuscarPorEtiquetaFuente(fuente.FuenteEtiqueta);
                Assert.NotNull(filaCatalogo);
                var valores = ValoresDe(fuente);

                foreach (var concepto in CatalogoConceptosAntExtRev.Conceptos)
                {
                    var destino = ResolvedorDestinoAntExtRev.Buscar(
                        destinos,
                        filaCatalogo!.EtiquetaBloque,
                        ase,
                        concepto.Fila,
                        concepto.Reporte);
                    Assert.NotNull(destino);

                    var esperado = valores[concepto.Indice];
                    if (!celdas.ContainsKey(destino!.Celda))
                    {
                        vacias.Add($"{periodo}!{destino.Celda}");
                        Assert.Equal(0m, esperado);
                    }

                    Assert.InRange(
                        ValorDeCelda(celdas, destino.Celda) - esperado,
                        -Tolerancia,
                        Tolerancia);
                }
            }
        }

        // §0.6 (hallazgo H1-manual): el único vacío del universo se declara; el resto son valores del manual.
        string[] esperadasVacias = periodo == "2026092" ? ["2026092!J20"] : [];
        Assert.Equal(esperadasVacias, vacias.ToArray());
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // REQ-06 — resolvedor por etiquetas (cero coordenadas en producción)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// REQ-06/D-05: la plantilla v4 resuelve 100 celdas de valor (5 bloques × 5 ASE × 2 filas × 2 reportes)
    /// únicas, derivadas de las etiquetas que la hoja declara. Las celdas citadas por T0 se afirman SOLO como
    /// evidencia del fixture (§1.2/§1.4): producción no las conoce.
    /// </summary>
    [Fact]
    public void Resolvedor_CienCeldasDeValorPorEtiquetas_ConEvidenciaT0()
    {
        var destinos = DestinosDeLaPlantilla(PlantillaV4());

        Assert.Equal(100, destinos.Count);
        Assert.Equal(100, destinos.Select(destino => destino.Celda).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.Equal(5, destinos.Select(destino => destino.EtiquetaBloque).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(5, destinos.Select(destino => destino.AseId).Distinct().Count());
        Assert.Equal([1, 2, 3, 4, 5], destinos.Select(destino => destino.AseId).Distinct().Order().ToArray());

        // Evidencia T0 §1.2/§1.4 (el fixture sí conoce las letras; producción no).
        Assert.Equal("C11", Celda(destinos, "BANCO DE OCCIDENTE", 1, "APLICADO", "REPORTE 3"));
        Assert.Equal("C12", Celda(destinos, "BANCO DE OCCIDENTE", 1, "ANTICIPOS", "REPORTE 3"));
        Assert.Equal("D11", Celda(destinos, "BANCO DE OCCIDENTE", 1, "APLICADO", "REPORTE 4"));
        Assert.Equal("D12", Celda(destinos, "BANCO DE OCCIDENTE", 1, "ANTICIPOS", "REPORTE 4"));
        Assert.Equal("I11", Celda(destinos, "BANCO DE OCCIDENTE", 4, "APLICADO", "REPORTE 3"));
        Assert.Equal("I19", Celda(destinos, "EAAB RECIPROCIDAD", 4, "APLICADO", "REPORTE 3"));
        Assert.Equal("J20", Celda(destinos, "EAAB RECIPROCIDAD", 4, "ANTICIPOS", "REPORTE 4"));
        Assert.Equal("D35", Celda(destinos, "ENEL", 1, "APLICADO", "REPORTE 4"));
        Assert.Equal("D36", Celda(destinos, "ENEL", 1, "ANTICIPOS", "REPORTE 4"));
        Assert.Equal("I35", Celda(destinos, "ENEL", 4, "APLICADO", "REPORTE 3"));
        Assert.Equal("I36", Celda(destinos, "ENEL", 4, "ANTICIPOS", "REPORTE 3"));
        Assert.Equal("J35", Celda(destinos, "ENEL", 4, "APLICADO", "REPORTE 4"));
        Assert.Equal("J36", Celda(destinos, "ENEL", 4, "ANTICIPOS", "REPORTE 4"));
    }

    /// <summary>
    /// REQ-06: la validación es plan-then-write —una plantilla a la que le falta una etiqueta de bloque NO
    /// resuelve destino y falla ANTES de escribir, nombrando hoja + etiqueta esperada.
    /// </summary>
    [Fact]
    public void Resolvedor_EtiquetaDeBloqueAusente_FallaAntesDeEscribir()
    {
        var copia = CopiarPlantilla();
        try
        {
            var destinos = DestinosDeLaPlantilla(copia);
            var celdaEtiqueta = CeldaDeEtiqueta(PlantillaV4(), "ENEL");
            EscribirTexto(copia, celdaEtiqueta, "ENEL ROTULO MOVIDO");
            Assert.NotEmpty(destinos);

            var excepcion = Assert.Throws<CalculoInvalidoException>(() => DestinosDeLaPlantilla(copia));

            Assert.Equal(CodigoError.Plantilla, excepcion.Codigo);
            Assert.Contains("ENEL", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("ANT EXT-REV", excepcion.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(copia);
        }
    }

    /// <summary>
    /// REQ-06/§2.5: una celda de VALOR que trae fórmula es una plantilla incompatible con la escritura
    /// values-only: el resolvedor la detecta antes de escribir y nombra la celda.
    /// </summary>
    [Fact]
    public void Resolvedor_CeldaDeValorConFormula_FallaAntesDeEscribir()
    {
        var copia = CopiarPlantilla();
        try
        {
            var destino = ResolvedorDestinoAntExtRev.Buscar(
                DestinosDeLaPlantilla(copia), "BANCO DE OCCIDENTE", 1, "APLICADO", "REPORTE 3");
            Assert.NotNull(destino);

            using (var documento = SpreadsheetDocument.Open(copia, true))
            {
                var hoja = HojaAbierta(documento, ResolvedorDestinoAntExtRev.Hoja);
                var celda = hoja.Descendants<Cell>()
                    .Single(c => string.Equals(c.CellReference?.Value, destino!.Celda, StringComparison.OrdinalIgnoreCase));
                celda.CellFormula = new CellFormula("1+1");
                hoja.Save();
            }

            var excepcion = Assert.Throws<CalculoInvalidoException>(() => DestinosDeLaPlantilla(copia));

            Assert.Equal(CodigoError.Plantilla, excepcion.Codigo);
            Assert.Contains(destino!.Celda, excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("BANCO DE OCCIDENTE", excepcion.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(copia);
        }
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // REQ-04 — fail-fast del lector (insumos sintéticos: solo se manipula la FORMA)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>§2.4 paso 2: marcadores invertidos (bloque 3 antes que bloque 2) = fail-fast.</summary>
    [Fact]
    public void Lectura_MarcadoresInvertidos_Falla()
    {
        var ruta = CrearInsumo(
        [
            ["Resumen Recaudo Aplicado Por Servicio", null, null, null, null, null],
            ["RECAUDO OTROS", "TOTAL", null, null, null, null],
            ["Total", null, null, 300m, 300m, 300m],
            ["Pagos Pendientes por Aplicar", "Banco", "Fecha Recaudo", "Cantidad", "Valor", null],
            [null, null, null, null, "ASEO", "Total"],
            [null, null, null, "66", "Total", null],
            [null, null, "1-APLICADOS A FACTURACION", 300m, 300m, 300m],
            ["ENEL", "Total", null, 300m, 300m, 300m]
        ]);

        try
        {
            var excepcion = Assert.Throws<CalculoInvalidoException>(() => LeerSintetico(ruta));
            Assert.Equal(CodigoError.FormatoFuente, excepcion.Codigo);
            Assert.Contains("invertidos", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("ASE 1", excepcion.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    /// <summary>§2.4 paso 2: marcador ausente = fail-fast que nombra el marcador y el archivo.</summary>
    [Fact]
    public void Lectura_MarcadorAusente_Falla()
    {
        var ruta = CrearInsumo(
        [
            ["Pagos Pendientes por Aplicar", "Banco", "Fecha Recaudo", "Cantidad", "Valor", null],
            [null, null, null, null, "ASEO", "Total"],
            [null, null, null, "66", "Total", null],
            [null, null, "1-APLICADOS A FACTURACION", 300m, 300m, 300m],
            ["ENEL", "Total", null, 300m, 300m, 300m]
        ]);

        try
        {
            var excepcion = Assert.Throws<CalculoInvalidoException>(() => LeerSintetico(ruta));
            Assert.Equal(CodigoError.FormatoFuente, excepcion.Codigo);
            Assert.Contains("Resumen Recaudo Aplicado Por Servicio", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("RecaudosReversados", excepcion.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    /// <summary>
    /// §2.4 paso 6 (A3-gemelo): una etiqueta NO tabulada que coexiste con una fuente del catálogo AUSENTE es
    /// estructura rota — puede ser la ausente con otro nombre y no se asume 0 ni se mapea por parecido.
    /// </summary>
    [Fact]
    public void Lectura_EtiquetaDesconocidaConFuenteAusente_Falla()
    {
        var ruta = CrearInsumoSintetico(
        [
            Grupo("ENEL", 300m),
            Grupo("EMPRESA FANTASMA", 200m)
        ]);

        try
        {
            var excepcion = Assert.Throws<CalculoInvalidoException>(() => LeerSintetico(ruta));
            Assert.Equal(CodigoError.FormatoFuente, excepcion.Codigo);
            Assert.Contains("EMPRESA FANTASMA", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("ANTEXTREV:OCCIDENTE", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("ANTEXTREV:NUEVO ESQUEMA", excepcion.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    /// <summary>
    /// §2.4 paso 6 (gemelo Plan 50/51): sin ausencias en el mismo archivo, la etiqueta desconocida es SOLO
    /// evidencia — no bloquea y viaja como <c>ANTEXTREV:{etiqueta}</c>.
    /// </summary>
    [Fact]
    public void Lectura_EtiquetaDesconocidaSinAusencia_EsSoloEvidencia()
    {
        var ruta = CrearInsumoSintetico(
        [
            Grupo("OCCIDENTE", 100m),
            Grupo("NUEVO ESQUEMA", 200m),
            Grupo("ENEL", 300m),
            Grupo("EMPRESA FANTASMA", 400m)
        ]);

        try
        {
            var leido = LeerSintetico(ruta);

            Assert.Equal(3, leido.Fuentes.Count);
            Assert.Empty(leido.FuentesAusentes);
            Assert.Equal(["ANTEXTREV:EMPRESA FANTASMA"], leido.EtiquetasNoReconocidas.ToArray());

            // El bloque 3 (control) sí incluye la etiqueta desconocida: el Δ queda registrado, no bloquea.
            Assert.NotNull(leido.TotalControlBloque3);
            Assert.NotNull(leido.HallazgoControlBloque3);
            Assert.Contains("Δ", leido.HallazgoControlBloque3!, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    /// <summary>§2.4 paso 6: la misma fuente rotulada dos veces hace ambiguo el destino ⇒ fail-fast.</summary>
    [Fact]
    public void Lectura_FuenteDuplicada_Falla()
    {
        var ruta = CrearInsumoSintetico(
        [
            Grupo("ENEL", 300m),
            Grupo("ENEL", 100m),
            Grupo("OCCIDENTE", 50m),
            Grupo("NUEVO ESQUEMA", 25m)
        ]);

        try
        {
            var excepcion = Assert.Throws<CalculoInvalidoException>(() => LeerSintetico(ruta));
            Assert.Equal(CodigoError.FormatoFuente, excepcion.Codigo);
            Assert.Contains("MÁS DE UNA VEZ", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("ENEL", excepcion.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    /// <summary>
    /// Enmienda M4 (BLOQUEANTE): la terna de valores del insumo son tres copias del mismo monto. Si divergen
    /// más de ±0,5 el insumo está roto y el mensaje muestra los tres valores (ASE + fuente + concepto + D/E/F
    /// observados). Prohibido "elegir una columna" o leer por letra fija.
    /// </summary>
    [Fact]
    public void Lectura_TernaDeValoresQueDiverge_FallaConLosTresValores()
    {
        var ruta = CrearInsumo(
        [
            ["Pagos Pendientes por Aplicar", "Banco", "Fecha Recaudo", "Cantidad", "Valor", null],
            [null, null, null, null, "ASEO", "Total"],
            [null, null, null, "66", "Total", null],
            [null, null, "1-APLICADOS A FACTURACION", 100m, 100m, 101m],
            ["ENEL", "Total", null, 100m, 100m, 101m],
            ["Resumen Recaudo Aplicado Por Servicio", null, null, null, null, null],
            ["RECAUDO OTROS", "TOTAL", null, null, null, null],
            ["Total", null, null, 100m, 100m, 100m]
        ]);

        try
        {
            var excepcion = Assert.Throws<CalculoInvalidoException>(() => LeerSintetico(ruta));
            Assert.Equal(CodigoError.FormatoFuente, excepcion.Codigo);
            Assert.Contains("101", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("ASE 1", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("1-APLICADOS A FACTURACION", excepcion.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    /// <summary>
    /// §2.4 paso 4 / D-06: Σ de los cuatro conceptos ≠ total de la fuente &gt; 0,5 ⇒ fail-fast con la fuente,
    /// la Σ y el total observado (el dinero del grupo no se puede atribuir sin cuadrar).
    /// </summary>
    [Fact]
    public void Lectura_SigmaDistintaDelTotalDeFuente_Falla()
    {
        var ruta = CrearInsumo(
        [
            ["Pagos Pendientes por Aplicar", "Banco", "Fecha Recaudo", "Cantidad", "Valor", null],
            [null, null, null, null, "ASEO", "Total"],
            [null, null, null, "66", "Total", null],
            [null, null, "1-APLICADOS A FACTURACION", 100m, 100m, 100m],
            ["ENEL", "Total", null, 200m, 200m, 200m],
            ["Resumen Recaudo Aplicado Por Servicio", null, null, null, null, null],
            ["RECAUDO OTROS", "TOTAL", null, null, null, null],
            ["Total", null, null, 200m, 200m, 200m]
        ]);

        try
        {
            var excepcion = Assert.Throws<CalculoInvalidoException>(() => LeerSintetico(ruta));
            Assert.Equal(CodigoError.FormatoFuente, excepcion.Codigo);
            Assert.Contains("ENEL", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("100", excepcion.Message, StringComparison.Ordinal);
            Assert.Contains("200", excepcion.Message, StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    /// <summary>
    /// D-04: un CONCEPTO ausente dentro de una fuente PRESENTE viaja como 0 explícito (nunca se inventa una
    /// fila ni se arrastra el monto de otro concepto).
    /// </summary>
    [Fact]
    public void Lectura_ConceptoAusenteEnFuentePresente_EsCeroExplicito()
    {
        var ruta = CrearInsumo(
        [
            ["Pagos Pendientes por Aplicar", "Banco", "Fecha Recaudo", "Cantidad", "Valor", null],
            [null, null, null, null, "ASEO", "Total"],
            [null, null, null, "66", "Total", null],
            [null, null, "1-APLICADOS A FACTURACION", 125m, 125m, 125m],
            ["ENEL", "Total", null, 125m, 125m, 125m],
            ["Resumen Recaudo Aplicado Por Servicio", null, null, null, null, null],
            ["RECAUDO OTROS", "TOTAL", null, null, null, null],
            ["Total", null, null, 125m, 125m, 125m]
        ]);

        try
        {
            var leido = LeerSintetico(ruta);
            var fuente = Assert.Single(leido.Fuentes);

            Assert.Equal("ENEL", fuente.FuenteEtiqueta);
            Assert.Equal(125m, fuente.AplicadoR3);
            Assert.Equal(0m, fuente.AnticiposR3);
            Assert.Equal(0m, fuente.AplicadoR4);
            Assert.Equal(0m, fuente.AnticiposR4);
            Assert.Equal(125m, fuente.TotalFuente);
            // H5.3 (Plan 64, doctrina de fixtures): el assert sobre 'SumaConceptos' era TAUTOLÓGICO — la
            // propiedad es la suma de los cuatro campos que ya se midieron arriba, no una medición
            // independiente. El cierre Σ ≡ total lo valida el lector (fail-fast) y el total de fuente sí se
            // contrasta con el valor esperado.

            // Traza de las otras dos fuentes del catálogo (ausencia registrada, no es error).
            Assert.Equal(["ANTEXTREV:OCCIDENTE", "ANTEXTREV:NUEVO ESQUEMA"], leido.FuentesAusentes.ToArray());
        }
        finally
        {
            File.Delete(ruta);
        }
    }

    /// <summary>
    /// Regresión M-B / enmienda M4: si la terna de valores se DESPLAZA una columna a la derecha (misma
    /// gramática de marcadores, concepto y cierre), el lector devuelve EXACTAMENTE los mismos montos — porque
    /// la terna se identifica por su invariante (tres copias del mismo valor) y no por una letra congelada.
    /// </summary>
    [Fact]
    public void Lectura_ColumnaDeValoresMovida_DevuelveElMismoValor()
    {
        var normal = CrearInsumoSintetico(
        [
            Grupo("OCCIDENTE", 110m, 220m, -330m, 0m),
            Grupo("NUEVO ESQUEMA", 10m, 20m, -30m, 0m),
            Grupo("ENEL", 1m, 2m, -3m, 0m)
        ]);
        var movida = CrearInsumoSintetico(
        [
            Grupo("OCCIDENTE", 110m, 220m, -330m, 0m, desplazamiento: 1),
            Grupo("NUEVO ESQUEMA", 10m, 20m, -30m, 0m, desplazamiento: 1),
            Grupo("ENEL", 1m, 2m, -3m, 0m, desplazamiento: 1)
        ]);

        try
        {
            var leidoNormal = LeerSintetico(normal);
            var leidoMovida = LeerSintetico(movida);

            Assert.Equal(NormalizarFuentes(leidoNormal), NormalizarFuentes(leidoMovida));
            Assert.Equal(
                "OCCIDENTE=110/220/-330/0;NUEVO ESQUEMA=10/20/-30/0;ENEL=1/2/-3/0",
                NormalizarFuentes(leidoNormal));
        }
        finally
        {
            File.Delete(normal);
            File.Delete(movida);
        }
    }

    private static string NormalizarFuentes(AntExtRevAseInputs leido) =>
        string.Join(
            ";",
            leido.Fuentes.Select(fuente =>
                $"{fuente.FuenteEtiqueta}=" + string.Join(
                    "/",
                    ValoresDe(fuente).Select(valor => valor.ToString(CultureInfo.InvariantCulture)))));

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Insumos sintéticos (OpenXml: solo forma, nunca un monto real del negocio)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Insumo sintético completo: marcadores, cabecera del bloque 2, los grupos indicados y un bloque 3 de
    /// control cuyo total es la Σ de los grupos (así el control cierra salvo que la prueba quiera otra cosa).
    /// </summary>
    private static string CrearInsumoSintetico(IReadOnlyList<object?[][]> grupos)
    {
        var filas = new List<object?[]>
        {
            new object?[] { "Pagos Pendientes por Aplicar", "Banco", "Fecha Recaudo", "Cantidad", "Valor" },
            new object?[] { null, null, null, null, "ASEO", "Total" },
            new object?[] { null, null, null, "66", "Total" }
        };

        decimal total = 0;
        foreach (var grupo in grupos)
        {
            foreach (var fila in grupo)
            {
                filas.Add(fila);
            }

            total += grupo[^1].OfType<decimal>().First();
        }

        filas.Add(["Resumen Recaudo Aplicado Por Servicio"]);
        filas.Add(["RECAUDO OCCIDENTE", "RECAUDO OTROS", "Total"]);
        filas.Add(["Total", null, total]);

        return CrearInsumo(filas);
    }

    /// <summary>
    /// Grupo de una fuente con la gramática medida en T0 §4.2: filas de concepto (etiqueta en la columna del
    /// concepto y la terna de valores a su derecha) y la fila de total rotulada con la etiqueta de la fuente.
    /// <paramref name="desplazamiento"/> corre la terna N columnas a la derecha (regresión M-B).
    /// </summary>
    private static object?[][] Grupo(
        string fuente,
        decimal aplicadoR3,
        decimal anticiposR3 = 0m,
        decimal aplicadoR4 = 0m,
        decimal anticiposR4 = 0m,
        int desplazamiento = 0)
    {
        var valores = new (string Etiqueta, decimal Valor)[]
        {
            ("1-APLICADOS A FACTURACION", aplicadoR3),
            ("2-SALDOS A FAVOR GENERADOS", anticiposR3),
            ("9-REVERSION PAGOS A COMPONENTES", aplicadoR4),
            ("9-REVERSION SALDOS A FAVOR", anticiposR4)
        };

        var filas = new List<object?[]>();
        foreach (var (etiqueta, valor) in valores)
        {
            filas.Add(FilaConValores(etiqueta, valor, desplazamiento));
        }

        var total = aplicadoR3 + anticiposR3 + aplicadoR4 + anticiposR4;
        var cierre = new object?[6 + desplazamiento];
        cierre[0] = fuente;
        cierre[1] = "Total";
        cierre[3 + desplazamiento] = total;
        cierre[4 + desplazamiento] = total;
        cierre[5 + desplazamiento] = total;
        filas.Add(cierre);

        return filas.ToArray();
    }

    private static object?[] FilaConValores(string concepto, decimal valor, int desplazamiento)
    {
        var fila = new object?[6 + desplazamiento];
        fila[2] = concepto;
        fila[3 + desplazamiento] = valor;
        fila[4 + desplazamiento] = valor;
        fila[5 + desplazamiento] = valor;
        return fila;
    }

    private static AntExtRevAseInputs LeerSintetico(string ruta) =>
        new ExcelDataReaderWorkbookLeafInputReader().LeerAntExtRev(
            new Ase { Id = 1, NombreCorto = "ASE1" },
            Periodo.Parse("2026091"),
            ruta);

    /// <summary>Escribe un .xlsx mínimo (una hoja <c>Sheet1</c>) con las filas indicadas.</summary>
    private static string CrearInsumo(IReadOnlyList<object?[]> filas)
    {
        var ruta = Path.Combine(Path.GetTempPath(), $"p64-h1-{Guid.NewGuid():N}.xlsx");
        using var documento = SpreadsheetDocument.Create(ruta, SpreadsheetDocumentType.Workbook);
        var parteWorkbook = documento.AddWorkbookPart();
        parteWorkbook.Workbook = new Workbook();
        var parteHoja = parteWorkbook.AddNewPart<WorksheetPart>();
        var datos = new SheetData();

        for (var i = 0; i < filas.Count; i++)
        {
            var fila = new Row { RowIndex = (uint)(i + 1) };
            var celdas = filas[i];
            for (var j = 0; j < celdas.Length; j++)
            {
                if (celdas[j] is null)
                {
                    continue;
                }

                var referencia = $"{LetraDeColumna(j + 1)}{i + 1}";
                if (celdas[j] is string texto)
                {
                    fila.Append(new Cell
                    {
                        CellReference = referencia,
                        DataType = CellValues.InlineString,
                        InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(texto))
                    });
                    continue;
                }

                fila.Append(new Cell
                {
                    CellReference = referencia,
                    CellValue = new CellValue(
                        Convert.ToDecimal(celdas[j], CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture))
                });
            }

            datos.Append(fila);
        }

        parteHoja.Worksheet = new Worksheet(datos);
        parteWorkbook.Workbook.AppendChild(new Sheets());
        parteWorkbook.Workbook.GetFirstChild<Sheets>()!.Append(new Sheet
        {
            Id = parteWorkbook.GetIdOfPart(parteHoja),
            SheetId = 1,
            Name = "Sheet1"
        });
        parteWorkbook.Workbook.Save();
        return ruta;
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════════
    // Utilidades de plantilla/manual (OpenXml + bytes)
    // ═════════════════════════════════════════════════════════════════════════════════════════════

    private static IReadOnlyList<ResolvedorDestinoAntExtRev.Destino> DestinosDeLaPlantilla(string ruta)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        return ResolvedorDestinoAntExtRev.Resolver(documento.WorkbookPart!);
    }

    private static IReadOnlyList<ResolvedorDestinoAntExtRev.Destino> DestinosDelManual(
        string periodo,
        out Dictionary<string, string> celdas)
    {
        var ruta = Manual(periodo);
        var destinos = DestinosDeLaPlantilla(ruta);
        celdas = LeerValoresDeHoja(ruta, ResolvedorDestinoAntExtRev.Hoja);
        return destinos;
    }

    private static string Celda(
        IReadOnlyList<ResolvedorDestinoAntExtRev.Destino> destinos,
        string etiquetaBloque,
        int aseId,
        string fila,
        string reporte)
    {
        var destino = ResolvedorDestinoAntExtRev.Buscar(destinos, etiquetaBloque, aseId, fila, reporte);
        Assert.NotNull(destino);
        Assert.Equal(etiquetaBloque, destino!.EtiquetaBloque);
        return destino.Celda;
    }

    private static string CopiarPlantilla()
    {
        var destino = Path.Combine(Path.GetTempPath(), $"p64-h1-plantilla-{Guid.NewGuid():N}.xlsx");
        File.Copy(PlantillaV4(), destino, overwrite: true);
        return destino;
    }

    private static Worksheet HojaAbierta(SpreadsheetDocument documento, string nombre)
    {
        var parte = documento.WorkbookPart!;
        var workbook = parte.Workbook ?? throw new InvalidOperationException("El workbook no tiene metadata válida.");
        var hoja = workbook.Descendants<Sheet>()
            .Single(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)parte.GetPartById(hoja.Id!.Value!)).Worksheet
            ?? throw new InvalidOperationException($"La hoja '{nombre}' no tiene Worksheet válido.");
    }

    private static string CeldaDeEtiqueta(string ruta, string etiqueta)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart!;
        var hoja = HojaAbierta(documento, ResolvedorDestinoAntExtRev.Hoja);
        var celda = hoja.Descendants<Cell>()
            .Single(c =>
                string.Equals(CeldaColumna(c.CellReference?.Value), "B", StringComparison.OrdinalIgnoreCase)
                && string.Equals(LectorTextoHoja.Texto(parte, c).Trim(), etiqueta, StringComparison.OrdinalIgnoreCase));
        return celda.CellReference!.Value!;
    }

    private static void EscribirTexto(string ruta, string referencia, string texto)
    {
        using var documento = SpreadsheetDocument.Open(ruta, true);
        var hoja = HojaAbierta(documento, ResolvedorDestinoAntExtRev.Hoja);
        var celda = hoja.Descendants<Cell>()
            .Single(c => string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        celda.CellValue = new CellValue(texto);
        celda.DataType = CellValues.String;
        hoja.Save();
    }

    private static string CeldaColumna(string? referencia) =>
        referencia is null ? string.Empty : new string(referencia.Where(char.IsLetter).ToArray());

    /// <summary>Valores (<c>&lt;v&gt;</c>) de una hoja leída por bytes; celda ausente = no está en el mapa.</summary>
    private static Dictionary<string, string> LeerValoresDeHoja(string ruta, string nombreHoja)
    {
        using var zip = ZipFile.OpenRead(ruta);

        XDocument workbook;
        using (var stream = zip.GetEntry("xl/workbook.xml")!.Open())
        {
            workbook = XDocument.Load(stream);
        }

        XDocument relaciones;
        using (var stream = zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open())
        {
            relaciones = XDocument.Load(stream);
        }

        var destinos = relaciones.Root!.Elements(Prns + "Relationship")
            .ToDictionary(r => r.Attribute("Id")!.Value, r => r.Attribute("Target")!.Value);
        var hoja = workbook.Root!.Element(Ns + "sheets")!.Elements(Ns + "sheet")
            .SingleOrDefault(s => (string?)s.Attribute("name") == nombreHoja);
        Assert.True(hoja is not null, $"El archivo '{ruta}' no trae la hoja '{nombreHoja}'.");

        var destino = destinos[hoja!.Attribute(Rns + "id")!.Value].Replace("worksheets/", "xl/worksheets/", StringComparison.Ordinal);
        if (!destino.StartsWith("xl/", StringComparison.Ordinal))
        {
            destino = "xl/" + destino;
        }

        XDocument hojaXml;
        using (var stream = zip.GetEntry(destino)!.Open())
        {
            hojaXml = XDocument.Load(stream);
        }

        return hojaXml.Descendants(Ns + "c")
            .Where(celda => celda.Attribute("r") is not null
                && !string.IsNullOrEmpty(celda.Element(Ns + "v")?.Value))
            .ToDictionary(
                celda => (string)celda.Attribute("r")!,
                celda => celda.Element(Ns + "v")!.Value!,
                StringComparer.OrdinalIgnoreCase);
    }

    private static decimal ValorDeCelda(IReadOnlyDictionary<string, string> celdas, string referencia)
    {
        if (!celdas.TryGetValue(referencia, out var valor))
        {
            return 0m;
        }

        return decimal.TryParse(valor, NumberStyles.Any, CultureInfo.InvariantCulture, out var numero) ? numero : 0m;
    }

    private static string LetraDeColumna(int columna)
    {
        var letra = string.Empty;
        while (columna > 0)
        {
            var resto = (columna - 1) % 26;
            letra = (char)('A' + resto) + letra;
            columna = (columna - 1) / 26;
        }

        return letra;
    }
}
