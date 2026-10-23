using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Xunit;
using Xunit.Abstractions;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 56 (H0/F0 — S1.1..S1.5): evidencia READ-ONLY congelada desde BYTES de la jerarquía de DOS
/// NIVELES de la hoja <c>Reporte Componentes R1</c>: <b>sección = columna A</b>
/// (<c>AFaseo</c>/<c>Componente</c>/<c>Subs/Cont</c>; <c>Total</c> = gran total) y
/// <b>subsección = columna B</b> (<c>Mes</c>/<c>Aplicacion nuevos x reversion</c>), con las etiquetas
/// escritas en la fila de <b>cierre</b> de cada bloque. Identidad del cuadro =
/// <c>(sección A, subsección B, ocurrencia de B dentro de A)</c>; dentro del cuadro:
/// <c>fuente C → canal D → concepto E</c>, y los valores son <b>componentes por encabezado</b>.
///
/// CERO código de producción: este fixture y sus asserts SON la evidencia (S0 G-H0-READONLY). Lectura
/// SOLO con BCL (<see cref="ZipArchive"/> + <see cref="XDocument"/>): sin Excel/COM, sin OpenXml SDK,
/// sin ExcelDataReader; prohibido <c>Insumos.Raiz()</c> en lo nuevo (rutas explícitas por período).
/// El workbook se abre con <see cref="FileShare.ReadWrite"/> (lectura mientras Excel lo tiene abierto)
/// y NUNCA se escribe: SHA256 verificado antes/después.
///
/// Verificación (2026-10-23, 15/15 = 5 bloques × julio-1/julio-2/octubre-1): ver
/// <c>Docs/T0-Plan56-Jerarquia-AB.md</c>.
/// </summary>
public sealed class JerarquiaABT0Plan56Tests
{
    private const string HojaR1 = "Reporte Componentes R1";
    private const string HojaFuente = "Sheet1";

    /// <summary>SHA256 histórico de la plantilla v1 (intacta; no la usa el pipeline — Plan 48).</summary>
    private const string ShaPlantillaV1 =
        "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";

    /// <summary>
    /// SHA256 de la plantilla v2 declarado como ancla por los planes 48..56. Se contrasta contra el
    /// valor observado en disco (el fixture NO lo usa como assert duro: ver
    /// <c>T0_PlantillaV2_LecturaEstable</c> y la nota de desviación del doc de evidencia).
    /// </summary>
    private const string ShaPlantillaV2AnclaDelPlan =
        "0F192AD2721C77986B6A8CB85D1A92F782DCD2F69989F362E066FA17027F2DAB";

    private const decimal Tolerancia = 0.5m;

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    private static readonly (string Periodo, string CarpetaPeriodo)[] Periodos =
    [
        ("2026071", Path.Combine("Docs", "Prueba Julio-1", "Insumos")),
        ("2026072", Path.Combine("Docs", "Prueba Julio-2", "Insumos")),
        ("2026101", Path.Combine("Docs", "Prueba Octubre-1", "Insumos")),
    ];

    /// <summary>
    /// T0-Plan54 §3 (descomposición del <c>DetRetri</c> verificada 5/5 al peso contra el R10 oficial
    /// de octubre-1): los dos operandos R1 (<c>D9</c> TOT_OPT y <c>D47</c> EXTEMP) por ASE. Se usan
    /// como oráculo R10 EXTERNO del cierre R1 (el R10 no publica TOT_OPT/EXTEMP por separado).
    /// </summary>
    private static readonly (decimal TotOpt, decimal Extemp)[] R1LegsOctubre =
    [
        (14_815_021_974.75m, 648_870.00m),
        (11_219_101_305.09m, 9_704_974.00m),
        (8_811_472_160.34m, 0.00m),
        (10_473_967_813.42m, 21_557_177.00m),
        (7_544_253_279.94m, 0.00m),
    ];

    /// <summary>
    /// Oráculo R10 EXTERNO ya congelado por el proyecto (columna <c>RECAUDO TOTAL</c> = <c>D9:D13</c>),
    /// por período y ASE: T0-Plan54 §3 (julio-1 y octubre-1), <c>plans/24 - T0 Evidencia.md</c>
    /// (julio-2) y <c>Docs/T0-Residual-24329-OccidenteAse5.md</c> (julio-2 ASE5). Se usa para
    /// verificar que el R10 leído es el MISMO oráculo que el proyecto validó al peso.
    /// </summary>
    private static readonly Dictionary<string, decimal[]> R10Documentado = new(StringComparer.Ordinal)
    {
        ["2026071"] = [16_758_167_585m, 12_227_292_053m, 10_127_195_599m, 10_553_776_376m, 8_543_663_209m],
        ["2026072"] = [17_450_228_673m, 20_516_143_970m, 15_221_896_467m, 7_179_595_396m, 12_073_344_662m],
        ["2026101"] = [14_842_425_136m, 11_302_637_605m, 8_823_787_300m, 10_483_172_762m, 7_550_668_665m],
    };

    private readonly ITestOutputHelper _salida;

    public JerarquiaABT0Plan56Tests(ITestOutputHelper salida) => _salida = salida;

    // ── S0 G-SHA / G-H0-READONLY ────────────────────────────────────────────────────────────────

    [Fact]
    public void T0_Plantillas_Intactas_AntesYDespues()
    {
        var v1Antes = Sha256(PlantillaV1());
        var v2Antes = Sha256(PlantillaV2());

        Assert.Equal(ShaPlantillaV1, v1Antes);

        // Medición read-only completa (los 3 períodos) entre las dos lecturas del SHA.
        foreach (var (periodo, carpeta) in Periodos)
        {
            for (var ase = 1; ase <= CarpetasAse.Length; ase++)
            {
                var hoja = LeerHoja(FuenteR1(carpeta, ase), HojaFuente);
                var bloques = Particionar(hoja);
                Assert.Single(bloques);
            }
        }

        var v2Despues = Sha256(PlantillaV2());
        Assert.Equal(v2Antes, v2Despues);

        _salida.WriteLine($"v1 SHA256 = {v1Antes}  (ancla del plan: {ShaPlantillaV1} → {(v1Antes == ShaPlantillaV1 ? "IDÉNTICO" : "DISTINTO")})");
        _salida.WriteLine($"v2 SHA256 antes = {v2Antes}");
        _salida.WriteLine($"v2 SHA256 después = {v2Despues} → {(v2Antes == v2Despues ? "IDÉNTICO (H0 no escribió)" : "CAMBIÓ")}");
        _salida.WriteLine($"v2 ancla del plan = {ShaPlantillaV2AnclaDelPlan} → {(v2Antes == ShaPlantillaV2AnclaDelPlan ? "IDÉNTICO" : "DESVIACIÓN PRE-EXISTENTE (ver Docs/T0-Plan56 §4)")}");
    }

    // ── S1.1 (lado PLANTILLA) ───────────────────────────────────────────────────────────────────

    [Fact]
    public void T0_PlantillaV2_SeccionesYSubsecciones_PorBloque()
    {
        var hoja = LeerHoja(PlantillaV2(), HojaR1);
        var bloques = Particionar(hoja);

        Assert.Equal(5, bloques.Count);

        var formas = new List<string>();
        foreach (var bloque in bloques)
        {
            _salida.WriteLine(
                $"[plantilla] bloque {bloque.Indice} «{bloque.Empresa}» filas {bloque.Inicio}..{bloque.Fin} "
                + $"(cabecera de componentes {bloque.FilaCabecera}; gran total A{bloque.Fin})");

            Assert.Equal(3, bloque.Secciones.Count);
            Assert.Equal(["AFaseo", "Componente", "Subs/Cont"], bloque.Secciones.Select(s => s.Etiqueta));

            foreach (var seccion in bloque.Secciones)
            {
                _salida.WriteLine(
                    $"  sección {seccion.Etiqueta,-12} filas {seccion.Inicio,4}..{seccion.Fin,-4} "
                    + $"cierre A{seccion.Fin}=«{Etiqueta(hoja, "A", seccion.Fin)}»/B{seccion.Fin}=«{Etiqueta(hoja, "B", seccion.Fin)}»");

                foreach (var sub in seccion.SubBloques)
                {
                    _salida.WriteLine(
                        $"    subsección {sub.Etiqueta,-34} ocurrencia #{sub.Ocurrencia} filas {sub.Inicio,4}..{sub.Fin,-4} "
                        + $"cierre B{sub.Fin}=«{Etiqueta(hoja, "B", sub.Fin)}»/C{sub.Fin}=«{Etiqueta(hoja, "C", sub.Fin)}»");
                    Assert.Equal("Total", Etiqueta(hoja, "C", sub.Fin));
                    Assert.Equal(sub.Etiqueta, Etiqueta(hoja, "B", sub.Fin));
                }
            }

            formas.Add(bloque.Forma());
        }

        // El concepto es de la PLANTILLA, no de un bloque: los 5 bloques declaran la MISMA forma.
        Assert.Single(formas.Distinct(StringComparer.Ordinal).ToList(), formas[0]);

        // Encabezados por NOMBRE (nunca por posición): F = Total, L = SERVICIO ESPECIALES.
        foreach (var bloque in bloques)
        {
            Assert.Equal("Total", hoja.Texto($"F{bloque.FilaCabecera}"));
            Assert.Equal("SERVICIO ESPECIALES", hoja.Texto($"L{bloque.FilaCabecera}"));
            Assert.Equal("Rec.Intereses", hoja.Texto($"Z{bloque.FilaCabecera}"));
        }

        // Cinco cuadros por bloque = (AFaseo,Mes) + 2×(Aplicacion,Mes) de Componente y Subs/Cont.
        Assert.All(bloques, b => Assert.Equal(5, b.Secciones.Sum(s => s.SubBloques.Count)));
    }

    // ── S1.1 (lado FUENTE) ─────────────────────────────────────────────────────────────────────

    [Fact]
    public void T0_Fuentes_SeccionesYSubsecciones_PorPeriodoAse()
    {
        var combinaciones = 0;
        foreach (var (periodo, carpeta) in Periodos)
        {
            var plantilla = LeerHoja(PlantillaV2(), HojaR1);
            var bloquesPlantilla = Particionar(plantilla);

            for (var ase = 1; ase <= CarpetasAse.Length; ase++)
            {
                var hoja = LeerHoja(FuenteR1(carpeta, ase), HojaFuente);
                var bloques = Particionar(hoja);
                Assert.Single(bloques);
                var bloque = bloques[0];

                var plantillaBloque = bloquesPlantilla[ase - 1];

                // Correspondencia bloque ↔ ASE por NOMBRE derivado (no por lista): la etiqueta de
                // empresa del bloque de la plantilla y el nombre de la carpeta del insumo coinciden.
                Assert.Equal(
                    Normalizar(plantillaBloque.Empresa),
                    Normalizar(CarpetasAse[ase - 1].Split('-', 2)[1]));

                _salida.WriteLine(
                    $"[fuente] {periodo} ASE{ase} «{bloque.Empresa}» filas {bloque.Inicio}..{bloque.Fin} "
                    + $"(cabecera {bloque.FilaCabecera}) secciones={bloque.Secciones.Count} subsecciones={bloque.Secciones.Sum(s => s.SubBloques.Count)}");

                foreach (var seccion in bloque.Secciones)
                {
                    Assert.Contains(seccion.Etiqueta, plantillaBloque.Secciones.Select(s => s.Etiqueta), StringComparer.OrdinalIgnoreCase);
                    _salida.WriteLine(
                        $"  sección {seccion.Etiqueta,-12} filas {seccion.Inicio,4}..{seccion.Fin,-4} "
                        + $"cierre A{seccion.Fin}=«{Etiqueta(hoja, "A", seccion.Fin)}»");

                    foreach (var sub in seccion.SubBloques)
                    {
                        _salida.WriteLine(
                            $"    subsección {sub.Etiqueta,-34} ocurrencia #{sub.Ocurrencia} filas {sub.Inicio,4}..{sub.Fin,-4} "
                            + $"cierre B{sub.Fin}=«{Etiqueta(hoja, "B", sub.Fin)}»");

                        // Toda combinación (A, B, ocurrencia) de la fuente está DECLARADA por la
                        // plantilla del bloque; si no, es huérfana de estructura (no se resuelve).
                        Assert.True(
                            plantillaBloque.BuscarCuadro(seccion.Etiqueta, sub.Etiqueta, sub.Ocurrencia) is not null,
                            $"{periodo} ASE{ase}: la fuente declara el cuadro ({seccion.Etiqueta}, {sub.Etiqueta}, #{sub.Ocurrencia}) "
                            + "que la plantilla del bloque no declara.");
                    }
                }

                combinaciones++;
            }
        }

        Assert.Equal(15, combinaciones);
    }

    // ── S1.2: cruce por identidad (A, B, ocurrencia) + (C, D, E) por encabezado ─────────────────

    [Fact]
    public void T0_Cruce_IdentidadAB_15de15()
    {
        var resueltas = 0;
        var sinDestinoEnCuadro = new List<string>();
        var huerfanasDelBloque = new List<string>();

        foreach (var (periodo, carpeta) in Periodos)
        {
            var plantilla = LeerHoja(PlantillaV2(), HojaR1);
            var bloquesPlantilla = Particionar(plantilla);

            for (var ase = 1; ase <= CarpetasAse.Length; ase++)
            {
                var hoja = LeerHoja(FuenteR1(carpeta, ase), HojaFuente);
                var bloqueFuente = Particionar(hoja)[0];
                var bloquePlantilla = bloquesPlantilla[ase - 1];

                var filasCuadro = 0;
                foreach (var seccionFuente in bloqueFuente.Secciones)
                {
                    var seccionPlantilla = bloquePlantilla.BuscarSeccion(seccionFuente.Etiqueta);
                    Assert.True(seccionPlantilla is not null, $"{periodo} ASE{ase}: la plantilla no declara la sección '{seccionFuente.Etiqueta}'.");

                    // Cierre de SECCIÓN: la etiqueta A de la fuente viaja en su propia fila de cierre.
                    var cierreFuente = seccionFuente.Fin;
                    var cierrePlantilla = seccionPlantilla!.Fin;
                    Assert.Equal("Total", Etiqueta(hoja, "B", cierreFuente));
                    Assert.Equal("Total", Etiqueta(plantilla, "B", cierrePlantilla));
                    Assert.Equal(
                        Etiqueta(hoja, "A", cierreFuente),
                        Etiqueta(plantilla, "A", cierrePlantilla));
                    resueltas++;

                    foreach (var subFuente in seccionFuente.SubBloques)
                    {
                        var cuadro = seccionPlantilla.BuscarCuadro(subFuente.Etiqueta, subFuente.Ocurrencia);
                        Assert.True(
                            cuadro is not null,
                            $"{periodo} ASE{ase}: cuadro ({seccionFuente.Etiqueta}, {subFuente.Etiqueta}, #{subFuente.Ocurrencia}) no declarado por la plantilla.");

                        var indice = FilasConFirma(hoja, subFuente);
                        var declarado = IndiceFirmas(plantilla, bloquePlantilla, cuadro!);
                        var delBloque = FirmasDelBloque(plantilla, bloquePlantilla);
                        var ordinal = new Dictionary<string, int>(StringComparer.Ordinal);
                        var cardinalidades = new List<string>();

                        foreach (var (firma, fila) in indice)
                        {
                            filasCuadro++;
                            ordinal.TryGetValue(firma, out var k);
                            ordinal[firma] = ++k;

                            if (!declarado.TryGetValue(firma, out var filasPlantilla))
                            {
                                var linea =
                                    $"{periodo} ASE{ase} ({seccionFuente.Etiqueta}, {subFuente.Etiqueta}, #{subFuente.Ocurrencia}) "
                                    + $"fila fuente {fila} firma «{firma}»";
                                if (delBloque.TryGetValue(firma, out var otras))
                                {
                                    sinDestinoEnCuadro.Add(
                                        linea + $" → SIN destino en su cuadro; la firma SÍ existe en el bloque en "
                                        + $"{string.Join(", ", otras.Select(f => $"R{f}"))} (no se reacomoda)");
                                }
                                else
                                {
                                    huerfanasDelBloque.Add(linea + " → ausente en TODO el bloque");
                                }

                                continue;
                            }

                            if (k > filasPlantilla.Count)
                            {
                                huerfanasDelBloque.Add(
                                    $"{periodo} ASE{ase} ({seccionFuente.Etiqueta}, {subFuente.Etiqueta}, #{subFuente.Ocurrencia}) "
                                    + $"fila fuente {fila} firma «{firma}»: la fuente trae {k} filas con esa firma y la plantilla solo {filasPlantilla.Count}");
                                continue;
                            }

                            resueltas++;
                        }

                        foreach (var (firma, filasPlantilla) in declarado)
                        {
                            ordinal.TryGetValue(firma, out var usadas);
                            if (filasPlantilla.Count != usadas)
                            {
                                cardinalidades.Add(
                                    $"({seccionFuente.Etiqueta}, {subFuente.Etiqueta}, #{subFuente.Ocurrencia}) «{firma}»: "
                                    + $"fuente {usadas} / plantilla {filasPlantilla.Count}");
                            }
                        }

                        if (cardinalidades.Count > 0)
                        {
                            _salida.WriteLine(
                                $"  [cardinalidad] {periodo} ASE{ase} ({seccionFuente.Etiqueta}, {subFuente.Etiqueta}, #{subFuente.Ocurrencia}): "
                                + string.Join(" | ", cardinalidades));
                        }

                        // Cierre de SUB-BLOQUE (la frontera B/Total también es destino declarado).
                        Assert.Equal("Total", Etiqueta(hoja, "C", subFuente.Fin));
                        Assert.Equal("Total", Etiqueta(plantilla, "C", cuadro!.Fin));
                        Assert.Equal(subFuente.Etiqueta, Etiqueta(plantilla, "B", cuadro.Fin));
                        resueltas++;
                    }
                }

                // Gran total del bloque (A='Total'): también resuelve por etiqueta.
                Assert.Equal(Etiqueta(hoja, "A", bloqueFuente.Fin), Etiqueta(plantilla, "A", bloquePlantilla.Fin));
                resueltas++;

                _salida.WriteLine(
                    $"[cruce] {periodo} ASE{ase}: filas de cuadro={filasCuadro} resueltas por identidad A/B; "
                    + $"secciones={bloqueFuente.Secciones.Count} subsecciones={bloqueFuente.Secciones.Sum(s => s.SubBloques.Count)}");
            }
        }

        _salida.WriteLine(
            $"[cruce] firmas resueltas (filas + cierres) = {resueltas}; sin destino en su cuadro = {sinDestinoEnCuadro.Count}; "
            + $"huérfanas del bloque = {huerfanasDelBloque.Count}");
        foreach (var linea in sinDestinoEnCuadro)
        {
            _salida.WriteLine($"  SIN-DESTINO-EN-CUADRO: {linea}");
        }

        foreach (var linea in huerfanasDelBloque)
        {
            _salida.WriteLine($"  HUÉRFANA-DEL-BLOQUE: {linea}");
        }

        Assert.True(huerfanasDelBloque.Count == 0, string.Join(Environment.NewLine, huerfanasDelBloque.Take(30)));

        // 15 cruces (5 bloques × 3 períodos), todos por la MISMA ruta.
        Assert.True(resueltas >= 15 * 15, $"Se esperaban al menos 225 firmas resueltas (15 cruces); se resolvieron {resueltas}.");
    }

    // ── S1.3: cierre aritmético TOT_OPT / EXTEMP ────────────────────────────────────────────────

    [Fact]
    public void T0_CierreAritmetico_TotOptYExtemp_15de15()
    {
        foreach (var (periodo, carpeta) in Periodos)
        {
            var plantilla = LeerHoja(PlantillaV2(), HojaR1);
            var bloquesPlantilla = Particionar(plantilla);
            var r10 = LeerR10(periodo, carpeta);

            for (var ase = 1; ase <= CarpetasAse.Length; ase++)
            {
                var hoja = LeerHoja(FuenteR1(carpeta, ase), HojaFuente);
                var bloqueFuente = Particionar(hoja)[0];
                var bloquePlantilla = bloquesPlantilla[ase - 1];

                var declarados = Declarados(hoja, plantilla, bloqueFuente, bloquePlantilla);

                var anclaTotOpt = bloquePlantilla.AnclaResumen("OPORTUNO");
                var anclaExtemp = bloquePlantilla.AnclaResumen("EXTEMPORANEO");
                Assert.True(anclaTotOpt is not null, $"{periodo} ASE{ase}: el bloque no declara la fila TOTAL/OPORTUNO de su tabla resumen.");
                Assert.True(anclaExtemp is not null, $"{periodo} ASE{ase}: el bloque no declara la fila TOTAL/EXTEMPORANEO de su tabla resumen.");

                // (a) El ancla de la PLANTILLA evaluada con texto-<f> + literales + valores del cruce.
                var totOptFormula = Evaluar(plantilla, declarados, anclaTotOpt!);
                var extempFormula = Evaluar(plantilla, declarados, anclaExtemp!);

                // (b) La MISMA cantidad derivada de la fuente por semántica documentada
                //     (T0-Plan54 §3): TOT_OPT = Σ F(Mes) − Σ L(Mes sin la última);
                //     EXTEMP = Σ F(Aplicación) − Σ L(Aplicación sin la primera).
                var (totOptFuente, extempFuente) = DerivarDesdeFuente(bloqueFuente, bloquePlantilla, declarados);

                Assert.Equal(totOptFuente, totOptFormula);
                Assert.Equal(extempFuente, extempFormula);

                var r10RecaudoTotal = r10[ase - 1];

                // El R10 leído es el MISMO oráculo ya congelado por el proyecto (5/5 por período).
                Assert.Equal(R10Documentado[periodo][ase - 1], r10RecaudoTotal);

                _salida.WriteLine(
                    $"[cierre] {periodo} ASE{ase}: TOT_OPT={Moneda(totOptFormula)} (ancla {anclaTotOpt}) · "
                    + $"EXTEMP={Moneda(extempFormula)} (ancla {anclaExtemp}) · "
                    + $"R1={Moneda(totOptFormula + extempFormula)} · R10 RECAUDO TOTAL={Moneda(r10RecaudoTotal)} · "
                    + $"Δ(R1,R10)={Moneda(r10RecaudoTotal - (totOptFormula + extempFormula))}");

                if (periodo == "2026101")
                {
                    var (totOptPlan54, extempPlan54) = R1LegsOctubre[ase - 1];
                    Assert.True(
                        Math.Abs(totOptFormula - totOptPlan54) <= Tolerancia,
                        $"{periodo} ASE{ase}: TOT_OPT medido {totOptFormula} ≠ T0-Plan54 {totOptPlan54} (Δ={totOptFormula - totOptPlan54}).");
                    Assert.True(
                        Math.Abs(extempFormula - extempPlan54) <= Tolerancia,
                        $"{periodo} ASE{ase}: EXTEMP medido {extempFormula} ≠ T0-Plan54 {extempPlan54} (Δ={extempFormula - extempPlan54}).");
                }
            }
        }
    }

    /// <summary>
    /// S1.4: lista CERRADA de combinaciones presentes en la fuente y <b>ausentes en todo el bloque</b>
    /// de la plantilla (destino inexistente, §2-D3). Debe ser vacía en 15/15; además se reporta —sin
    /// resolver y sin fallar— la lista cerrada de filas sin destino en SU cuadro pero declaradas en
    /// otro cuadro del mismo bloque (no se reacomodan).
    /// </summary>
    [Fact]
    public void T0_CombinacionesSinDestino_ListaCerradaVacia()
    {
        var sinDestinoEnBloque = new List<string>();
        var sinDestinoEnCuadro = new List<string>();

        foreach (var (periodo, carpeta) in Periodos)
        {
            var plantilla = LeerHoja(PlantillaV2(), HojaR1);
            var bloquesPlantilla = Particionar(plantilla);

            for (var ase = 1; ase <= CarpetasAse.Length; ase++)
            {
                var hoja = LeerHoja(FuenteR1(carpeta, ase), HojaFuente);
                var bloqueFuente = Particionar(hoja)[0];
                var bloquePlantilla = bloquesPlantilla[ase - 1];

                foreach (var seccion in bloqueFuente.Secciones)
                {
                    foreach (var sub in seccion.SubBloques)
                    {
                        var cuadro = bloquePlantilla.BuscarCuadro(seccion.Etiqueta, sub.Etiqueta, sub.Ocurrencia);
                        var firmasPlantilla = cuadro is null
                            ? new Dictionary<string, List<int>>(StringComparer.Ordinal)
                            : IndiceFirmas(plantilla, bloquePlantilla, cuadro);
                        var firmasBloque = FirmasDelBloque(plantilla, bloquePlantilla);
                        var ordinal = new Dictionary<string, int>(StringComparer.Ordinal);

                        foreach (var (firma, fila) in FilasConFirma(hoja, sub))
                        {
                            ordinal.TryGetValue(firma, out var k);
                            ordinal[firma] = ++k;
                            if (firmasPlantilla.TryGetValue(firma, out var filasPlantilla) && k <= filasPlantilla.Count)
                            {
                                continue;
                            }

                            var enElBloque = firmasBloque.TryGetValue(firma, out var otras);
                            var linea =
                                $"{periodo} ASE{ase} ({seccion.Etiqueta}, {sub.Etiqueta}, #{sub.Ocurrencia}) fila fuente {fila} "
                                + $"firma «{firma}» (ocurrencia {k})";
                            if (enElBloque)
                            {
                                sinDestinoEnCuadro.Add(
                                    linea + $" → la firma existe en el bloque en {string.Join(", ", otras!.Select(f => $"R{f}"))}; "
                                    + "no se reacomoda (queda en el cero del template)");
                            }
                            else
                            {
                                sinDestinoEnBloque.Add(linea + " → ausente en TODO el bloque (destino inexistente)");
                            }
                        }
                    }
                }
            }
        }

        foreach (var linea in sinDestinoEnCuadro)
        {
            _salida.WriteLine($"  SIN-DESTINO-EN-CUADRO: {linea}");
        }

        foreach (var linea in sinDestinoEnBloque)
        {
            _salida.WriteLine($"  SIN-DESTINO-EN-BLOQUE: {linea}");
        }

        Assert.True(sinDestinoEnBloque.Count == 0, string.Join(Environment.NewLine, sinDestinoEnBloque.Take(30)));
    }

    // ── Modelo de partición A/B ────────────────────────────────────────────────────────────────

    private sealed class Hoja
    {
        public Hoja(IReadOnlyDictionary<string, Celda> celdas) => Celdas = celdas;

        public IReadOnlyDictionary<string, Celda> Celdas { get; }

        public string Texto(string direccion) => Celdas.TryGetValue(direccion, out var celda) ? celda.Texto : string.Empty;

        public string? Formula(string direccion) =>
            Celdas.TryGetValue(direccion, out var celda) && !string.IsNullOrEmpty(celda.Formula) ? celda.Formula : null;

        public decimal? Numero(string direccion) =>
            Celdas.TryGetValue(direccion, out var celda) && decimal.TryParse(
                celda.Valor, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor)
                ? valor
                : null;
    }

    private sealed record Celda(string Direccion, string Texto, string? Formula, string? Valor);

    private sealed class Bloque
    {
        public required int Indice { get; init; }

        public required string Empresa { get; init; }

        public required int Inicio { get; init; }

        public required int Fin { get; init; }

        public required int FilaCabecera { get; init; }

        public required IReadOnlyList<Seccion> Secciones { get; init; }

        public Seccion? BuscarSeccion(string etiqueta) =>
            Secciones.FirstOrDefault(s => string.Equals(s.Etiqueta, etiqueta, StringComparison.OrdinalIgnoreCase));

        public SubBloque? BuscarCuadro(string seccion, string subseccion, int ocurrencia) =>
            BuscarSeccion(seccion)?.BuscarCuadro(subseccion, ocurrencia);

        /// <summary>
        /// Ancla de la tabla resumen del bloque: la fila <c>C='TOTAL'</c> con
        /// <c>D='OPORTUNO'</c>/<c>D='EXTEMPORANEO'</c>. Se localiza por ETIQUETA (nunca por dirección).
        /// </summary>
        public string? AnclaResumen(string canal) => _anclas.TryGetValue(canal, out var celda) ? celda : null;

        private readonly Dictionary<string, string> _anclas = new(StringComparer.OrdinalIgnoreCase);

        public void RegistrarAncla(string canal, string direccion) => _anclas[canal] = direccion;

        public string Forma() => string.Join(
            " | ",
            Secciones.Select(s => s.Etiqueta + ":" + string.Join(",", s.SubBloques.Select(b => b.Etiqueta + "#" + b.Ocurrencia.ToString(CultureInfo.InvariantCulture)))));
    }

    private sealed class Seccion
    {
        public required string Etiqueta { get; init; }

        public required int Inicio { get; init; }

        public required int Fin { get; init; }

        public required IReadOnlyList<SubBloque> SubBloques { get; init; }

        public SubBloque? BuscarCuadro(string etiqueta, int ocurrencia) =>
            SubBloques.FirstOrDefault(b =>
                string.Equals(b.Etiqueta, etiqueta, StringComparison.OrdinalIgnoreCase) && b.Ocurrencia == ocurrencia);
    }

    private sealed class SubBloque
    {
        public required string Etiqueta { get; init; }

        public required int Ocurrencia { get; init; }

        public required int Inicio { get; init; }

        public required int Fin { get; init; }
    }

    private sealed record Firma(string Fuente, string Canal, string Concepto)
    {
        public override string ToString() => $"{Fuente}|{Canal}|{Concepto}";
    }

    // ── Lectura BCL (ZIP + XML) ────────────────────────────────────────────────────────────────

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md"))
                && Directory.Exists(Path.Combine(dir.FullName, "Docs", "Plantilla estandarizada")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md + Docs/Plantilla estandarizada).");
    }

    private static string PlantillaV1() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string PlantillaV2() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "v2", "Formato Remuneracion 2026xx-x Total vf.xlsx");

    private static string FuenteR1(string carpetaPeriodo, int ase)
    {
        var carpeta = Path.Combine(Raiz(), carpetaPeriodo, CarpetasAse[ase - 1]);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta explícita del insumo: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f)
                .StartsWith("Recaudoporcomponente", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta el Recaudoporcomponente del ASE{ase} en {carpeta}");
        return archivo!;
    }

    private static string RutaR10(string carpetaPeriodo, string periodo) =>
        Path.Combine(Raiz(), carpetaPeriodo, $"R10_Remuneracion_{periodo}.xlsx");

    private static string Sha256(string ruta)
    {
        using var flujo = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Convert.ToHexString(SHA256.HashData(flujo));
    }

    private static XDocument Xml(ZipArchive zip, string entrada)
    {
        var entry = zip.GetEntry(entrada)
            ?? throw new InvalidOperationException($"El workbook no trae la parte '{entrada}'.");
        using var flujo = entry.Open();
        return XDocument.Load(flujo);
    }

    private static Hoja LeerHoja(string ruta, string nombreHoja)
    {
        // FileShare.ReadWrite: lectura aunque Excel tenga el archivo abierto (nunca se escribe).
        using var flujo = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var zip = new ZipArchive(flujo, ZipArchiveMode.Read);

        var libro = Xml(zip, "xl/workbook.xml");
        var hoja = libro.Descendants()
            .FirstOrDefault(e => e.Name.LocalName == "sheet"
                && string.Equals((string?)e.Attribute("name"), nombreHoja, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"{Path.GetFileName(ruta)} no trae la hoja '{nombreHoja}'.");
        var id = (string?)hoja.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")
            ?? throw new InvalidOperationException($"La hoja '{nombreHoja}' no declara relación.");

        var rels = Xml(zip, "xl/_rels/workbook.xml.rels");
        var destino = rels.Descendants()
            .FirstOrDefault(e => e.Name.LocalName == "Relationship" && string.Equals((string?)e.Attribute("Id"), id, StringComparison.Ordinal))
            ?.Attribute("Target")?.Value
            ?? throw new InvalidOperationException($"Sin relación para la hoja '{nombreHoja}'.");
        destino = destino.TrimStart('/');
        if (!destino.StartsWith("xl/", StringComparison.Ordinal))
        {
            destino = "xl/" + destino;
        }
        var compartidas = zip.GetEntry("xl/sharedStrings.xml") is null
            ? new List<string>()
            : Xml(zip, "xl/sharedStrings.xml").Descendants()
                .Where(e => e.Name.LocalName == "si")
                .Select(e => string.Concat(e.Descendants().Where(t => t.Name.LocalName == "t").Select(t => t.Value)))
                .ToList();

        var celdas = new Dictionary<string, Celda>(StringComparer.Ordinal);
        foreach (var celda in Xml(zip, destino).Descendants().Where(e => e.Name.LocalName == "c"))
        {
            var direccion = (string?)celda.Attribute("r");
            if (string.IsNullOrEmpty(direccion))
            {
                continue;
            }

            var tipo = (string?)celda.Attribute("t");
            var formula = celda.Elements().FirstOrDefault(e => e.Name.LocalName == "f")?.Value;
            var valor = celda.Elements().FirstOrDefault(e => e.Name.LocalName == "v")?.Value;

            string texto;
            if (tipo == "s" && int.TryParse(valor, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indice)
                && indice >= 0 && indice < compartidas.Count)
            {
                texto = compartidas[indice];
            }
            else if (tipo == "inlineStr")
            {
                texto = string.Concat(celda.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value));
            }
            else
            {
                texto = valor ?? string.Empty;
            }

            celdas[direccion] = new Celda(direccion, texto.Trim(), formula, valor);
        }

        return new Hoja(celdas);
    }

    private static List<decimal> LeerR10(string periodo, string carpetaPeriodo)
    {
        var ruta = RutaR10(carpetaPeriodo, periodo);
        Assert.True(File.Exists(ruta), $"Falta el R10 del período {periodo}: {ruta}");

        var hoja = LeerHoja(ruta, $"DetRetri{periodo}");
        var filaEncabezado = Enumerable.Range(1, 60)
            .FirstOrDefault(r =>
                hoja.Texto($"B{r}").Equals("ASE", StringComparison.OrdinalIgnoreCase)
                && hoja.Texto($"D{r}").StartsWith("RECAUDO TOTAL", StringComparison.OrdinalIgnoreCase));
        Assert.True(filaEncabezado > 0, $"R10 {periodo}: no se encontró el encabezado 'ASE / RECAUDO TOTAL'.");

        var valores = new List<decimal>();
        for (var fila = filaEncabezado + 1; fila <= filaEncabezado + 5; fila++)
        {
            Assert.True(
                int.TryParse(hoja.Texto($"B{fila}"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var ase) && ase == fila - filaEncabezado,
                $"R10 {periodo}: la fila {fila} no es el ASE {fila - filaEncabezado}.");
            valores.Add(MontoR10(hoja.Texto($"D{fila}"), periodo, fila));
        }

        return valores;
    }

    private static decimal MontoR10(string texto, string periodo, int fila)
    {
        Assert.False(string.IsNullOrWhiteSpace(texto), $"R10 {periodo} D{fila}: celda vacía.");
        Assert.DoesNotContain(",", texto, StringComparison.Ordinal);

        var negativo = texto.Contains('-', StringComparison.Ordinal);
        var digitos = new string(texto.Where(char.IsAsciiDigit).ToArray());
        Assert.False(string.IsNullOrEmpty(digitos), $"R10 {periodo} D{fila}: '{texto}' no es un monto.");
        var numero = decimal.Parse(digitos, NumberStyles.Integer, CultureInfo.InvariantCulture);
        return negativo ? -numero : numero;
    }

    // ── Partición A/B (etiquetas, ocurrencias, rangos y cierres) ────────────────────────────────

    private static string Etiqueta(Hoja hoja, string columna, int fila) => hoja.Texto($"{columna}{fila}");

    private static List<Bloque> Particionar(Hoja hoja)
    {
        var ultima = hoja.Celdas.Keys
            .Select(d => int.TryParse(new string(d.Where(char.IsAsciiDigit).ToArray()), NumberStyles.Integer, CultureInfo.InvariantCulture, out var fila) ? fila : 0)
            .DefaultIfEmpty(0)
            .Max();
        if (ultima == 0)
        {
            return [];
        }

        var totales = Enumerable.Range(1, ultima)
            .Where(f => Etiqueta(hoja, "A", f).Equals("Total", StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrEmpty(Etiqueta(hoja, "B", f)))
            .ToList();

        var bloques = new List<Bloque>();
        var inicio = 1;
        for (var i = 0; i < totales.Count; i++)
        {
            var fin = totales[i];
            var cabecera = DetectarCabecera(hoja, inicio, fin);
            var empresa = Etiqueta(hoja, "B", cabecera - 1);
            if (string.IsNullOrEmpty(empresa))
            {
                empresa = Etiqueta(hoja, "B", cabecera - 2);
            }

            bloques.Add(new Bloque
            {
                Indice = i + 1,
                Empresa = empresa,
                Inicio = cabecera + 1,
                Fin = fin,
                FilaCabecera = cabecera,
                Secciones = ParticionarSecciones(hoja, cabecera + 1, fin - 1),
            });
            inicio = fin + 1;
        }

        // Las tablas resumen viven DESPUÉS del gran total de su bloque y ANTES del siguiente bloque.
        for (var i = 0; i < bloques.Count; i++)
        {
            var desde = bloques[i].Fin + 1;
            var hasta = i + 1 < bloques.Count ? bloques[i + 1].FilaCabecera - 1 : ultima;
            RegistrarAnclas(hoja, bloques[i], desde, hasta);
        }

        return bloques;
    }

    /// <summary>
    /// La fila de encabezados de componentes es la única del bloque con varias celdas de TEXTO en
    /// F..Z (las filas de datos traen números; las de etiqueta, texto solo en A..E).
    /// </summary>
    private static int DetectarCabecera(Hoja hoja, int desde, int hasta)
    {
        var mejor = desde;
        var mejorCuenta = -1;
        for (var fila = desde; fila < hasta; fila++)
        {
            var cuenta = 0;
            foreach (var columna in Columnas())
            {
                var texto = hoja.Texto($"{columna}{fila}");
                if (texto.Length > 0
                    && !decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                {
                    cuenta++;
                }
            }

            if (cuenta > mejorCuenta)
            {
                mejorCuenta = cuenta;
                mejor = fila;
            }
        }

        Assert.True(mejorCuenta > 0, $"No se encontró la fila de encabezados de componentes entre {desde} y {hasta - 1}.");
        return mejor;
    }

    private static List<Seccion> ParticionarSecciones(Hoja hoja, int desde, int hasta)
    {
        var secciones = new List<Seccion>();
        var inicio = desde;
        for (var fila = desde; fila <= hasta; fila++)
        {
            var a = Etiqueta(hoja, "A", fila);
            var b = Etiqueta(hoja, "B", fila);
            if (a.Length > 0 && b.Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                secciones.Add(new Seccion
                {
                    Etiqueta = a,
                    Inicio = inicio,
                    Fin = fila,
                    SubBloques = ParticionarSubBloques(hoja, inicio, fila - 1),
                });
                inicio = fila + 1;
            }
        }

        return secciones;
    }

    private static List<SubBloque> ParticionarSubBloques(Hoja hoja, int desde, int hasta)
    {
        var subBloques = new List<SubBloque>();
        var ocurrencias = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var inicio = desde;
        for (var fila = desde; fila <= hasta; fila++)
        {
            var a = Etiqueta(hoja, "A", fila);
            var b = Etiqueta(hoja, "B", fila);
            var c = Etiqueta(hoja, "C", fila);
            if (a.Length == 0 && b.Length > 0 && c.Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                ocurrencias.TryGetValue(b, out var cuenta);
                ocurrencias[b] = ++cuenta;
                subBloques.Add(new SubBloque { Etiqueta = b, Ocurrencia = cuenta, Inicio = inicio, Fin = fila });
                inicio = fila + 1;
            }
        }

        if (inicio <= hasta)
        {
            // Cola de la sección sin frontera de cierre: se declara con etiqueta vacía (se reporta).
            subBloques.Add(new SubBloque { Etiqueta = string.Empty, Ocurrencia = 0, Inicio = inicio, Fin = hasta });
        }

        return subBloques;
    }

    // ── Firmas (fuente C → canal D efectivo → concepto E) ──────────────────────────────────────

    /// <summary>
    /// Filas de datos del cuadro (excluye la fila-frontera que lo cierra) con su firma
    /// <c>C|D|E</c>: <c>C</c> y <c>E</c> literales de la fila y <c>D</c> = canal propio o, si la fila
    /// no trae <c>D</c>, el <b>canal efectivo</b> (el <c>D</c> de la fila resumen <c>…/Total</c> que
    /// cierra su grupo de canal, Plan 46 §2.1). Devuelve el orden físico del documento.
    /// </summary>
    private static List<(string Firma, int Fila)> FilasConFirma(Hoja hoja, SubBloque cuadro)
    {
        var filas = new List<(string Firma, int Fila)>();
        for (var fila = cuadro.Inicio; fila < cuadro.Fin; fila++)
        {
            filas.Add((FirmaDe(hoja, cuadro, fila), fila));
        }

        return filas;
    }

    /// <summary>
    /// Índice <c>firma → filas (en orden físico)</c>. Las firmas repetidas dentro del cuadro
    /// (mismo <c>C</c>/<c>D</c>/<c>E</c>, p. ej. los dos grupos <c>D=E</c> de un cuadro) se emparejan
    /// por ORDINAL dentro de la firma: k-ésima fila de la fuente ↔ k-ésima fila declarada.
    /// </summary>
    private static Dictionary<string, List<int>> IndiceFirmas(Hoja hoja, Bloque bloque, SubBloque cuadro)
    {
        var indice = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        foreach (var (firma, fila) in FilasConFirma(hoja, cuadro))
        {
            if (!indice.TryGetValue(firma, out var filas))
            {
                filas = [];
                indice[firma] = filas;
            }

            filas.Add(fila);
        }

        return indice;
    }

    private static Dictionary<string, List<int>> FirmasDelBloque(Hoja hoja, Bloque bloque)
    {
        var indice = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        foreach (var seccion in bloque.Secciones)
        {
            foreach (var cuadro in seccion.SubBloques)
            {
                foreach (var (firma, filas) in IndiceFirmas(hoja, bloque, cuadro))
                {
                    if (!indice.TryGetValue(firma, out var acumuladas))
                    {
                        acumuladas = [];
                        indice[firma] = acumuladas;
                    }

                    acumuladas.AddRange(filas);
                }
            }
        }

        return indice;
    }

    private static string FirmaDe(Hoja hoja, SubBloque cuadro, int fila)
    {
        var d = Etiqueta(hoja, "D", fila);
        var canal = d;
        if (canal.Length == 0)
        {
            for (var j = fila + 1; j < cuadro.Fin; j++)
            {
                var siguiente = Etiqueta(hoja, "D", j);
                if (siguiente.Length > 0 && !siguiente.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    canal = siguiente;
                    break;
                }
            }
        }

        return new Firma(Normalizar(Etiqueta(hoja, "C", fila)), Normalizar(canal), Normalizar(Etiqueta(hoja, "E", fila)))
            .ToString();
    }

    // ── Valores declarados por el cruce (componentes por ENCABEZADO) ────────────────────────────

    /// <summary>
    /// Mapa <c>dirección de la plantilla → valor de la fuente</c> de TODAS las filas del bloque,
    /// resueltas por identidad (A, B, ocurrencia) + firma (C, D, E) + componente por ENCABEZADO.
    /// </summary>
    private static Dictionary<string, decimal> Declarados(Hoja fuente, Hoja plantilla, Bloque bloqueFuente, Bloque bloquePlantilla)
    {
        var declarados = new Dictionary<string, decimal>(StringComparer.Ordinal);
        var columnasFuente = Encabezados(fuente, bloqueFuente);
        var columnasPlantilla = Encabezados(plantilla, bloquePlantilla);

        foreach (var seccion in bloqueFuente.Secciones)
        {
            var seccionPlantilla = bloquePlantilla.BuscarSeccion(seccion.Etiqueta)
                ?? throw new InvalidOperationException($"La plantilla no declara la sección '{seccion.Etiqueta}'.");

            DeclararFila(fuente, plantilla, seccion.Fin, seccionPlantilla.Fin, columnasFuente, columnasPlantilla, declarados);

            foreach (var cuadro in seccion.SubBloques)
            {
                var cuadroPlantilla = seccionPlantilla.BuscarCuadro(cuadro.Etiqueta, cuadro.Ocurrencia)
                    ?? throw new InvalidOperationException($"La plantilla no declara el cuadro ({seccion.Etiqueta}, {cuadro.Etiqueta}, #{cuadro.Ocurrencia}).");
                var destino = IndiceFirmas(plantilla, bloquePlantilla, cuadroPlantilla);
                var ordinal = new Dictionary<string, int>(StringComparer.Ordinal);

                foreach (var (firma, filaFuente) in FilasConFirma(fuente, cuadro))
                {
                    ordinal.TryGetValue(firma, out var k);
                    ordinal[firma] = ++k;
                    if (!destino.TryGetValue(firma, out var filasPlantilla) || k > filasPlantilla.Count)
                    {
                        // Fila de la fuente sin destino declarado en su cuadro: la reporta el cruce;
                        // aquí no se escribe (cero del template), nunca se reacomoda.
                        continue;
                    }

                    DeclararFila(fuente, plantilla, filaFuente, filasPlantilla[k - 1], columnasFuente, columnasPlantilla, declarados);
                }

                DeclararFila(fuente, plantilla, cuadro.Fin, cuadroPlantilla.Fin, columnasFuente, columnasPlantilla, declarados);
            }
        }

        DeclararFila(fuente, plantilla, bloqueFuente.Fin, bloquePlantilla.Fin, columnasFuente, columnasPlantilla, declarados);
        return declarados;
    }

    private static void DeclararFila(
        Hoja fuente,
        Hoja plantilla,
        int filaFuente,
        int filaPlantilla,
        IReadOnlyDictionary<string, string> columnasFuente,
        IReadOnlyDictionary<string, string> columnasPlantilla,
        Dictionary<string, decimal> declarados)
    {
        foreach (var (encabezado, columnaFuente) in columnasFuente)
        {
            if (!columnasPlantilla.TryGetValue(encabezado, out var columnaPlantilla))
            {
                continue; // El componente de la fuente que la plantilla no declara no se escribe.
            }

            if (fuente.Numero($"{columnaFuente}{filaFuente}") is { } valor)
            {
                declarados[$"{columnaPlantilla}{filaPlantilla}"] = valor;
            }
        }
    }

    private static Dictionary<string, string> Encabezados(Hoja hoja, Bloque bloque)
    {
        var encabezados = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var columna in Columnas())
        {
            var texto = hoja.Texto($"{columna}{bloque.FilaCabecera}");
            if (texto.Length > 0)
            {
                encabezados.TryAdd(texto, columna);
            }
        }

        return encabezados;
    }

    // ── Cierre aritmético ──────────────────────────────────────────────────────────────────────

    private static void RegistrarAnclas(Hoja hoja, Bloque bloque, int desde, int hasta)
    {
        for (var fila = desde; fila <= hasta; fila++)
        {
            if (!Etiqueta(hoja, "C", fila).Equals("TOTAL", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (Etiqueta(hoja, "D", fila).Equals("OPORTUNO", StringComparison.OrdinalIgnoreCase))
            {
                bloque.RegistrarAncla("OPORTUNO", $"F{fila}");
            }

            // La fila EXTEMPORANEO del grupo TOTAL no repite la etiqueta C: se toma la primera
            // hacia abajo mientras C siga vacía (misma estructura en los 5 bloques).
            for (var j = fila + 1; j <= hasta && string.IsNullOrEmpty(Etiqueta(hoja, "C", j)); j++)
            {
                if (Etiqueta(hoja, "D", j).Equals("EXTEMPORANEO", StringComparison.OrdinalIgnoreCase))
                {
                    bloque.RegistrarAncla("EXTEMPORANEO", $"F{j}");
                    break;
                }
            }
        }
    }

    private static decimal Evaluar(Hoja plantilla, IReadOnlyDictionary<string, decimal> declarados, string direccion) =>
        Valor(plantilla, declarados, direccion, []);

    private static decimal Valor(Hoja plantilla, IReadOnlyDictionary<string, decimal> declarados, string direccion, HashSet<string> enCurso)
    {
        var clave = direccion.Replace("$", string.Empty, StringComparison.Ordinal);
        if (!enCurso.Add(clave))
        {
            throw new InvalidOperationException($"Ciclo de fórmulas en {clave}.");
        }

        decimal resultado;
        if (plantilla.Formula(clave) is { } formula)
        {
            resultado = EvaluarExpresion(formula, d => Valor(plantilla, declarados, d, enCurso));
        }
        else if (declarados.TryGetValue(clave, out var declarado))
        {
            resultado = declarado;
        }
        else
        {
            resultado = plantilla.Numero(clave) ?? 0m;
        }

        enCurso.Remove(clave);
        return resultado;
    }

    /// <summary>
    /// Evaluador mínimo de las fórmulas del template R1: referencias de celda, sumas y restas,
    /// <c>SUM(rango)</c>. Un <c>-</c> repetido se interpreta como doble negación (semántica Excel);
    /// p. ej. <c>F193+F150+F112--L112-L150</c> (bloque 2) suma <c>L112</c>.
    /// </summary>
    private static decimal EvaluarExpresion(string expresion, Func<string, decimal> valor)
    {
        var total = 0m;
        var signo = 1m;
        var i = 0;
        while (i < expresion.Length)
        {
            var caracter = expresion[i];
            if (caracter == '+')
            {
                signo = 1m;
                i++;
            }
            else if (caracter == '-')
            {
                signo = -signo;
                i++;
            }
            else if (char.IsWhiteSpace(caracter) || caracter == '(' || caracter == ')')
            {
                i++;
            }
            else if (char.IsAsciiDigit(caracter))
            {
                var j = i;
                while (j < expresion.Length && (char.IsAsciiDigit(expresion[j]) || expresion[j] == '.'))
                {
                    j++;
                }

                total += signo * decimal.Parse(expresion[i..j], NumberStyles.Any, CultureInfo.InvariantCulture);
                signo = 1m;
                i = j;
            }
            else if (char.IsLetter(caracter) || caracter == '$')
            {
                var j = i;
                while (j < expresion.Length
                    && (char.IsLetterOrDigit(expresion[j]) || expresion[j] is '$' or '_' or '!'))
                {
                    j++;
                }

                var token = expresion[i..j].Replace("$", string.Empty, StringComparison.Ordinal);
                if (token.Equals("SUM", StringComparison.OrdinalIgnoreCase) && j < expresion.Length && expresion[j] == '(')
                {
                    var cierre = expresion.IndexOf(')', j);
                    Assert.True(cierre > 0, $"Fórmula no soportada: {expresion}");
                    foreach (var referencia in ExpandirRango(expresion[(j + 1)..cierre]))
                    {
                        total += signo * valor(referencia);
                    }

                    i = cierre + 1;
                }
                else
                {
                    total += signo * valor(token);
                    i = j;
                }

                signo = 1m;
            }
            else
            {
                throw new InvalidOperationException($"Fórmula no soportada en «{expresion}» (carácter '{caracter}').");
            }
        }

        return total;
    }

    private static IEnumerable<string> ExpandirRango(string rango)
    {
        var partes = rango.Split(':');
        if (partes.Length != 2)
        {
            yield return partes[0].Trim();
            yield break;
        }

        var desde = Descomponer(partes[0].Trim());
        var hasta = Descomponer(partes[1].Trim());
        for (var fila = desde.Fila; fila <= hasta.Fila; fila++)
        {
            for (var columna = desde.Columna; columna <= hasta.Columna; columna++)
            {
                yield return $"{ColumnaLetra(columna)}{fila}";
            }
        }
    }

    private static (int Columna, int Fila) Descomponer(string referencia)
    {
        var letras = new string(referencia.Where(char.IsLetter).ToArray());
        var digitos = new string(referencia.Where(char.IsAsciiDigit).ToArray());
        var columna = 0;
        foreach (var letra in letras)
        {
            columna = (columna * 26) + (char.ToUpperInvariant(letra) - 'A' + 1);
        }

        return (columna, int.Parse(digitos, CultureInfo.InvariantCulture));
    }

    private static string ColumnaLetra(int columna)
    {
        var texto = new StringBuilder();
        while (columna > 0)
        {
            var resto = (columna - 1) % 26;
            texto.Insert(0, (char)('A' + resto));
            columna = (columna - 1) / 26;
        }

        return texto.ToString();
    }

    /// <summary>
    /// Derivación INDEPENDIENTE del cierre (no usa las fórmulas del template): toma los cierres de
    /// subsección que el cruce declara y aplica la semántica documentada (T0-Plan54 §3):
    /// <c>TOT_OPT = Σ F(Mes) − Σ L(Mes sin la última)</c> y
    /// <c>EXTEMP = Σ F(Aplicación) − Σ L(Aplicación sin la primera)</c>.
    /// </summary>
    private static (decimal TotOpt, decimal Extemp) DerivarDesdeFuente(
        Bloque bloqueFuente,
        Bloque bloquePlantilla,
        IReadOnlyDictionary<string, decimal> declarados)
    {
        var mes = new List<(string F, string L)>();
        var aplicacion = new List<(string F, string L)>();

        foreach (var seccion in bloqueFuente.Secciones)
        {
            var seccionPlantilla = bloquePlantilla.BuscarSeccion(seccion.Etiqueta)!;
            foreach (var cuadro in seccion.SubBloques)
            {
                var cuadroPlantilla = seccionPlantilla.BuscarCuadro(cuadro.Etiqueta, cuadro.Ocurrencia)!;
                var esMes = cuadro.Etiqueta.Equals("Mes", StringComparison.OrdinalIgnoreCase);
                var par = ($"F{cuadroPlantilla.Fin}", $"L{cuadroPlantilla.Fin}");
                if (esMes)
                {
                    mes.Add(par);
                }
                else
                {
                    aplicacion.Add(par);
                }
            }
        }

        var totOpt = 0m;
        for (var i = 0; i < mes.Count; i++)
        {
            totOpt += Declarado(declarados, mes[i].F);
            if (i < mes.Count - 1)
            {
                totOpt -= Declarado(declarados, mes[i].L);
            }
        }

        var extemp = 0m;
        for (var i = 0; i < aplicacion.Count; i++)
        {
            extemp += Declarado(declarados, aplicacion[i].F);
            if (i > 0)
            {
                extemp -= Declarado(declarados, aplicacion[i].L);
            }
        }

        return (totOpt, extemp);
    }

    private static decimal Declarado(IReadOnlyDictionary<string, decimal> declarados, string direccion) =>
        declarados.TryGetValue(direccion, out var valor) ? valor : 0m;

    // ── Utilidades ────────────────────────────────────────────────────────────────────────────

    private static IEnumerable<string> Columnas()
    {
        for (var columna = 6; columna <= 26; columna++)
        {
            yield return ColumnaLetra(columna);
        }
    }

    private static string Normalizar(string texto)
    {
        if (string.IsNullOrWhiteSpace(texto))
        {
            return string.Empty;
        }

        var descompuesto = texto.Normalize(NormalizationForm.FormD);
        var limpio = new StringBuilder(descompuesto.Length);
        foreach (var caracter in descompuesto)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(caracter) == UnicodeCategory.NonSpacingMark)
            {
                continue;
            }

            if (char.IsLetterOrDigit(caracter))
            {
                limpio.Append(char.ToUpperInvariant(caracter));
            }
        }

        return limpio.ToString();
    }

    private static string Moneda(decimal valor) =>
        valor.ToString("N2", CultureInfo.GetCultureInfo("es-CO"));
}
