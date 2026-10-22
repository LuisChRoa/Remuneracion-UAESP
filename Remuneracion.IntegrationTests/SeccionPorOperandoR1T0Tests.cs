using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 53 (T0 / H0, read-only BLOQUEANTE) — contrato celda → sección esperada de los operandos
/// empresa de R1. Evidencia: <c>Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md</c>.
///
/// Lee las 20 fuentes R1 reales por BCL (ZIP+XML, sin ExcelDataReader) en ventana explícita, tipa
/// cada fila-empresa con el predicado vigente
/// <see cref="ContratoSeccionPeriodFreeR1.SeccionPorCierre"/> (Plan 46) y contrasta la resolución
/// VIEJA (ocurrencia / <c>|F|</c>, <c>ExcelDataReaderWorkbookLeafInputReader.cs:1585-1637</c>) con
/// la NUEVA (sección declarada por celda). Cero código de producción modificado, cero escrituras.
///
/// Gates: S1.1 matriz 52 grupos/180 filas-empresa + partición/orden; S1.2 tabla celda→sección
/// (56 entradas) + evidencia de fórmula consumidora de la v2; S1.3 casos cardinales; S1.4 inventario
/// R2/R4; S1.5 lista cerrada de Δ exacta; S1.6 tipabilidad/empates/firma; S1.7 oráculo del manual.
/// </summary>
public sealed class SeccionPorOperandoR1T0Tests
{
    private static readonly (string Carpeta, string Codigo)[] Periodos =
    [
        ("Prueba Julio-1", "2026071"),
        ("Prueba Julio-2", "2026072"),
        ("Prueba Agosto-2", "2026082"),
        ("Prueba Octubre-1", "2026101"),
    ];

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    // ── S1.1 — matriz ───────────────────────────────────────────────────────────────────────────

    [Fact]
    public void S1_1_Matriz_52Grupos_180FilasEmpresa_FirmaYSecciones()
    {
        var grupos = 0;
        var filasEmpresa = 0;
        var cardinalidades = new List<int>();

        foreach (var (carpeta, codigo) in Periodos)
        {
            foreach (var ase in new[] { 1, 2, 3, 4, 5 })
            {
                var fuente = Fuente(carpeta, codigo, ase);
                foreach (var empresa in EmpresaFacturacion.Catalogo)
                {
                    if (!WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.ContainsKey((empresa.Id, ase)))
                    {
                        continue;
                    }

                    grupos++;
                    var totales = Bloque(fuente, empresa, subs: false);
                    var subs = Bloque(fuente, empresa, subs: true);
                    filasEmpresa += totales.Count + subs.Count;

                    // Firma A–E de TODA fila-empresa: A/B/E vacíos, C=label, D='Total' (R3 no se dispara).
                    foreach (var indice in totales.Concat(subs).Select(b => b.Indice))
                    {
                        var fila = fuente.Filas[indice];
                        Assert.True(string.IsNullOrEmpty(fila.A), $"{codigo} ASE {ase} f{fila.Fila}: A debe estar vacío.");
                        Assert.True(string.IsNullOrEmpty(fila.B), $"{codigo} ASE {ase} f{fila.Fila}: B debe estar vacío.");
                        Assert.True(string.IsNullOrEmpty(fila.E), $"{codigo} ASE {ase} f{fila.Fila}: E debe estar vacío (concepto 'Total').");
                        Assert.True(fila.D.Equals("Total", StringComparison.OrdinalIgnoreCase), $"{codigo} ASE {ase} f{fila.Fila}: D debe ser 'Total'.");
                    }

                    foreach (var bloque in new[] { totales, subs })
                    {
                        if (bloque.Count > 0)
                        {
                            cardinalidades.Add(bloque.Count);
                        }
                        // Gate R4-riesgo: formas de 4+ secciones = STOP (no hay regla aprobada).
                        var secciones = Secciones(bloque);
                        Assert.True(secciones.Count <= 3, $"{codigo} ASE {ase} empresa {empresa.Id}: {secciones.Count} secciones (4+ = STOP).");

                        // Gate R2-riesgo: empate (2 filas-empresa en la MISMA sección) = STOP.
                        foreach (var seccion in secciones)
                        {
                            var enSeccion = bloque.Count(b => b.Seccion == seccion);
                            Assert.True(enSeccion == 1, $"{codigo} ASE {ase} empresa {empresa.Id}: empate en {Fmt(seccion)} ({enSeccion} filas).");
                        }

                        // Gate de intercalado: una sección no puede reaparecer tras otra.
                        var orden = new List<SeccionFronteraArriba>();
                        foreach (var b in bloque)
                        {
                            if (!orden.Contains(b.Seccion))
                            {
                                orden.Add(b.Seccion);
                            }
                        }

                        for (var i = 0; i < bloque.Count; i++)
                        {
                            var esperada = orden.IndexOf(bloque[i].Seccion);
                            var previa = i == 0 ? 0 : orden.IndexOf(bloque[i - 1].Seccion);
                            Assert.True(esperada >= previa, $"{codigo} ASE {ase} empresa {empresa.Id}: secciones intercaladas.");
                        }
                    }
                }
            }
        }

        Assert.Equal(52, grupos);
        Assert.Equal(180, filasEmpresa);
        Assert.DoesNotContain(cardinalidades, c => c >= 4);
        // Cardinalidades observadas (T0-52 §2): Totales 1/2/3 filas = 14/19/18 · Subs = 28/23/0.
        Assert.Equal(42, cardinalidades.Count(c => c == 1));
        Assert.Equal(42, cardinalidades.Count(c => c == 2));
        Assert.Equal(18, cardinalidades.Count(c => c == 3));
    }

    [Fact]
    public void S1_1_ParticionYOrden_SeccionPorCierre_vs_PorFronteraArriba()
    {
        var etiquetasDistintas = 0;
        var filas = 0;

        foreach (var (carpeta, codigo) in Periodos)
        {
            foreach (var ase in new[] { 1, 2, 3, 4, 5 })
            {
                var fuente = Fuente(carpeta, codigo, ase);
                foreach (var empresa in EmpresaFacturacion.Catalogo)
                {
                    if (!WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.ContainsKey((empresa.Id, ase)))
                    {
                        continue;
                    }

                    foreach (var bloque in new[] { Bloque(fuente, empresa, false), Bloque(fuente, empresa, true) })
                    {
                        for (var i = 0; i < bloque.Count; i++)
                        {
                            filas++;
                            if (bloque[i].Seccion != fuente.Arriba[bloque[i].Indice])
                            {
                                etiquetasDistintas++;
                            }

                            // Partición idéntica: dos filas comparten sección-por-cierre ⇔ comparten
                            // sección-por-arriba (mismo intervalo).
                            for (var j = i + 1; j < bloque.Count; j++)
                            {
                                var mismaCierre = bloque[i].Seccion == bloque[j].Seccion;
                                var mismaArriba = fuente.Arriba[bloque[i].Indice] == fuente.Arriba[bloque[j].Indice];
                                Assert.True(mismaCierre == mismaArriba,
                                    $"{codigo} ASE {ase}: partición distinta entre f{fuente.Filas[bloque[i].Indice].Fila} y f{fuente.Filas[bloque[j].Indice].Fila}.");
                            }

                            // Orden idéntico: la posición de la sección-por-cierre de cada fila dentro de
                            // su propia lista de secciones distintas == la de su sección-por-arriba.
                            Assert.Equal(
                                Posicion(bloque.Select(b => b.Seccion).ToList(), bloque[i].Seccion),
                                Posicion(fuente.Arriba.Where((_, k) => bloque.Any(b => b.Indice == k)).ToList(), fuente.Arriba[bloque[i].Indice]));
                        }
                    }
                }
            }
        }

        Assert.Equal(180, filas);
        // El T0-52 §7 midió etiquetas distintas en 180/180 (la sección se abre por una frontera y se
        // cierra por otra) pero partición y orden idénticos: el desacuerdo de etiqueta no altera ningún
        // operando. Aquí se congela el número medido (partición/orden ya asertados arriba).
        Assert.Equal(180, etiquetasDistintas);
    }

    private static int Posicion<T>(List<T> orden, T valor)
    {
        var vista = new List<T>();
        foreach (var item in orden)
        {
            if (!vista.Contains(item))
            {
                vista.Add(item);
            }
        }

        return vista.IndexOf(valor);
    }

    // ── S1.2 — contrato celda → sección esperada ────────────────────────────────────────────────

    [Fact]
    public void S1_2_ContratoCeldaSeccion_CubreExactamenteElMapa()
    {
        var delMapa = WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa
            .SelectMany(kv => kv.Value.Select(c => (Ase: kv.Key.AseId, Empresa: kv.Key.EmpresaId, c.Celda)))
            .ToHashSet();
        var deLaTabla = TablaContrato.Keys.ToHashSet();

        Assert.Equal(delMapa.Count, deLaTabla.Count);
        Assert.True(delMapa.SetEquals(deLaTabla),
            $"Celdas fuera del contrato T0: {string.Join(", ", delMapa.Except(deLaTabla))} · sobrantes: {string.Join(", ", deLaTabla.Except(delMapa))}");

        // Coherencia rol ↔ operando (la tabla es DATOS, no una regla derivada).
        foreach (var ((ase, empresa, celda), rol) in TablaContrato)
        {
            var operando = WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa[(empresa, ase)]
                .Single(c => c.Celda == celda).Fuente;

            var esperado = operando switch
            {
                WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor
                    or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMenor
                    or WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMenor => RolSeccionT0.Primera,
                WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMedio
                    or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMedio => RolSeccionT0.Intermedia,
                WorkbookLeafCellMapPorEmpresa.FuenteR1.Block5Total => RolSeccionT0.SinSeccion,
                WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMain
                    or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMain
                    or WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMain => ase == 5 && operando != WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMain
                        ? RolSeccionT0.UnicaRecaudo
                        : RolSeccionT0.UltimaOportuna,
                _ => throw new InvalidOperationException($"Operando no previsto: {operando}"),
            };

            Assert.True(rol == esperado, $"{celda} (ASE {ase}, empresa {empresa}): rol {rol} ≠ esperado {esperado} para {operando}.");
        }

        // ASE5 (2,5)/(3,5)/(4,5) Main/Especiales = UnicaRecaudo (§9 #13–#17, D5): (3,5) sin TotalMenor
        // es la prueba de fuego del mapa (F487 = 160.186.886,97 intacto en octubre-1).
        foreach (var celda in new[] { "F483", "L483", "F487", "L487", "F497", "L497" })
        {
            var empresa = celda is "F487" or "L487" ? 3 : celda is "F497" or "L497" ? 4 : 2;
            Assert.Equal(RolSeccionT0.UnicaRecaudo, TablaContrato[(5, empresa, celda)]);
        }
    }

    [Fact]
    public void S1_2_EvidenciaV2_CeldasInertesYSusConsumidores()
    {
        var plantilla = Insumos.PlantillaOficial();
        var celdas = CeldasCrudas(plantilla, "Reporte Componentes R1");

        // Consumidores presentes en los bytes de la v2 (leg de corroboración D2.2, solo donde existe <f>).
        Assert.Equal("F69+F46+F15-L15-L46", celdas["F95"].Formula);
        Assert.Equal("F62", celdas["F86"].Formula);
        Assert.Equal("F353", celdas["F416"].Formula);
        Assert.Equal("F342+F315-L315", celdas["F417"].Formula);
        Assert.Equal("F193+F150+F112--L112-L150", celdas["F264"].Formula);
        Assert.Equal("F330+F353+F299-L299-L330", celdas["F415"].Formula);

        // Hojas inertes (D4): celdas del contrato SIN <f> consumidora en la v2 (slot de valor).
        foreach (var celda in new[] { "F222", "F242", "F29", "F68", "F338", "F237", "F216", "F253", "F474", "F483", "F487", "F497" })
        {
            Assert.True(celdas.ContainsKey(celda), $"{celda} no existe en la v2.");
            Assert.Null(celdas[celda].Formula);
        }

        // Ninguna <f> de la hoja consume las hojas inertes declaradas.
        foreach (var celda in new[] { "F222", "F242", "F29", "F68", "F338" })
        {
            var consumidores = celdas
                .Where(kv => kv.Value.Formula is not null && kv.Key != celda
                             && System.Text.RegularExpressions.Regex.IsMatch(kv.Value.Formula, $@"(^|[^A-Z0-9]){celda}([^0-9]|$)"))
                .Select(kv => $"{kv.Key}={kv.Value.Formula}")
                .ToList();
            Assert.Empty(consumidores);
        }
    }

    // ── S1.3 — casos cardinales ─────────────────────────────────────────────────────────────────

    [Fact]
    public void S1_3_CasosCardinales()
    {
        // Oct-1 ASE5: 3 empresas con UNA fila en la única sección de recaudo → Main(UnicaRecaudo)=valor,
        // Menor(es) ausente (D5). Sin pérdida: ENERBIT no declara TotalMenor.
        foreach (var (empresa, celdaMain, main, celdaMenor) in new[]
        {
            (2, "F483", 7414210944.89m, (string?)"F474"),
            (3, "F487", 160186886.97m, null),
            (4, "F497", 134612117.14m, "F477"),
        })
        {
            var f = Fuente("Prueba Octubre-1", "2026101", 5);
            var e = EmpresaFacturacion.Obtener(empresa);
            var totales = Bloque(f, e, subs: false);
            var subs = Bloque(f, e, subs: true);
            Assert.Single(totales);
            Assert.Single(Secciones(totales));
            Assert.Equal(main, Resolver(f, WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMain, RolSeccionT0.UnicaRecaudo, totales, subs));
            Assert.Equal(main, Valor(f, celdaMain, empresa, totales, subs));
            if (celdaMenor is not null)
            {
                Assert.Equal(0m, Valor(f, celdaMenor, empresa, totales, subs) ?? 0m);
                Assert.Null(Resolver(f, WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor, RolSeccionT0.Primera, totales, subs));
            }
        }

        // Oct-1 ASE3 OCCIDENTE: solo el bloque oportuno (f23 Oportuno#2) → Main = 286.711.907,53, resto 0.
        // La empresa tiene UNA sola sección (la hoja tiene 2: Oportuno#1 solo con ENEL) ⇒ `Primera`
        // no existe para ella (por eso TotalMenor/SubsMenor = 0) y `Intermedia` tampoco.
        {
            var f = Fuente("Prueba Octubre-1", "2026101", 3);
            var empresa = EmpresaFacturacion.Obtener(4);
            var totales = Bloque(f, empresa, subs: false);
            var subs = Bloque(f, empresa, subs: true);
            Assert.Single(Secciones(totales));
            Assert.Single(Secciones(subs));
            Assert.Equal(2, Secciones(Bloque(f, EmpresaFacturacion.Obtener(2), subs: false)).Count); // ENEL sí trae las 2
            Assert.Equal(286711907.53m, Resolver(f, WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMain, RolSeccionT0.UltimaOportuna, totales, subs));
            Assert.Null(Resolver(f, WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor, RolSeccionT0.Primera, totales, subs));
            Assert.Null(Resolver(f, WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMedio, RolSeccionT0.Intermedia, totales, subs));
            Assert.Null(Resolver(f, WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMenor, RolSeccionT0.Primera, totales, subs));
        }

        // Ago-2 ASE2: 3 secciones con magnitud ≠ ordinal → Menor/Medio por POSICIÓN de sección.
        {
            var f = Fuente("Prueba Agosto-2", "2026082", 2);
            foreach (var (empresa, menorEsperado, medioEsperado) in new[] { (2, -1414859.14m, 131543.19m), (4, -28650.84m, 24890.00m) })
            {
                var e = EmpresaFacturacion.Obtener(empresa);
                var totales = Bloque(f, e, subs: false);
                Assert.Equal(3, Secciones(totales).Count);
                Assert.Equal(menorEsperado, Resolver(f, WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor, RolSeccionT0.Primera, totales, []));
                Assert.Equal(medioEsperado, Resolver(f, WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMedio, RolSeccionT0.Intermedia, totales, []));
            }
        }

        // Jul-2 ASE5 ENEL: UNA fila Total no-primera (Oportuno#1) → Main = valor, Menor ausente.
        {
            var f = Fuente("Prueba Julio-2", "2026072", 5);
            var totales = Bloque(f, EmpresaFacturacion.Obtener(2), subs: false);
            Assert.Single(totales);
            Assert.Equal(12823196468.26m, Resolver(f, WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMain, RolSeccionT0.UnicaRecaudo, totales, []));
            Assert.Null(Resolver(f, WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor, RolSeccionT0.Primera, totales, []));
        }

        // 2 secciones ⇒ Intermedia ausente (19 grupos); 3 secciones ⇒ Intermedia = la de en medio.
        foreach (var (carpeta, codigo, ase, empresa, secciones, intermedia) in new[]
        {
            ("Prueba Julio-2", "2026072", 3, 4, 2, (decimal?)null),
            ("Prueba Agosto-2", "2026082", 3, 4, 2, null),
            ("Prueba Julio-1", "2026071", 3, 4, 3, 5698824.85m),
        })
        {
            var f = Fuente(carpeta, codigo, ase);
            var totales = Bloque(f, EmpresaFacturacion.Obtener(empresa), subs: false);
            Assert.Equal(secciones, Secciones(totales).Count);
            Assert.Equal(intermedia, Resolver(f, WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMedio, RolSeccionT0.Intermedia, totales, []));
        }
    }

    private static decimal? Valor(FuenteT0 f, string celda, int empresa, List<(int Indice, SeccionFronteraArriba Seccion)> totales, List<(int, SeccionFronteraArriba)> subs)
    {
        var operando = WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa[(empresa, f.Ase)].Single(c => c.Celda == celda).Fuente;
        var rol = TablaContrato[(f.Ase, empresa, celda)];
        return Resolver(f, operando, rol, totales, subs);
    }

    // ── S1.4 — inventario R2/R4 (D6) ────────────────────────────────────────────────────────────

    [Fact]
    public void S1_4_InventarioR2R4()
    {
        var r2 = new List<int>();
        var r4 = new List<int>();
        var detalleCeros = new List<string>();

        foreach (var (carpeta, codigo) in Periodos)
        {
            for (var ase = 1; ase <= 5; ase++)
            {
                var (filasR2, _) = LeerHoja(RutaReporte(carpeta, ase, "RerpoteDetalleSaldosaFavor"), null);
                var (filasR4, _) = LeerHoja(RutaReporte(carpeta, ase, "Reversi"), null);
                foreach (var empresa in EmpresaFacturacion.Catalogo)
                {
                    // Mismo universo de pares empresa×ASE que el contrato R1 (13 pares × 4 períodos = 52).
                    if (!WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.ContainsKey((empresa.Id, ase)))
                    {
                        continue;
                    }

                    var label = empresa.LabelTemplate;
                    bool Misma(string c) => SinonimosEmpresaR1.SonMismaEmpresa(c, label);

                    var cuentaR2 = filasR2.Count(f =>
                        Misma(Txt(f.Celdas, 1)) && Txt(f.Celdas, 2).Equals("Total", StringComparison.OrdinalIgnoreCase));
                    var cuentaR4 = filasR4.Count(f =>
                        Misma(Txt(f.Celdas, 0)) && Txt(f.Celdas, 1).Equals("Total", StringComparison.OrdinalIgnoreCase));
                    r2.Add(cuentaR2);
                    r4.Add(cuentaR4);
                    if (cuentaR2 == 0)
                    {
                        detalleCeros.Add($"{codigo}/ASE{ase}/{label}");
                    }
                    Assert.True(cuentaR2 is 0 or 2, $"{codigo} ASE {ase} {label}: R2 con {cuentaR2} filas-empresa (D6 exige 0 o 2). Filas leídas={filasR2.Count}.");
                    Assert.True(cuentaR4 <= 1, $"{codigo} ASE {ase} {label}: R4 con {cuentaR4} filas-empresa (D6 exige ≤ 1).");
                }
            }
        }

        Assert.Equal(52, r2.Count);
        var detalle = string.Join(", ", detalleCeros);
        Assert.True(r2.Count(c => c == 0) == 2, $"R2 con 0 filas: {r2.Count(c => c == 0)} → {detalle}");
    }

    // ── S1.5 — lista cerrada de Δ ───────────────────────────────────────────────────────────────

    private static readonly (string Codigo, int Ase, int Empresa, string Celda, decimal Viejo, decimal Nuevo)[] DeltasCerrados =
    [
        ("2026072", 1, 4, "F29", 168024399.69m, 0m),
        ("2026072", 3, 4, "F242", 404748758.01m, 0m),
        ("2026072", 3, 4, "F222", -16458.46m, 0m),
        ("2026072", 5, 2, "F474", 12823196468.26m, 0m),
        ("2026082", 1, 4, "F29", 206290801.38m, 0m),
        ("2026082", 2, 2, "F83", 131543.19m, -1414859.14m),
        ("2026082", 2, 4, "F89", 24890.00m, -28650.84m),
        ("2026082", 3, 4, "F242", 360488591.23m, 0m),
        ("2026082", 3, 4, "F222", -522.80m, 0m),
        ("2026082", 5, 2, "F474", 13260342962.39m, 0m),
        ("2026101", 3, 4, "F216", 286711907.53m, 0m),
        ("2026101", 3, 4, "F242", 145489038.47m, 0m),
        // Filas 13/14 del contrato por celda: D5 manda Menor = 0 en la sección única de oct-1 ASE5
        // (misma regla que §9 #1/#2/#5/#8/#12). No están en la tabla §9 de 17 filas: ver §9 del T0.
        ("2026101", 5, 2, "F474", 7414210944.89m, 0m),
        ("2026101", 5, 4, "F477", 134612117.14m, 0m),
    ];

    [Fact]
    public void S1_5_ListaCerradaDelta_Exacta()
    {
        var observados = new List<(string Codigo, int Ase, int Empresa, string Celda, decimal Viejo, decimal Nuevo)>();

        foreach (var (carpeta, codigo) in Periodos)
        {
            for (var ase = 1; ase <= 5; ase++)
            {
                var fuente = Fuente(carpeta, codigo, ase);
                foreach (var empresa in EmpresaFacturacion.Catalogo)
                {
                    if (!WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.TryGetValue((empresa.Id, ase), out var mapa))
                    {
                        continue;
                    }

                    var totales = Bloque(fuente, empresa, subs: false);
                    var subs = Bloque(fuente, empresa, subs: true);
                    foreach (var (celda, operando, _) in mapa)
                    {
                        var rol = TablaContrato[(ase, empresa.Id, celda)];
                        var nuevo = Resolver(fuente, operando, rol, totales, subs) ?? 0m;
                        var viejo = ViejoValor(fuente, operando, totales, subs) ?? 0m;
                        if (nuevo != viejo)
                        {
                            observados.Add((codigo, ase, empresa.Id, celda, viejo, nuevo));
                        }
                    }
                }
            }
        }

        Assert.Equal(
            DeltasCerrados.OrderBy(d => d.Codigo).ThenBy(d => d.Ase).ThenBy(d => d.Empresa).ThenBy(d => d.Celda).ToList(),
            observados.OrderBy(d => d.Codigo).ThenBy(d => d.Ase).ThenBy(d => d.Empresa).ThenBy(d => d.Celda).ToList());

        // Identidades blindadas (§9 #13–#17 + L487): oct-1 ASE5 Main/Especiales no cambian.
        var oct5 = Fuente("Prueba Octubre-1", "2026101", 5);
        foreach (var (empresa, celda, esperado) in new[]
        {
            (2, "F483", 7414210944.89m), (2, "L483", 10471696.87m),
            (3, "F487", 160186886.97m), (3, "L487", 0m),
            (4, "F497", 134612117.14m), (4, "L497", 1908241.19m),
        })
        {
            var totales = Bloque(oct5, EmpresaFacturacion.Obtener(empresa), subs: false);
            var subs = Bloque(oct5, EmpresaFacturacion.Obtener(empresa), subs: true);
            var valor = Valor(oct5, celda, empresa, totales, subs);
            Assert.Equal(esperado, valor);
        }
    }

    [Fact]
    public void S1_5_VisibleR1_PorCasoArbitrado()
    {
        foreach (var (carpeta, codigo, ase, empresa, viejo, nuevo) in new[]
        {
            ("Prueba Julio-2", "2026072", 5, 2, 24524354313.16m, 11701157844.90m),
            ("Prueba Agosto-2", "2026082", 5, 2, 25008738395.88m, 11748395433.49m),
            ("Prueba Agosto-2", "2026082", 2, 2, 19239204490.16m, 19237658087.83m),
            ("Prueba Agosto-2", "2026082", 2, 4, 2217333974.79m, 2217280433.95m),
            ("Prueba Octubre-1", "2026101", 5, 2, 14647388373.02m, 7233177428.13m),
            ("Prueba Octubre-1", "2026101", 5, 3, 139510580.00m, 139510580.00m),
            ("Prueba Octubre-1", "2026101", 5, 4, 306177388.95m, 171565271.81m),
            ("Prueba Octubre-1", "2026101", 3, 4, 714101241.72m, 427389334.19m),
        })
        {
            var fuente = Fuente(carpeta, codigo, ase);
            var e = EmpresaFacturacion.Obtener(empresa);
            var mapa = WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa[(empresa, ase)];
            var totales = Bloque(fuente, e, subs: false);
            var subs = Bloque(fuente, e, subs: true);

            Assert.Equal(viejo, Visible(ase, empresa, mapa, (_, operando) => ViejoValor(fuente, operando, totales, subs)));
            Assert.Equal(nuevo, Visible(ase, empresa, mapa, (celda, operando) => Resolver(fuente, operando, TablaContrato[(ase, empresa, celda)], totales, subs)));
        }

        // Julio-1 (Q1 con conciliación activa): identidad total, 0 Δ en los 5 ASE.
        foreach (var ase in new[] { 1, 2, 3, 4, 5 })
        {
            var fuente = Fuente("Prueba Julio-1", "2026071", ase);
            foreach (var e in EmpresaFacturacion.Catalogo)
            {
                if (!WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.TryGetValue((e.Id, ase), out var mapa))
                {
                    continue;
                }

                var totales = Bloque(fuente, e, subs: false);
                var subs = Bloque(fuente, e, subs: true);
                Assert.Equal(
                    Visible(ase, e.Id, mapa, (_, operando) => ViejoValor(fuente, operando, totales, subs)),
                    Visible(ase, e.Id, mapa, (celda, operando) => Resolver(fuente, operando, TablaContrato[(ase, e.Id, celda)], totales, subs)));
            }
        }
    }

    [Fact]
    public void S1_5_GateSigma_NoEmpeora()
    {
        // Contraste con el gate Σ (Σ visibles-empresa == visible de bloque) del lector de producción:
        // ningún caso que hoy cierra puede pasar a fallar con el contrato por celda.
        var casos = new List<string>();
        foreach (var (carpeta, codigo) in Periodos)
        {
            for (var ase = 1; ase <= 5; ase++)
            {
                var rutaR1 = RutaReporte(carpeta, ase, "Recaudoporcomponente");
                var rutaR2 = RutaReporte(carpeta, ase, "RerpoteDetalleSaldosaFavor");
                var rutaR4 = RutaReporte(carpeta, ase, "Reversi");
                WorkbookLeafInputs leaf;
                try
                {
                    leaf = new ExcelDataReaderWorkbookLeafInputReader().LeerLeafInputs(
                        Insumos.Ase(ase), Periodo.Parse(codigo), rutaR1, rutaR2, rutaR4);
                }
                catch (Exception)
                {
                    // El bloque no se puede calcular hoy (residuos declarados Q1/oct-1): el gate no se
                    // evalúa → no puede empeorar por este plan.
                    continue;
                }

                var fuente = Fuente(carpeta, codigo, ase);
                decimal sumaVieja = 0m, sumaNueva = 0m;
                foreach (var e in EmpresaFacturacion.Catalogo)
                {
                    if (!WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.TryGetValue((e.Id, ase), out var mapa))
                    {
                        continue;
                    }

                    var totales = Bloque(fuente, e, subs: false);
                    var subs = Bloque(fuente, e, subs: true);
                    sumaVieja += Visible(ase, e.Id, mapa, (_, operando) => ViejoValor(fuente, operando, totales, subs));
                    sumaNueva += Visible(ase, e.Id, mapa, (celda, operando) => Resolver(fuente, operando, TablaContrato[(ase, e.Id, celda)], totales, subs));
                }

                var bloque = leaf.R1.TotalOportunoEsperadoPorAse;
                var okVieja = Math.Abs(sumaVieja - bloque) <= 0.5m;
                var okNueva = Math.Abs(sumaNueva - bloque) <= 0.5m;
                casos.Add($"{codigo} ASE {ase}: viejo={okVieja} nuevo={okNueva}");
                Assert.False(okVieja && !okNueva, $"{codigo} ASE {ase}: el gate Σ que hoy cierra pasa a fallar (viejo={sumaVieja}, nuevo={sumaNueva}, bloque={bloque}).");
            }
        }

        Assert.Contains("2026082 ASE 2: viejo=False nuevo=True", casos);
        Assert.Contains("2026082 ASE 5: viejo=False nuevo=True", casos);
    }

    // ── S1.6 — tipabilidad ──────────────────────────────────────────────────────────────────────

    [Fact]
    public void S1_6_TipabilidadTotal_SinEmpatesNiFormasDe4()
    {
        var tipadas = 0;
        foreach (var (carpeta, codigo) in Periodos)
        {
            for (var ase = 1; ase <= 5; ase++)
            {
                var fuente = Fuente(carpeta, codigo, ase);
                foreach (var empresa in EmpresaFacturacion.Catalogo)
                {
                    if (!WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.ContainsKey((empresa.Id, ase)))
                    {
                        continue;
                    }

                    foreach (var bloque in new[] { Bloque(fuente, empresa, false), Bloque(fuente, empresa, true) })
                    {
                        foreach (var (indice, seccion) in bloque)
                        {
                            tipadas++;
                            Assert.NotNull(fuente.Cierre[indice]);
                            Assert.Equal(seccion, fuente.Cierre[indice]!.Value);
                        }
                    }
                }
            }
        }

        Assert.Equal(180, tipadas);
    }

    // ── S1.7 — oráculo del manual ───────────────────────────────────────────────────────────────

    [Fact]
    public void S1_7_ManualArbitra_4Casos_AlCentavo()
    {
        var jul2 = Path.Combine(Insumos.Raiz(), "Docs", "Prueba Julio-2", "Resultado", "Remuneracion 202607-2 Total Administrativo.xlsx");
        var ago2 = Path.Combine(Insumos.Raiz(), "Docs", "Prueba Agosto-2", "Resultado", "Resultado Manual por el administrativo", "Remuneracion 202608-2 Total_7721.xlsx");
        Assert.True(File.Exists(jul2), $"Falta el manual de julio-2: {jul2}");
        Assert.True(File.Exists(ago2), $"Falta el manual de agosto-2: {ago2}");

        // D{8+ase} de la hoja REMUNERACION_<empresa> del manual = visible R1 de esa empresa en ese ASE.
        // El manual trae los montos como double (ruido de coma flotante): se contrasta con la
        // tolerancia del proyecto (±0,5) — la misma que usa la paridad app-vs-manual.
        ManualCerca(jul2, "REMUNERACION_ENEL", "D13", 11701157844.90m);
        ManualCerca(ago2, "REMUNERACION_ENEL", "D13", 11748395433.49m);
        ManualCerca(ago2, "REMUNERACION_ENEL", "D10", 19237658087.83m);
        ManualCerca(ago2, "REMUNERACION_OCCIDENTE_ Directa", "D10", 2217280433.95m);
    }

    private static void ManualCerca(string ruta, string hoja, string celda, decimal esperado)
    {
        var observado = Manual(ruta, hoja, celda);
        Assert.True(Math.Abs(observado - esperado) <= Insumos.Tolerancia,
            $"{hoja}!{celda} del manual ({Path.GetFileName(ruta)}): observado={observado} esperado={esperado} (Δ={observado - esperado} > ±{Insumos.Tolerancia}).");
    }

    private static decimal Manual(string ruta, string hoja, string celda)
    {
        var celdas = CeldasCrudas(ruta, hoja);
        Assert.True(celdas.TryGetValue(celda, out var e), $"Falta {hoja}!{celda} en {Path.GetFileName(ruta)}.");
        Assert.True(decimal.TryParse(e.Valor, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor),
            $"{hoja}!{celda} no es numérico en {Path.GetFileName(ruta)} (v=[{e.Valor}]).");
        return Math.Round(valor, 2);
    }

    // ── Tabla del contrato (DATOS congelados por el T0; H1.3 la lleva al mapa) ──────────────────

    private enum RolSeccionT0
    {
        UnicaRecaudo,
        Primera,
        UltimaOportuna,
        Intermedia,
        SinSeccion,
    }

    private static readonly Dictionary<(int Ase, int Empresa, string Celda), RolSeccionT0> TablaContrato = new()
    {
        [(1, 2, "F33")] = RolSeccionT0.UltimaOportuna,
        [(1, 2, "F14")] = RolSeccionT0.UltimaOportuna,
        [(1, 2, "L14")] = RolSeccionT0.UltimaOportuna,
        [(1, 4, "F40")] = RolSeccionT0.UltimaOportuna,
        [(1, 4, "F24")] = RolSeccionT0.UltimaOportuna,
        [(1, 4, "L24")] = RolSeccionT0.UltimaOportuna,
        [(1, 4, "F29")] = RolSeccionT0.Primera,
        [(1, 4, "F9")] = RolSeccionT0.Primera,
        [(1, 4, "L9")] = RolSeccionT0.Primera,
        [(2, 1, "F102")] = RolSeccionT0.UltimaOportuna,
        [(2, 1, "F122")] = RolSeccionT0.UltimaOportuna,
        [(2, 1, "F86")] = RolSeccionT0.SinSeccion,
        [(2, 2, "F117")] = RolSeccionT0.UltimaOportuna,
        [(2, 2, "F95")] = RolSeccionT0.UltimaOportuna,
        [(2, 2, "F83")] = RolSeccionT0.Primera,
        [(2, 2, "L95")] = RolSeccionT0.UltimaOportuna,
        [(2, 4, "F129")] = RolSeccionT0.UltimaOportuna,
        [(2, 4, "F112")] = RolSeccionT0.UltimaOportuna,
        [(2, 4, "F89")] = RolSeccionT0.Primera,
        [(2, 4, "L112")] = RolSeccionT0.UltimaOportuna,
        [(3, 2, "F246")] = RolSeccionT0.UltimaOportuna,
        [(3, 2, "F227")] = RolSeccionT0.UltimaOportuna,
        [(3, 2, "F213")] = RolSeccionT0.Primera,
        [(3, 2, "L227")] = RolSeccionT0.UltimaOportuna,
        [(3, 4, "F253")] = RolSeccionT0.UltimaOportuna,
        [(3, 4, "F237")] = RolSeccionT0.UltimaOportuna,
        [(3, 4, "F216")] = RolSeccionT0.Primera,
        [(3, 4, "L237")] = RolSeccionT0.UltimaOportuna,
        [(3, 4, "F242")] = RolSeccionT0.Primera,
        [(3, 4, "F222")] = RolSeccionT0.Intermedia,
        [(4, 1, "F409")] = RolSeccionT0.UltimaOportuna,
        [(4, 1, "F380")] = RolSeccionT0.UltimaOportuna,
        [(4, 1, "L380")] = RolSeccionT0.UltimaOportuna,
        [(4, 2, "F404")] = RolSeccionT0.UltimaOportuna,
        [(4, 2, "F373")] = RolSeccionT0.UltimaOportuna,
        [(4, 2, "F353")] = RolSeccionT0.Primera,
        [(4, 2, "L373")] = RolSeccionT0.UltimaOportuna,
        [(4, 2, "F395")] = RolSeccionT0.Primera,
        [(4, 2, "F362")] = RolSeccionT0.Intermedia,
        [(4, 4, "F416")] = RolSeccionT0.UltimaOportuna,
        [(4, 4, "F390")] = RolSeccionT0.UltimaOportuna,
        [(4, 4, "F356")] = RolSeccionT0.Primera,
        [(4, 4, "L390")] = RolSeccionT0.UltimaOportuna,
        [(4, 4, "F400")] = RolSeccionT0.Primera,
        [(4, 4, "F368")] = RolSeccionT0.Intermedia,
        [(5, 2, "F502")] = RolSeccionT0.UltimaOportuna,
        [(5, 2, "F483")] = RolSeccionT0.UnicaRecaudo,
        [(5, 2, "F474")] = RolSeccionT0.Primera,
        [(5, 2, "L483")] = RolSeccionT0.UnicaRecaudo,
        [(5, 3, "F505")] = RolSeccionT0.UltimaOportuna,
        [(5, 3, "F487")] = RolSeccionT0.UnicaRecaudo,
        [(5, 3, "L487")] = RolSeccionT0.UnicaRecaudo,
        [(5, 4, "F512")] = RolSeccionT0.UltimaOportuna,
        [(5, 4, "F497")] = RolSeccionT0.UnicaRecaudo,
        [(5, 4, "F477")] = RolSeccionT0.Primera,
        [(5, 4, "L497")] = RolSeccionT0.UnicaRecaudo,
    };

    // ── Análisis compartido (BCL ZIP+XML, sin producción) ───────────────────────────────────────

    private sealed record FilaT0(int Indice, int Fila, string A, string B, string C, string D, string E, decimal? Total, decimal? Especiales);

    private sealed record FuenteT0(
        string Periodo,
        string Codigo,
        int Ase,
        List<FilaT0> Filas,
        int IndiceComponente,
        List<SeccionFronteraArriba?> Cierre,
        List<SeccionFronteraArriba> Arriba);

    private static readonly Dictionary<(string, int), FuenteT0> _cache = [];

    private static FuenteT0 Fuente(string carpeta, string codigo, int ase)
    {
        if (_cache.TryGetValue((codigo, ase), out var existente))
        {
            return existente;
        }

        var ruta = RutaReporte(carpeta, ase, "Recaudoporcomponente");
        var (filas, _) = LeerHoja(ruta, "Sheet1");

        var idxIni = filas.FindIndex(f => Normalizar(Txt(f.Celdas, 4)) == "vlrservicio");
        Assert.True(idxIni > 0, $"{codigo} ASE {ase}: falta la fila 'Vlr Servicio'.");

        var idxHeader = idxIni - 1;
        while (idxHeader > 0 && FilaVacia(filas[idxHeader].Celdas))
        {
            idxHeader--;
        }

        var encabezados = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var j = 0; j < filas[idxHeader].Celdas.Count; j++)
        {
            var texto = Normalizar(Txt(filas[idxHeader].Celdas, j));
            if (texto.Length > 0 && !encabezados.ContainsKey(texto))
            {
                encabezados[texto] = j;
            }
        }

        Assert.True(encabezados.ContainsKey("total"), $"{codigo} ASE {ase}: el header no trae la columna 'Total'.");
        Assert.Equal(5, encabezados["total"]); // columna F = índice 5 (contrato del mapa)
        var colEspeciales = encabezados.TryGetValue("servicioespeciales", out var c1)
            ? c1
            : encabezados.TryGetValue("especiales", out var c2) ? c2 : -1;

        var idxTotal = -1;
        for (var i = idxIni; i < filas.Count; i++)
        {
            if (Txt(filas[i].Celdas, 0).Equals("Total", StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrWhiteSpace(Txt(filas[i].Celdas, 1)))
            {
                idxTotal = i;
                break;
            }
        }

        Assert.True(idxTotal > idxIni, $"{codigo} ASE {ase}: falta la fila invariante 'Total'.");

        var secuencia = new List<FilaT0>();
        for (var i = idxIni; i <= idxTotal; i++)
        {
            if (FilaVacia(filas[i].Celdas))
            {
                continue;
            }

            secuencia.Add(new FilaT0(
                secuencia.Count,
                filas[i].Numero,
                Txt(filas[i].Celdas, 0),
                Txt(filas[i].Celdas, 1),
                Txt(filas[i].Celdas, 2),
                Txt(filas[i].Celdas, 3),
                Txt(filas[i].Celdas, 4),
                Numero(filas[i].Celdas, 5),
                colEspeciales >= 0 ? Numero(filas[i].Celdas, colEspeciales) : 0m));
        }

        var espejo = secuencia.Select(f => new FilaEspejoR1
        {
            A = f.A, B = f.B, C = f.C, D = f.D, E = f.E,
            ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase)
            {
                ["Total"] = f.Total,
                ["SERVICIO ESPECIALES"] = f.Especiales,
            },
        }).ToList();

        var cierre = new List<SeccionFronteraArriba?>();
        var arriba = new List<SeccionFronteraArriba>();
        for (var i = 0; i < espejo.Count; i++)
        {
            cierre.Add(ContratoSeccionPeriodFreeR1.SeccionPorCierre(espejo, i));
            arriba.Add(SeccionEspejoR1.PorFronteraArriba(espejo, i));
        }

        var indiceComponente = secuencia.FindIndex(f => f.A.Equals("Componente", StringComparison.OrdinalIgnoreCase));
        Assert.True(indiceComponente > 0, $"{codigo} ASE {ase}: falta la fila 'Componente'.");
        var fuente = new FuenteT0(carpeta, codigo, ase, secuencia, indiceComponente, cierre, arriba);
        _cache[(codigo, ase)] = fuente;
        return fuente;
    }

    private static List<(int Indice, SeccionFronteraArriba Seccion)> Bloque(
        FuenteT0 fuente, EmpresaFacturacion empresa, bool subs)
    {
        var bloque = new List<(int, SeccionFronteraArriba)>();
        for (var i = 0; i < fuente.Filas.Count; i++)
        {
            var enBloque = subs ? i > fuente.IndiceComponente : i < fuente.IndiceComponente;
            if (!enBloque)
            {
                continue;
            }

            var fila = fuente.Filas[i];
            if (!SinonimosEmpresaR1.SonMismaEmpresa(fila.C, empresa.LabelTemplate)
                || !fila.D.Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var seccion = fuente.Cierre[i];
            Assert.True(seccion is not null, $"{fuente.Codigo} ASE {fuente.Ase}: fila-empresa no tipable (f{fila.Fila}).");
            bloque.Add((i, seccion!.Value));
        }

        return bloque;
    }

    private static List<SeccionFronteraArriba> Secciones(List<(int Indice, SeccionFronteraArriba Seccion)> bloque)
    {
        var secciones = new List<SeccionFronteraArriba>();
        foreach (var (_, seccion) in bloque)
        {
            if (!secciones.Contains(seccion))
            {
                secciones.Add(seccion);
            }
        }

        return secciones;
    }

    /// <summary>
    /// Resolución NUEVA (contrato por celda, D1/D4/D5): el bloque es el conjunto de filas-empresa de
    /// la empresa en el bloque (Totales o Subs) y sus secciones en orden de aparición.
    /// <list type="bullet">
    /// <item><c>Primera</c> → la primera sección SOLO si hay ≥ 2 (con una sola no hay "otro bloque");</item>
    /// <item><c>UltimaOportuna</c> / <c>UnicaRecaudo</c> → la última sección del bloque;</item>
    /// <item><c>Intermedia</c> → la sección de en medio SOLO con exactamente 3;</item>
    /// <item>sección declarada ausente → 0 explícito (nunca throw, nunca duplicado).</item>
    /// </list>
    /// </summary>
    private static decimal? Resolver(
        FuenteT0 fuente,
        WorkbookLeafCellMapPorEmpresa.FuenteR1 operando,
        RolSeccionT0 rol,
        List<(int Indice, SeccionFronteraArriba Seccion)> totales,
        List<(int Indice, SeccionFronteraArriba Seccion)> subs)
    {
        if (rol == RolSeccionT0.SinSeccion)
        {
            return null; // Block5Total: canal D='5', sin sección (semántica intacta)
        }

        var bloque = operando is WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMain
            or WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMenor
            ? subs
            : totales;

        var secciones = Secciones(bloque);
        Assert.True(secciones.Count <= 3, $"{fuente.Codigo} ASE {fuente.Ase}: {secciones.Count} secciones (4+ = STOP).");

        SeccionFronteraArriba? objetivo = rol switch
        {
            RolSeccionT0.Primera => secciones.Count >= 2 ? secciones[0] : null,
            RolSeccionT0.UltimaOportuna or RolSeccionT0.UnicaRecaudo => secciones.Count >= 1 ? secciones[^1] : null,
            RolSeccionT0.Intermedia => secciones.Count == 3 ? secciones[1] : null,
            _ => null,
        };

        if (objetivo is null)
        {
            return null;
        }

        var filas = bloque.Where(b => b.Seccion == objetivo).ToList();
        Assert.True(filas.Count == 1, $"{fuente.Codigo} ASE {fuente.Ase}: empate en {Fmt(objetivo)}.");
        var indice = filas[0].Indice;
        return operando is WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMain
            or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMenor
            or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMedio
            ? fuente.Filas[indice].Especiales
            : fuente.Filas[indice].Total;
    }

    /// <summary>Resolución VIEJA (código vigente hasta el Plan 53): orden por <c>|F|</c> descendente.</summary>
    private static decimal? ViejoValor(
        FuenteT0 fuente,
        WorkbookLeafCellMapPorEmpresa.FuenteR1 operando,
        List<(int Indice, SeccionFronteraArriba Seccion)> totales,
        List<(int Indice, SeccionFronteraArriba Seccion)> subs)
    {
        if (operando == WorkbookLeafCellMapPorEmpresa.FuenteR1.Block5Total)
        {
            return null;
        }

        var bloque = operando is WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMain
            or WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMenor
            ? subs
            : totales;
        if (bloque.Count == 0)
        {
            return null;
        }

        var ordenado = bloque.OrderByDescending(f => Math.Abs(fuente.Filas[f.Indice].Total ?? 0m)).ToList();
        var indice = operando switch
        {
            WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMain
                or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMain
                or WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMain => ordenado[0].Indice,
            WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor
                or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMenor
                or WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMenor => ordenado[^1].Indice,
            WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMedio
                or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMedio => ordenado.Count > 1 ? ordenado[1].Indice : -1,
            _ => -1,
        };

        if (indice < 0)
        {
            return null; // el código vigente LANZA aquí (índice fuera de rango)
        }

        return operando is WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMain
            or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMenor
            or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMedio
            ? fuente.Filas[indice].Especiales
            : fuente.Filas[indice].Total;
    }

    /// <summary>Cadena del visible por empresa (<c>CalcularVisibleR1</c>, intacta en el plan).</summary>
    private static decimal Visible(
        int aseId,
        int empresaId,
        (string Celda, WorkbookLeafCellMapPorEmpresa.FuenteR1 Fuente, SeccionEsperadaR1 Seccion)[] mapa,
        Func<string, WorkbookLeafCellMapPorEmpresa.FuenteR1, decimal?> valorDe)
    {
        decimal main = 0m, espMain = 0m, subsMain = 0m, menor = 0m, block5 = 0m;
        foreach (var (celda, fuente, _) in mapa)
        {
            var valor = valorDe(celda, fuente) ?? 0m;
            switch (fuente)
            {
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMain: main = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMain: espMain = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMain: subsMain = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor: menor = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.Block5Total: block5 = valor; break;
            }
        }

        return aseId switch
        {
            1 => main + subsMain - espMain,
            2 when empresaId == 1 => main + subsMain + block5,
            2 => main + subsMain + menor - espMain,
            3 => main + subsMain + menor - espMain,
            4 when empresaId == 1 => main + subsMain - espMain,
            4 => main + subsMain + menor - espMain,
            5 => main + subsMain + menor - espMain,
            _ => 0m,
        };
    }

    private static string Fmt(SeccionFronteraArriba? s) => s is null ? "null" : $"{s.Value.Tipo}#{s.Value.Ordinal}";

    // ── Lectura BCL ─────────────────────────────────────────────────────────────────────────────

    private static string RutaReporte(string carpeta, int ase, string prefijo)
    {
        var dir = Path.Combine(Insumos.Raiz(), "Docs", carpeta, "Insumos", CarpetasAse[ase - 1]);
        Assert.True(Directory.Exists(dir), $"Falta la carpeta explícita de insumos: {dir}");
        var archivo = Directory.EnumerateFiles(dir, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta '{prefijo}*' del ASE {ase} en {dir}");
        return archivo!;
    }

    private sealed record FilaCruda(int Numero, Dictionary<int, object?> Celdas);

    private static readonly Dictionary<string, List<FilaCruda>> _cacheHoja = [];

    private static (List<FilaCruda> Filas, List<string> Compartidas) LeerHoja(string ruta, string? hoja)
    {
        if (_cacheHoja.TryGetValue(ruta + "|" + hoja, out var cache))
        {
            return (cache, []);
        }

        using var zip = ZipFile.OpenRead(ruta);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var compartidas = new List<string>();
        var entradaCompartidas = zip.GetEntry("xl/sharedStrings.xml");
        if (entradaCompartidas is not null)
        {
            using var stream = entradaCompartidas.Open();
            var doc = XDocument.Load(stream);
            foreach (var si in doc.Root!.Elements(ns + "si"))
            {
                compartidas.Add(string.Concat(si.Descendants(ns + "t").Select(t => t.Value)));
            }
        }

        using var wbStream = zip.GetEntry("xl/workbook.xml")!.Open();
        var wb = XDocument.Load(wbStream);
        var hojas = wb.Root!.Element(ns + "sheets")!.Elements(ns + "sheet").ToList();
        var hojaXml = hoja is null
            ? hojas[0]
            : hojas.FirstOrDefault(s => string.Equals((string?)s.Attribute("name"), hoja, StringComparison.OrdinalIgnoreCase));
        Assert.True(hojaXml is not null, $"Falta la hoja '{hoja}' en {Path.GetFileName(ruta)}");
        var relId = (string?)hojaXml!.Attribute(rns + "id");
        using var relStream = zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open();
        var rels = XDocument.Load(relStream);
        var destino = (string?)rels.Root!.Elements().First(r => (string?)r.Attribute("Id") == relId).Attribute("Target");
        var rutaHoja = destino!.StartsWith("xl/", StringComparison.OrdinalIgnoreCase) ? destino : "xl/" + destino.TrimStart('/');

        using var hojaStream = zip.GetEntry(rutaHoja)!.Open();
        var hojaDoc = XDocument.Load(hojaStream);
        var filas = new List<FilaCruda>();
        foreach (var row in hojaDoc.Root!.Descendants(ns + "row"))
        {
            var celdas = new Dictionary<int, object?>();
            foreach (var c in row.Elements(ns + "c"))
            {
                var indice = IndiceColumna((string?)c.Attribute("r") ?? string.Empty);
                var tipo = (string?)c.Attribute("t");
                var v = c.Element(ns + "v")?.Value;
                object? valor = tipo switch
                {
                    "s" when v is not null => compartidas[int.Parse(v, CultureInfo.InvariantCulture)],
                    "inlineStr" => string.Concat(c.Descendants(ns + "t").Select(t => t.Value)),
                    _ => v is null
                        ? null
                        : decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var n) ? n : v,
                };
                if (valor is not null)
                {
                    celdas[indice] = valor;
                }
            }

            filas.Add(new FilaCruda(int.Parse((string?)row.Attribute("r") ?? "0", CultureInfo.InvariantCulture), celdas));
        }

        _cacheHoja[ruta + "|" + hoja] = filas;
        return (filas, compartidas);
    }

    private static readonly Dictionary<(string, string), Dictionary<string, (string? Formula, string? Valor)>> _cacheCeldas = [];

    private static Dictionary<string, (string? Formula, string? Valor)> CeldasCrudas(string ruta, string hoja)
    {
        if (_cacheCeldas.TryGetValue((ruta, hoja), out var cache))
        {
            return cache;
        }

        using var zip = ZipFile.OpenRead(ruta);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var compartidas = new List<string>();
        var entradaCompartidas = zip.GetEntry("xl/sharedStrings.xml");
        if (entradaCompartidas is not null)
        {
            using var stream = entradaCompartidas.Open();
            var doc = XDocument.Load(stream);
            foreach (var si in doc.Root!.Elements(ns + "si"))
            {
                compartidas.Add(string.Concat(si.Descendants(ns + "t").Select(t => t.Value)));
            }
        }

        using var wbStream = zip.GetEntry("xl/workbook.xml")!.Open();
        var wb = XDocument.Load(wbStream);
        var hojaXml = wb.Root!.Element(ns + "sheets")!.Elements(ns + "sheet")
            .FirstOrDefault(s => string.Equals((string?)s.Attribute("name"), hoja, StringComparison.OrdinalIgnoreCase));
        Assert.True(hojaXml is not null, $"Falta la hoja '{hoja}' en {Path.GetFileName(ruta)}");
        var relId = (string?)hojaXml!.Attribute(rns + "id");
        using var relStream = zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open();
        var rels = XDocument.Load(relStream);
        var destino = (string?)rels.Root!.Elements().First(r => (string?)r.Attribute("Id") == relId).Attribute("Target");
        var rutaHoja = destino!.StartsWith("xl/", StringComparison.OrdinalIgnoreCase) ? destino : "xl/" + destino.TrimStart('/');

        using var hojaStream = zip.GetEntry(rutaHoja)!.Open();
        var hojaDoc = XDocument.Load(hojaStream);
        var mapa = new Dictionary<string, (string?, string?)>(StringComparer.OrdinalIgnoreCase);
        foreach (var c in hojaDoc.Root!.Descendants(ns + "c"))
        {
            var referencia = (string?)c.Attribute("r");
            if (referencia is null)
            {
                continue;
            }

            var formula = c.Element(ns + "f")?.Value;
            var v = c.Element(ns + "v")?.Value;
            var t = (string?)c.Attribute("t");
            if (t == "s" && v is not null)
            {
                v = compartidas[int.Parse(v, CultureInfo.InvariantCulture)];
            }

            mapa[referencia] = (formula, v);
        }

        _cacheCeldas[(ruta, hoja)] = mapa;
        return mapa;
    }

    private static int IndiceColumna(string referencia)
    {
        var n = 0;
        foreach (var ch in referencia)
        {
            if (ch is < 'A' or > 'Z')
            {
                break;
            }

            n = (n * 26) + (ch - 'A' + 1);
        }

        return n - 1;
    }

    private static string Txt(Dictionary<int, object?> fila, int columna) => fila.GetValueOrDefault(columna) switch
    {
        null => string.Empty,
        string s => s.Trim(),
        decimal d => d.ToString(CultureInfo.InvariantCulture),
        var o => o.ToString()?.Trim() ?? string.Empty,
    };

    private static decimal? Numero(Dictionary<int, object?> fila, int columna) => fila.GetValueOrDefault(columna) switch
    {
        decimal d => d,
        string s when decimal.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var n) => n,
        _ => null,
    };

    private static bool FilaVacia(Dictionary<int, object?> fila) =>
        fila.Values.All(v => v is null || (v is string s && string.IsNullOrWhiteSpace(s)));

    private static string Normalizar(string texto)
    {
        var descompuesto = texto.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder(descompuesto.Length);
        foreach (var ch in descompuesto)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark && !char.IsWhiteSpace(ch))
            {
                sb.Append(char.ToLowerInvariant(ch));
            }
        }

        return sb.ToString();
    }
}
