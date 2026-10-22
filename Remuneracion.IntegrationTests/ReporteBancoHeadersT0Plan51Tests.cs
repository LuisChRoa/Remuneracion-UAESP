using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 51 (H0, S1.1-S1.7): T0 READ-ONLY del <c>ReportePagosxBanco</c>, con CERO código de
/// producción. Congela por BYTES (BCL <see cref="ZipArchive"/> + <see cref="XDocument"/>; sin
/// Excel/COM, sin OpenXml SDK, sin ExcelDataReader) los veredictos del §0.4 del plan:
/// <list type="bullet">
/// <item>S1.1 — matriz de headers por (período × ASE) en julio-1, julio-2, agosto-2 y octubre-1:
/// fila real del <c>Resumen Recaudo Aplicado Por Servicio</c>, fila de headers, texto crudo +
/// índice físico de columna (evidencia, no contrato) y presentes/ausentes contra
/// <see cref="WorkbookLeafCellMapReporteBanco.EditablesPorAse"/>;</item>
/// <item>S1.2 — <c>Total</c> se distingue de una fuente (exclusión explícita; ninguna fuente real
/// normaliza a <c>total</c>);</item>
/// <item>S1.3 — conceptos 1/2/3/7 y fila <c>Total</c> se siguen leyendo por etiqueta/prefijo
/// (misma forma en las 4 muestras; la fila 7 no viene en ninguna → 0 demostrable);</item>
/// <item>S1.4 — Σ conceptos == <c>TotalFuente</c> por empresa PRESENTE (±0.5) en las 40
/// combinaciones (20 muestras × empresas presentes);</item>
/// <item>S1.5 — veredicto octubre-1: ASE2 (y ASE4, medido) traen 2 fuentes porque
/// <c>NUEVO ESQUEMA</c> no viene; ASE5 trae 3; paridad global esperado↔observado;</item>
/// <item>S1.6 — inventario de headers no reconocidos: VACÍO en las 20 muestras (cero STOP).</item>
/// </list>
///
/// La OBSERVACIÓN es una derivación independiente (espejo local de las reglas del plan §2.1); la
/// EXPECTATIVA por ASE se lee de <c>EditablesPorAse</c> (la plantilla manda, D1-A). El fixture no
/// toca <c>ColumnasFuentePorAse</c> (mapa posicional que el plan retira). Puerta H0 → H1: este
/// fixture verde.
/// </summary>
public sealed class ReporteBancoHeadersT0Plan51Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>Prefijo normalizado del Resumen (espejo de <c>PrefijoEtiquetaResumen</c>).</summary>
    private const string PrefijoResumen = "resumenrecaudoaplicadoporservicio";

    /// <summary>Prefijos normalizados de los conceptos (espejo de <c>WorkbookLeafCellMapReporteBanco.Conceptos</c>).</summary>
    private static readonly (int Indice, string Prefijo)[] Conceptos =
    [
        (0, "1-aplicadosafacturacion"),
        (1, "2-saldosafavor"),
        (2, "3-aplicadosafinanciaciones"),
        (3, "7-aplicados"),
    ];

    // ── S1.1/S1.2/S1.3/S1.6 — matriz congelada por bytes ────────────────────────────────────────

    /// <summary>
    /// Líneas congeladas:
    /// <c>período|ASE|filaResumen|filaHeaders|headers|presentes|ausentes|conceptos</c>.
    /// <c>headers</c> = texto crudo + índice físico de columna entre corchetes (evidencia).
    /// <c>conceptos</c> = índices 1-based de <see cref="Conceptos"/> observados como fila del
    /// Resumen (la fila 7 no viene en ninguna muestra: 0 demostrable).
    /// </summary>
    public static TheoryData<string, int, int, int, string, string, string, string> MatrizHeadersData()
    {
        const string P1 = "Prueba Julio-1";
        const string P2 = "Prueba Julio-2";
        const string P3 = "Prueba Agosto-2";
        const string P4 = "Prueba Octubre-1";
        string[] lineas =
        [
            // Julio-1 (Q1)
            $"{P1}|1|414|415|[1]=ENEL [2]=OCCIDENTE [3]=Total|ENEL,OCCIDENTE||1,2,3",
            $"{P1}|2|440|441|[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total|ENEL,NUEVO ESQUEMA,OCCIDENTE||1,2",
            $"{P1}|3|406|407|[1]=ENEL [2]=OCCIDENTE [3]=Total|ENEL,OCCIDENTE||1,2,3",
            $"{P1}|4|433|434|[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total|ENEL,NUEVO ESQUEMA,OCCIDENTE||1,2,3",
            $"{P1}|5|418|419|[1]=ENEL [2]=ENERBIT [3]=OCCIDENTE [4]=Total|ENEL,ENERBIT,OCCIDENTE||1,2,3",
            // Julio-2 (Q2)
            $"{P2}|1|451|452|[1]=ENEL [2]=OCCIDENTE [3]=Total|ENEL,OCCIDENTE||1,2,3",
            $"{P2}|2|468|469|[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total|ENEL,NUEVO ESQUEMA,OCCIDENTE||1,2",
            $"{P2}|3|434|435|[1]=ENEL [2]=OCCIDENTE [3]=Total|ENEL,OCCIDENTE||1,2,3",
            $"{P2}|4|456|457|[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total|ENEL,NUEVO ESQUEMA,OCCIDENTE||1,2,3",
            $"{P2}|5|448|449|[1]=ENEL [2]=ENERBIT [3]=OCCIDENTE [4]=Total|ENEL,ENERBIT,OCCIDENTE||1,2,3",
            // Agosto-2 (Q2)
            $"{P3}|1|450|451|[1]=ENEL [2]=OCCIDENTE [3]=Total|ENEL,OCCIDENTE||1,2,3",
            $"{P3}|2|473|474|[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total|ENEL,NUEVO ESQUEMA,OCCIDENTE||1,2",
            $"{P3}|3|427|428|[1]=ENEL [2]=OCCIDENTE [3]=Total|ENEL,OCCIDENTE||1,2,3",
            $"{P3}|4|445|446|[1]=ENEL [2]=NUEVO ESQUEMA [3]=OCCIDENTE [4]=Total|ENEL,NUEVO ESQUEMA,OCCIDENTE||1,2,3",
            $"{P3}|5|434|435|[1]=ENEL [2]=ENERBIT [3]=OCCIDENTE [4]=Total|ENEL,ENERBIT,OCCIDENTE||1,2,3",
            // Octubre-1 (Q1) — la corrida roja del plan: ASE2 y ASE4 SIN NUEVO ESQUEMA
            $"{P4}|1|404|405|[1]=ENEL [2]=OCCIDENTE [3]=Total|ENEL,OCCIDENTE||1,2,3",
            $"{P4}|2|391|392|[1]=ENEL [2]=OCCIDENTE [3]=Total|ENEL,OCCIDENTE|NUEVO ESQUEMA|1,2,3",
            $"{P4}|3|378|379|[1]=ENEL [2]=OCCIDENTE [3]=Total|ENEL,OCCIDENTE||1,2,3",
            $"{P4}|4|394|395|[1]=ENEL [2]=OCCIDENTE [3]=Total|ENEL,OCCIDENTE|NUEVO ESQUEMA|1,2,3",
            $"{P4}|5|409|410|[1]=ENEL [2]=ENERBIT [3]=OCCIDENTE [4]=Total|ENEL,ENERBIT,OCCIDENTE||1,2,3",
        ];

        var data = new TheoryData<string, int, int, int, string, string, string, string>();
        foreach (var linea in lineas)
        {
            var partes = linea.Split('|');
            data.Add(
                partes[0],
                int.Parse(partes[1], CultureInfo.InvariantCulture),
                int.Parse(partes[2], CultureInfo.InvariantCulture),
                int.Parse(partes[3], CultureInfo.InvariantCulture),
                partes[4],
                partes[5],
                partes[6],
                partes[7]);
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(MatrizHeadersData))]
    public void T0P51_MatrizHeaders_PorPeriodoAse(
        string periodo, int ase, int filaResumen, int filaHeaders, string headersEsperados,
        string presentesEsperados, string ausentesEsperados, string conceptosEsperados)
    {
        var filas = LeerFilas(RutaFuente(periodo, ase));
        var indiceResumen = BuscarResumen(filas);
        Assert.True(indiceResumen >= 0, $"{periodo} ASE {ase}: falta la etiqueta del Resumen.");
        Assert.Equal(filaResumen, filas[indiceResumen].Fila);

        var header = filas[indiceResumen + 1];
        Assert.Equal(filaHeaders, header.Fila);
        Assert.Equal(headersEsperados, HeadersCrudos(header));

        var clasificacion = ClasificarFuentes(header, ase);
        Assert.Equal(presentesEsperados, string.Join(",", clasificacion.Presentes.Select(p => p.Label)));
        Assert.Equal(ausentesEsperados, string.Join(",", clasificacion.Ausentes));

        var conceptos = ConceptosObservados(filas, indiceResumen);
        Assert.Equal(conceptosEsperados, string.Join(",", conceptos));
    }

    // ── S1.2 — Total nunca es fuente ────────────────────────────────────────────────────────────

    [Theory]
    [MemberData(nameof(MatrizHeadersData))]
    public void T0P51_Total_SeDistingueDeUnaFuente_EnLas20Muestras(
        string periodo, int ase, int filaResumen, int filaHeaders, string headersEsperados,
        string presentesEsperados, string ausentesEsperados, string conceptosEsperados)
    {
        _ = filaResumen;
        _ = filaHeaders;
        _ = ausentesEsperados;
        _ = conceptosEsperados;

        var filas = LeerFilas(RutaFuente(periodo, ase));
        var header = filas[BuscarResumen(filas) + 1];

        // Exactamente UNA columna "Total" en el header; ninguna fuente real normaliza a "total".
        var columnasTotal = header.Celdas
            .Where(kv => NormalizarEtiqueta(Texto(kv.Value)) == "total")
            .Select(kv => kv.Key)
            .ToList();
        Assert.Single(columnasTotal);
        Assert.Contains($"[{columnasTotal[0]}]=Total", headersEsperados, StringComparison.Ordinal);

        var fuentesDeclaradas = WorkbookLeafCellMapReporteBanco.EditablesPorAse[ase].Keys
            .Select(NormalizarEtiqueta)
            .ToList();
        Assert.DoesNotContain("total", fuentesDeclaradas);

        // El "Total" NO entra en la clasificación de fuentes (exclusión explícita antes del
        // vocabulario): las presentes observadas son exactamente las del contrato congelado.
        Assert.Equal(
            presentesEsperados,
            string.Join(",", ClasificarFuentes(header, ase).Presentes.Select(p => p.Label)));
    }

    // ── S1.3 — conceptos + fila Total intactos ─────────────────────────────────────────────────

    [Theory]
    [MemberData(nameof(MatrizHeadersData))]
    public void T0P51_ConceptosYFilaTotal_SeSiguenLeyendoIgual(
        string periodo, int ase, int filaResumen, int filaHeaders, string headersEsperados,
        string presentesEsperados, string ausentesEsperados, string conceptosEsperados)
    {
        _ = filaResumen;
        _ = filaHeaders;
        _ = headersEsperados;
        _ = presentesEsperados;
        _ = ausentesEsperados;

        var filas = LeerFilas(RutaFuente(periodo, ase));
        var indiceResumen = BuscarResumen(filas);

        // Los conceptos se localizan por prefijo normalizado sobre la columna A y la fila "Total"
        // por normalizado == "total" (misma regla del reader, T0-0.4).
        var observados = ConceptosObservados(filas, indiceResumen);
        Assert.Equal(conceptosEsperados, string.Join(",", observados));

        // La fila Total existe y aporta valor a TODAS las columnas de fuente presentes.
        var filaTotal = FilaTotal(filas, indiceResumen);
        Assert.NotNull(filaTotal);
        foreach (var (_, columna) in ColumnasFuente(header: filas[indiceResumen + 1], ase))
        {
            Assert.NotEqual(0m, Numero(filaTotal!, columna));
        }
    }

    [Fact]
    public void T0P51_Concepto7_NoVieneEnNingunaDeLas20Muestras_EsCeroDemostrable()
    {
        string[] periodos = ["Prueba Julio-1", "Prueba Julio-2", "Prueba Agosto-2", "Prueba Octubre-1"];
        foreach (var periodo in periodos)
        {
            for (var ase = 1; ase <= 5; ase++)
            {
                var filas = LeerFilas(RutaFuente(periodo, ase));
                var observados = ConceptosObservados(filas, BuscarResumen(filas));
                Assert.DoesNotContain(4, observados);
            }
        }
    }

    // ── S1.4 — Σ conceptos == TotalFuente por empresa presente ─────────────────────────────────

    [Theory]
    [MemberData(nameof(MatrizHeadersData))]
    public void T0P51_SigmaConceptosVsTotalFuente_PorEmpresaPresente(
        string periodo, int ase, int filaResumen, int filaHeaders, string headersEsperados,
        string presentesEsperados, string ausentesEsperados, string conceptosEsperados)
    {
        _ = filaResumen;
        _ = filaHeaders;
        _ = headersEsperados;
        _ = ausentesEsperados;
        _ = conceptosEsperados;

        var filas = LeerFilas(RutaFuente(periodo, ase));
        var indiceResumen = BuscarResumen(filas);
        var header = filas[indiceResumen + 1];
        var filaTotal = FilaTotal(filas, indiceResumen);
        Assert.NotNull(filaTotal);

        var presentes = ClasificarFuentes(header, ase).Presentes;
        Assert.Equal(
            presentesEsperados,
            string.Join(",", presentes.Select(p => p.Label)));

        foreach (var (label, columna) in presentes)
        {
            var suma = 0m;
            for (var i = indiceResumen + 2; i < filas.Count; i++)
            {
                var fila = filas[i];
                var etiqueta = NormalizarEtiqueta(Texto(fila.Celdas.GetValueOrDefault(0)));
                if (etiqueta == "total")
                {
                    break;
                }

                if (etiqueta.Length == 0)
                {
                    continue;
                }

                if (Conceptos.Any(c => etiqueta.StartsWith(c.Prefijo, StringComparison.Ordinal)))
                {
                    suma += Numero(fila, columna);
                }
            }

            var totalFuente = Numero(filaTotal!, columna);
            Assert.True(
                Math.Abs(suma - totalFuente) <= 0.5m,
                $"{periodo} ASE {ase} · {label}: Σ conceptos ({suma}) != Total fuente ({totalFuente}).");
        }
    }

    // ── S1.5 — veredicto octubre-1 + paridad global esperado↔observado ─────────────────────────

    [Theory]
    [MemberData(nameof(MatrizHeadersData))]
    public void T0P51_ParidadEsperadoVsObservado_SubconjuntoValido(
        string periodo, int ase, int filaResumen, int filaHeaders, string headersEsperados,
        string presentesEsperados, string ausentesEsperados, string conceptosEsperados)
    {
        _ = filaResumen;
        _ = filaHeaders;
        _ = headersEsperados;
        _ = conceptosEsperados;

        var filas = LeerFilas(RutaFuente(periodo, ase));
        var clasificacion = ClasificarFuentes(filas[BuscarResumen(filas) + 1], ase);
        var presentes = clasificacion.Presentes;
        var ausentes = clasificacion.Ausentes;

        var esperadas = WorkbookLeafCellMapReporteBanco.EditablesPorAse[ase].Keys
            .OrderBy(k => k, StringComparer.Ordinal)
            .ToList();
        var observadas = presentes.Select(p => p.Label).OrderBy(k => k, StringComparer.Ordinal).ToList();
        Assert.Equal(presentesEsperados, string.Join(",", observadas));

        // Observado ⊆ esperado (una fuente reconocida sin celdas destino es desync → STOP) y
        // ausente == esperado − observado (la ausencia es del label completo, nunca inventada).
        Assert.All(observadas, o => Assert.Contains(o, esperadas));
        Assert.Equal(esperadas.Count - observadas.Count, ausentes.Count);
        Assert.Equal(
            string.Join(",", esperadas.Except(observadas, StringComparer.Ordinal).OrderBy(k => k, StringComparer.Ordinal)),
            ausentesEsperados);
    }

    [Fact]
    public void T0P51_Octubre1_Ase2_TieneDosFuentes_NuevoEsquemaAusente()
    {
        // El rojo del plan (§0.1): el Resumen de octubre-1 ASE2 trae {ENEL, OCCIDENTE, Total};
        // "NUEVO ESQUEMA" (la tercera fuente esperada por la plantilla) NO viene en el banco.
        var filas = LeerFilas(RutaFuente("Prueba Octubre-1", 2));
        var indiceResumen = BuscarResumen(filas);
        Assert.Equal(391, filas[indiceResumen].Fila);

        var clasificacion = ClasificarFuentes(filas[indiceResumen + 1], 2);
        Assert.Equal("ENEL,OCCIDENTE", string.Join(",", clasificacion.Presentes.Select(p => p.Label)));
        Assert.Equal("NUEVO ESQUEMA", string.Join(",", clasificacion.Ausentes));

        // Evidencia de valores (fila Total, R396): ENEL 10.309.260.573 / OCCIDENTE 971.950.730.
        var filaTotal = FilaTotal(filas, indiceResumen);
        Assert.NotNull(filaTotal);
        Assert.Equal(396, filaTotal!.Fila);
        Assert.Equal(10309260573m, Numero(filaTotal, 1));
        Assert.Equal(971950730m, Numero(filaTotal, 2));
    }

    [Fact]
    public void T0P51_Octubre1_Ase4_TambienOmiteNuevoEsquema_VeredictoMedido()
    {
        // Desviación DOCUMENTADA del plan (§3-S1.5 esperaba ASE4 con 3 fuentes en octubre-1):
        // la muestra real de octubre-1 ASE4 trae {ENEL, OCCIDENTE} (sin NUEVO ESQUEMA). Es la
        // MISMA familia de ausencia legítima (esperada no observada) y se absorbe con la regla
        // genérica ratificada (sin ramas): no hay header desconocido ni desync.
        var filas = LeerFilas(RutaFuente("Prueba Octubre-1", 4));
        var indiceResumen = BuscarResumen(filas);
        var clasificacion = ClasificarFuentes(filas[indiceResumen + 1], 4);

        Assert.Equal("ENEL,OCCIDENTE", string.Join(",", clasificacion.Presentes.Select(p => p.Label)));
        Assert.Equal("NUEVO ESQUEMA", string.Join(",", clasificacion.Ausentes));
        Assert.Equal(394, filas[indiceResumen].Fila);
    }

    [Fact]
    public void T0P51_Octubre1_Ase5_TresFuentes_TresEsperadas()
    {
        var filas = LeerFilas(RutaFuente("Prueba Octubre-1", 5));
        var clasificacion = ClasificarFuentes(filas[BuscarResumen(filas) + 1], 5);

        Assert.Equal("ENEL,ENERBIT,OCCIDENTE", string.Join(",", clasificacion.Presentes.Select(p => p.Label)));
        Assert.Empty(clasificacion.Ausentes);
    }

    // ── S1.6 — inventario de headers no reconocidos ────────────────────────────────────────────

    [Fact]
    public void T0P51_HeadersNoReconocidos_VacioEnLas20Muestras()
    {
        string[] periodos = ["Prueba Julio-1", "Prueba Julio-2", "Prueba Agosto-2", "Prueba Octubre-1"];
        foreach (var periodo in periodos)
        {
            for (var ase = 1; ase <= 5; ase++)
            {
                var filas = LeerFilas(RutaFuente(periodo, ase));
                var header = filas[BuscarResumen(filas) + 1];
                var desconocidas = ClasificarFuentes(header, ase).Desconocidas;
                Assert.True(
                    desconocidas.Count == 0,
                    $"{periodo} ASE {ase}: headers no reconocidos [{string.Join(",", desconocidas)}] → STOP del T0.");
            }
        }
    }

    // ── Helpers de dominio (espejo local, independiente de producción) ─────────────────────────

    private static readonly string[] VocabularioGlobal =
    [
        "RECIPROCIDAD", "ENEL", "ENERBIT", "OCCIDENTE", "CIUDAD LIMPIA-ACUEDUCTO",
    ];

    /// <summary>Label vigente por ASE (espejo de <c>ObtenerLabelFuente</c>: empresa 1 en ASE2/4).</summary>
    private static string LabelVigente(string label, int ase) =>
        label == "RECIPROCIDAD" && ase is 2 or 4 ? "NUEVO ESQUEMA"
        : label == "RECIPROCIDAD" ? "RECIPROCIDAD"
        : label;

    private static (List<(string Label, int Columna)> Presentes, List<string> Ausentes, List<string> Desconocidas) ClasificarFuentes(
        FilaExcel header, int ase)
    {
        var esperadas = WorkbookLeafCellMapReporteBanco.EditablesPorAse[ase].Keys.ToList();
        var presentes = new List<(string Label, int Columna)>();
        var desconocidas = new List<string>();

        foreach (var (columna, valor) in header.Celdas.OrderBy(kv => kv.Key))
        {
            if (columna == 0)
            {
                continue; // la columna 0 es la etiqueta de concepto, nunca una fuente.
            }

            var texto = Texto(valor);
            if (texto.Length == 0)
            {
                continue; // celda vacía: no es evidencia (regla §2.3.2).
            }

            if (NormalizarEtiqueta(texto) == "total")
            {
                continue; // "Total" se excluye ANTES del vocabulario (regla §2.3.1).
            }

            var resuelta = ResolverFuente(texto, ase, esperadas);
            if (resuelta is null)
            {
                desconocidas.Add(texto);
                continue;
            }

            presentes.Add((resuelta, columna));
        }

        var observadas = presentes.Select(p => p.Label).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var ausentes = esperadas
            .Where(e => !observadas.Contains(e))
            .OrderBy(e => e, StringComparer.Ordinal)
            .ToList();

        return (presentes, ausentes, desconocidas);
    }

    /// <summary>
    /// Resuelve un header contra el vocabulario global (catálogo + rótulo vigente + sinonimia
    /// RECIPROCIDAD↔NUEVO ESQUEMA) y lo devuelve como la CLAVE de la plantilla para ese ASE.
    /// </summary>
    private static string? ResolverFuente(string header, int ase, IReadOnlyList<string> esperadas)
    {
        var vocabulario = VocabularioGlobal
            .Select(l => LabelVigente(l, ase))
            .Append("NUEVO ESQUEMA")
            .Append("RECIPROCIDAD")
            .Distinct(StringComparer.OrdinalIgnoreCase);

        var normalizado = NormalizarEtiqueta(header);
        if (!vocabulario.Any(v => NormalizarEtiqueta(v) == normalizado))
        {
            return null;
        }

        // La clave se toma de la plantilla (D1-A): la sinonimia solo reconcilia, no crea fuentes.
        var clave = esperadas.FirstOrDefault(e => NormalizarEtiqueta(e) == normalizado)
            ?? esperadas.FirstOrDefault(e =>
                (normalizado == NormalizarEtiqueta("RECIPROCIDAD") && NormalizarEtiqueta(e) == NormalizarEtiqueta("NUEVO ESQUEMA"))
                || (normalizado == NormalizarEtiqueta("NUEVO ESQUEMA") && NormalizarEtiqueta(e) == NormalizarEtiqueta("RECIPROCIDAD")));

        return clave;
    }

    private static List<(string Label, int Columna)> ColumnasFuente(FilaExcel header, int ase) =>
        ClasificarFuentes(header, ase).Presentes;

    private static string HeadersCrudos(FilaExcel header) =>
        string.Join(
            " ",
            header.Celdas
                .Where(kv => kv.Key > 0 && Texto(kv.Value).Length > 0)
                .OrderBy(kv => kv.Key)
                .Select(kv => $"[{kv.Key}]={Texto(kv.Value)}"));

    private static List<int> ConceptosObservados(List<FilaExcel> filas, int indiceResumen)
    {
        var observados = new List<int>();
        for (var i = indiceResumen + 2; i < filas.Count; i++)
        {
            var etiqueta = NormalizarEtiqueta(Texto(filas[i].Celdas.GetValueOrDefault(0)));
            if (etiqueta.Length == 0)
            {
                continue;
            }

            if (etiqueta == "total")
            {
                break;
            }

            foreach (var (indice, prefijo) in Conceptos)
            {
                if (etiqueta.StartsWith(prefijo, StringComparison.Ordinal))
                {
                    observados.Add(indice + 1);
                    break;
                }
            }
        }

        return observados;
    }

    private static FilaExcel? FilaTotal(List<FilaExcel> filas, int indiceResumen)
    {
        for (var i = indiceResumen + 2; i < filas.Count; i++)
        {
            if (NormalizarEtiqueta(Texto(filas[i].Celdas.GetValueOrDefault(0))) == "total")
            {
                return filas[i];
            }
        }

        return null;
    }

    /// <summary>
    /// Normalización del reader de banco (minúsculas, solo letras/dígitos y <c>-</c>;
    /// conserva el guion de los prefijos de concepto).
    /// </summary>
    private static string NormalizarEtiqueta(string texto)
    {
        var sb = new System.Text.StringBuilder(texto.Length);
        foreach (var ch in texto)
        {
            if (char.IsLetterOrDigit(ch) || ch == '-')
            {
                sb.Append(char.ToLowerInvariant(ch));
            }
        }

        return sb.ToString();
    }

    // ── Helpers de bytes (BCL, read-only) ─────────────────────────────────────────────────────

    private sealed record FilaExcel(int Fila, Dictionary<int, object?> Celdas);

    private static string Raiz()
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

    private static string RutaFuente(string periodo, int ase)
    {
        var insumos = Path.Combine(Raiz(), "Docs", periodo, "Insumos");
        Assert.True(Directory.Exists(insumos), $"Falta la carpeta de insumos: {insumos}");
        var carpeta = Directory.EnumerateDirectories(insumos)
            .FirstOrDefault(d => Path.GetFileName(d).StartsWith($"{ase}-", StringComparison.Ordinal));
        Assert.True(carpeta is not null, $"Falta la carpeta del ASE {ase} en {insumos}");
        var archivo = Directory.EnumerateFiles(carpeta!, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("ReportePagosxBanco", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta ReportePagosxBanco del ASE {ase} en {carpeta}");
        return archivo!;
    }

    private static int BuscarResumen(List<FilaExcel> filas)
    {
        for (var i = filas.Count - 1; i >= 0; i--)
        {
            foreach (var (_, valor) in filas[i].Celdas)
            {
                var texto = Texto(valor);
                if (texto.Length > 0 && NormalizarEtiqueta(texto).StartsWith(PrefijoResumen, StringComparison.Ordinal))
                {
                    return i;
                }
            }
        }

        return -1;
    }

    private static List<FilaExcel> LeerFilas(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var compartidas = new List<string>();
        var entrada = zip.GetEntry("xl/sharedStrings.xml");
        if (entrada is not null)
        {
            using var stream = entrada.Open();
            var doc = XDocument.Load(stream);
            foreach (var si in doc.Root!.Elements(Ns + "si"))
            {
                compartidas.Add(string.Concat(si.Descendants(Ns + "t").Select(t => t.Value)));
            }
        }

        var hoja = zip.Entries.First(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                                          && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        using var hojaStream = hoja.Open();
        var hojaDoc = XDocument.Load(hojaStream);
        var filas = new List<FilaExcel>();
        foreach (var row in hojaDoc.Root!.Descendants(Ns + "row"))
        {
            var numero = int.Parse((string?)row.Attribute("r") ?? "0", CultureInfo.InvariantCulture);
            var celdas = new Dictionary<int, object?>();
            foreach (var c in row.Elements(Ns + "c"))
            {
                var referencia = (string?)c.Attribute("r") ?? string.Empty;
                var indice = IndiceColumna(referencia);
                var tipo = (string?)c.Attribute("t");
                var v = c.Element(Ns + "v")?.Value;
                object? valor = tipo switch
                {
                    "s" when v is not null => compartidas[int.Parse(v, CultureInfo.InvariantCulture)],
                    "inlineStr" => string.Concat(c.Descendants(Ns + "t").Select(t => t.Value)),
                    _ => v is null ? null : decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var numeroCelda) ? numeroCelda : v,
                };
                if (valor is not null)
                {
                    celdas[indice] = valor;
                }
            }

            filas.Add(new FilaExcel(numero, celdas));
        }

        return filas;
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

            n = n * 26 + (ch - 'A' + 1);
        }

        return n - 1;
    }

    private static string Texto(object? valor) => valor switch
    {
        null => string.Empty,
        string s => s.Trim(),
        decimal d => d.ToString(CultureInfo.InvariantCulture),
        var o => o.ToString()?.Trim() ?? string.Empty,
    };

    private static decimal Numero(FilaExcel fila, int columna) => fila.Celdas.GetValueOrDefault(columna) switch
    {
        decimal d => d,
        double d => (decimal)d,
        string s when decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var r) => r,
        _ => 0m,
    };
}
