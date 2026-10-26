using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 56 (H2.1/H2.3 — S3.1/S3.2): cobertura de la identidad de cuadro
/// <c>(sección A, subsección B, ocurrencia de B dentro de A)</c> del escritor values-only de
/// <c>Reporte Componentes R1</c>.
///
/// <list type="bullet">
/// <item>15 cruces (5 bloques x julio-1 / julio-2 / octubre-1) con insumos reales por ruta explícita:
/// cada fila resuelve a una fila de la MISMA identidad y la escritura lleva el cierre de cada cuadro
/// a la fila que la plantilla declara para ese cuadro;</item>
/// <item>regla de destino inexistente (§2-D3): combinación ausente en todo el bloque implica mensaje
/// accionable (A+B+fuente+canal+concepto+hoja+rango) y CERO escritura; fila fuera de su cuadro pero
/// declarada en el bloque implica cero documentado si el cuadro cierra, error si no cierra;</item>
/// <item>candado anti-quemado de dos capas (§2-D5): estático (allow-list del camino de resolución +
/// prueba de injerto) y conductual (mismas reglas en los 5 bloques y reorden dentro de la subsección).
/// </item>
/// </list>
/// </summary>
public sealed class JerarquiaABH1Plan56Tests
{
    private const string HojaR1 = "Reporte Componentes R1";
    private const string InicioMarcador = "[P56-JERARQUIA-AB-INICIO]";
    private const string FinMarcador = "[P56-JERARQUIA-AB-FIN]";

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    private static readonly (string Periodo, string Carpeta)[] Periodos =
    [
        ("2026071", Path.Combine("Docs", "Prueba Julio-1", "Insumos")),
        ("2026072", Path.Combine("Docs", "Prueba Julio-2", "Insumos")),
        ("2026101", Path.Combine("Docs", "Prueba Octubre-1", "Insumos")),
    ];

    /// <summary>
    /// CA-JER-1/CA-JER-2 (S3.1): los 15 cruces resuelven por la MISMA ruta (identidad A/B), sin
    /// fallback posicional: toda fila fuente queda colocada o en cero documentado, y cada colocación
    /// aterriza en una fila de su MISMA identidad (sección/subsección/ocurrencia) y su MISMA firma.
    /// </summary>
    [Theory]
    [MemberData(nameof(Cruces15))]
    public void H2_IdentidadAB_15de15(string periodo, string carpeta, int ase)
    {
        var bloque = LeerFuente(carpeta, ase);
        var plantilla = LeerPlantilla(ase);

        var plan = ContratoSeccionPeriodFreeR1.PlanificarJerarquiaR1(bloque.Filas, plantilla.Filas);

        Assert.True(plan.AplicaJerarquia, $"{periodo} ASE{ase}: la fuente real no declara jerarquía A/B.");
        Assert.Null(plan.Fallo);
        Assert.Equal(bloque.Filas.Count, plan.Colocaciones.Count + plan.Ceros.Count);

        var jerarquiaFuente = ContratoSeccionPeriodFreeR1.ParticionarJerarquia(bloque.Filas);
        var jerarquiaPlantilla = ContratoSeccionPeriodFreeR1.ParticionarJerarquia(plantilla.Filas);
        AssertSoloIdentidades(jerarquiaFuente, bloque.Filas, jerarquiaPlantilla, plantilla.Filas, plan.Colocaciones);

        // El cierre de cada cuadro de la fuente es el que la plantilla declara para ese cuadro: la
        // colocación de los cierres es la que alimenta TOT_OPT/EXTEMP (nunca se mueve un peso).
        foreach (var seccion in jerarquiaFuente.Secciones)
        {
            var seccionPlantilla = jerarquiaPlantilla.BuscarSeccion(seccion.Etiqueta, seccion.Ocurrencia);
            Assert.NotNull(seccionPlantilla);
            Assert.Contains(plan.Colocaciones, c => c.Fuente == seccion.Fin && c.Slot == seccionPlantilla!.Fin);

            foreach (var cuadro in seccion.Subsecciones)
            {
                var cuadroPlantilla = seccionPlantilla!.BuscarSubseccion(cuadro.Etiqueta, cuadro.Ocurrencia);
                Assert.NotNull(cuadroPlantilla);
                Assert.Contains(plan.Colocaciones, c => c.Fuente == cuadro.Fin && c.Slot == cuadroPlantilla!.Fin);
            }
        }

        if (jerarquiaFuente.FilaTotal is { } totalFuente)
        {
            Assert.Equal(jerarquiaPlantilla.FilaTotal, plan.Colocaciones.Single(c => c.Fuente == totalFuente).Slot);
        }
    }

    /// <summary>
    /// S3.1/S3.2: escritura REAL de los 5 bloques del período. El valor que entra en cada celda de
    /// cierre de cuadro/sección/gran total es el de la fuente (identidad), las combinaciones que la
    /// fuente no trae quedan en el CERO del template y no hay error.
    /// </summary>
    [Theory]
    [InlineData("2026071")]
    [InlineData("2026072")]
    [InlineData("2026101")]
    public void H2_Escritura_LosCierresCaenEnSuCuadro_5Bloques(string periodo)
    {
        var carpeta = Periodos.Single(p => p.Periodo == periodo).Carpeta;
        using var salida = new CopiaTemporal();

        var bloques = new List<(BloqueEspejoAseInputs Bloque, BloquePlantilla Plantilla)>();
        using (var documento = SpreadsheetDocument.Open(salida.Ruta, true))
        {
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            for (var ase = 1; ase <= 5; ase++)
            {
                var bloque = LeerFuente(carpeta, ase);
                var plantilla = LeerPlantilla(ase);
                EscritorValoresSlotR1.Escribir(parte, bloque);
                bloques.Add((bloque, plantilla));
            }

            parte.Workbook?.Save();
        }

        for (var ase = 1; ase <= 5; ase++)
        {
            var (bloque, plantillaCompleta) = bloques[ase - 1];
            var plantilla = plantillaCompleta.Filas;
            var inicio = plantillaCompleta.Inicio;
            var jerarquiaFuente = ContratoSeccionPeriodFreeR1.ParticionarJerarquia(bloque.Filas);
            var jerarquiaPlantilla = ContratoSeccionPeriodFreeR1.ParticionarJerarquia(plantilla);
            var declarados = new HashSet<int>();

            foreach (var seccion in jerarquiaFuente.Secciones)
            {
                var seccionPlantilla = jerarquiaPlantilla.BuscarSeccion(seccion.Etiqueta, seccion.Ocurrencia);
                Assert.NotNull(seccionPlantilla);
                AssertCierre(
                    salida.Ruta,
                    inicio + seccionPlantilla!.Fin,
                    bloque.Filas,
                    seccion.Fin,
                    $"{periodo} ASE{ase} sección {seccion.Etiqueta}");
                declarados.Add(seccionPlantilla.Fin);

                foreach (var cuadro in seccion.Subsecciones)
                {
                    var cuadroPlantilla = seccionPlantilla.BuscarSubseccion(cuadro.Etiqueta, cuadro.Ocurrencia);
                    Assert.NotNull(cuadroPlantilla);
                    AssertCierre(
                        salida.Ruta,
                        inicio + cuadroPlantilla!.Fin,
                        bloque.Filas,
                        cuadro.Fin,
                        $"{periodo} ASE{ase} cuadro ({seccion.Etiqueta}, {cuadro.Etiqueta}, #{cuadro.Ocurrencia})");
                    declarados.Add(cuadroPlantilla.Fin);
                }
            }

            if (jerarquiaFuente.FilaTotal is { } total)
            {
                Assert.NotNull(jerarquiaPlantilla.FilaTotal);
                AssertCierre(
                    salida.Ruta,
                    inicio + jerarquiaPlantilla.FilaTotal!.Value,
                    bloque.Filas,
                    total,
                    $"{periodo} ASE{ase} gran total");
                declarados.Add(jerarquiaPlantilla.FilaTotal.Value);
            }

            // Ausencia en la fuente implica cero del template, sin error: los cierres declarados por
            // la plantilla que la fuente no trae quedan intactos en su 0.
            for (var i = 0; i < plantilla.Count; i++)
            {
                if (declarados.Contains(i)
                    || !EsCierre(plantilla[i]))
                {
                    continue;
                }

                Assert.Equal(0m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, $"F{inicio + i}"));
            }
        }
    }

    /// <summary>
    /// CA-JER-4 (§2-D3): combinación presente en la fuente y ausente en TODO el bloque implica
    /// <c>ERR-SIN-SLOT</c> accionable con sección + subsección + fuente + canal + concepto + hoja +
    /// rango, y CERO escritura (el archivo queda byte-idéntico a la copia previa).
    /// </summary>
    [Fact]
    public void H2_CombinacionAusenteEnTodoElBloque_ErrorAccionableSinEscribir()
    {
        var bloque = LeerFuente(Periodos[0].Carpeta, 1);
        var filas = bloque.Filas.ToList();
        filas.Insert(1, new FilaEspejoR1
        {
            C = "EMPRESA FANTASMA",
            D = "Total",
            ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase) { ["Total"] = 7m },
        });

        var conFantasma = new BloqueEspejoAseInputs
        {
            Ase = bloque.Ase,
            Encabezados = bloque.Encabezados,
            TieneColumnaEspeciales = bloque.TieneColumnaEspeciales,
            Filas = filas,
        };

        using var salida = new CopiaTemporal();
        var antes = FotoCeldas(salida.Ruta);
        string mensaje;
        using (var documento = SpreadsheetDocument.Open(salida.Ruta, true))
        {
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var ex = Assert.Throws<Remuneracion.Core.Exceptions.CalculoInvalidoException>(
                () => EscritorValoresSlotR1.Escribir(parte, conFantasma));
            Assert.Equal(Remuneracion.Core.Errors.CodigoError.SinSlot, ex.Codigo);
            mensaje = ex.Message;
        }

        Assert.Contains("ERR-SIN-SLOT", mensaje, StringComparison.Ordinal);
        Assert.Contains("ASE 1", mensaje, StringComparison.Ordinal);
        Assert.Contains("sección 'Componente'", mensaje, StringComparison.Ordinal);
        Assert.Contains("subsección 'Aplicacion nuevos x reversion'", mensaje, StringComparison.Ordinal);
        Assert.Contains("EMPRESA FANTASMA", mensaje, StringComparison.Ordinal);
        Assert.Contains("canal (D) 'Total'", mensaje, StringComparison.Ordinal);
        Assert.Contains(HojaR1, mensaje, StringComparison.Ordinal);

        // Plan 59 (H1.2): el rango que viaja en el mensaje es el bloque DELIMITADO de la propia
        // plantilla (cabecera por rótulo + gran total), no la ventana congelada `6-72` del contrato
        // viejo. Se contrasta contra el delimitador sobre LA MISMA copia y se exige que la ventana
        // histórica ya no aparezca.
        int cabeceraBloque;
        int granTotalBloque;
        using (var verificacion = SpreadsheetDocument.Open(salida.Ruta, false))
        {
            var parteVerificacion = verificacion.WorkbookPart
                ?? throw new InvalidOperationException("WorkbookPart null");
            var bloqueHoja = DelimitadorBloquesHoja.DelimitarBloque(parteVerificacion, DelimitadorBloquesHoja.R1, 1);
            cabeceraBloque = bloqueHoja.Cabecera;
            granTotalBloque = bloqueHoja.GranTotal;
        }

        Assert.Contains(
            $"rango del bloque {cabeceraBloque}-{granTotalBloque}",
            mensaje,
            StringComparison.Ordinal);
        Assert.DoesNotContain("6-72", mensaje, StringComparison.Ordinal);

        Assert.Equal(antes, FotoCeldas(salida.Ruta));
    }

    /// <summary>
    /// §2-D3 rama 2: la fila no tiene destino en su cuadro y el subtotal del cuadro NO cierra sin
    /// ella (se perdería dinero) implica error accionable + alto, sin escribir.
    /// </summary>
    [Fact]
    public void H2_FilaSinDestinoEnSuCuadro_SubtotalNoCierra_ErrorSinEscribir()
    {
        var bloque = LeerFuente(Periodos[0].Carpeta, 2);
        var filas = bloque.Filas.ToList();

        // Rompe el subtotal del cuadro (AFaseo, Mes, #1): la fila de cierre B='Mes'/C='Total' deja de
        // cuadrar con la suma de sus filas-hoja (los -16.473,56 de las 3 filas sin destino).
        var cierre = filas.FindIndex(f =>
            string.Equals(f.B.Trim(), "Mes", StringComparison.OrdinalIgnoreCase)
            && string.Equals(f.C.Trim(), "Total", StringComparison.OrdinalIgnoreCase));
        Assert.True(cierre > 0, "La fuente de julio-1 ASE2 no trae el cierre del cuadro (AFaseo, Mes, #1).");
        filas[cierre] = new FilaEspejoR1
        {
            A = filas[cierre].A,
            B = filas[cierre].B,
            C = filas[cierre].C,
            D = filas[cierre].D,
            E = filas[cierre].E,
            ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase) { ["Total"] = 0m },
        };

        var roto = new BloqueEspejoAseInputs
        {
            Ase = bloque.Ase,
            Encabezados = bloque.Encabezados,
            TieneColumnaEspeciales = bloque.TieneColumnaEspeciales,
            Filas = filas,
        };

        using var salida = new CopiaTemporal();
        var antes = FotoCeldas(salida.Ruta);
        string mensaje;
        using (var documento = SpreadsheetDocument.Open(salida.Ruta, true))
        {
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            var ex = Assert.Throws<Remuneracion.Core.Exceptions.CalculoInvalidoException>(
                () => EscritorValoresSlotR1.Escribir(parte, roto));
            Assert.Equal(Remuneracion.Core.Errors.CodigoError.SinSlot, ex.Codigo);
            mensaje = ex.Message;
        }

        Assert.Contains("subtotal del cuadro NO cierra", mensaje, StringComparison.Ordinal);
        Assert.Contains("sección 'AFaseo'", mensaje, StringComparison.Ordinal);
        Assert.Contains("subsección 'Mes'", mensaje, StringComparison.Ordinal);
        Assert.Equal(antes, FotoCeldas(salida.Ruta));
    }

    /// <summary>
    /// §2-D3 rama 1 (puro, sin I/O): el caso medido de julio-1 ASE2 queda en CERO DOCUMENTADO con
    /// registro durable (cuadro + firma + canal + concepto + valor), sin abortar.
    /// </summary>
    [Fact]
    public void H2_FilaSinDestinoEnSuCuadro_SubtotalCierra_CeroDocumentado()
    {
        var bloque = LeerFuente(Periodos[0].Carpeta, 2);
        var plantilla = LeerPlantilla(2);

        var plan = ContratoSeccionPeriodFreeR1.PlanificarJerarquiaR1(bloque.Filas, plantilla.Filas);

        Assert.Null(plan.Fallo);
        Assert.Equal(3, plan.Ceros.Count);
        Assert.All(plan.Ceros, cero =>
        {
            Assert.Equal(new ContratoSeccionPeriodFreeR1.CuadroR1("AFaseo", "Mes", 1), cero.Cuadro);
            Assert.Equal(-16473.56m, cero.Total);
        });
        Assert.Contains(plan.Ceros, c => c.Canal == "5" && c.Concepto == "Vlr Servicio");
        Assert.Contains(plan.Ceros, c => c.Canal == "Total" && c.Fuente == "NUEVO ESQUEMA");
        Assert.Equal(bloque.Filas.Count, plan.Colocaciones.Count + plan.Ceros.Count);
    }

    // ── §2-D5 candado capa 1: ESTÁTICO (allow-list del camino de resolución + injerto) ──────────────

    [Fact]
    public void Candado_Estatico_ElCaminoDeResolucionNoQuemaBloqueNiPeriodo()
    {
        var core = ExtraerRegion(Path.Combine("Remuneracion.Core", "Models", "ContratoSeccionPeriodFreeR1.cs"));
        var writer = ExtraerRegion(Path.Combine("Remuneracion.Infrastructure", "Excel", "EscritorValoresSlotR1.cs"));

        Assert.NotEmpty(core);
        Assert.NotEmpty(writer);
        Assert.Empty(Violaciones(core));
        Assert.Empty(Violaciones(writer));

        // Prueba de INJERTO: cualquier id de bloque, nombre de empresa, período, quincena o dirección
        // congelada inyectada en el camino de resolución debe ENROJECER el candado.
        var injertos = new[]
        {
            "var bloque = 5; ASE 5", "if (empresa == \"AREA LIMPIA\")", "periodo == \"2026071\"",
            "if (quincena == 2)", "var celda = \"F637\";", "if (rotulo == \"ENERBIT\")",
        };
        foreach (var injerto in injertos)
        {
            Assert.NotEmpty(Violaciones(core + "\n" + injerto));
            Assert.NotEmpty(Violaciones(writer + "\n" + injerto));
        }
    }

    // ── §2-D5 candado capa 2: CONDUCTUAL (mismas reglas en 5 bloques + reorden) ─────────────────────

    [Fact]
    public void Candado_Conductual_ElMismoFixtureProduceLasMismasReglasEnLos5Bloques()
    {
        var plantillas = Enumerable.Range(1, 5).Select(LeerPlantilla).ToList();
        var comun = FirmasComunes(plantillas);
        Assert.NotEmpty(comun);

        // Fixture ÚNICO: las filas comunes a los 5 bloques (mismo orden del bloque 1) más los cierres.
        var fixture = ConstruirFixtureComun(plantillas[0], comun);
        Assert.NotEmpty(fixture);

        var jerarquiaFixture = ContratoSeccionPeriodFreeR1.ParticionarJerarquia(fixture);
        Assert.True(jerarquiaFixture.Secciones.Count >= 3);

        foreach (var (plantilla, indice) in plantillas.Select((p, i) => (p, i + 1)))
        {
            var plan = ContratoSeccionPeriodFreeR1.PlanificarJerarquiaR1(fixture, plantilla.Filas);
            var jerarquiaPlantilla = ContratoSeccionPeriodFreeR1.ParticionarJerarquia(plantilla.Filas);

            Assert.True(plan.AplicaJerarquia, $"Bloque {indice}: la jerarquía no aplica al fixture común.");
            Assert.Null(plan.Fallo);
            Assert.Empty(plan.Ceros);
            Assert.Equal(fixture.Count, plan.Colocaciones.Count);
            AssertSoloIdentidades(jerarquiaFixture, fixture, jerarquiaPlantilla, plantilla.Filas, plan.Colocaciones);
        }
    }

    [Fact]
    public void Candado_Conductual_ReordenDentroDeLaSubseccionColocaIgual()
    {
        var bloque = LeerFuente(Periodos[1].Carpeta, 5);
        var plantilla = LeerPlantilla(5);
        var original = bloque.Filas.ToList();

        var planOriginal = ContratoSeccionPeriodFreeR1.PlanificarJerarquiaR1(original, plantilla.Filas);
        Assert.Null(planOriginal.Fallo);

        var jerarquia = ContratoSeccionPeriodFreeR1.ParticionarJerarquia(original);
        var seccion = jerarquia.BuscarSeccion("Componente", 1);
        Assert.NotNull(seccion);
        var cuadro = seccion!.BuscarSubseccion("Aplicacion nuevos x reversion", 1);
        Assert.NotNull(cuadro);

        // Reordena GRUPOS de canal completos (VS/VI/D=... Total) dentro de la subsección: las firmas y
        // los canales efectivos se conservan, la posición física cambia.
        var reordenado = new List<FilaEspejoR1>(original);
        var grupos = GruposDeCanal(original, cuadro!);
        Assert.True(grupos.Count >= 2, "El cuadro no trae al menos 2 grupos de canal para reordenar.");
        var primero = grupos[0];
        var segundo = grupos[1];
        var bloqueUno = original.GetRange(primero.Inicio, primero.Fin - primero.Inicio + 1);
        var bloqueDos = original.GetRange(segundo.Inicio, segundo.Fin - segundo.Inicio + 1);
        reordenado.RemoveRange(primero.Inicio, bloqueUno.Count + bloqueDos.Count);
        reordenado.InsertRange(primero.Inicio, bloqueDos.Concat(bloqueUno));

        var planReordenado = ContratoSeccionPeriodFreeR1.PlanificarJerarquiaR1(reordenado, plantilla.Filas);
        Assert.Null(planReordenado.Fallo);

        Assert.Equal(
            MapaPorFirma(original, jerarquia, planOriginal.Colocaciones),
            MapaPorFirma(reordenado, ContratoSeccionPeriodFreeR1.ParticionarJerarquia(reordenado), planReordenado.Colocaciones));
    }

    // ── Datos de prueba ─────────────────────────────────────────────────────────────────────────────

    public static IEnumerable<object[]> Cruces15()
    {
        foreach (var (periodo, carpeta) in Periodos)
        {
            for (var ase = 1; ase <= 5; ase++)
            {
                yield return [periodo, carpeta, ase];
            }
        }
    }

    private sealed record BloquePlantilla(IReadOnlyList<FilaEspejoR1> Filas, int Inicio, int Fin);

    private static BloqueEspejoAseInputs LeerFuente(string carpeta, int ase) =>
        new ExcelDataReaderWorkbookLeafInputReader().LeerEspejoR1(AseFactory.DesdeId(ase), FuenteR1(carpeta, ase));

    /// <summary>
    /// Bloque de la plantilla leído desde su geometría DECLARADA (Plan 59 H1.2): el rango sale del
    /// delimitador por cabecera de bloque + gran total, no de una ventana congelada. El índice 0 de
    /// <see cref="Filas"/> es la primera fila de slot del bloque (<see cref="Inicio"/>).
    /// </summary>
    private static BloquePlantilla LeerPlantilla(int ase)
    {
        using var documento = SpreadsheetDocument.Open(PlantillaV2(), false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, HojaR1, StringComparison.OrdinalIgnoreCase));
        var worksheet = ((WorksheetPart)parte.GetPartById(hoja.Id!)).Worksheet!;
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(parte, DelimitadorBloquesHoja.R1, ase, worksheet);
        var filas = new List<FilaEspejoR1>();
        foreach (var filaBloque in bloque.FilasSlot)
        {
            var numero = filaBloque.Fila;
            var row = worksheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex?.Value == numero);
            Assert.True(row is not null, $"La plantilla no trae la fila {numero} del bloque ASE {ase}.");
            filas.Add(new FilaEspejoR1
            {
                A = filaBloque.Etiquetas.A,
                B = filaBloque.Etiquetas.B,
                C = filaBloque.Etiquetas.C,
                D = filaBloque.Etiquetas.D,
                E = filaBloque.Etiquetas.E,
                ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["Total"] = Numero(parte, row!, $"F{numero}"),
                },
            });
        }

        Assert.NotEmpty(filas);
        return new BloquePlantilla(filas, bloque.FilasSlot[0].Fila, bloque.GranTotal);
    }

    private static string Texto(WorkbookPart parte, Row row, string referencia)
    {
        var celda = row.Elements<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        if (celda is null)
        {
            return string.Empty;
        }

        if (celda.DataType?.Value == CellValues.SharedString
            && celda.CellValue is not null
            && int.TryParse(celda.CellValue.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indice))
        {
            return parte.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(indice)?.InnerText
                ?? string.Empty;
        }

        return celda.CellValue?.InnerText ?? string.Empty;
    }

    private static decimal? Numero(WorkbookPart parte, Row row, string referencia)
    {
        var texto = Texto(parte, row, referencia);
        return decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor) ? valor : null;
    }

    // ── Identidad (aserción independiente del plan) ────────────────────────────────────────────────

    private static Dictionary<int, (string Rol, string? Seccion, string? Subseccion, int Ocurrencia)> Roles(
        ContratoSeccionPeriodFreeR1.JerarquiaR1 jerarquia)
    {
        var roles = new Dictionary<int, (string, string?, string?, int)>();
        foreach (var seccion in jerarquia.Secciones)
        {
            roles[seccion.Fin] = ("seccion", seccion.Etiqueta, null, seccion.Ocurrencia);
            foreach (var cuadro in seccion.Subsecciones)
            {
                roles[cuadro.Fin] = ("cierre-cuadro", seccion.Etiqueta, cuadro.Etiqueta, cuadro.Ocurrencia);
                for (var i = cuadro.Inicio; i < cuadro.Fin; i++)
                {
                    roles[i] = ("dato", seccion.Etiqueta, cuadro.Etiqueta, cuadro.Ocurrencia);
                }
            }
        }

        if (jerarquia.FilaTotal is { } total)
        {
            roles[total] = ("total", null, null, 0);
        }

        return roles;
    }

    private static void AssertSoloIdentidades(
        ContratoSeccionPeriodFreeR1.JerarquiaR1 jerarquiaFuente,
        IReadOnlyList<FilaEspejoR1> fuente,
        ContratoSeccionPeriodFreeR1.JerarquiaR1 jerarquiaPlantilla,
        IReadOnlyList<FilaEspejoR1> plantilla,
        IReadOnlyList<(int Fuente, int Slot)> colocaciones)
    {
        var rolesFuente = Roles(jerarquiaFuente);
        var rolesPlantilla = Roles(jerarquiaPlantilla);

        foreach (var (iFuente, iSlot) in colocaciones)
        {
            Assert.True(rolesFuente.ContainsKey(iFuente), $"La fila fuente {iFuente + 1} no tiene identidad A/B.");
            Assert.True(rolesPlantilla.ContainsKey(iSlot), $"La fila destino {iSlot + 1} no tiene identidad A/B.");
            var origen = rolesFuente[iFuente];
            var destino = rolesPlantilla[iSlot];

            Assert.Equal(origen.Rol, destino.Rol);
            Assert.Equal(origen.Seccion, destino.Seccion);
            Assert.Equal(origen.Subseccion, destino.Subseccion);
            Assert.Equal(origen.Ocurrencia, destino.Ocurrencia);

            if (origen.Rol != "dato")
            {
                continue;
            }

            // Dentro del cuadro: mismo concepto (E), mismo canal efectivo y misma fuente (C).
            var cuadroFuente = jerarquiaFuente.BuscarSeccion(origen.Seccion!, origen.Ocurrencia)!
                .BuscarSubseccion(origen.Subseccion!, origen.Ocurrencia)!;
            var cuadroPlantilla = jerarquiaPlantilla.BuscarSeccion(destino.Seccion!, destino.Ocurrencia)!
                .BuscarSubseccion(destino.Subseccion!, destino.Ocurrencia)!;
            var firmaFuente = ContratoSeccionPeriodFreeR1.FirmaCuadroR1(fuente, iFuente, cuadroFuente.Fin);
            var firmaSlot = ContratoSeccionPeriodFreeR1.FirmaCuadroR1(plantilla, iSlot, cuadroPlantilla.Fin);
            Assert.Equal(firmaFuente, firmaSlot);
        }
    }

    private static bool EsCierre(FilaEspejoR1 fila) =>
        (fila.A.Trim().Length > 0 && string.Equals(fila.B.Trim(), "Total", StringComparison.OrdinalIgnoreCase))
        || (fila.A.Trim().Length == 0 && fila.B.Trim().Length > 0
            && string.Equals(fila.C.Trim(), "Total", StringComparison.OrdinalIgnoreCase))
        || (string.Equals(fila.A.Trim(), "Total", StringComparison.OrdinalIgnoreCase) && fila.B.Trim().Length == 0);

    private static void AssertCierre(
        string ruta,
        int celdaFila,
        IReadOnlyList<FilaEspejoR1> fuente,
        int filaFuente,
        string etiqueta)
    {
        var esperado = fuente[filaFuente].Valor("Total");
        Assert.True(esperado is not null, $"{etiqueta}: la fila fuente {filaFuente + 1} no trae 'Total'.");
        Assert.Equal(esperado!.Value, TestHelpers.LeerCeldaNumerica(ruta, HojaR1, $"F{celdaFila}"));
    }

    // ── Candado conductual: fixture común y reorden ────────────────────────────────────────────────

    private static HashSet<string> FirmasComunes(IReadOnlyList<BloquePlantilla> plantillas)
    {
        HashSet<string>? comunes = null;
        foreach (var plantilla in plantillas)
        {
            var jerarquia = ContratoSeccionPeriodFreeR1.ParticionarJerarquia(plantilla.Filas);
            var firmas = new HashSet<string>(StringComparer.Ordinal);
            foreach (var seccion in jerarquia.Secciones)
            {
                foreach (var cuadro in seccion.Subsecciones)
                {
                    for (var i = cuadro.Inicio; i < cuadro.Fin; i++)
                    {
                        firmas.Add(ClaveFirma(plantilla.Filas, seccion.Etiqueta, cuadro, i));
                    }
                }
            }

            comunes = comunes is null
                ? firmas
                : comunes.Intersect(firmas, StringComparer.Ordinal).ToHashSet(StringComparer.Ordinal);
        }

        return comunes ?? [];
    }

    private static string ClaveFirma(
        IReadOnlyList<FilaEspejoR1> filas,
        string seccion,
        ContratoSeccionPeriodFreeR1.SubseccionR1 cuadro,
        int indice) =>
        $"{seccion.Trim()}\u001f{cuadro.Etiqueta.Trim()}\u001f{cuadro.Ocurrencia}\u001f{ContratoSeccionPeriodFreeR1.FirmaCuadroR1(filas, indice, cuadro.Fin)}";

    private static List<FilaEspejoR1> ConstruirFixtureComun(BloquePlantilla plantilla, HashSet<string> comunes)
    {
        var jerarquia = ContratoSeccionPeriodFreeR1.ParticionarJerarquia(plantilla.Filas);
        var fixture = new List<FilaEspejoR1>();
        foreach (var seccion in jerarquia.Secciones)
        {
            foreach (var cuadro in seccion.Subsecciones)
            {
                for (var i = cuadro.Inicio; i < cuadro.Fin; i++)
                {
                    if (comunes.Contains(ClaveFirma(plantilla.Filas, seccion.Etiqueta, cuadro, i)))
                    {
                        fixture.Add(plantilla.Filas[i]);
                    }
                }

                fixture.Add(new FilaEspejoR1 { B = cuadro.Etiqueta, C = "Total", ValoresPorColumna = Solo(0m) });
            }

            fixture.Add(new FilaEspejoR1 { A = seccion.Etiqueta, B = "Total", ValoresPorColumna = Solo(0m) });
        }

        fixture.Add(new FilaEspejoR1 { A = "Total", ValoresPorColumna = Solo(0m) });
        return fixture;
    }

    private sealed record GrupoCanal(int Inicio, int Fin);

    private static List<GrupoCanal> GruposDeCanal(
        IReadOnlyList<FilaEspejoR1> filas,
        ContratoSeccionPeriodFreeR1.SubseccionR1 cuadro)
    {
        var grupos = new List<GrupoCanal>();
        var inicio = cuadro.Inicio;
        for (var i = cuadro.Inicio; i < cuadro.Fin; i++)
        {
            if (filas[i].D.Trim().Length > 0
                && string.Equals(filas[i].E.Trim(), "Total", StringComparison.OrdinalIgnoreCase))
            {
                grupos.Add(new GrupoCanal(inicio, i));
                inicio = i + 1;
            }
        }

        return grupos;
    }

    private static Dictionary<string, int> MapaPorFirma(
        IReadOnlyList<FilaEspejoR1> filas,
        ContratoSeccionPeriodFreeR1.JerarquiaR1 jerarquia,
        IReadOnlyList<(int Fuente, int Slot)> colocaciones)
    {
        var destinoPorFila = colocaciones.ToDictionary(c => c.Fuente, c => c.Slot);
        var mapa = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var seccion in jerarquia.Secciones)
        {
            foreach (var cuadro in seccion.Subsecciones)
            {
                var ocurrencias = new Dictionary<string, int>(StringComparer.Ordinal);
                for (var i = cuadro.Inicio; i < cuadro.Fin; i++)
                {
                    var firma = ContratoSeccionPeriodFreeR1.FirmaCuadroR1(filas, i, cuadro.Fin);
                    ocurrencias.TryGetValue(firma, out var k);
                    ocurrencias[firma] = ++k;
                    var clave = $"({seccion.Etiqueta},{cuadro.Etiqueta},#{cuadro.Ocurrencia})|{firma}|{k}";
                    mapa[clave] = destinoPorFila.TryGetValue(i, out var slot) ? slot : -1;
                }
            }
        }

        return mapa;
    }

    private static Dictionary<string, decimal?> Solo(decimal? valor) =>
        new(StringComparer.OrdinalIgnoreCase) { ["Total"] = valor };

    // ── Candado estático ───────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Allow-list cerrada del camino de resolución: ninguna de estas formas puede aparecer en las
    /// regiones <c>[P56-JERARQUIA-AB-*]</c> (ids de bloque, nombres de empresa, quincenas, períodos o
    /// direcciones de celda congeladas).
    /// </summary>
    private static readonly Regex[] Prohibidos =
    [
        new(@"\bASE\s*-?\s*\d", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant),
        new(
            @"\b(PROMOAMBIENTAL|LIME|CIUDAD\s+LIMPIA|BOGOT[AÁ]\s+LIMPIA|[AÁ]REA\s+LIMPIA|ENEL|OCCIDENTE|ENERBIT|NUEVO\s+ESQUEMA|RECIPROCIDAD)\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant),
        new(@"\b(quincena|quincenal|Q1|Q2)\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant),
        new(@"\b20\d{5}\b", RegexOptions.CultureInvariant),
        new(@"(?<![A-Za-z0-9_$§/-])[A-Z]{1,2}\$?\d{1,4}(?![A-Za-z0-9/])", RegexOptions.CultureInvariant),
    ];

    private static IReadOnlyList<string> Violaciones(string texto) =>
        Prohibidos
            .SelectMany(p => p.Matches(texto).Select(m => m.Value))
            .Distinct(StringComparer.Ordinal)
            .ToList();

    private static string ExtraerRegion(string rutaRelativa)
    {
        var texto = File.ReadAllText(Path.Combine(Raiz(), rutaRelativa));
        var inicio = texto.IndexOf(InicioMarcador, StringComparison.Ordinal);
        var fin = texto.IndexOf(FinMarcador, StringComparison.Ordinal);
        Assert.True(
            inicio >= 0 && fin > inicio,
            $"No se encontró la región {InicioMarcador}..{FinMarcador} en {rutaRelativa}.");
        return texto[(inicio + InicioMarcador.Length)..fin];
    }

    // ── Rutas y utilidades ─────────────────────────────────────────────────────────────────────────

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

        throw new DirectoryNotFoundException(
            "No se encontró la raíz del repositorio (AGENTS.md + Docs/Plantilla estandarizada).");
    }

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

    /// <summary>
    /// Foto SEMÁNTICA de la hoja R1 (referencia + texto de fórmula + valor) para probar "cero
    /// escritura": abrir el paquete en modo edición re-serializa el ZIP aunque no se toque ninguna
    /// celda, así que el hash de bytes no es un oráculo válido aquí; las celdas sí lo son.
    /// </summary>
    private static string FotoCeldas(string ruta)
    {
        using var documento = SpreadsheetDocument.Open(ruta, false);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var hoja = parte.Workbook!.Descendants<Sheet>()
            .First(s => string.Equals(s.Name?.Value, HojaR1, StringComparison.OrdinalIgnoreCase));
        var worksheet = ((WorksheetPart)parte.GetPartById(hoja.Id!)).Worksheet!;
        var foto = new StringBuilder();
        foreach (var celda in worksheet.Descendants<Cell>().OrderBy(c => c.CellReference?.Value, StringComparer.Ordinal))
        {
            foto.Append(celda.CellReference?.Value)
                .Append('=')
                .Append(celda.CellFormula?.Text ?? string.Empty)
                .Append('|')
                .Append(celda.CellValue?.InnerText ?? string.Empty)
                .Append(';');
        }

        return foto.ToString();
    }

    private sealed class CopiaTemporal : IDisposable
    {
        private readonly string _dir;

        public CopiaTemporal()
        {
            _dir = Path.Combine(Path.GetTempPath(), "p56-h1-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
            Ruta = Path.Combine(_dir, "salida.xlsx");
            File.Copy(PlantillaV2(), Ruta, overwrite: false);
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
