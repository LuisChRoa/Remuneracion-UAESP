using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 50 (H1) — A1 (opcionalidad de fuentes de recaudo) y A3 sobre las fuentes REALES del período.
/// Trazabilidad: S2.1 (ausencia total → 0 + huella), S2.3 (repro del rojo 2026101 sin throw en ASE2),
/// S2.5 (visibles coherentes con el 0) y la neutralidad de período del reader (A1 es period-free).
///
/// **A2 (levantamiento del recorte Q2) NO se habilita**: la escritura de la conciliación en Q2 aborta en
/// `Rem. Anticipos R2!E285` porque en la v2 esa celda destino es fórmula → residuo declarado
/// `R-MAPA-V2-R2-ASE4` (`Docs/T0-Plan50-FuentesRecaudoOpcionales.md` §4.3). El recorte del procesador se
/// mantiene y su test (`ProcesadorPeriodoTests.Ejecutar_Periodo2026072_SalidaCertificada5De5_ConDetRetri`)
/// lo re-congela con esa evidencia.
///
/// La política ratificada (§7.1/§7.3) es:
/// - label ausente (CERO filas) → 0 explícito por celda del mapa + huella <c>{Reporte}:{label}</c>;
/// - label presente con un bloque faltante (parcial) → fail-fast vigente (residuo declarado en
///   `Docs/T0-Plan50-FuentesRecaudoOpcionales.md` §2.2), y
/// - invariantes de cierre (T0e) → fail-fast intacto.
/// </summary>
public sealed class FuentesAusentesPlan50Tests
{
    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    [Fact]
    public void S2_1_Ase2Octubre_R1SinNuevoEsquema_QuedaEnCeroConHuella()
    {
        var leafs = LeerConciliacion("Prueba Octubre-1", "2026101", 2);

        var recip = leafs.Single(c => c.Empresa.Id == 1);
        Assert.Equal(new[] { "R1:NUEVO ESQUEMA", "R4:NUEVO ESQUEMA" }, recip.FuentesAusentes.OrderBy(x => x, StringComparer.Ordinal));

        // Plan 55 (H1.1, D4): los destinos se resuelven por CRUCE DE ETIQUETAS contra el cuadro; si la
        // fuente está ausente, sus operandos quedan en el CERO del template y no hay celda que escribir.
        Assert.Empty(recip.CeldasR1);
        Assert.Empty(recip.CeldasR4);

        // S2.5: el visible de una fuente ausente es 0 (no rompe el gate Σ).
        Assert.Equal(0m, recip.VisibleR1);
        Assert.Equal(0m, recip.VisibleR4);

        // Las fuentes presentes del mismo ASE siguen leyéndose normalmente.
        var enel = leafs.Single(c => c.Empresa.Id == 2);
        Assert.Empty(enel.FuentesAusentes);
        Assert.Equal(13705253661.25m, enel.CeldasR1["F150"]);
    }

    [Fact]
    public void S2_1_Ase5Octubre_R4SinOccidente_QuedaEnCeroConHuella()
    {
        // Octubre-1 ASE5: R4 solo trae ENEL (OCCIDENTE ausente). (ASE3 de octubre es el parcial
        // declarado del R1 — §2.2 del T0 — y se verifica en su propio test.)
        var occidente = LeerConciliacion("Prueba Octubre-1", "2026101", 5).Single(c => c.Empresa.Id == 4);

        Assert.Equal(new[] { "R4:OCCIDENTE" }, occidente.FuentesAusentes);
        Assert.All(occidente.CeldasR4.Values, valor => Assert.Equal(0m, valor));
        Assert.Equal(0m, occidente.VisibleR4);
    }

    [Fact]
    public void S2_1_Ase3Agosto_R4SinOccidente_QuedaEnCeroConHuella()
    {
        var occidente = LeerConciliacion("Prueba Agosto-2", "2026082", 3).Single(c => c.Empresa.Id == 4);
        Assert.Equal(new[] { "R4:OCCIDENTE" }, occidente.FuentesAusentes);
        Assert.All(occidente.CeldasR4.Values, valor => Assert.Equal(0m, valor));
        Assert.Equal(0m, occidente.VisibleR4);
    }

    [Fact]
    public void S2_2_Reader_LeeLaConciliacionPorEmpresa_EnQ1yQ2_PeriodFree()
    {
        // A1 es period-free (la ausencia se predica del resultado observado, no de la quincena):
        // el reader materializa 0 + huella igual en Q1 y Q2. El RECORTE del procesador en Q2 sigue
        // vigente por el residuo `R-MAPA-V2-R2-ASE4` (la celda R2!E285 es fórmula en la v2), NO por
        // esta política — ver `Docs/T0-Plan50-FuentesRecaudoOpcionales.md` §4.3.
        foreach (var ase in new[] { 1, 2, 3, 4, 5 })
        {
            var q1 = LeerConciliacion("Prueba Julio-1", "2026071", ase);
            var q2 = LeerConciliacion("Prueba Julio-2", "2026072", ase);

            // Plan 55 (H1.1): el catálogo (5) + las etiquetas que el propio CUADRO declara y la fuente
            // trae (p. ej. EAB en ASE2 octubre) — nunca una lista por ASE en código.
            Assert.True(q1.Count >= 5, $"{ase}: Q1 trae {q1.Count} entradas (catálogo = 5).");
            Assert.True(q2.Count >= 5, $"{ase}: Q2 trae {q2.Count} entradas (catálogo = 5).");
        }

        // La causa del recorte (R4-por-empresa Q2 sin NUEVO ESQUEMA en ASE2) ya no aborta: 0 + huella.
        var ase2 = LeerConciliacion("Prueba Julio-2", "2026072", 2);
        var recip = ase2.Single(c => c.Empresa.Id == 1);
        Assert.Equal(new[] { "R4:NUEVO ESQUEMA" }, recip.FuentesAusentes);
        Assert.Empty(recip.CeldasR4);
    }

    [Fact]
    public void S2_2_Agosto_Ase1YAse3_R4SinOccidente_QuedaEnCeroConHuella()
    {
        foreach (var ase in new[] { 1, 3 })
        {
            var leafs = LeerConciliacion("Prueba Agosto-2", "2026082", ase);
            var occidente = leafs.Single(c => c.Empresa.Id == 4);

            Assert.Contains("R4:OCCIDENTE", occidente.FuentesAusentes);
            Assert.Equal(0m, occidente.VisibleR4);
        }

        var ase2 = LeerConciliacion("Prueba Agosto-2", "2026082", 2);
        Assert.Contains("R4:NUEVO ESQUEMA", ase2.Single(c => c.Empresa.Id == 1).FuentesAusentes);
    }

    [Fact]
    public void S2_1_ParcialConLabelPresente_SigueSiendoFailFast()
    {
        // SUPERSEDIDO POR EL PLAN 53 (H1.4): el caso "parcial con label presente" de octubre-1 ASE3
        // R1 OCCIDENTE (una sola fila `Total` en la sección oportuna) YA NO es fail-fast: la sección
        // declarada del contrato (`TotalMenor`/`SubsMenor`/`TotalMedio`) está ausente en la fuente y su
        // celda queda en **0 explícito** (reglas 6 y 7 del project-context; residuo
        // `R-PARCIAL-OCT-ASE3-R1-OCC` cerrado). El método se conserva con aserciones de valor
        // (Plan 53 §7-H1.4 / §9 "baja+alta declarada").
        var leafs = LeerConciliacion("Prueba Octubre-1", "2026101", 3);
        var occidente = leafs.Single(c => c.Empresa.Id == 4);

        // Plan 55 (H1.1, D4): los destinos se resuelven por el cuadro del template. OCCIDENTE trae una
        // sola sección Oportuno en el cuadro Componente (F329) y otra en Subs/Cont (F352).
        Assert.Equal(286711907.53m, occidente.CeldasR1["F329"]);
        Assert.Equal(145489038.47m, occidente.CeldasR1["F352"]);
        // Celdas: F329 (Total de la sección), L329 (Especiales) y F352 (Subs/Cont de la sección).
        Assert.Equal(3, occidente.CeldasR1.Count);
        Assert.Equal(427389334.19m, occidente.VisibleR1);
        Assert.DoesNotContain(occidente.FuentesAusentes, f => f.StartsWith("R1:", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void S2_1_InvarianteComponenteAusente_SigueSiendoFailFast()
    {
        // Invariante de estructura (no es una fuente opcional): la fila A='Componente' debe seguir
        // abortando. Se elimina la fila de una COPIA en temp (la fuente real no se toca).
        var original = RutaFuente("Prueba Octubre-1", 1, "Recaudoporcomponente");
        var copia = Path.Combine(Path.GetTempPath(), "p50-r1-sin-componente-" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.Copy(original, copia);

        try
        {
            var filas = LeerFilas(copia);
            var indice = filas.FindIndex(f => Txt(f, 0).Equals("Componente", StringComparison.OrdinalIgnoreCase));
            Assert.True(indice > 0, "La fuente de octubre-1 ASE1 debe traer la fila invariante 'Componente'.");
            EliminarFilaPorIndice(copia, indice);

            var ex = Assert.Throws<CalculoInvalidoException>(() =>
                new ExcelDataReaderWorkbookLeafInputReader { RutaPlantilla = Insumos.PlantillaOficial() }.LeerConciliacionEmpresas(
                    Insumos.Ase(1),
                    Periodo.Parse("2026101"),
                    copia,
                    RutaFuente("Prueba Octubre-1", 1, "RerpoteDetalleSaldosaFavor"),
                    RutaFuente("Prueba Octubre-1", 1, "Reversi")));

            Assert.Contains("no se encontró la fila 'Componente'", ex.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            File.Delete(copia);
        }
    }

    private static IReadOnlyList<ConciliacionEmpresaInputs> LeerConciliacion(string periodo, string codigo, int ase)
    {
        // Plan 55 (H1.1/D4): la PLANTILLA es la autoridad de los destinos (resolución por etiquetas).
        var reader = new ExcelDataReaderWorkbookLeafInputReader { RutaPlantilla = Insumos.PlantillaOficial() };
        return reader.LeerConciliacionEmpresas(
            Insumos.Ase(ase),
            Periodo.Parse(codigo),
            RutaFuente(periodo, ase, "Recaudoporcomponente"),
            RutaFuente(periodo, ase, "RerpoteDetalleSaldosaFavor"),
            RutaFuente(periodo, ase, "Reversi"));
    }

    /// <summary>
    /// Elimina la fila XML en la posición <paramref name="indice"/> del orden de <c>&lt;row&gt;</c>
    /// (el mismo orden denso que devuelve el lector local; las hojas omiten filas vacías, así que el
    /// número de fila de Excel no coincide con el índice).
    /// </summary>
    private static void EliminarFilaPorIndice(string ruta, int indice)
    {
        using var zip = ZipFile.Open(ruta, ZipArchiveMode.Update);
        var entrada = zip.Entries.First(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                                             && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        XDocument doc;
        using (var stream = entrada.Open())
        {
            doc = XDocument.Load(stream);
        }

        XNamespace ns = doc.Root!.Name.Namespace;
        var filas = doc.Root.Descendants(ns + "row").ToList();
        Assert.True(indice < filas.Count, $"La hoja trae {filas.Count} filas; no existe el índice {indice}.");
        filas[indice].Remove();

        var nombre = entrada.FullName;
        entrada.Delete();
        var nueva = zip.CreateEntry(nombre);
        using var escritura = nueva.Open();
        doc.Save(escritura);
    }

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md"))) return dir.FullName;
            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (AGENTS.md).");
    }

    private static string RutaFuente(string periodo, int ase, string prefijo)
    {
        var dir = Path.Combine(Raiz(), "Docs", periodo, "Insumos", CarpetasAse[ase - 1]);
        Assert.True(Directory.Exists(dir), $"Falta la carpeta explícita de insumos: {dir}");
        var archivo = Directory.EnumerateFiles(dir, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta el reporte '{prefijo}*' del ASE {ase} en {dir}");
        return archivo!;
    }

    private static List<Dictionary<int, object?>> LeerFilas(string ruta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var compartidas = new List<string>();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var entrada = zip.GetEntry("xl/sharedStrings.xml");
        if (entrada is not null)
        {
            using var stream = entrada.Open();
            var doc = XDocument.Load(stream);
            foreach (var si in doc.Root!.Elements(ns + "si"))
            {
                compartidas.Add(string.Concat(si.Descendants(ns + "t").Select(t => t.Value)));
            }
        }

        var hoja = zip.Entries.First(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                                         && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        using var hojaStream = hoja.Open();
        var hojaDoc = XDocument.Load(hojaStream);
        var filas = new List<Dictionary<int, object?>>();
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
                    _ => v is null ? null : decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var numero) ? numero : v,
                };
                if (valor is not null) celdas[indice] = valor;
            }

            filas.Add(celdas);
        }

        return filas;
    }

    private static int IndiceColumna(string referencia)
    {
        var n = 0;
        foreach (var ch in referencia)
        {
            if (ch is < 'A' or > 'Z') break;
            n = n * 26 + (ch - 'A' + 1);
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
}
