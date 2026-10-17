using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 45 (H0, S1.1–S1.5): T0 multi-período READ-ONLY del contrato de sección PERIOD-FREE
/// (<c>plans/45 - contrato-seccion-periodfree-sinc.md</c>), evidencia congelada desde bytes para
/// derivar la regla sin-C sobre las TRES formas reales (julio-2 + agosto-2026082 + Q1-2026071).
///
/// CERO código de producción: el fixture + sus asserts SON la evidencia. Lectura SOLO con BCL
/// (<see cref="ZipArchive"/> + <see cref="XDocument"/>): sin Excel/COM, sin OpenXml SDK, sin
/// ExcelDataReader. El OFICIAL (única plantilla) nunca se modifica (SHA256 antes/después) y Q1 se
/// lee READ-ONLY (no se resuelve: su contrato es otro frente — RESTRICCIÓN DURA del plan §0.2).
///
/// Hallazgos congelados (verificación 2026-10-16 por ZIP+XML directo):
/// E1: el OFICIAL declara SIEMPRE 5 fronteras por ASE en el orden <c>Mes · Aplic · Mes · Aplic · Mes</c>
///     → 6 secciones <c>[1.ª, Oportuno#1, Aplic#1, Oportuno#2, Aplic#2, Oportuno#3]</c> (uniforme 5/5).
/// E2: las 5 secciones-1.ª del OFICIAL traen VS + canales E/H/O/T + OCCIDENTE y CERO <c>Vlr Intereses</c>.
/// E3: la FUENTE julio-2 ASE5 es la ÚNICA con <c>Vlr Intereses</c> en su 1.ª (R8 F=402,40; R11 F=7.697,23)
///     y además su forma arranca con <c>Aplicacion</c>: secciones <c>[1.ª, Aplic#1, Oportuno#1, Aplic#2, Oportuno#2]</c>.
/// E4: el MANUAL legacy ASE5 registraba esas dos filas (R506/R509) → hueco de PLANTILLA (no se toca).
/// E5: los 3 períodos exponen 3 formas de sección (subsecuencias del contrato oficial):
///     <c>[P,O,A,O,A,O]</c> (mayoría), <c>[P,A,O,A,O]</c> (ASE5-julio2, ASE1-Q1), <c>[P,O,O,O]</c> (sin Aplicación).
/// E6: Q1-ASE2 trae el canal <c>D=5</c> y NUEVO ESQUEMA en la 1.ª (sin slot en el oficial) — se CITA, no se resuelve.
/// E7: micro-punto F654/F661: el OFICIAL los trae en 0 sin <c>&lt;f&gt;</c>; la FUENTE emite filas
///     <c>D=H/Total</c> y <c>Subsidio</c> una sola vez → el duplicado es artefacto de ESCRITURA (cursor físico).
///
/// Contrato de sección PERIOD-FREE derivado (regla sin ramas, geometría pura): la k-ésima sección
/// NO-frontera de la fuente se alinea con la próxima sección del OFICIAL del MISMO TIPO aún no
/// consumida (emparejamiento monotónico/tipo), lo que reproduce el mapeo del plan §2.2:
/// <c>[P,A,O,A,O] → [1.ª, Aplic#1, Oportuno#2, Aplic#2, Oportuno#3]</c> y <c>[P,O,O,O] → [1.ª, Oportuno#1, Oportuno#2, Oportuno#3]</c>.
///
/// <para>
/// <b>ERRATA Plan 46 (W-2, 2026-10-17) — doctrina SUPERSEDIDA en el tipado de sección:</b> este T0
/// tipa por la frontera inmediatamente ARRIBA (la 1.ª sección es siempre <c>Primera</c>); la regla
/// VIGENTE tipa por la frontera que CIERRA (<c>ContratoSeccionPeriodFreeR1.SeccionPorCierre</c>), según
/// <c>Docs/T0-Plan45-Contrato-Seccion-PeriodFree.md §12</c> y el fixture
/// <see cref="ContratoSeccionPeriodFreeT0Plan46Tests"/>. Las expectativas con <c>Tipo.Primera</c> de
/// <c>T0_ReglaPeriodFree_FormasSonSubsecuencias_YEmparejanMonotonicoPorTipo</c> se conservan como
/// congelamiento HISTÓRICO de la aritmética de <c>Alinear</c>; el contrato actual está en el fixture del 46.
/// </para>
/// </summary>
public sealed class ContratoSeccionPeriodFreeT0Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string HojaR1 = "Reporte Componentes R1";
    private const string Sha256Oficial = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    // Tipo de frontera observado en la columna B de la fila de frontera.
    public enum Frontera { Mes, Aplicacion }

    // Tipo de sección del bloque (contexto de la frontera inmediatamente arriba).
    public enum Tipo { Primera, Oportuno, Aplicacion }

    // ── Períodos del banco de pruebas (rutas explícitas; prohibido Docs/Insumos raíz) ────────────
    private const string PeriodoJulio2 = "Prueba Julio-2";
    private const string PeriodoAgosto = "Prueba Agosto-2";
    private const string PeriodoQ1 = "Prueba Julio-1";

    // ── E1: fronteras del OFICIAL por ASE (fila, tipo) — congeladas desde bytes ───────────────────
    private static readonly (int Ase, int Fila, Frontera Tipo)[] FronterasOficial =
    [
        (1, 16, Frontera.Mes), (1, 32, Frontera.Aplicacion), (1, 47, Frontera.Mes), (1, 59, Frontera.Aplicacion), (1, 70, Frontera.Mes),
        (2, 120, Frontera.Mes), (2, 146, Frontera.Aplicacion), (2, 171, Frontera.Mes), (2, 190, Frontera.Aplicacion), (2, 208, Frontera.Mes),
        (3, 299, Frontera.Mes), (3, 315, Frontera.Aplicacion), (3, 330, Frontera.Mes), (3, 342, Frontera.Aplicacion), (3, 353, Frontera.Mes),
        (4, 460, Frontera.Mes), (4, 486, Frontera.Aplicacion), (4, 511, Frontera.Mes), (4, 530, Frontera.Aplicacion), (4, 548, Frontera.Mes),
        (5, 613, Frontera.Mes), (5, 629, Frontera.Aplicacion), (5, 648, Frontera.Mes), (5, 660, Frontera.Aplicacion), (5, 674, Frontera.Mes),
    ];

    // ── E1: 1.ª sección del OFICIAL (inicio/fin inclusive) ───────────────────────────────────────
    private static readonly (int Ase, int Inicio, int Fin)[] PrimeraSeccionOficial =
    [
        (1, 6, 15), (2, 110, 119), (3, 289, 298), (4, 450, 459), (5, 603, 612),
    ];

    // ── E5: forma de sección esperada por (período, ASE) — fronteras de la FUENTE (fila, tipo) ────
    // El tipo de la 1.ª sección es siempre "Primera"; las demás derivan de la frontera que las abre.
    private static readonly (string Periodo, int Ase, Frontera[] Fronteras)[] FormasFuente =
    [
        // Julio-2 (2026072): ASE1–4 forma completa; ASE5 invertida (arranca con Aplicación).
        (PeriodoJulio2, 1, [Frontera.Mes, Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        (PeriodoJulio2, 2, [Frontera.Mes, Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        (PeriodoJulio2, 3, [Frontera.Mes, Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        (PeriodoJulio2, 4, [Frontera.Mes, Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        (PeriodoJulio2, 5, [Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        // Agosto (2026082): ASE1/2/4/5 forma completa; ASE3 sin Aplicación (3 Mes).
        (PeriodoAgosto, 1, [Frontera.Mes, Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        (PeriodoAgosto, 2, [Frontera.Mes, Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        (PeriodoAgosto, 3, [Frontera.Mes, Frontera.Mes, Frontera.Mes]),
        (PeriodoAgosto, 4, [Frontera.Mes, Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        (PeriodoAgosto, 5, [Frontera.Mes, Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        // Q1 (2026071, SOLO LECTURA — citado sin resolver): ASE1 invertida; ASE2/ASE5 sin Aplicación.
        (PeriodoQ1, 1, [Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        (PeriodoQ1, 2, [Frontera.Mes, Frontera.Mes, Frontera.Mes]),
        (PeriodoQ1, 3, [Frontera.Mes, Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        (PeriodoQ1, 4, [Frontera.Mes, Frontera.Aplicacion, Frontera.Mes, Frontera.Aplicacion, Frontera.Mes]),
        (PeriodoQ1, 5, [Frontera.Mes, Frontera.Mes, Frontera.Mes]),
    ];

    // ── E1: SHA del oficial intacto ──────────────────────────────────────────────────────────────

    [Fact]
    public void T0_Sha256_OficialIntacto()
    {
        using var stream = File.OpenRead(Oficial());
        Assert.Equal(Sha256Oficial, Convert.ToHexString(SHA256.HashData(stream)));
    }

    // ── S1.2: el OFICIAL declara SIEMPRE 5 fronteras [Mes,Aplic,Mes,Aplic,Mes] (6 secciones) ──────

    [Fact]
    public void T0_Oficial_FronterasUniformes_5PorAse()
    {
        using var zip = ZipFile.OpenRead(Oficial());
        var hoja = Hoja(zip, HojaR1);
        var ss = Compartidas(zip);
        foreach (var ase in new[] { 1, 2, 3, 4, 5 })
        {
            var esperadas = FronterasOficial.Where(f => f.Ase == ase).ToArray();
            var observadas = LeerFronteras(hoja, ss, InicioBloqueOficial(ase), FinBloqueOficial(ase));
            Assert.Equal(esperadas.Length, observadas.Count);
            for (var i = 0; i < esperadas.Length; i++)
            {
                Assert.Equal(esperadas[i].Fila, observadas[i].Fila);
                Assert.Equal(esperadas[i].Tipo, observadas[i].Tipo);
            }
        }
    }

    // ── S1.2: la forma de fronteras de la FUENTE por (período, ASE) — base de la regla PERIOD-FREE ─

    [Theory]
    [MemberData(nameof(FormasFuenteData))]
    public void T0_Fuente_FormaDeFronteras_PorPeriodoAse(string periodo, int aseId, Frontera[] esperadas)
    {
        using var zip = ZipFile.OpenRead(FuenteR1(periodo, aseId));
        var hoja = Hoja(zip, "Sheet1");
        var ss = Compartidas(zip);
        var observadas = LeerFronteras(hoja, ss, 1, 400);
        Assert.Equal(esperadas, observadas.Select(o => o.Tipo).ToArray());
    }

    public static TheoryData<string, int, Frontera[]> FormasFuenteData()
    {
        var data = new TheoryData<string, int, Frontera[]>();
        foreach (var (periodo, ase, fronteras) in FormasFuente)
        {
            data.Add(periodo, ase, fronteras);
        }

        return data;
    }

    // ── S1.4 (regla PERIOD-FREE): las 3 formas son subsecuencias del contrato oficial y el
    //    emparejamiento monotónico por tipo produce el mapeo del plan §2.2 ─────────────────────────

    [Fact]
    public void T0_ReglaPeriodFree_FormasSonSubsecuencias_YEmparejanMonotonicoPorTipo()
    {
        // Contrato del OFICIAL: 6 secciones cuyos tipos son [P,O,A,O,A,O].
        var tiposOficial = TiposSeccion(FronterasOficialDeBloque(5)).ToArray();
        Assert.Equal(
            [Tipo.Primera, Tipo.Oportuno, Tipo.Aplicacion, Tipo.Oportuno, Tipo.Aplicacion, Tipo.Oportuno],
            tiposOficial);

        // Forma invertida de ASE5-julio2 [P,A,O,A,O]: empareja con [1.ª, Aplic#1, Oportuno#2, Aplic#2, Oportuno#3]
        // — exactamente el mapeo del plan §2.2 (Aplic#1→Aplic#1 oficial 630-647; Componente/Subsidio→Oportuno#2 649-659).
        var inversa = EmparejarMonotonicoPorTipo(
            [Tipo.Primera, Tipo.Aplicacion, Tipo.Oportuno, Tipo.Aplicacion, Tipo.Oportuno],
            tiposOficial);
        Assert.Equal([0, 2, 3, 4, 5], inversa);

        // Forma sin Aplicación [P,O,O,O] (agosto-ASE3 / Q1-ASE2/5): empareja con [1.ª, Oportuno#1, Oportuno#2, Oportuno#3].
        var sinAplic = EmparejarMonotonicoPorTipo(
            [Tipo.Primera, Tipo.Oportuno, Tipo.Oportuno, Tipo.Oportuno],
            tiposOficial);
        Assert.Equal([0, 1, 3, 5], sinAplic);

        // Forma completa [P,O,A,O,A,O]: identidad.
        var completa = EmparejarMonotonicoPorTipo(tiposOficial, tiposOficial);
        Assert.Equal([0, 1, 2, 3, 4, 5], completa);
    }

    // ── E2/E3/E4: hueco de la 1.ª sección (NO se toca la plantilla; se reporta al dueño) ─────────

    [Fact]
    public void T0_Oficial_PrimeraSeccion_SinVlrIntereses_CincoAse()
    {
        using var zip = ZipFile.OpenRead(Oficial());
        var hoja = Hoja(zip, HojaR1);
        var ss = Compartidas(zip);
        foreach (var (ase, inicio, fin) in PrimeraSeccionOficial)
        {
            var etiquetas = new List<string>();
            for (var fila = inicio; fila <= fin; fila++)
            {
                etiquetas.Add(TextoCelda(hoja, ss, $"E{fila}"));
            }

            Assert.True(!etiquetas.Any(e => string.Equals(e, "Vlr Intereses", StringComparison.OrdinalIgnoreCase)), $"ASE {ase}: la 1.ª sección trae 'Vlr Intereses'.");
            Assert.True(etiquetas.Any(e => string.Equals(e, "Vlr Servicio", StringComparison.OrdinalIgnoreCase)), $"ASE {ase}: la 1.ª sección no trae 'Vlr Servicio'.");
        }
    }

    [Fact]
    public void T0_FuenteJulio2Ase5_PrimeraSeccion_TraeVlrIntereses_ConDesglose()
    {
        using var zip = ZipFile.OpenRead(FuenteR1(PeriodoJulio2, 5));
        var hoja = Hoja(zip, "Sheet1");
        var ss = Compartidas(zip);
        Assert.Equal("Vlr Intereses", TextoCelda(hoja, ss, "E8"));
        Assert.Equal("402.4", Celda(hoja, "F8")?.Element(Ns + "v")?.Value);
        Assert.Equal("32.54", Celda(hoja, "G8")?.Element(Ns + "v")?.Value);
        Assert.Equal("Vlr Intereses", TextoCelda(hoja, ss, "E11"));
        Assert.Equal("7697.23", Celda(hoja, "F11")?.Element(Ns + "v")?.Value);
    }

    [Fact]
    public void T0_FuenteJulio2Ase1_Ase2_Ase3_Ase4_PrimeraSeccion_SinVlrIntereses()
    {
        // Solo ASE5-julio2 trae intereses en la 1.ª (hueco por ASE, no generalizable — T0 §4 del padre).
        foreach (var aseId in new[] { 1, 2, 3, 4 })
        {
            using var zip = ZipFile.OpenRead(FuenteR1(PeriodoJulio2, aseId));
            var hoja = Hoja(zip, "Sheet1");
            var ss = Compartidas(zip);
            var primeraFrontera = LeerFronteras(hoja, ss, 1, 400).First().Fila;
            for (var fila = 5; fila < primeraFrontera; fila++)
            {
                Assert.NotEqual("Vlr Intereses", TextoCelda(hoja, ss, $"E{fila}"));
            }
        }
    }

    [Fact]
    public void T0_ManualLegacyAse5_SiRegistrabaPrimeraSeccionIntereses()
    {
        using var zip = ZipFile.OpenRead(Manual());
        var hoja = Hoja(zip, HojaR1);
        var ss = Compartidas(zip);
        Assert.Equal("Vlr Intereses", TextoCelda(hoja, ss, "E506"));
        Assert.Equal("402.4", Celda(hoja, "F506")?.Element(Ns + "v")?.Value);
        Assert.Equal("Vlr Intereses", TextoCelda(hoja, ss, "E509"));
        Assert.Equal("7697.23", Celda(hoja, "F509")?.Element(Ns + "v")?.Value);
    }

    // ── E6: Q1-ASE2 (canal D=5 / NUEVO ESQUEMA en la 1.ª) — CITA read-only, no se resuelve ───────

    [Fact]
    public void T0_Q1Ase2_PrimeraSeccion_TraeCanal5YNuevoEsquema()
    {
        using var zip = ZipFile.OpenRead(FuenteR1(PeriodoQ1, 2));
        var hoja = Hoja(zip, "Sheet1");
        var ss = Compartidas(zip);
        var primeraFrontera = LeerFronteras(hoja, ss, 1, 400).First().Fila;
        var canales = new List<string>();
        var empresas = new List<string>();
        for (var fila = 5; fila < primeraFrontera; fila++)
        {
            canales.Add(TextoCelda(hoja, ss, $"D{fila}"));
            empresas.Add(TextoCelda(hoja, ss, $"C{fila}"));
        }

        Assert.Contains("5", canales);
        Assert.Contains("NUEVO ESQUEMA", empresas);
    }

    // ── E7: micro-punto F654/F661 (aislamiento lector vs escritor) ───────────────────────────────

    [Fact]
    public void T0_MicroPunto_F654_F661_EsArtefactoDeEscritura()
    {
        using (var zip = ZipFile.OpenRead(Oficial()))
        {
            var hoja = Hoja(zip, HojaR1);
            var ss = Compartidas(zip);
            Assert.Equal("H", TextoCelda(hoja, ss, "D654"));
            Assert.Equal("Total", TextoCelda(hoja, ss, "E654"));
            Assert.Equal("0", Celda(hoja, "F654")?.Element(Ns + "v")?.Value);
            Assert.Equal("Subsidio(-)/Contribucion(+)", TextoCelda(hoja, ss, "E661"));
            Assert.Equal("0", Celda(hoja, "F661")?.Element(Ns + "v")?.Value);
        }

        using (var zip = ZipFile.OpenRead(FuenteR1(PeriodoJulio2, 5)))
        {
            var hoja = Hoja(zip, "Sheet1");
            var ss = Compartidas(zip);
            Assert.Equal("Subsidio(-)/Contribucion(+)", TextoCelda(hoja, ss, "E47"));
            Assert.Equal("-187855.98", Celda(hoja, "F47")?.Element(Ns + "v")?.Value);
            Assert.Equal("H", TextoCelda(hoja, ss, "D48"));
            Assert.Equal("Total", TextoCelda(hoja, ss, "E48"));
            Assert.Equal("-187855.98", Celda(hoja, "F48")?.Element(Ns + "v")?.Value);
        }
    }

    // ── Utilidades ──────────────────────────────────────────────────────────────────────────────

    private static IEnumerable<Tipo> TiposSeccion(IReadOnlyList<Frontera> fronteras)
    {
        yield return Tipo.Primera;
        foreach (var f in fronteras)
        {
            yield return f == Frontera.Mes ? Tipo.Oportuno : Tipo.Aplicacion;
        }
    }

    /// <summary>
    /// Emparejamiento monotónico por tipo: la i-ésima sección fuente se alinea con la próxima
    /// sección del oficial del MISMO tipo aún no consumida (greedy, sin ramas por período/empresa).
    /// Devuelve el índice oficial (0-based) por cada sección fuente.
    /// </summary>
    private static int[] EmparejarMonotonicoPorTipo(IReadOnlyList<Tipo> fuente, IReadOnlyList<Tipo> oficial)
    {
        var resultado = new int[fuente.Count];
        var cursor = 0;
        for (var i = 0; i < fuente.Count; i++)
        {
            var encontrado = -1;
            for (var j = cursor; j < oficial.Count; j++)
            {
                if (oficial[j] == fuente[i])
                {
                    encontrado = j;
                    break;
                }
            }

            Assert.True(encontrado >= 0, $"La sección fuente #{i} ({fuente[i]}) no tiene contraparte oficial del mismo tipo tras el cursor {cursor}.");
            resultado[i] = encontrado;
            cursor = encontrado + 1;
        }

        return resultado;
    }

    private static IReadOnlyList<Frontera> FronterasOficialDeBloque(int ase) =>
        FronterasOficial.Where(f => f.Ase == ase).Select(f => f.Tipo).ToList();

    private static int InicioBloqueOficial(int ase) => ase switch
    {
        1 => 5,
        2 => 109,
        3 => 288,
        4 => 449,
        5 => 602,
        _ => throw new ArgumentOutOfRangeException(nameof(ase)),
    };

    private static int FinBloqueOficial(int ase) => ase switch
    {
        1 => 72,
        2 => 210,
        3 => 355,
        4 => 550,
        5 => 676,
        _ => throw new ArgumentOutOfRangeException(nameof(ase)),
    };

    private static List<(int Fila, Frontera Tipo)> LeerFronteras(XDocument hoja, List<string> compartidas, int inicio, int fin)
    {
        var fronteras = new List<(int, Frontera)>();
        for (var fila = inicio; fila <= fin; fila++)
        {
            var b = TextoCelda(hoja, compartidas, $"B{fila}");
            if (string.Equals(b, "Mes", StringComparison.OrdinalIgnoreCase))
            {
                fronteras.Add((fila, Frontera.Mes));
            }
            else if (b.Contains("Aplicacion nuevos x reversion", StringComparison.OrdinalIgnoreCase))
            {
                fronteras.Add((fila, Frontera.Aplicacion));
            }
        }

        return fronteras;
    }

    private static string Raiz()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("No se encontró la raíz del repositorio (plantilla estandarizada única).");
    }

    private static string Oficial() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string Manual() =>
        Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Resultado", "Remuneracion 202607-2 Total Administrativo.xlsx");

    private static string FuenteR1(string periodo, int aseId)
    {
        var carpeta = Path.Combine(Raiz(), "Docs", periodo, "Insumos", CarpetasAse[aseId - 1]);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta explícita: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("Recaudoporcomponente", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Recaudoporcomponente del ASE{aseId} en {carpeta}");
        return archivo!;
    }

    private static XDocument LeerXml(ZipArchive zip, string entrada)
    {
        var entry = zip.GetEntry(entrada) ?? throw new InvalidOperationException($"El xlsx no trae '{entrada}'.");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static XDocument Hoja(ZipArchive zip, string nombre)
    {
        var libro = LeerXml(zip, "xl/workbook.xml");
        var sheet = libro.Root?.Element(Ns + "sheets")?.Elements(Ns + "sheet")
            .FirstOrDefault(s => string.Equals(s.Attribute("name")?.Value, nombre, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"El workbook no trae la hoja '{nombre}'.");
        var id = sheet.Attribute(RelNs + "id")!.Value;
        var rels = LeerXml(zip, "xl/_rels/workbook.xml.rels");
        var destino = rels.Root?.Elements()
            .FirstOrDefault(r => r.Attribute("Id")?.Value == id)?.Attribute("Target")?.Value
            ?? throw new InvalidOperationException($"Sin rel para la hoja '{nombre}'.");
        if (!destino.StartsWith("xl/", StringComparison.Ordinal))
        {
            destino = "xl/" + destino.TrimStart('/');
        }

        return LeerXml(zip, destino);
    }

    private static List<string> Compartidas(ZipArchive zip)
    {
        var doc = LeerXml(zip, "xl/sharedStrings.xml");
        return doc.Root?.Elements(Ns + "si").Select(si => si.Value).ToList() ?? [];
    }

    private static XElement? Fila(XDocument hoja, string numero) =>
        hoja.Root?.Element(Ns + "sheetData")?.Elements(Ns + "row")
            .FirstOrDefault(r => r.Attribute("r")?.Value == numero);

    private static XElement? Celda(XDocument hoja, string referencia)
    {
        var numero = new string(referencia.Where(char.IsAsciiDigit).ToArray());
        return Fila(hoja, numero)?.Elements(Ns + "c")
            .FirstOrDefault(c => string.Equals(c.Attribute("r")?.Value, referencia, StringComparison.OrdinalIgnoreCase));
    }

    private static string TextoCelda(XDocument hoja, List<string> compartidas, string referencia)
    {
        var celda = Celda(hoja, referencia);
        if (celda is null)
        {
            return string.Empty;
        }

        if (celda.Attribute("t")?.Value == "s" && int.TryParse(celda.Element(Ns + "v")?.Value, out var indice)
            && indice >= 0 && indice < compartidas.Count)
        {
            return compartidas[indice];
        }

        return celda.Element(Ns + "inlineStr")?.Value ?? celda.Element(Ns + "v")?.Value ?? string.Empty;
    }
}
