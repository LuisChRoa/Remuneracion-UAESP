using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 46 (H0, S1/S2 — §2.1/§2.2 del plan): T0 CORREGIDO, READ-ONLY, con CERO código de producción.
///
/// Re-deriva CONTRA BYTES (BCL <see cref="ZipArchive"/> + <see cref="XDocument"/>; sin Excel/COM,
/// sin OpenXml SDK, sin ExcelDataReader) la MATRIZ corregida por ASE × período bajo las reglas del
/// Plan 46:
/// <list type="bullet">
/// <item>§2.2 — cada sección (fuente y oficial) se tipa por la frontera que la <b>CIERRA</b>
/// (inmediatamente debajo: su fila resumen <c>Mes</c>/<c>Aplicacion…</c>): <c>Mes</c>→Oportuno,
/// <c>Aplicacion</c>→Aplicación, ordinal = ocurrencia de esa frontera; la cola sin frontera debajo se
/// queda con la de arriba; AFaseo → sin sección;</item>
/// <item>§2.1 — CANAL EFECTIVO de una fila sin <c>D</c> propio (y sin <c>A</c>/<c>B</c>): el <c>D</c>
/// de la próxima fila <c>…/Total</c> hacia abajo con <c>D</c> no vacío y ≠ literal <c>"Total"</c>.</item>
/// </list>
/// Cada celda de la matriz se cita con fila-fuente + celda-slot + sección-contrato + canal efectivo.
/// El OFICIAL (plantilla única) nunca se modifica (SHA256 antes/después) y Q1 se lee READ-ONLY.
///
/// Este fixture NO usa <c>EscritorValoresSlotR1</c> ni <c>ContratoSeccionPeriodFreeR1</c>: la
/// derivación es INDEPENDIENTE (espejo local del algoritmo) y su matriz congelada es la evidencia del
/// contrato corregido. La palanca H0→H1 exige: fixture verde + matriz citada.
/// </summary>
public sealed class ContratoSeccionPeriodFreeT0Plan46Tests
{
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string Sha256Oficial = "3086AC317595C2F8E84ED24B1C5EE3E52565DD5193394E3062878E7FA49C5214";

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    private sealed class Fila
    {
        public int Num { get; init; }
        public string A { get; init; } = "";
        public string B { get; init; } = "";
        public string C { get; init; } = "";
        public string D { get; init; } = "";
        public string E { get; init; } = "";
        public decimal? F { get; init; }
        public bool Consumida { get; set; }
    }

    private sealed record Sec(string Tipo, int Ord)
    {
        public override string ToString() => $"{Tipo}#{Ord}";
    }

    // ── S: oficial — secciones por CIERRE uniformes [O#1,A#1,O#2,A#2,O#3] ─────────────────────────

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void T0P46_Oficial_SeccionesPorCierre_Uniformes(int ase)
    {
        var oficial = OficialBloque(ase);
        var secciones = Distintas(oficial.Select((_, i) => SeccionCierre(oficial, i)));

        Assert.Equal(
            ["Oportuno#1", "Aplicacion#1", "Oportuno#2", "Aplicacion#2", "Oportuno#3"],
            secciones.Select(s => s.ToString()).ToArray());
    }

    // ── S: fuente — secciones por CIERRE por (período, ASE) → las 3 formas probadas ──────────────

    public static TheoryData<string, int, string[]> FormasFuenteData()
    {
        var data = new TheoryData<string, int, string[]>();
        var completa = new[] { "Oportuno#1", "Aplicacion#1", "Oportuno#2", "Aplicacion#2", "Oportuno#3" };
        var invertida = new[] { "Aplicacion#1", "Oportuno#1", "Aplicacion#2", "Oportuno#2" };
        var sinAplic = new[] { "Oportuno#1", "Oportuno#2", "Oportuno#3" };

        data.Add("Prueba Julio-2", 1, completa);
        data.Add("Prueba Julio-2", 2, completa);
        data.Add("Prueba Julio-2", 3, completa);
        data.Add("Prueba Julio-2", 4, completa);
        data.Add("Prueba Julio-2", 5, invertida);
        data.Add("Prueba Agosto-2", 1, completa);
        data.Add("Prueba Agosto-2", 2, completa);
        data.Add("Prueba Agosto-2", 3, sinAplic);
        data.Add("Prueba Agosto-2", 4, completa);
        data.Add("Prueba Agosto-2", 5, completa);
        data.Add("Prueba Julio-1", 1, invertida);
        data.Add("Prueba Julio-1", 2, sinAplic);
        data.Add("Prueba Julio-1", 3, completa);
        data.Add("Prueba Julio-1", 4, completa);
        data.Add("Prueba Julio-1", 5, sinAplic);
        return data;
    }

    [Theory]
    [MemberData(nameof(FormasFuenteData))]
    public void T0P46_Fuente_SeccionesPorCierre(string periodo, int ase, string[] esperadas)
    {
        var fuente = FuenteBloque(periodo, ase);
        var secciones = Distintas(fuente.Select((_, i) => SeccionCierre(fuente, i)));

        Assert.Equal(esperadas, secciones.Select(s => s.ToString()).ToArray());
    }

    // ── S: el MAPEO monotónico por tipo de las 3 formas (regla sin ramas) ─────────────────────────

    [Fact]
    public void T0P46_Mapeo_MonotonicoPorTipo_TresFormas()
    {
        var oficial = new[] { "Oportuno#1", "Aplicacion#1", "Oportuno#2", "Aplicacion#2", "Oportuno#3" }
            .Select(Parsear).ToList();

        // [M,A,M,A,M] → identidad.
        AssertMapeo(
            ["Oportuno#1", "Aplicacion#1", "Oportuno#2", "Aplicacion#2", "Oportuno#3"],
            ["Oportuno#1", "Aplicacion#1", "Oportuno#2", "Aplicacion#2", "Oportuno#3"],
            oficial);

        // [A,M,A,M] (ASE5-julio2, ASE1-Q1): A#1→A#1, O#1→O#2, A#2→A#2, O#2→O#3 (salta el O#1 oficial).
        AssertMapeo(
            ["Aplicacion#1", "Oportuno#1", "Aplicacion#2", "Oportuno#2"],
            ["Aplicacion#1", "Oportuno#2", "Aplicacion#2", "Oportuno#3"],
            oficial);

        // [M,M,M] (agosto-ASE3, Q1-ASE2/5): O#1→O#1, O#2→O#2, O#3→O#3 (salta A#1/A#2 oficiales).
        AssertMapeo(
            ["Oportuno#1", "Oportuno#2", "Oportuno#3"],
            ["Oportuno#1", "Oportuno#2", "Oportuno#3"],
            oficial);

        void AssertMapeo(string[] forma, string[] esperado, IReadOnlyList<Sec> ofi)
        {
            var mapa = Alinear(forma.Select(Parsear).ToList(), ofi);
            var obtenido = forma.Select(f => mapa.TryGetValue(Parsear(f), out var d) ? d.ToString() : "SIN-CONTRAPARTE").ToArray();
            Assert.Equal(esperado, obtenido);
        }
    }

    // ── S1.1: matriz citada ASE1-julio2 (fila-fuente → celda-slot) ────────────────────────────────

    [Fact]
    public void T0P46_Julio2_Ase1_MatrizCitada()
    {
        // Fuente: Docs/Prueba Julio-2/Insumos/1-Promoambiental/Recaudoporcomponente_*.xlsx.
        // Slot:   Docs/Plantilla estandarizada/... 'Reporte Componentes R1' filas 6..72.
        var plan = Planificar("Prueba Julio-2", 1);
        AssertPlan(plan,
        [
            (5, 6), (6, 7), (7, 8), (8, 11), (9, 12), (10, 15), (11, 16), (12, 17),
            (13, 18), (14, 20), (15, 21), (16, 32), (17, 33), (18, 34), (19, 35), (20, 36),
            (21, 37), (22, 38), (23, 39), (24, 40), (25, 41), (26, 42), (27, 43), (28, 44),
            (29, 45), (30, 46), (31, 47), (32, 48), (33, 49), (34, 50), (35, 51), (36, 59),
            (37, 60), (38, 61), (39, 62), (40, 63), (41, 64), (42, 65), (43, 66), (44, 67),
            (45, 68), (46, 69), (47, 70), (48, 71), (49, 72),
        ]);

        // Cita de valor: la fila-fuente R8 (Vlr Servicio, canal O) cae en F11 (no en F9): su
        // predecesora R5 (canal E) va a F6. F9/F10/F13/F14 quedan en cero.
        var fuente = FuenteBloque("Prueba Julio-2", 1);
        Assert.Equal("O", CanalEfectivo(fuente, Indice(fuente, 8)));
        Assert.Equal("E", CanalEfectivo(fuente, Indice(fuente, 5)));
        Assert.Equal(-4108.55m, fuente[Indice(fuente, 8)].F);
    }

    // ── S1.2/S1.3: matriz citada ASE5-julio2 (fila-fuente → celda-slot) ───────────────────────────

    [Fact]
    public void T0P46_Julio2_Ase5_MatrizCitada()
    {
        var plan = Planificar("Prueba Julio-2", 5);
        AssertPlan(plan,
        [
            (5, 619), (6, 621), (7, 622), (8, 623), (9, 624), (10, 625), (11, 626), (12, 627),
            (13, 628), (14, 629), (15, 630), (16, 631), (17, 632), (18, 633), (19, 634), (20, 635),
            (21, 636), (22, 637), (23, 638), (24, 639), (25, 640), (26, 641), (27, 642), (28, 643),
            (29, 644), (30, 645), (31, 646), (32, 647), (33, 648), (34, 649), (35, 653), (36, 654),
            (37, 655), (38, 656), (39, 659), (40, 660), (41, 661), (42, 662), (43, 663), (44, 664),
            (45, 665), (46, 666), (47, 667), (48, 668), (49, 669), (50, 670), (51, 671), (52, 672),
            (53, 673), (54, 674), (55, 675), (56, 676),
        ]);

        // Citas de valor (bytes de la fuente ASE5-julio2): R5 canal H 52.706,16 → F619; R8/Vlr
        // Intereses 402,40 canal O → F623; R11 7.697,23 canal T → F626; R13 OCCIDENTE 417.519,62
        // → F628 (NO F612); R14 Aplicacion frontier → F629; R33 Mes 13.084.233.937,44 → F648;
        // R54 Mes −1.036.320.522,40 → F674.
        var fuente = FuenteBloque("Prueba Julio-2", 5);
        Assert.Equal("H", CanalEfectivo(fuente, Indice(fuente, 5)));
        Assert.Equal("O", CanalEfectivo(fuente, Indice(fuente, 8)));
        Assert.Equal("T", CanalEfectivo(fuente, Indice(fuente, 11)));
        Assert.Equal(52706.16m, fuente[Indice(fuente, 5)].F);
        Assert.Equal(402.4m, fuente[Indice(fuente, 8)].F);
        Assert.Equal(7697.23m, fuente[Indice(fuente, 11)].F);
        Assert.Equal(417519.62m, fuente[Indice(fuente, 13)].F);
        Assert.Equal(13084233937.44m, fuente[Indice(fuente, 33)].F);
        Assert.Equal(-1036320522.40m, fuente[Indice(fuente, 54)].F);
        // La 1.ª sección OFICIAL (603-612) no tiene contraparte fuente en ASE5-julio2 → queda en cero.
        Assert.DoesNotContain(plan, p => p.Slot is not null && p.Slot.Num is >= 603 and <= 613);
        // 614 (AFaseo oficial) tampoco recibe (la fuente ASE5-julio2 no trae AFaseo).
        Assert.DoesNotContain(plan, p => p.Slot is not null && p.Slot.Num == 614);
    }

    // ── S: SIN-SLOT documentados por (período, ASE) — el resto coloca todo ────────────────────────

    [Fact]
    public void T0P46_SinSlot_Documentados()
    {
        // agosto-ASE3 y Q1-ASE2/5 traen la forma [M,M,M]: su fila `Componente` vive en una sección
        // cerrada por `Mes` (Oportuno#3) mientras la fila `Componente` del OFICIAL vive en la sección
        // cerrada por `Aplicacion` (Aplicacion#2) → esa fila NO tiene slot por el contrato (el bloque
        // NO es colocable → el escritor reproduce el comportamiento histórico del Plan 41). Q1-ASE2
        // además trae el canal D=5/NUEVO ESQUEMA en la 1.ª (bloqueo Q1 declarado, read-only).
        var esperados = new Dictionary<(string, int), int>
        {
            [("Prueba Agosto-2", 3)] = 1,
            [("Prueba Julio-1", 2)] = 4,
            [("Prueba Julio-1", 5)] = 1,
        };

        foreach (var (periodo, ase, _) in FormasFuenteData().Select(d => ((string)d[0]!, (int)d[1]!, 0)))
        {
            var plan = Planificar(periodo, ase);
            var sinSlot = plan.Count(p => p.Slot is null);
            var esperado = esperados.GetValueOrDefault((periodo, ase), 0);
            Assert.True(sinSlot == esperado, $"{periodo} ASE{ase}: SIN-SLOT={sinSlot}, esperado={esperado}.");
        }
    }

    // ── S: SHA del oficial + Q1 read-only (cita) ─────────────────────────────────────────────────

    [Fact]
    public void T0P46_Sha256_OficialIntacto()
    {
        using var stream = File.OpenRead(Oficial());
        Assert.Equal(Sha256Oficial, Convert.ToHexString(SHA256.HashData(stream)));
    }

    [Fact]
    public void T0P46_Q1Ase2_Canal5NuevoEsquema_ReadOnly()
    {
        // Q1 solo se LEE: la 1.ª sección de ASE2 trae el canal 5 y NUEVO ESQUEMA (sin slot) — se cita,
        // no se resuelve (RESTRICCIÓN DURA del plan §0.2).
        var fuente = FuenteBloque("Prueba Julio-1", 2);
        var canales = fuente.Where(f => f.Num >= 5 && f.D is "5" or "").Select(f => f.D).ToList();
        Assert.Contains("5", canales);
        Assert.Contains(fuente, f => string.Equals(f.C, "NUEVO ESQUEMA", StringComparison.OrdinalIgnoreCase));
    }

    // ── Derivación (espejo local del algoritmo §2.1/§2.2; SIN producción) ────────────────────────

    private sealed record Colocacion(Fila Fuente, Fila? Slot);

    private static List<Colocacion> Planificar(string periodo, int ase)
    {
        var fuente = FuenteBloque(periodo, ase);
        var oficial = OficialBloque(ase);
        var secF = fuente.Select((_, i) => SeccionCierre(fuente, i)).ToList();
        var secO = oficial.Select((_, i) => SeccionCierre(oficial, i)).ToList();
        var mapa = Alinear(Distintas(secF), Distintas(secO));

        var plan = new List<Colocacion>(fuente.Count);
        for (var i = 0; i < fuente.Count; i++)
        {
            var fila = fuente[i];
            var seccion = secF[i] is { } sf && mapa.TryGetValue(sf, out var dst) ? dst : secF[i];
            var canal = CanalEfectivo(fuente, i);
            var slot = BuscarSlot(oficial, secO, fila, seccion, canal);
            plan.Add(new Colocacion(fila, slot));
        }

        return plan;
    }

    private static void AssertPlan(List<Colocacion> plan, (int Fuente, int Slot)[] esperado)
    {
        Assert.Equal(esperado.Length, plan.Count);
        for (var i = 0; i < esperado.Length; i++)
        {
            Assert.Equal(esperado[i].Fuente, plan[i].Fuente.Num);
            Assert.NotNull(plan[i].Slot);
            Assert.Equal(esperado[i].Slot, plan[i].Slot!.Num);
        }
    }

    private static Fila? BuscarSlot(List<Fila> oficial, List<Sec?> secO, Fila fila, Sec? seccion, string? canal)
    {
        for (var j = 0; j < oficial.Count; j++)
        {
            var o = oficial[j];
            if (o.Consumida) continue;
            if (!IgualEtiqueta(fila.A, o.A) || !IgualEtiqueta(fila.B, o.B) || !IgualEtiqueta(fila.D, o.D) || !IgualEtiqueta(fila.E, o.E)) continue;
            if (!CCompatibles(fila.C, o.C)) continue;
            if (seccion is not null && secO[j] is not null && !seccion.Equals(secO[j])) continue;
            if (canal is not null)
            {
                var canalO = CanalEfectivo(oficial, j);
                if (!string.Equals(canal, canalO, StringComparison.OrdinalIgnoreCase)) continue;
            }

            o.Consumida = true;
            return o;
        }

        return null;
    }

    /// <summary>§2.2 — sección por la frontera que CIERRA (espejo local, sin producción).</summary>
    private static Sec? SeccionCierre(List<Fila> filas, int i)
    {
        var f = filas[i];
        if (EsAFaseo(f)) return null;
        if (EsMes(f)) return new Sec("Oportuno", CuentaTipoHasta(filas, i, mes: true));
        if (EsAplic(f)) return new Sec("Aplicacion", CuentaTipoHasta(filas, i, mes: false));

        for (var j = i + 1; j < filas.Count; j++)
        {
            if (EsFront(filas[j])) return SeccionCierre(filas, j);
        }

        return PorFronteraArriba(filas, i);
    }

    private static int CuentaTipoHasta(List<Fila> filas, int i, bool mes)
    {
        var n = 0;
        for (var j = 0; j <= i; j++)
        {
            if (mes ? EsMes(filas[j]) : EsAplic(filas[j])) n++;
        }

        return n;
    }

    private static Sec? PorFronteraArriba(List<Fila> filas, int i)
    {
        var mes = 0;
        var apl = 0;
        var tipo = "Primera";
        var ord = 0;
        for (var j = 0; j < i; j++)
        {
            if (EsMes(filas[j])) { mes++; tipo = "Oportuno"; ord = mes; }
            else if (EsAplic(filas[j])) { apl++; tipo = "Aplicacion"; ord = apl; }
        }

        return new Sec(tipo, ord);
    }

    /// <summary>§2.1 — canal efectivo (espejo local, sin producción).</summary>
    private static string? CanalEfectivo(List<Fila> filas, int i)
    {
        var f = filas[i];
        if (!string.IsNullOrWhiteSpace(f.D)) return null;
        if (!string.IsNullOrWhiteSpace(f.A) || !string.IsNullOrWhiteSpace(f.B)) return null;
        for (var j = i + 1; j < filas.Count; j++)
        {
            var d = filas[j].D;
            if (!string.IsNullOrWhiteSpace(d) && !IgualEtiqueta(d, "Total")) return d;
        }

        return null;
    }

    private static List<Sec> Distintas(IEnumerable<Sec?> secs)
    {
        var list = new List<Sec>();
        Sec? prev = null;
        foreach (var s in secs)
        {
            if (s is null) continue;
            if (prev is null || !prev.Equals(s)) list.Add(s);
            prev = s;
        }

        return list;
    }

    private static Dictionary<Sec, Sec> Alinear(IReadOnlyList<Sec> fuente, IReadOnlyList<Sec> oficial)
    {
        var mapa = new Dictionary<Sec, Sec>();
        var cursor = 0;
        foreach (var f in fuente)
        {
            var enc = -1;
            for (var j = cursor; j < oficial.Count; j++)
            {
                if (oficial[j].Tipo == f.Tipo) { enc = j; break; }
            }

            if (enc < 0) continue;
            mapa[f] = oficial[enc];
            cursor = enc + 1;
        }

        return mapa;
    }

    private static bool EsAFaseo(Fila f) =>
        (IgualEtiqueta(f.A, "AFaseo") && IgualEtiqueta(f.B, "Total")) || (IgualEtiqueta(f.D, "AFASEO") && IgualEtiqueta(f.E, "Total"));

    private static bool EsMes(Fila f) => IgualEtiqueta(f.B, "Mes") && IgualEtiqueta(f.C, "Total");
    private static bool EsAplic(Fila f) => f.B.Contains("Aplicacion nuevos x reversion", StringComparison.OrdinalIgnoreCase) && IgualEtiqueta(f.C, "Total");
    private static bool EsFront(Fila f) => EsMes(f) || EsAplic(f);
    private static bool IgualEtiqueta(string? x, string? y) => string.Equals(x?.Trim(), y?.Trim(), StringComparison.OrdinalIgnoreCase);

    private static bool CCompatibles(string? x, string? y)
    {
        if (string.IsNullOrWhiteSpace(x) && string.IsNullOrWhiteSpace(y)) return true;
        if (string.IsNullOrWhiteSpace(x) || string.IsNullOrWhiteSpace(y)) return false;
        if (IgualEtiqueta(x, y)) return true;
        static bool EsRec(string s) => IgualEtiqueta(s, "RECIPROCIDAD") || IgualEtiqueta(s, "NUEVO ESQUEMA");
        return EsRec(x) && EsRec(y);
    }

    private static Sec Parsear(string etiqueta)
    {
        var partes = etiqueta.Split('#');
        return new Sec(partes[0], int.Parse(partes[1], CultureInfo.InvariantCulture));
    }

    private static int Indice(List<Fila> filas, int numero) =>
        filas.FindIndex(f => f.Num == numero) is var idx && idx >= 0 ? idx : throw new InvalidOperationException($"R{numero} no está en el bloque.");

    // ── Lectura (BCL ZIP+XML read-only) ─────────────────────────────────────────────────────────

    private static List<Fila> FuenteBloque(string periodo, int ase) =>
        LeerSubBloque(FuenteR1(periodo, ase), "Sheet1", 5, 400);

    private static List<Fila> OficialBloque(int ase)
    {
        var (inicio, fin) = ase switch
        {
            1 => (6, 72), 2 => (110, 210), 3 => (289, 355), 4 => (450, 550), 5 => (603, 676),
            _ => throw new ArgumentOutOfRangeException(nameof(ase)),
        };
        return LeerSubBloque(Oficial(), "Reporte Componentes R1", inicio, fin);
    }

    private static List<Fila> LeerSubBloque(string ruta, string hoja, int desde, int hasta)
    {
        using var zip = ZipFile.OpenRead(ruta);
        var doc = Hoja(zip, hoja);
        var ss = Compartidas(zip);
        var filas = new List<Fila>();
        for (var n = desde; n <= hasta; n++)
        {
            if (FilaXml(doc, n.ToString(CultureInfo.InvariantCulture)) is null) continue;
            filas.Add(new Fila
            {
                Num = n,
                A = Texto(doc, ss, $"A{n}"),
                B = Texto(doc, ss, $"B{n}"),
                C = Texto(doc, ss, $"C{n}"),
                D = Texto(doc, ss, $"D{n}"),
                E = Texto(doc, ss, $"E{n}"),
                F = Numero(doc, $"F{n}"),
            });
        }

        var corte = filas.FindLastIndex(x => IgualEtiqueta(x.A, "Total") && string.IsNullOrWhiteSpace(x.B));
        return corte >= 0 ? filas.Take(corte + 1).ToList() : filas;
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

    private static string Oficial() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string FuenteR1(string periodo, int aseId)
    {
        var dir = Path.Combine(Raiz(), "Docs", periodo, "Insumos", CarpetasAse[aseId - 1]);
        Assert.True(Directory.Exists(dir), $"Falta la carpeta explícita: {dir}");
        var archivo = Directory.EnumerateFiles(dir, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("Recaudoporcomponente", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Recaudoporcomponente del ASE{aseId} en {dir}");
        return archivo!;
    }

    private static XDocument LeerXml(ZipArchive zip, string entrada)
    {
        var e = zip.GetEntry(entrada) ?? throw new InvalidOperationException($"El xlsx no trae '{entrada}'.");
        using var s = e.Open();
        return XDocument.Load(s);
    }

    private static XDocument Hoja(ZipArchive zip, string nombre)
    {
        var libro = LeerXml(zip, "xl/workbook.xml");
        var sheet = libro.Root!.Element(Ns + "sheets")!.Elements(Ns + "sheet")
            .First(s => string.Equals(s.Attribute("name")?.Value, nombre, StringComparison.Ordinal));
        var id = sheet.Attribute(RelNs + "id")!.Value;
        var rels = LeerXml(zip, "xl/_rels/workbook.xml.rels");
        var destino = rels.Root!.Elements().First(r => r.Attribute("Id")?.Value == id).Attribute("Target")!.Value;
        if (!destino.StartsWith("xl/", StringComparison.Ordinal)) destino = "xl/" + destino.TrimStart('/');
        return LeerXml(zip, destino);
    }

    private static List<string> Compartidas(ZipArchive zip)
    {
        var doc = LeerXml(zip, "xl/sharedStrings.xml");
        return doc.Root!.Elements(Ns + "si").Select(si => si.Value).ToList();
    }

    private static XElement? FilaXml(XDocument hoja, string n) =>
        hoja.Root!.Element(Ns + "sheetData")!.Elements(Ns + "row").FirstOrDefault(r => r.Attribute("r")?.Value == n);

    private static XElement? Celda(XDocument hoja, string referencia)
    {
        var n = new string(referencia.Where(char.IsAsciiDigit).ToArray());
        return FilaXml(hoja, n)?.Elements(Ns + "c").FirstOrDefault(c => string.Equals(c.Attribute("r")?.Value, referencia, StringComparison.OrdinalIgnoreCase));
    }

    private static string Texto(XDocument hoja, List<string> compartidas, string referencia)
    {
        var c = Celda(hoja, referencia);
        if (c is null) return string.Empty;
        if (c.Attribute("t")?.Value == "s" && int.TryParse(c.Element(Ns + "v")?.Value, out var idx) && idx >= 0 && idx < compartidas.Count)
            return compartidas[idx];
        return c.Element(Ns + "inlineStr")?.Value ?? c.Element(Ns + "v")?.Value ?? string.Empty;
    }

    private static decimal? Numero(XDocument hoja, string referencia)
    {
        var v = Celda(hoja, referencia)?.Element(Ns + "v")?.Value;
        return v is not null && decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : null;
    }
}
