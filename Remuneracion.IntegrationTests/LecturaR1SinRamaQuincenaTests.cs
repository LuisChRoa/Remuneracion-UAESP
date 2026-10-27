using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 54 (H2.1 / T4, D5): CANDADO anti-recaída de la doble resolución por quincena. Dos capas:
///
/// 1. ESTÁTICA (allow-list cerrada): el cuerpo de
///    <c>ExcelDataReaderWorkbookLeafInputReader.cs</c> no puede contener <c>NumeroQuincena</c> salvo
///    los usos declarados de D7 (fuentes/visibles), con su cita. El test lleva la lista como DATOS:
///    si el archivo se reordena, se re-ancla; si aparece un uso nuevo, enrojece.
/// 2. CONDUCTUAL (identidad): la MISMA fuente resuelta con <see cref="Periodo"/> Q1 y Q2 produce
///    <c>WorkbookLeafInputsR1</c> idénticos miembro a miembro. La quincena NO selecciona algoritmo.
///
/// La capa estática incluye su propia PRUEBA DE CANDADO: un texto sintético con
/// <c>NumeroQuincena</c> injertado en el mapeo debe producir exactamente una violación (sin tocar el
/// código de producción).
/// </summary>
public sealed class LecturaR1SinRamaQuincenaTests
{
    private const string ReaderRelativo = "Remuneracion.Infrastructure/Excel/ExcelDataReaderWorkbookLeafInputReader.cs";

    /// <summary>
    /// Allow-list CERRADA (D7). Cada entrada cita el uso legítimo que la justifica:
    /// - <c>var esQuincena2 = ...</c> ×1: <c>LeerRecaudosEmpresa</c> — par de columnas D/E (Q1) vs
    ///   F/G (Q2) del RESUMEN MES de la conciliación por empresa: declaración de visibles, no de
    ///   algoritmo. (Plan 59 §11-bis: la segunda ocurrencia —el gate de <c>DetalleR2/DetalleR4</c>
    ///   solo-Q2— quedó RETIRADA; el detalle es del período, no de la quincena, y su ausencia en Q1
    ///   dejaba el bloque de anticipos en cero.)
    /// - <c>Quincena = periodo.NumeroQuincena,</c> ×1: dato de PROCEDENCIA en <c>ReporteBancoInputs</c>
    ///   (no ramifica ni resuelve).
    /// </summary>
    private static readonly IReadOnlyDictionary<string, int> AllowList = new Dictionary<string, int>(StringComparer.Ordinal)
    {
        ["var esQuincena2 = periodo.NumeroQuincena == 2;"] = 1,
        ["Quincena = periodo.NumeroQuincena,"] = 1
    };

    // ── Capa 1: estática ────────────────────────────────────────────────────────────────────

    [Fact]
    public void T4_Estatica_ElReaderNoRamificaPorQuincena()
    {
        var fuente = File.ReadAllText(Ruta(ReaderRelativo));
        var violaciones = Violaciones(fuente);

        Assert.True(violaciones.Count == 0,
            "El lector leaf volvió a usar Periodo.NumeroQuincena fuera de la allow-list (D5/D7): " +
            string.Join(" | ", violaciones));

        Assert.Equal(2, Lineas(fuente).Select(l => l.Trim())
            .Count(l => !EsComentario(l) && l.Contains("NumeroQuincena", StringComparison.Ordinal)));
    }

    [Fact]
    public void T4_Estatica_EnrojeceAnteUnUsoInjertadoEnElMapeo()
    {
        var fuente = File.ReadAllText(Ruta(ReaderRelativo));
        var inyectada = fuente.Replace(
            "    private static WorkbookLeafInputsR1 MapearR1(Ase ase, List<object?[]> filas)",
            "    private static WorkbookLeafInputsR1 MapearR1(Ase ase, List<object?[]> filas)\n" +
            "    {\n        if (periodo.NumeroQuincena == 2) { }");

        var violaciones = Violaciones(inyectada);

        Assert.Single(violaciones);
        Assert.Contains("if (periodo.NumeroQuincena == 2)", violaciones[0], StringComparison.Ordinal);
    }

    // ── Capa 2: conductual (Q1 ≡ Q2 sobre la MISMA fuente) ──────────────────────────────────

    public static TheoryData<int, string, string> Fixtures => new()
    {
        { 3, "Octubre-1", "2026101" }, // caso rojo del plan (0 filas Aplicacion)
        { 1, "Julio-1", "2026071" },   // 2 filas Mes (enmienda 2)
        { 5, "Julio-2", "2026072" },   // ASE5 con 2 filas Mes (antecedente de la enmienda)
        { 4, "Agosto-2", "2026082" }   // 3 filas Mes / 2 Aplicacion (forma "completa")
    };

    [Theory]
    [MemberData(nameof(Fixtures))]
    public void T4_Conductual_Q1YQ2_ResuelvenIdentico(int aseId, string carpetaPeriodo, string codigo)
    {
        var periodo = Periodo.Parse(codigo);
        var (r1, r2, r4) = Rutas(carpetaPeriodo, aseId);
        var reader = new ExcelDataReaderWorkbookLeafInputReader();
        var ase = AseFactory.DesdeId(aseId);

        var q1 = reader.LeerLeafInputs(ase, ConQuincena(periodo, 1), r1, r2, r4);
        var q2 = reader.LeerLeafInputs(ase, ConQuincena(periodo, 2), r1, r2, r4);

        Assert.Equal(q2.R1.F25, q1.R1.F25);
        Assert.Equal(q2.R1.F41, q1.R1.F41);
        Assert.Equal(q2.R1.L25, q1.R1.L25);
        Assert.Equal(q2.R1.F30, q1.R1.F30);
        Assert.Equal(q2.R1.F10, q1.R1.F10);
        Assert.Equal(q2.R1.L10, q1.R1.L10);
        Assert.Equal(q2.R1.TotalOportunoEsperadoPorAse, q1.R1.TotalOportunoEsperadoPorAse);
        Assert.Equal(q2.R1.ExtemporaneoEsperadoPorAse, q1.R1.ExtemporaneoEsperadoPorAse);

        Assert.Equal(q2.R1.ValoresPorRol.Count, q1.R1.ValoresPorRol.Count);
        foreach (var (rol, valor) in q2.R1.ValoresPorRol)
        {
            Assert.Equal(valor, q1.R1.ValoresPorRol[rol]);
        }

        // La quincena NO decide si hay detalle: el desglose-detalle de R2/R4 es del PERÍODO y se
        // construye en AMBAS quincenas (Plan 59 §11-bis). Antes Q1 lo omitía y el entregable de la
        // primera quincena salía con el bloque de anticipos en el literal 0 de la plantilla.
        Assert.NotNull(q1.DetalleR2);
        Assert.NotNull(q1.DetalleR4);

        // Sobre la MISMA fuente, la quincena no cambia ni una fila del detalle: se compara la
        // secuencia completa (firma + columna 'Total'), no solo su existencia.
        Assert.Equal(q2.DetalleR2!.Filas.Count, q1.DetalleR2!.Filas.Count);
        for (var i = 0; i < q2.DetalleR2.Filas.Count; i++)
        {
            Assert.Equal(q2.DetalleR2.Filas[i].Firma, q1.DetalleR2.Filas[i].Firma);
            Assert.Equal(q2.DetalleR2.Filas[i].Valor("Total"), q1.DetalleR2.Filas[i].Valor("Total"));
        }

        Assert.Equal(q2.DetalleR4!.Filas.Count, q1.DetalleR4!.Filas.Count);
        for (var i = 0; i < q2.DetalleR4.Filas.Count; i++)
        {
            Assert.Equal(q2.DetalleR4.Filas[i].Firma, q1.DetalleR4.Filas[i].Firma);
            Assert.Equal(q2.DetalleR4.Filas[i].Valor("Total"), q1.DetalleR4.Filas[i].Valor("Total"));
        }
    }

    // ── Infraestructura del candado ─────────────────────────────────────────────────────────

    /// <summary>
    /// Devuelve las líneas que usan <c>NumeroQuincena</c> fuera de la allow-list (o que exceden su
    /// cupo). Puro sobre el texto: el candado se prueba a sí mismo sin tocar producción.
    /// </summary>
    private static List<string> Violaciones(string fuente)
    {
        var violaciones = new List<string>();
        var vistos = new Dictionary<string, int>(StringComparer.Ordinal);

        foreach (var linea in Lineas(fuente))
        {
            if (!linea.Contains("NumeroQuincena", StringComparison.Ordinal))
            {
                continue;
            }

            var normalizada = linea.Trim();

            // El candado vigila CÓDIGO (ramas/resolución), no documentación: un comentario que nombra
            // la quincena no selecciona algoritmo.
            if (EsComentario(normalizada))
            {
                continue;
            }

            if (!AllowList.ContainsKey(normalizada))
            {
                violaciones.Add($"uso no declarado: {normalizada}");
                continue;
            }

            vistos[normalizada] = vistos.GetValueOrDefault(normalizada) + 1;
        }

        foreach (var (texto, cupo) in AllowList)
        {
            var observado = vistos.GetValueOrDefault(texto);
            if (observado > cupo)
            {
                violaciones.Add($"'{texto}' aparece {observado} veces (cupo {cupo}).");
            }
        }

        return violaciones;
    }

    private static string[] Lineas(string fuente) => fuente.Split('\n');

    private static bool EsComentario(string linea) =>
        linea.StartsWith("//", StringComparison.Ordinal)
        || linea.StartsWith("*", StringComparison.Ordinal)
        || linea.StartsWith("/*", StringComparison.Ordinal);

    /// <summary>Mismo período con la quincena forzada (la fuente no cambia: solo el <see cref="Periodo"/>).</summary>
    private static Periodo ConQuincena(Periodo basePeriodo, int quincena) =>
        new() { CodigoAAAAMM = basePeriodo.CodigoAAAAMM, NumeroQuincena = quincena };
    // ── Rutas / fixtures ────────────────────────────────────────────────────────────────────

    private static (string R1, string R2, string R4) Rutas(string carpetaPeriodo, int aseId)
    {
        var carpeta = Ruta("Docs", $"Prueba {carpetaPeriodo}", "Insumos", CarpetasAse(aseId));

        return (
            Unico(carpeta, "Recaudoporcomponente"),
            Unico(carpeta, "RerpoteDetalleSaldosaFavor"),
            Unico(carpeta, "ReversiónPorComponente"));
    }

    private static string CarpetasAse(int aseId) => aseId switch
    {
        1 => "1-Promoambiental",
        2 => "2-Lime",
        3 => "3-Ciudad Limpia",
        4 => "4-Bogota Limpia",
        5 => "5-Área Limpia",
        _ => throw new ArgumentOutOfRangeException(nameof(aseId))
    };

    private static string Unico(string carpeta, string prefijo)
    {
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .SingleOrDefault(f => Path.GetFileName(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta '{prefijo}*' en {carpeta}.");
        return archivo!;
    }

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

    private static string Ruta(params string[] partes) => Path.Combine(partes.Prepend(Raiz()).ToArray());
}
