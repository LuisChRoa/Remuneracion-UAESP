using DocumentFormat.OpenXml.Packaging;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 45 (H1) + Plan 46 (H1/H2, S2.1/S2.2/S2.4): regla de SECCIÓN-CONTRATO por frontera de CIERRE
/// (<see cref="ContratoSeccionPeriodFreeR1.SeccionPorCierre"/>) + CANAL EFECTIVO
/// (<see cref="ContratoSeccionPeriodFreeR1.CanalEfectivo"/>) en el escritor values-only.
///
/// Evidencia (bytes): Plan 46 §S1 + fixture read-only
/// <see cref="ContratoSeccionPeriodFreeT0Plan46Tests"/>. Caso crítico ASE5-julio-2: su 1.ª sección
/// (filas 5-13) la cierra <c>Aplicacion</c> (R14, 417.519,62) → sección Aplicación → cae en
/// <c>Aplic#1</c> oficial (615-628); 603-614 queda en cero. Cada assert de este archivo cita
/// fila-fuente + celda-slot + sección-contrato + canal efectivo.
/// </summary>
public sealed class ContratoSeccionPeriodFreeH1Tests
{
    private const string HojaR1 = "Reporte Componentes R1";

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    // ── S1.1: ASE1-julio2 — canal efectivo en las filas sin D (VS/VI) ─────────────────────────────

    [Fact]
    public void S1_1_Ase1Julio2_CanalEfectivo()
    {
        var bloque = LeerJulio(1);
        var salida = Escribir(1, bloque);
        try
        {
            // Fuente R5 (Vlr Servicio, canal E por R6 D=E), R6 (D=E), R7 (C=ENEL) → F6/F7/F8.
            // Fuente R8 (Vlr Servicio, canal O por R9 D=O) → F11 (NO F9); R9 (D=O) → F12;
            // R10 (C=OCCIDENTE) → F15. F9/F10/F13/F14 quedan en el cero del template.
            Assert.Equal(-7108.82m, Celda(salida.Ruta, "F6"));
            Assert.Equal(-7108.82m, Celda(salida.Ruta, "F7"));
            Assert.Equal(-7108.82m, Celda(salida.Ruta, "F8"));
            Assert.Equal(0m, Celda(salida.Ruta, "F9"));
            Assert.Equal(0m, Celda(salida.Ruta, "F10"));
            Assert.Equal(-4108.55m, Celda(salida.Ruta, "F11"));
            Assert.Equal(-4108.55m, Celda(salida.Ruta, "F12"));
            Assert.Equal(0m, Celda(salida.Ruta, "F13"));
            Assert.Equal(0m, Celda(salida.Ruta, "F14"));
            Assert.Equal(-4108.55m, Celda(salida.Ruta, "F15"));

            // Frontera Mes (R11) y AFaseo (R12) sin cambio: F16/F17 = -11.217,37.
            Assert.Equal(-11217.37m, Celda(salida.Ruta, "F16"));
            Assert.Equal(-11217.37m, Celda(salida.Ruta, "F17"));
        }
        finally
        {
            salida.Dispose();
        }
    }

    // ── S1.2/S1.3: ASE5-julio2 — bloque de la 1.ª y resto del bloque por sección de CIERRE ────────

    [Fact]
    public void S1_2_S1_3_Ase5Julio2_SeccionPorCierre()
    {
        var bloque = LeerJulio(5);
        var salida = Escribir(5, bloque);
        try
        {
            // §S1.2: la 1.ª oficial (603-613) y el AFaseo (614) quedan en CERO — la fuente ASE5-julio2
            // no trae bloque para esa sección (su dinero se ve en 615-628).
            for (var fila = 603; fila <= 614; fila++)
            {
                Assert.Equal(0m, Celda(salida.Ruta, $"F{fila}"));
            }

            // Fuente R5 (VS, canal H por R6) → F619; R6 (D=H) → F621.
            Assert.Equal(52706.16m, Celda(salida.Ruta, "F619"));
            Assert.Equal(0m, Celda(salida.Ruta, "F620"));
            Assert.Equal(52706.16m, Celda(salida.Ruta, "F621"));

            // Fuente R7 (VS, canal O) → F622; R8 (Vlr Intereses, canal O) → F623 (FP-3: TENÍA slot,
            // ya no es cero documentado); R9 (D=O) → F624.
            Assert.Equal(49921.06m, Celda(salida.Ruta, "F622"));
            Assert.Equal(402.40m, Celda(salida.Ruta, "F623"));
            Assert.Equal(50323.46m, Celda(salida.Ruta, "F624"));

            // Fuente R10 (VS, canal T) → F625; R11 (Vlr Intereses, canal T) → F626; R12 (D=T) → F627.
            Assert.Equal(306792.77m, Celda(salida.Ruta, "F625"));
            Assert.Equal(7697.23m, Celda(salida.Ruta, "F626"));
            Assert.Equal(314490m, Celda(salida.Ruta, "F627"));

            // Fuente R13 (OCCIDENTE) → F628 y R14 (frontera Aplicacion, cierra la 1.ª) → F629.
            Assert.Equal(417519.62m, Celda(salida.Ruta, "F628"));
            Assert.Equal(417519.62m, Celda(salida.Ruta, "F629"));
            // El dinero ya NO se ve en el OCCIDENTE de la 1.ª oficial (F612): el bloque migró.
            Assert.Equal(0m, Celda(salida.Ruta, "F612"));

            // §S1.3: bloque grande (O#2 oficial, 630-647) y cierre.
            Assert.Equal(12823196468.26m, Celda(salida.Ruta, "F632"));
            Assert.Equal(13084233937.44m, Celda(salida.Ruta, "F648"));
            // Componente (A#2) y subsidios con H/O en su canal.
            Assert.Equal(13084651457.06m, Celda(salida.Ruta, "F649"));
            Assert.Equal(-7906.16m, Celda(salida.Ruta, "F653"));
            Assert.Equal(32235.54m, Celda(salida.Ruta, "F655"));
            Assert.Equal(24329.38m, Celda(salida.Ruta, "F659"));
            // O#3 (661-673) y cola.
            Assert.Equal(-187855.98m, Celda(salida.Ruta, "F667"));
            Assert.Equal(-1036320522.40m, Celda(salida.Ruta, "F674"));
            Assert.Equal(-1036296193.02m, Celda(salida.Ruta, "F675"));
            Assert.Equal(12048355264.04m, Celda(salida.Ruta, "F676"));

            // Fronteras Mes: R33 → F648 (O#2) y R54 → F674 (O#3) — ya cubiertas arriba.
            // La 1.ª oficial (F613) quedó en cero (sin contraparte fuente).
            Assert.Equal(0m, Celda(salida.Ruta, "F613"));
        }
        finally
        {
            salida.Dispose();
        }
    }

    // ── S2.3: fuera de la lista de ceros → ERR-SIN-SLOT (sin remap) ───────────────────────────────

    [Fact]
    public void S2_3_FueraDeListaErrSinSlot()
    {
        using var otra = new CopiaTemporal();
        var inventada = new BloqueEspejoAseInputs
        {
            Ase = AseFactory.DesdeId(1),
            TieneColumnaEspeciales = false,
            Filas =
            [
                new FilaEspejoR1
                {
                    A = "ComponenteRaro",
                    B = "Total",
                    ValoresPorColumna = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase) { ["Total"] = 1m },
                },
            ],
        };

        using var documento = SpreadsheetDocument.Open(otra.Ruta, true);
        var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
        var ex = Assert.Throws<CalculoInvalidoException>(() => EscritorValoresSlotR1.Escribir(parte, inventada));
        Assert.Equal(CodigoError.SinSlot, ex.Codigo);
        Assert.Contains("ERR-SIN-SLOT", ex.Message, StringComparison.Ordinal);
        Assert.Contains("ASE 1", ex.Message, StringComparison.Ordinal);
    }

    // ── S2.2: filas de empresa (C=…) intactas y reubicadas por cierre (OCCIDENTE ASE5) ────────────

    [Fact]
    public void S2_2_FilasEmpresaIntactas_OccidenteAse5()
    {
        var bloque = LeerJulio(5);
        var salida = Escribir(5, bloque);
        try
        {
            // Plan 46 §S1.2: el OCCIDENTE de la 1.ª fuente (R13, 417.519,62) vive ahora en el
            // OCCIDENTE de la Aplic#1 oficial (F628); el de la 1.ª oficial (F612) queda en cero.
            Assert.Equal(417519.62m, Celda(salida.Ruta, "F628"));
            Assert.Equal(0m, Celda(salida.Ruta, "F612"));

            // Los otros 2 anclajes OCCIDENTE de ASE5 (Aplic#1 647 / Aplic#2 673) siguen en su slot.
            AssertSlot(salida.Ruta, "F647", bloque.Filas[32 - 5]);
            AssertSlot(salida.Ruta, "F673", bloque.Filas[53 - 5]);
            Assert.NotEqual(0m, Celda(salida.Ruta, "F647"));
            Assert.NotEqual(0m, Celda(salida.Ruta, "F673"));
        }
        finally
        {
            salida.Dispose();
        }
    }

    // ── Unitarios del contrato puro (sección por CIERRE + canal efectivo) ────────────────────────

    [Fact]
    public void Contrato_Alinear_MapeosProbados()
    {
        var oficial = new[]
        {
            new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1),
            new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 1),
            new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2),
            new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 2),
            new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 3),
        };

        // Forma por CIERRE de ASE5-julio2 [A#1,O#1,A#2,O#2] → [A#1,O#2,A#2,O#3] (salta el O#1 oficial).
        var inversa = ContratoSeccionPeriodFreeR1.Alinear(
            [
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 1),
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1),
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 2),
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2),
            ],
            oficial);

        Assert.Equal(oficial[1], inversa[new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 1)]);
        Assert.Equal(oficial[2], inversa[new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1)]);
        Assert.Equal(oficial[3], inversa[new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 2)]);
        Assert.Equal(oficial[4], inversa[new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2)]);

        // Forma por CIERRE [O#1,O#2,O#3] (agosto-ASE3 / Q1-ASE2/5) → [O#1,O#2,O#3] (salta A#1/A#2).
        var sinAplic = ContratoSeccionPeriodFreeR1.Alinear(
            [
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1),
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2),
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 3),
            ],
            oficial);
        Assert.Equal(oficial[0], sinAplic[new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1)]);
        Assert.Equal(oficial[2], sinAplic[new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2)]);
        Assert.Equal(oficial[4], sinAplic[new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 3)]);

        // Forma completa (julio2 ASE1-4) → identidad.
        var completa = ContratoSeccionPeriodFreeR1.Alinear(oficial, oficial);
        for (var i = 0; i < oficial.Length; i++)
        {
            Assert.Equal(oficial[i], completa[oficial[i]]);
        }
    }

    [Fact]
    public void Contrato_SeccionPorCierre_AbsorbeLasTresFormas_SinRamas()
    {
        // Secuencia mínima: [VS, Mes, AFaseo, VS, Aplic, VS, Mes] → secciones por CIERRE.
        FilaEspejoR1 F(string a = "", string b = "", string c = "", string d = "", string e = "") =>
            new() { A = a, B = b, C = c, D = d, E = e };

        var filas = new List<FilaEspejoR1>
        {
            F(e: "Vlr Servicio"),                                   // 0 cierra con Mes@1 → Oportuno#1
            F(b: "Mes", c: "Total"),                                // 1 cierra la sección de arriba → Oportuno#1
            F(a: "AFaseo", b: "Total"),                             // 2 sin sección
            F(e: "Vlr Servicio"),                                   // 3 cierra con Aplic@4 → Aplicacion#1
            F(b: "Aplicacion nuevos x reversion", c: "Total"),      // 4 Aplicacion#1
            F(e: "Vlr Servicio"),                                   // 5 cierra con Mes@6 → Oportuno#1
            F(b: "Mes", c: "Total"),                                // 6 Oportuno#1
        };

        Assert.Equal(new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1), ContratoSeccionPeriodFreeR1.SeccionPorCierre(filas, 0));
        Assert.Equal(new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1), ContratoSeccionPeriodFreeR1.SeccionPorCierre(filas, 1));
        Assert.Null(ContratoSeccionPeriodFreeR1.SeccionPorCierre(filas, 2));
        Assert.Equal(new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 1), ContratoSeccionPeriodFreeR1.SeccionPorCierre(filas, 3));
        Assert.Equal(new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 1), ContratoSeccionPeriodFreeR1.SeccionPorCierre(filas, 4));
        Assert.Equal(new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2), ContratoSeccionPeriodFreeR1.SeccionPorCierre(filas, 5));
        Assert.Equal(new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2), ContratoSeccionPeriodFreeR1.SeccionPorCierre(filas, 6));
    }

    [Fact]
    public void Contrato_CanalEfectivo_SoloFilasDeCanalSinD()
    {
        FilaEspejoR1 F(string a = "", string b = "", string c = "", string d = "", string e = "") =>
            new() { A = a, B = b, C = c, D = d, E = e };

        var filas = new List<FilaEspejoR1>
        {
            F(e: "Vlr Servicio"),              // 0 → canal H (próximo D=H)
            F(d: "H", e: "Total"),             // 1 → con D propio, no se recalcula
            F(c: "ENEL", d: "Total"),          // 2 → C=… D=Total, conserva su D
            F(e: "Vlr Intereses"),             // 3 → canal O (próximo D=O)
            F(d: "O", e: "Total"),             // 4
            F(a: "Componente", b: "Total"),    // 5 → no es fila de canal
            F(b: "Mes", c: "Total"),           // 6 → no es fila de canal
        };

        Assert.Equal("H", ContratoSeccionPeriodFreeR1.CanalEfectivo(filas, 0));
        Assert.Null(ContratoSeccionPeriodFreeR1.CanalEfectivo(filas, 1));
        Assert.Null(ContratoSeccionPeriodFreeR1.CanalEfectivo(filas, 2));
        Assert.Equal("O", ContratoSeccionPeriodFreeR1.CanalEfectivo(filas, 3));
        Assert.Null(ContratoSeccionPeriodFreeR1.CanalEfectivo(filas, 5));
        Assert.Null(ContratoSeccionPeriodFreeR1.CanalEfectivo(filas, 6));
    }

    [Fact]
    public void CerosDocumentados_Vacios_TrasWithdrawFp3()
    {
        // Plan 46 (FP-3, S2.4): las 2 entradas de `Vlr Intereses` se retiraron; el mecanismo sigue
        // existiendo pero sin patrones activos → ninguna fila es cero documentado.
        var intereses = new FilaEspejoR1 { E = "Vlr Intereses" };
        Assert.False(CerosDocumentadosR1.EsCeroDocumentado(intereses, new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 1)));
        Assert.False(CerosDocumentadosR1.EsCeroDocumentado(intereses, SeccionEspejoR1.Primera));
        Assert.False(CerosDocumentadosR1.EsCeroDocumentado(new FilaEspejoR1 { E = "Vlr Servicio" }, SeccionEspejoR1.Primera));
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────────────────────

    private static decimal Celda(string ruta, string celda) =>
        TestHelpers.LeerCeldaNumerica(ruta, HojaR1, celda);

    private static void AssertSlot(string ruta, string celda, FilaEspejoR1 fila)
    {
        var esperado = fila.Valor("Total");
        Assert.True(esperado is not null, $"La fila fuente esperada para {celda} no trae 'Total'.");
        Assert.Equal(esperado!.Value, Celda(ruta, celda));
    }

    private static BloqueEspejoAseInputs LeerJulio(int aseId) =>
        new ExcelDataReaderWorkbookLeafInputReader().LeerEspejoR1(AseFactory.DesdeId(aseId), FuenteR1Julio(aseId));

    private static CopiaTemporal Escribir(int aseId, BloqueEspejoAseInputs bloque)
    {
        var salida = new CopiaTemporal();
        using (var documento = SpreadsheetDocument.Open(salida.Ruta, true))
        {
            var parte = documento.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart null");
            EscritorValoresSlotR1.Escribir(parte, bloque);
            parte.Workbook?.Save();
        }

        return salida;
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

    private static string Plantilla() =>
        Path.Combine(Raiz(), "Docs", "Plantilla estandarizada", "Formato Remuneracion 2026xx-x Total.xlsx");

    private static string FuenteR1Julio(int aseId)
    {
        var carpeta = Path.Combine(Raiz(), "Docs", "Prueba Julio-2", "Insumos", CarpetasAse[aseId - 1]);
        Assert.True(Directory.Exists(carpeta), $"Falta la carpeta explícita: {carpeta}");
        var archivo = Directory.EnumerateFiles(carpeta, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("Recaudoporcomponente", StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta Recaudoporcomponente del ASE{aseId} en {carpeta}");
        return archivo!;
    }

    private sealed class CopiaTemporal : IDisposable
    {
        private readonly string _dir;

        public CopiaTemporal()
        {
            _dir = Path.Combine(Path.GetTempPath(), "p46-h1-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
            Ruta = Path.Combine(_dir, "salida.xlsx");
            File.Copy(Plantilla(), Ruta, overwrite: false);
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
