using DocumentFormat.OpenXml.Packaging;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 45 (H1, S2.1/S2.2/S2.3/S2.4): regla sin-C + cero documentado con el contrato de sección
/// PERIOD-FREE por ASE (<see cref="ContratoSeccionPeriodFreeR1"/>).
///
/// El caso crítico es ASE5-julio-2: su fuente arranca con <c>Aplicacion</c> (5 secciones) mientras el
/// oficial declara 6 (<c>1.ª·Oportuno#1·Aplic#1·Oportuno#2·Aplic#2·Oportuno#3</c>). La regla alinea
/// monotónicamente por tipo: <c>[1.ª,Aplic#1,Oportuno#1,Aplic#2,Oportuno#2] →
/// [1.ª,Aplic#1,Oportuno#2,Aplic#2,Oportuno#3]</c>. Rutas explícitas; regen a temp FRESCA; el
/// <c>&lt;f&gt;</c> del oficial queda intacto.
/// </summary>
public sealed class ContratoSeccionPeriodFreeH1Tests
{
    private const string HojaR1 = "Reporte Componentes R1";

    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    // ── S2.1: colocación corregida de ASE5-julio-2 (fila fuente → slot de su sección-contrato) ────

    [Fact]
    public void S2_1_Ase5Julio2_SinC_ColocaPorSeccionContrato()
    {
        var bloque = LeerJulio(5);
        var salida = Escribir(5, bloque);
        try
        {
            // Índice de Filas = fila-fuente − 5 (la primera fila de datos es la R5 'Vlr Servicio').
            FilaEspejoR1 F(int fila) => bloque.Filas[fila - 5];

            // 1.ª sección (fuente → oficial 1.ª): VS/D=H/D=O/D=T/OCCIDENTE a su slot.
            AssertSlot(salida.Ruta, "F603", F(5));
            AssertSlot(salida.Ruta, "F607", F(6));
            AssertSlot(salida.Ruta, "F609", F(9));
            AssertSlot(salida.Ruta, "F611", F(12));
            Assert.Equal(417519.62m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F612"));

            // Plan 45 S2.4: la 1.ª oficial NO tiene VI → los dos 'Vlr Intereses' de la fuente (R8/R11)
            // caen a CERO DOCUMENTADO (el hueco de plantilla); no ensucian ningún otro slot.
            Assert.Equal(402.4m, F(8).Valor("Total"));
            Assert.Equal(7697.23m, F(11).Valor("Total"));

            // 'Aplic#1' fuente → 'Aplic#1' oficial (630-647): el D=E/Total (R17) va a 632, NO a 604.
            Assert.Equal(0m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F604"));
            Assert.Equal(0m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F610"));
            AssertSlot(salida.Ruta, "F632", F(17));
            Assert.Equal(12823196468.26m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F632"));

            // 'Oportuno#1' fuente (Componente/Subsidio) → 'Oportuno#2' oficial (649-659): el bloque NO
            // se desplaza a la Oportuno#1 oficial (615-628), que queda en su cero.
            AssertSlot(salida.Ruta, "F649", F(34));
            AssertSlot(salida.Ruta, "F650", F(35));
            AssertSlot(salida.Ruta, "F653", F(37));
            AssertSlot(salida.Ruta, "F654", F(36));
            AssertSlot(salida.Ruta, "F656", F(38));
            AssertSlot(salida.Ruta, "F659", F(39));
            Assert.Equal(24329.38m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F659"));
            Assert.Equal(0m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F628"));

            // S2.4: F654 (D=H de Oportuno#2) y F661 (Subsidio de Aplic#2) NO se duplican (era el
            // artefacto de ESCRITURA del cursor físico).
            AssertSlot(salida.Ruta, "F661", F(41));
            Assert.NotEqual(
                TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F654"),
                TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F661"));

            // Cierre: 'Subs/Cont' → 675, 'Total' → 676.
            AssertSlot(salida.Ruta, "F675", F(55));
            AssertSlot(salida.Ruta, "F676", F(56));

            // Fronteras: la fila `Mes/Total` viaja con la sección que ABRE por el contrato. ASE5-julio2
            // tiene 2 `Mes` fuente → Oportuno#2 (648) y Oportuno#3 (674) oficiales; el Oportuno#1
            // oficial (613) queda sin contraparte en su cero.
            Assert.Equal(0m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F613"));
            AssertSlot(salida.Ruta, "F648", F(33));
            AssertSlot(salida.Ruta, "F674", F(54));
        }
        finally
        {
            salida.Dispose();
        }
    }

    // ── S2.3: cero documentado sin ERR-SIN-SLOT; fila sin slot FUERA de la lista = ERR-SIN-SLOT ─────

    [Fact]
    public void S2_3_CeroDocumentado_Y_FueraDeListaErrSinSlot()
    {
        // El 'Vlr Intereses' de la 1.ª de ASE5 es cero documentado: la corrida completa no lanza.
        var bloque = LeerJulio(5);
        var salida = Escribir(5, bloque);
        salida.Dispose();

        // Fuera de la lista: una fila de componente con (firma, sección) sin slot → ERR-SIN-SLOT.
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

    // ── S2.2: las filas con empresa en C quedan intactas (ancla OCCIDENTE ASE5) ───────────────────

    [Fact]
    public void S2_2_FilasEmpresaIntactas_OccidenteAse5()
    {
        var bloque = LeerJulio(5);
        var salida = Escribir(5, bloque);
        try
        {
            // Los 3 anclajes OCCIDENTE de ASE5 (1.ª 612, Aplic#1 647, Aplic#2 673) quedan intactos.
            Assert.Equal(417519.62m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F612"));
            AssertSlot(salida.Ruta, "F647", bloque.Filas[32 - 5]);
            AssertSlot(salida.Ruta, "F673", bloque.Filas[53 - 5]);

            // El rótulo de empresa sigue viviendo en su celda (no se toca; solo se escriben valores F..Z).
            Assert.NotEqual(0m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F647"));
            Assert.NotEqual(0m, TestHelpers.LeerCeldaNumerica(salida.Ruta, HojaR1, "F673"));
        }
        finally
        {
            salida.Dispose();
        }
    }

    // ── Unitarios del contrato puro ─────────────────────────────────────────────────────────────

    [Fact]
    public void Contrato_Alinear_MapeosProbados()
    {
        var oficial = new[]
        {
            SeccionEspejoR1.Primera,
            new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1),
            new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 1),
            new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2),
            new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 2),
            new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 3),
        };

        // Forma invertida ASE5-julio2 [P,A,O,A,O] → [1.ª,Aplic#1,Oportuno#2,Aplic#2,Oportuno#3].
        var inversa = ContratoSeccionPeriodFreeR1.Alinear(
            [
                SeccionEspejoR1.Primera,
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 1),
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1),
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 2),
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2),
            ],
            oficial);

        Assert.Equal(oficial[0], inversa[SeccionEspejoR1.Primera]);
        Assert.Equal(oficial[2], inversa[new SeccionFronteraArriba(TipoSeccionEspejoR1.Aplicacion, 1)]);
        Assert.Equal(oficial[3], inversa[new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1)]);
        Assert.Equal(oficial[5], inversa[new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2)]);

        // Forma sin Aplicación [P,O,O,O] → [1.ª,Oportuno#1,Oportuno#2,Oportuno#3].
        var sinAplic = ContratoSeccionPeriodFreeR1.Alinear(
            [
                SeccionEspejoR1.Primera,
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1),
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2),
                new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 3),
            ],
            oficial);
        Assert.Equal(oficial[1], sinAplic[new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1)]);
        Assert.Equal(oficial[3], sinAplic[new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 2)]);
        Assert.Equal(oficial[5], sinAplic[new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 3)]);
    }

    [Fact]
    public void CerosDocumentados_SoloInteresesDePrimeraSeccion()
    {
        var intereses = new FilaEspejoR1 { E = "Vlr Intereses" };
        Assert.True(CerosDocumentadosR1.EsCeroDocumentado(intereses, SeccionEspejoR1.Primera));
        Assert.False(CerosDocumentadosR1.EsCeroDocumentado(intereses, new SeccionFronteraArriba(TipoSeccionEspejoR1.Oportuno, 1)));
        Assert.False(CerosDocumentadosR1.EsCeroDocumentado(new FilaEspejoR1 { E = "Vlr Servicio" }, SeccionEspejoR1.Primera));
        Assert.False(CerosDocumentadosR1.EsCeroDocumentado(new FilaEspejoR1 { C = "ENEL", D = "Total" }, SeccionEspejoR1.Primera));
    }

    // ── Helpers ──────────────────────────────────────────────────────────────────────────────────

    private static void AssertSlot(string ruta, string celda, FilaEspejoR1 fila)
    {
        var esperado = fila.Valor("Total");
        Assert.True(esperado is not null, $"La fila fuente esperada para {celda} no trae 'Total'.");
        Assert.Equal(esperado!.Value, TestHelpers.LeerCeldaNumerica(ruta, HojaR1, celda));
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
            _dir = Path.Combine(Path.GetTempPath(), "p45-h1-" + Guid.NewGuid().ToString("N"));
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
