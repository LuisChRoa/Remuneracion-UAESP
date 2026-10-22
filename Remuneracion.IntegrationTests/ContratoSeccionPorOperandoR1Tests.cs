using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 53 (H1.2/H1.4, A1/A2/D1/D5): contrato de sección por operando/celda sobre las fuentes R1
/// REALES (rutas explícitas; prohibido <c>Insumos.Raiz()</c> como fuente de datos). Gates:
/// <list type="bullet">
/// <item><b>no duplicación</b>: una sola fila en la sección declarada ⇒ `Menor` = 0, nunca el bloque
/// grande repetido (residuo `R-PARCIAL-OCT-ASE3-R1-OCC` del Plan 50);</item>
/// <item><b>no pérdida</b>: `Main(UnicaRecaudo)` alimenta la celda aunque el mapa no declare
/// `TotalMenor` (oct-1 ASE5 ENERBIT: `F487` intacto y `VisibleR1 ≥ 0`);</item>
/// <item><b>magnitud ≠ ordinal</b>: `Menor`/`Medio` por posición de sección (arbitrado por el manual);
/// </item>
/// <item><b>manual</b>: el `VisibleR1` por empresa cierra al valor del administrativo en los 4 casos
/// arbitrados por el T0.</item>
/// </list>
/// Evidencia: <c>Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md</c> §4/§5.5.
/// </summary>
public sealed class ContratoSeccionPorOperandoR1Tests
{
    private static readonly string[] CarpetasAse =
    [
        "1-Promoambiental", "2-Lime", "3-Ciudad Limpia", "4-Bogota Limpia", "5-Área Limpia",
    ];

    [Fact]
    public void NoDuplicado_Oct1_Ase3_Occidente_UnaFilaOportuna()
    {
        // Octubre-1 ASE3 OCCIDENTE: la empresa trae UNA fila (Oportuno#2, f23) y la hoja tiene otra
        // sección (Oportuno#1) solo con ENEL → `TotalMenor`/`SubsMenor`/`TotalMedio` = 0 y `TotalMain`
        // = el bloque grande. Antes (ocurrencia): `TotalMedio=[1]` lanzaba y `SubsMenor` duplicaba.
        var occidente = LeerConciliacion("Prueba Octubre-1", "2026101", 3).Single(c => c.Empresa.Id == 4);

        Assert.Equal(286711907.53m, occidente.CeldasR1["F237"]); // TotalMain (última sección presente)
        Assert.Equal(145489038.47m, occidente.CeldasR1["F253"]); // SubsMain
        Assert.Equal(0m, occidente.CeldasR1["F216"]);            // TotalMenor: sección declarada ausente
        Assert.Equal(0m, occidente.CeldasR1["F242"]);            // SubsMenor: ídem
        Assert.Equal(0m, occidente.CeldasR1["F222"]);            // TotalMedio: < 3 secciones
        Assert.NotEqual(occidente.CeldasR1["F237"], occidente.CeldasR1["F216"]);
        Assert.Equal(427389334.19m, occidente.VisibleR1);

        // La fuente R1 ESTÁ presente: la ausencia de sección es 0 explícito sin huella nueva
        // (la huella R4:OCCIDENTE de octubre-1 es de OTRO reporte, Plan 50 A1).
        Assert.DoesNotContain(occidente.FuentesAusentes, f => f.StartsWith("R1:", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void NoPerdida_Oct1_Ase5_Enerbit_UnicaRecaudo()
    {
        // Octubre-1 ASE5 ENERBIT: el mapa (3,5) NO declara `TotalMenor`; la única fila (Oportuno#1) va a
        // `Main(UnicaRecaudo)` → el importe no desaparece y el visible no queda negativo.
        var enerbit = LeerConciliacion("Prueba Octubre-1", "2026101", 5).Single(c => c.Empresa.Id == 3);

        Assert.Equal(160186886.97m, enerbit.CeldasR1["F487"]);
        Assert.Equal(0m, enerbit.CeldasR1["L487"]);
        Assert.False(enerbit.CeldasR1.ContainsKey("F474"));
        Assert.Equal(139510580.00m, enerbit.VisibleR1);
        Assert.True(enerbit.VisibleR1 >= 0m, "El visible de una fuente presente no puede ser negativo.");
    }

    [Fact]
    public void Oct1_Ase5_LasTresEmpresas_MainUnicaRecaudo()
    {
        var leafs = LeerConciliacion("Prueba Octubre-1", "2026101", 5);

        Assert.Equal(7414210944.89m, leafs.Single(c => c.Empresa.Id == 2).CeldasR1["F483"]);
        Assert.Equal(160186886.97m, leafs.Single(c => c.Empresa.Id == 3).CeldasR1["F487"]);
        Assert.Equal(134612117.14m, leafs.Single(c => c.Empresa.Id == 4).CeldasR1["F497"]);

        // `Menor` de ENEL/OCCIDENTE: la sección única no tiene "otro" bloque → 0 (sin duplicar).
        Assert.Equal(0m, leafs.Single(c => c.Empresa.Id == 2).CeldasR1["F474"]);
        Assert.Equal(0m, leafs.Single(c => c.Empresa.Id == 4).CeldasR1["F477"]);
    }

    [Fact]
    public void MagnitudNoOrdinal_Ago2_Ase2_PorPosicionDeSeccion()
    {
        var leafs = LeerConciliacion("Prueba Agosto-2", "2026082", 2);

        // ENEL: f7 Oportuno#1 = −1.414.859,14 (primera) vs f17 Aplicacion#1 = 131.543,19 (intermedia).
        var enel = leafs.Single(c => c.Empresa.Id == 2);
        Assert.Equal(-1414859.14m, enel.CeldasR1["F83"]); // TotalMenor por POSICIÓN, no por magnitud
        Assert.Equal(19237658087.83m, enel.VisibleR1);

        // OCCIDENTE: f12 Oportuno#1 = −28.650,84 vs f20 Aplicacion#1 = 24.890,00.
        var occidente = leafs.Single(c => c.Empresa.Id == 4);
        Assert.Equal(-28650.84m, occidente.CeldasR1["F89"]);
        Assert.Equal(2217280433.95m, occidente.VisibleR1);
    }

    [Fact]
    public void Jul2_Ase5_Enel_MenorCeroPorSeccionUnica()
    {
        var enel = LeerConciliacion("Prueba Julio-2", "2026072", 5).Single(c => c.Empresa.Id == 2);

        Assert.Equal(12823196468.26m, enel.CeldasR1["F483"]);
        Assert.Equal(0m, enel.CeldasR1["F474"]); // el bloque grande NO se duplica
        Assert.Equal(11701157844.90m, enel.VisibleR1);
    }

    [Fact]
    public void VisibleR1_CierraAlValorDelManual_EnLos4CasosArbitrados()
    {
        Assert.Equal(11701157844.90m, LeerConciliacion("Prueba Julio-2", "2026072", 5).Single(c => c.Empresa.Id == 2).VisibleR1);
        Assert.Equal(11748395433.49m, LeerConciliacion("Prueba Agosto-2", "2026082", 5).Single(c => c.Empresa.Id == 2).VisibleR1);
        Assert.Equal(19237658087.83m, LeerConciliacion("Prueba Agosto-2", "2026082", 2).Single(c => c.Empresa.Id == 2).VisibleR1);
        Assert.Equal(2217280433.95m, LeerConciliacion("Prueba Agosto-2", "2026082", 2).Single(c => c.Empresa.Id == 4).VisibleR1);
    }

    [Fact]
    public void SinEmpates_Los20ParesPeriodoAse_SeResuelvenSinThrow()
    {
        foreach (var (carpeta, codigo) in new[] { ("Prueba Julio-1", "2026071"), ("Prueba Julio-2", "2026072"), ("Prueba Agosto-2", "2026082"), ("Prueba Octubre-1", "2026101") })
        {
            for (var ase = 1; ase <= 5; ase++)
            {
                var leafs = LeerConciliacion(carpeta, codigo, ase);
                Assert.Equal(5, leafs.Count);

                // El mapa declara celdas solo para las empresas con detalle no-estático: esas deben
                // traer operandos resueltos (y ninguna debe lanzar por empate de sección).
                var conDetalle = WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa
                    .Count(kv => kv.Key.AseId == ase);
                Assert.Equal(conDetalle, leafs.Count(l => l.CeldasR1.Count > 0));
            }
        }
    }

    [Fact]
    public void ContratoMapa_DeclaraSeccionPorCelda_EnLas13Claves()
    {
        foreach (var ((empresaId, aseId), celdas) in WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa)
        {
            Assert.NotEmpty(celdas);
            foreach (var (celda, operando, seccion) in celdas)
            {
                Assert.False(string.IsNullOrWhiteSpace(celda), $"Celda vacía en ({empresaId},{aseId}).");
                var esperada = operando switch
                {
                    WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor
                        or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMenor
                        or WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMenor => SeccionEsperadaR1.Primera,
                    WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMedio
                        or WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMedio => SeccionEsperadaR1.Intermedia,
                    WorkbookLeafCellMapPorEmpresa.FuenteR1.Block5Total => seccion, // canal D=5: no participa
                    _ => aseId == 5 && operando != WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMain
                        ? SeccionEsperadaR1.UnicaRecaudo
                        : SeccionEsperadaR1.UltimaOportuna,
                };

                Assert.True(seccion == esperada, $"{celda} ({empresaId},{aseId}): {seccion} ≠ {esperada} para {operando}.");
            }
        }
    }

    private static IReadOnlyList<ConciliacionEmpresaInputs> LeerConciliacion(string periodo, string codigo, int ase)
    {
        return new ExcelDataReaderWorkbookLeafInputReader().LeerConciliacionEmpresas(
            Insumos.Ase(ase),
            Periodo.Parse(codigo),
            RutaFuente(periodo, ase, "Recaudoporcomponente"),
            RutaFuente(periodo, ase, "RerpoteDetalleSaldosaFavor"),
            RutaFuente(periodo, ase, "Reversi"));
    }

    private static string RutaFuente(string periodo, int ase, string prefijo)
    {
        var dir = Path.Combine(Insumos.Raiz(), "Docs", periodo, "Insumos", CarpetasAse[ase - 1]);
        Assert.True(Directory.Exists(dir), $"Falta la carpeta explícita de insumos: {dir}");
        var archivo = Directory.EnumerateFiles(dir, "*.xlsx", SearchOption.TopDirectoryOnly)
            .FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith(prefijo, StringComparison.OrdinalIgnoreCase));
        Assert.True(archivo is not null, $"Falta el reporte '{prefijo}*' del ASE {ase} en {dir}");
        return archivo!;
    }
}
