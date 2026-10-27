using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 53 (H1.2/H1.4, A1/A2/D1/D5) + Plan 55 (H1.1/H1.3, D4 — enmienda 2026-10-23): contrato de
/// sección por operando/celda sobre las fuentes R1 REALES (rutas explícitas; prohibido
/// <c>Insumos.Raiz()</c> como fuente de datos). Gates:
/// <list type="bullet">
/// <item><b>no duplicación</b>: una sola fila de empresa en su sección ⇒ el bloque NO se repite en otra
/// celda (Plan 55: el destino es la fila que el CUADRO declara para esa sección);</item>
/// <item><b>no pérdida</b>: el importe de una fuente presente entra al agregado por empresa aunque el
/// cuadro no declare su fila (exceso REGISTRADO, nunca descarte silencioso);</item>
/// <item><b>magnitud ≠ ordinal</b>: la sección la fija la frontera que la cierra, no la magnitud;</item>
/// <item><b>manual</b>: el <c>VisibleR1</c> por empresa sigue cerrando al valor del administrativo en
/// los 4 casos arbitrados por el T0 (Plan 53) — verificado tras el cambio de mecanismo del Plan 55.</item>
/// </list>
/// Evidencia: <c>Docs/T0-Plan53-Contrato-Seccion-Por-Operando-R1.md</c> §4/§5.5 y
/// <c>Docs/T0-Plan55-Conciliacion-Por-Empresa.md</c> §3.
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
        // Octubre-1 ASE3 OCCIDENTE: la empresa trae UNA fila en la sección Oportuno (cuadro Componente)
        // y otra en Subs/Cont. Plan 55 (H1.1, D4): el destino es la fila que el CUADRO del ASE3 declara
        // para esa (sección, etiqueta) — F329 (Componente) y F352 (Subs/Cont). No hay celda "Menor/Medio"
        // que duplique el bloque.
        var occidente = LeerConciliacion("Prueba Octubre-1", "2026101", 3).Single(c => c.Empresa.Id == 4);

        Assert.Equal(286711907.53m, occidente.CeldasR1["F329"]);
        Assert.Equal(145489038.47m, occidente.CeldasR1["F352"]);
        // Celdas: F329 (Total de la sección), L329 (Especiales) y F352 (Subs/Cont de la sección).
        Assert.Equal(3, occidente.CeldasR1.Count);
        Assert.NotEqual(occidente.CeldasR1["F329"], occidente.CeldasR1["F352"]);
        Assert.Equal(427389334.19m, occidente.VisibleR1);

        // La fuente R1 ESTÁ presente: la ausencia de sección no genera huella nueva.
        Assert.DoesNotContain(occidente.FuentesAusentes, f => f.StartsWith("R1:", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void NoPerdida_Oct1_Ase5_Enerbit_UnicaRecaudo()
    {
        // Octubre-1 ASE5 ENERBIT: el cuadro Subs/Cont del ASE5 declara su fila (F666). Su fila del
        // cuadro Componente NO tiene cuadro declarado en el template ⇒ EXCESO registrado (nunca celda
        // inventada), y el importe igual entra al agregado porque se deriva de la FUENTE.
        var enerbit = LeerConciliacion("Prueba Octubre-1", "2026101", 5).Single(c => c.Empresa.Id == 3);

        Assert.Equal(-20676306.97m, enerbit.CeldasR1["F666"]);
        Assert.Single(enerbit.CeldasR1);
        Assert.Contains(enerbit.DestinosSinCelda, d => d.StartsWith("R1:ENERBIT:Componente", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(139510580.00m, enerbit.VisibleR1);
        Assert.True(enerbit.VisibleR1 >= 0m, "El visible de una fuente presente no puede ser negativo.");
    }

    [Fact]
    public void Oct1_Ase5_LasTresEmpresas_ResuelvenSuFilaDeclarada()
    {
        var leafs = LeerConciliacion("Prueba Octubre-1", "2026101", 5);

        Assert.Equal(7414210944.89m, leafs.Single(c => c.Empresa.Id == 2).CeldasR1["F605"]);
        Assert.Equal(-20676306.97m, leafs.Single(c => c.Empresa.Id == 3).CeldasR1["F666"]);
        Assert.Equal(134612117.14m, leafs.Single(c => c.Empresa.Id == 4).CeldasR1["F612"]);

        // ENEL/OCCIDENTE: el bloque grande aparece UNA sola vez (no se duplica en otra celda).
        Assert.Equal(3, leafs.Single(c => c.Empresa.Id == 2).CeldasR1.Count);
        Assert.Equal(7233177428.13m, leafs.Single(c => c.Empresa.Id == 2).VisibleR1);
    }

    [Fact]
    public void MagnitudNoOrdinal_Ago2_Ase2_PorPosicionDeSeccion()
    {
        var leafs = LeerConciliacion("Prueba Agosto-2", "2026082", 2);

        // ENEL: primera sección Oportuno (cuadro Componente) = −1.414.859,14 (por POSICIÓN, no por
        // magnitud: el bloque grande es 24.836.019.402,94 y vive en otra fila del mismo cuadro).
        var enel = leafs.Single(c => c.Empresa.Id == 2);
        Assert.Equal(-1414859.14m, enel.CeldasR1["F112"]);
        Assert.Equal(19237658087.83m, enel.VisibleR1);

        // OCCIDENTE: primera sección = −28.650,84 (vs 24.890,00 de su sección Aplicación).
        var occidente = leafs.Single(c => c.Empresa.Id == 4);
        Assert.Equal(-28650.84m, occidente.CeldasR1["F119"]);
        Assert.Equal(2217280433.95m, occidente.VisibleR1);
    }

    [Fact]
    public void Jul2_Ase5_Enel_SeccionUnicaSinDuplicarElBloque()
    {
        var enel = LeerConciliacion("Prueba Julio-2", "2026072", 5).Single(c => c.Empresa.Id == 2);

        Assert.Equal(12823196468.26m, enel.CeldasR1["F633"]);
        Assert.DoesNotContain(enel.CeldasR1.Keys, k => k == "F474"); // el bloque grande NO se duplica
        Assert.Equal(11701157844.90m, enel.VisibleR1);
    }

    [Fact]
    public void VisibleR1_CierraAlValorDelManual_EnLos4CasosArbitrados()
    {
        // Plan 55 (H1.3, D4): el agregado por empresa se deriva de la FUENTE sobre el mismo conjunto de
        // roles del visible de bloque (secciones Oportuno; Especiales de todas menos la última). Los 4
        // valores arbitrados por el T0 del Plan 53 se reproducen al peso tras el cambio de mecanismo.
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

                // Plan 55 (H1.1, D4): el catálogo (5 empresas) + las etiquetas que el propio CUADRO
                // declara y el catálogo no cubre (p. ej. EAB) — nunca una lista por ASE en código.
                Assert.True(leafs.Count >= 5, $"{carpeta}/{codigo} ASE{ase}: se esperaban al menos 5 empresas.");
                Assert.All(leafs, c => Assert.Equal(ase, c.Ase.Id));
                Assert.Contains(leafs, c => c.CeldasR1.Count > 0 || c.CeldasR2.Count > 0 || c.CeldasR4.Count > 0);
            }
        }
    }

    /// <summary>
    /// El mapa congelado por (empresa, ASE) queda LEGACY (Plan 55 H1.2/D7) pero se conserva: este test
    /// verifica que su contrato interno (sección declarada por celda) sigue siendo el del Plan 53, para
    /// los fixtures históricos que lo citan.
    /// </summary>
    [Fact]
    public void ContratoMapaLegacy_DeclaraSeccionPorCelda_EnLas13Claves()
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
        // Plan 55 (H1.1/D4): la PLANTILLA es la autoridad de los destinos (resolución por etiquetas).
        var reader = new ExcelDataReaderWorkbookLeafInputReader { RutaPlantilla = Insumos.PlantillaHistoricaV2() };
        return reader.LeerConciliacionEmpresas(
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
