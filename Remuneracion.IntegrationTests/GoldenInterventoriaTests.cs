using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-16 (Plan 16 §2.7): Golden Capa A HU-16 contra AMBOS canónicos (Q1 golden y Q2 canónico
/// "Plantilla 8 agos…" con caché golden <c>Remuneracion 202607-2 Total.xlsx</c>).
/// Matriz: A1 (L-menores D3a escritas en la salida == leaf del golden ±0.5), A2 (dominio
/// L-menores == caché golden), A3 (INTERVENTORIA intacta: totales en fórmula, bloque por ASE en
/// valores DECLARADOS sin mutación; D3b siguen 0), A4 (SHA256 de ambos canónicos sin cambios),
/// A8 (stale-guard de ceros — cubierto por <see cref="InterventoriaTests.D3b…"/>). A5: este test
/// jamás compara caché de fórmula de la salida contra golden ni usa agregados HU-02 como oráculo.
/// </summary>
public sealed class GoldenInterventoriaTests
{
    private const decimal Tolerancia = Insumos.Tolerancia;

    [Fact]
    public void CapaA_InterventoriaYLEspeciales_AmbosCanonicos()
    {
        var hashQ1 = Sha256(Insumos.Plantilla);
        var hashQ2 = Sha256(Insumos.PlantillaQ2);

        foreach (var (periodo, carpeta, plantilla, goldenCache, etiqueta) in new[]
                 {
                     (Insumos.Periodo(), Insumos.CarpetaPeriodo, Insumos.Plantilla, Insumos.Plantilla, "Q1"),
                     (Insumos.PeriodoQ2(), Insumos.CarpetaPeriodoQ2, Insumos.PlantillaQ2, Insumos.GoldenQ2, "Q2")
                 })
        {
            var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-hu16-" + etiqueta + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(salidaDir);
            var salida = Path.Combine(salidaDir, periodo.NombreArchivo);

            var procesador = new ProcesadorPeriodo(
                new ExcelDataReaderRecaudoReader(),
                new ExcelDataReaderWorkbookLeafInputReader(),
                new CalculoRemuneracion(),
                new ValidadorBasico(),
                new OpenXmlPlantillaWriter(),
                new ArchivoFuenteLocator());

            var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = periodo,
                CarpetaPeriodo = carpeta,
                RutaPlantilla = plantilla,
                RutaSalida = salida
            });

            Assert.True(File.Exists(salida));

            // ---- A1: L-menores D3a escritas en la SALIDA == leaf del golden ±0.5 ----
            foreach (var leaf in resultado.Leafs.OrderBy(l => l.Ase.Id))
            {
                Assert.NotNull(leaf.LEspecialesMenores);
                foreach (var (celda, _) in leaf.LEspecialesMenores!.Celdas)
                {
                    var escrito = TestHelpers.LeerCeldaNumerica(salida, WorkbookLeafCellMapInterventoria.HojaR1, celda);
                    var golden = TestHelpers.LeerCeldaNumerica(goldenCache, WorkbookLeafCellMapInterventoria.HojaR1, celda);
                    Assert.InRange(escrito - golden, -Tolerancia, Tolerancia);
                }
            }

            // ---- A2: dominio L-menores == caché golden ±0.5 (aritmética D3a) ----
            foreach (var leaf in resultado.Leafs.OrderBy(l => l.Ase.Id))
            {
                foreach (var (celda, valor) in leaf.LEspecialesMenores!.Celdas)
                {
                    var golden = TestHelpers.LeerCeldaNumerica(goldenCache, WorkbookLeafCellMapInterventoria.HojaR1, celda);
                    Assert.InRange(valor - golden, -Tolerancia, Tolerancia);
                }
            }

            // ---- A3: INTERVENTORIA intacta (nunca escrita) ----
            // Totales siguen en fórmula; el bloque por ASE conserva los valores DECLARADOS
            // (copiados del canónico; el writer no los toca — prueba de no-mutación).
            foreach (var (_, celda, _) in WorkbookLeafCellMapInterventoria.FormulasProtegidas)
            {
                Assert.True(TestHelpers.CeldaEsFormula(salida, InterventoriaDeclarada.Hoja, celda), $"{InterventoriaDeclarada.Hoja}!{celda} debió seguir siendo fórmula.");
            }

            for (var i = 0; i < 5; i++)
            {
                var aseId = i + 1;
                var fila = InterventoriaDeclarada.FilaPrimerAse + i;
                Assert.InRange(TestHelpers.LeerCeldaNumerica(salida, InterventoriaDeclarada.Hoja, $"K{fila}") - InterventoriaDeclarada.ValorOficialMesPorAse[aseId], -Tolerancia, Tolerancia);
                Assert.InRange(TestHelpers.LeerCeldaNumerica(salida, InterventoriaDeclarada.Hoja, $"M{fila}") - InterventoriaDeclarada.SegundaQuincenaPorAse[aseId], -Tolerancia, Tolerancia);
                Assert.InRange(TestHelpers.LeerCeldaNumerica(salida, InterventoriaDeclarada.Hoja, $"N{fila}") - InterventoriaDeclarada.PrimeraQuincenaPorAse[aseId], -Tolerancia, Tolerancia);
            }

            // D3b: las L numéricas fuera del mapa D3a siguen en 0 en la SALIDA (no escritas).
            foreach (var (celda, valor) in TestHelpers.EnumerarCeldaLNumericas(salida, WorkbookLeafCellMapInterventoria.HojaR1))
            {
                if (MapeadaEnPeriodo(celda, periodo.NumeroQuincena))
                {
                    continue;
                }

                Assert.True(Math.Abs(valor) <= Tolerancia,
                    $"{etiqueta} salida {WorkbookLeafCellMapInterventoria.HojaR1}!{celda} fuera del set mapeado vale {valor} (D3b/A8: debe seguir 0).");
            }
        }

        // ---- A4: canónicos no mutados ----
        Assert.Equal(hashQ1, Sha256(Insumos.Plantilla));
        Assert.Equal(hashQ2, Sha256(Insumos.PlantillaQ2));
    }

    private static bool MapeadaEnPeriodo(string celda, int numeroQuincena)
    {
        // Mismo set que InterventoriaTests.CeldasLMenoresMapeadas (V4 + HU-08 + Q2-map + D3a).
        if (new[] { "L25", "L10", "L113", "L90", "L217", "L238", "L223", "L357", "L391", "L369", "L478", "L498" }
            .Contains(celda, StringComparer.OrdinalIgnoreCase))
        {
            return true;
        }

        foreach (var entrada in WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.Values)
        {
            foreach (var (c, _) in entrada)
            {
                if (c.StartsWith("L", StringComparison.OrdinalIgnoreCase) && string.Equals(c, celda, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        if (numeroQuincena == 2)
        {
            foreach (var entrada in WorkbookLeafCellMapQ2.R1Q2EditablesPorAse.Values)
            {
                foreach (var (c, _) in entrada)
                {
                    if (c.StartsWith("L", StringComparison.OrdinalIgnoreCase) && string.Equals(c, celda, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
        }

        foreach (var mapa in numeroQuincena == 2
                     ? WorkbookLeafCellMapInterventoria.LMenoresPorAseQ2.Values
                     : WorkbookLeafCellMapInterventoria.LMenoresPorAse.Values)
        {
            foreach (var (c, _, _) in mapa)
            {
                if (string.Equals(c, celda, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static string Sha256(string ruta)
    {
        using var stream = File.OpenRead(ruta);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}
