using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-07 Req 5 / plan §2.4: orquestación de período con las 5 carpetas reales Q1,
/// salida temporal, fail-fast sin salida certificada ante fallo.
/// </summary>
public sealed class ProcesadorPeriodoTests
{
    [Fact]
    public void CalcularConsolidado_CincoTuplasReales_GranTotalSumaYDuplicadosFallan()
    {
        var calculo = new CalculoRemuneracion();
        var lector = new ExcelDataReaderRecaudoReader();
        var periodo = Insumos.Periodo();

        var datos = Enumerable.Range(1, 5).Select(i =>
        {
            var r1 = lector.LeerR1(Insumos.R1(i));
            var r2 = lector.LeerR2(Insumos.R2(i));
            var r4 = lector.LeerR4(Insumos.R4(i));
            return (ase: Insumos.Ase(i), r1, r2, r4);
        }).ToList();

        var resultado = calculo.CalcularConsolidado(periodo, datos);
        Assert.Equal(5, resultado.Consolidados.Count);
        Assert.Equal(resultado.Consolidados.Sum(c => c.TotalAse), resultado.GranTotal);

        // Duplicados fallan (Req 1).
        var conDuplicado = datos.Append(datos[0]).ToList();
        Assert.Throws<CalculoInvalidoException>(() => calculo.CalcularConsolidado(periodo, conDuplicado));
    }

    [Fact]
    public void Ejecutar_Periodo2026071_SalidaGeneradaCon5Consolidados()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var procesador = CrearProcesador();
        var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.Periodo(),
            CarpetaPeriodo = Insumos.CarpetaPeriodo,
            RutaPlantilla = Insumos.Plantilla,
            RutaSalida = salida
        });

        Assert.True(File.Exists(salida), "Debe existir el archivo de salida.");
        Assert.Equal(5, resultado.Resultado.Consolidados.Count);
        Assert.Equal(5, resultado.Leafs.Count);
        Assert.Equal([1, 2, 3, 4, 5], resultado.Leafs.OrderBy(l => l.Ase.Id).Select(l => l.Ase.Id).ToArray());

        // HU-08 (2.2): cada leaf trae las 5 empresas de facturación y las 5 hojas Recaudo *.
        foreach (var leaf in resultado.Leafs)
        {
            Assert.Equal(5, leaf.Conciliacion.Count);
            Assert.Equal(5, leaf.Recaudos.Count);
        }

        // GranTotal de dominio = suma de TotalAse (±0.5) — §2.5 regla 5 (invariante interno).
        Assert.InRange(
            resultado.Resultado.GranTotal - resultado.Resultado.Consolidados.Sum(c => c.TotalAse),
            -0.5m, 0.5m);

        // Golden GranTotal post-Excel D109 = 58210094820.50 (cache del golden, §2.1).
        // El GranTotal honesto del CONSOLIDADO es Σ visibles leaf por ASE (nunca agregados HU-02).
        var granTotalPostExcel = resultado.Leafs.Sum(l =>
            l.R1.TotalOportunoEsperadoPorAse
            + l.R2.TotalOportunoEsperado
            + l.R1.ExtemporaneoEsperadoPorAse
            + l.R4.TotalReversionEsperada);
        Assert.InRange(granTotalPostExcel - 58210094820.50m, -0.5m, 0.5m);
    }

    [Fact]
    public void Ejecutar_FaltaConciliacionDeEnerbit_FallaNombrandoEmpresaYSinSalida()
    {
        // HU-08 (2.2): si falta el archivo de conciliación de una empresa, fail-fast nombra la
        // empresa y NO hay salida certificada (T0-0.6: Recaudo * ← Consolidado/Conciliaciones).
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-failconc-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-conc-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodo, carpetaPeriodo);
        var conciliaciones = Path.Combine(carpetaPeriodo, "Consolidado", "Conciliaciones");
        var enerbit = Directory.EnumerateFiles(conciliaciones, "*.xlsx", SearchOption.TopDirectoryOnly)
            .First(f => Path.GetFileNameWithoutExtension(f).StartsWith("Conjunta ENERBIT", StringComparison.OrdinalIgnoreCase));
        File.Delete(enerbit);

        var procesador = CrearProcesador();
        var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.Periodo(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.Plantilla,
                RutaSalida = salida
            }));

        Assert.Contains("Enerbit", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(salida), "No debe existir salida certificada ante fallo.");
    }

    private static void CopiarArbol(string origen, string destino)
    {
        Directory.CreateDirectory(destino);
        foreach (var archivo in Directory.EnumerateFiles(origen, "*.xlsx", SearchOption.AllDirectories))
        {
            var relativo = Path.GetRelativePath(origen, archivo);
            var destinoArchivo = Path.Combine(destino, relativo);
            Directory.CreateDirectory(Path.GetDirectoryName(destinoArchivo)!);
            File.Copy(archivo, destinoArchivo);
        }
    }

    [Fact]
    public void Ejecutar_FaltaCarpetaDeUnAse_FallaSinSalida()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-fail-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-carpeta-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(carpetaPeriodo);
        // Copiamos fuentes reales de 4 ASE; omitimos la carpeta 3 (Ciudad Limpia).
        foreach (var aseId in new[] { 1, 2, 4, 5 })
        {
            var origen = Insumos.CarpetasAse[aseId - 1];
            var destino = Path.Combine(carpetaPeriodo, Path.GetFileName(origen));
            Directory.CreateDirectory(destino);
            foreach (var archivo in Directory.EnumerateFiles(origen, "*.xlsx", SearchOption.TopDirectoryOnly))
            {
                File.Copy(archivo, Path.Combine(destino, Path.GetFileName(archivo)));
            }
        }

        var procesador = CrearProcesador();
        var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.Periodo(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.Plantilla,
                RutaSalida = salida
            }));

        Assert.Contains("ASE 3", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(salida), "No debe existir salida certificada ante fallo.");
    }

    [Fact]
    public void Ejecutar_FaltaR4EnAse4_FallaNombrandoElAseYSinSalida()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-failr4-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        // Copiamos las 5 carpetas reales pero sin R4 en ASE4.
        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-r4-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(carpetaPeriodo);
        for (var i = 1; i <= 5; i++)
        {
            var origen = Insumos.CarpetasAse[i - 1];
            var destino = Path.Combine(carpetaPeriodo, Path.GetFileName(origen));
            Directory.CreateDirectory(destino);
            foreach (var archivo in Directory.EnumerateFiles(origen, "*.xlsx", SearchOption.TopDirectoryOnly))
            {
                var nombre = Path.GetFileName(archivo);
                if (i == 4 && (nombre.StartsWith("ReversiónPorComponente", StringComparison.OrdinalIgnoreCase)
                    || nombre.StartsWith("ReversionPorComponente", StringComparison.OrdinalIgnoreCase)))
                {
                    continue; // quita R4 de ASE4
                }

                File.Copy(archivo, Path.Combine(destino, nombre));
            }
        }

        var procesador = CrearProcesador();
        var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.Periodo(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.Plantilla,
                RutaSalida = salida
            }));

        Assert.Contains("R4", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("ASE 4", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(salida), "No debe existir salida certificada ante fallo.");
    }

    [Fact]
    public void WriterMultiAse_GateDeCoherenciaRotoEnAse3_NoGeneraArchivo()
    {
        // Req 3 / §2.3: el writer multi-ASE valida coherencia por ASE con matcheo estricto ANTES
        // de escribir; si un leaf no cuadra con su consolidado, no hay archivo certificado.
        var lector = new ExcelDataReaderRecaudoReader();
        var leafReader = new ExcelDataReaderWorkbookLeafInputReader();
        var calculo = new CalculoRemuneracion();
        var periodo = Insumos.Periodo();

        var datos = Enumerable.Range(1, 5).Select(i =>
        {
            var r1 = lector.LeerR1(Insumos.R1(i));
            var r2 = lector.LeerR2(Insumos.R2(i));
            var r4 = lector.LeerR4(Insumos.R4(i));
            return (ase: Insumos.Ase(i), r1, r2, r4);
        }).ToList();
        var resultado = calculo.CalcularConsolidado(periodo, datos);

        var leafs = datos.Select(d =>
            leafReader.LeerLeafInputs(d.ase, periodo, Insumos.R1(d.ase.Id), Insumos.R2(d.ase.Id), Insumos.R4(d.ase.Id))).ToList();

        // Rompemos la coherencia SOLO del ASE 3.
        leafs[2].R2.E15 += 100m;

        var salida = Path.Combine(Path.GetTempPath(), "remuneracion-writer-gate-" + Guid.NewGuid().ToString("N"), "salida.xlsx");

        Assert.Throws<CalculoInvalidoException>(() =>
            new OpenXmlPlantillaWriter().GenerarWorkbook(Insumos.Plantilla, salida, resultado, leafs));

        Assert.False(File.Exists(salida), "No debe existir salida certificada ante gate roto.");
    }

    private static ProcesadorPeriodo CrearProcesador() =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator());
}