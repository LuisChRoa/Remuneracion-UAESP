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

        // HU-09 (2.3): cada leaf trae su bloque banco (mapa T0-0.7) con C59 = quincena.
        foreach (var leaf in resultado.Leafs)
        {
            Assert.NotNull(leaf.ReporteBanco);
            Assert.Single(leaf.ReporteBanco!.Ases);
            Assert.Equal(leaf.Ase.Id, leaf.ReporteBanco.Ases[0].Ase.Id);
            Assert.Equal(1, leaf.ReporteBanco.Quincena);
        }

        // HU-16 (D3a): cada leaf trae sus L-Especiales menores (leídas de la fuente R1; el
        // mapa T0-0.5 las declara para los 5 ASE en Q1). INTERVENTORIA queda declarada (D2b)
        // y la hoja intacta — verificado por Capa A (GoldenInterventoriaTests).
        foreach (var leaf in resultado.Leafs)
        {
            Assert.NotNull(leaf.LEspecialesMenores);
            Assert.True(leaf.LEspecialesMenores!.TieneCeldas,
                $"ASE {leaf.Ase.Id}: el mapa D3a Q1 exige celdas L-menores.");
        }

        // HU-10 (2.4): cada leaf trae su fila de balance SC (TotalBsc == TotalFuente; H D2(b)).
        foreach (var leaf in resultado.Leafs)
        {
            Assert.NotNull(leaf.BalanceSc);
            Assert.Single(leaf.BalanceSc!.Ases);
            Assert.Equal(leaf.Ase.Id, leaf.BalanceSc.Ases[0].Ase.Id);
            Assert.Null(leaf.BalanceSc.Ases[0].Sistema);
            Assert.InRange(
                leaf.BalanceSc.Ases[0].TotalBsc - leaf.BalanceSc.Ases[0].TotalFuente,
                -0.5m, 0.5m);
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
    public void Ejecutar_FaltaEtiquetaResumenEnAse3_FallaNombrandoAseYSinSalida()
    {
        // HU-09 (2.3, Requirement 1 / §4.3): si la fuente banco del ASE3 no trae la etiqueta
        // del Resumen, fail-fast nombra el ASE y NO hay salida certificada.
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-banco-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-banco-carpeta-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodo, carpetaPeriodo);

        // Reemplaza el ReportePagosxBanco de ASE3 por un archivo que no trae la etiqueta
        // (usa el R1 de ASE3 renombrado: mantiene el prefijo para que el locator lo encuentre).
        var carpetaAse3 = Path.Combine(carpetaPeriodo, "3-Ciudad Limpia");
        var bancoAse3 = Directory.EnumerateFiles(carpetaAse3, "*.xlsx", SearchOption.TopDirectoryOnly)
            .First(f => Path.GetFileNameWithoutExtension(f).StartsWith("ReportePagosxBanco", StringComparison.OrdinalIgnoreCase));
        File.Delete(bancoAse3);
        File.Copy(Insumos.R1(3), Path.Combine(carpetaAse3, "ReportePagosxBanco_sin_resumen.xlsx"));

        var procesador = CrearProcesador();
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.Periodo(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.Plantilla,
                RutaSalida = salida
            }));

        Assert.Contains("ASE 3", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Resumen Recaudo Aplicado Por Servicio", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(salida), "No debe existir salida certificada ante fallo.");
    }

    [Fact]
    public void Ejecutar_FaltaFuenteBancoEnAse2_FallaNombrandoAseYSinSalida()
    {
        // HU-09 (2.3, V13): si falta el ReportePagosxBanco de un ASE, fail-fast nombra el ASE.
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-bancorf-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-bancorf-carpeta-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodo, carpetaPeriodo);

        var carpetaAse2 = Path.Combine(carpetaPeriodo, "2-Lime");
        var bancoAse2 = Directory.EnumerateFiles(carpetaAse2, "*.xlsx", SearchOption.TopDirectoryOnly)
            .First(f => Path.GetFileNameWithoutExtension(f).StartsWith("ReportePagosxBanco", StringComparison.OrdinalIgnoreCase));
        File.Delete(bancoAse2);

        var procesador = CrearProcesador();
        var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.Periodo(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.Plantilla,
                RutaSalida = salida
            }));

        Assert.Contains("ReportePagosxBanco", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("ASE 2", ex.Message, StringComparison.OrdinalIgnoreCase);
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
    public void Ejecutar_FaltaTotalGeneralEnFuenteAse3_FallaNombrandoAseYSinSalida()
    {
        // HU-10 (2.4, Requirement 1 / §4.3): si la fuente Balance del ASE3 no trae la etiqueta
        // "Total General", fail-fast nombra el ASE y NO hay salida certificada.
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-bce-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-bce-carpeta-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodo, carpetaPeriodo);

        // Reemplaza el Balance de ASE3 por un archivo que no trae "Total General" (usa el R1 de
        // ASE3 renombrado: mantiene el prefijo para que el locator lo encuentre).
        var carpetaAse3 = Path.Combine(carpetaPeriodo, "3-Ciudad Limpia");
        var balanceAse3 = Directory.EnumerateFiles(carpetaAse3, "*.xlsx", SearchOption.TopDirectoryOnly)
            .First(f => Path.GetFileNameWithoutExtension(f).StartsWith("R4-BalanceSubsidioyContribuciones", StringComparison.OrdinalIgnoreCase));
        File.Delete(balanceAse3);
        File.Copy(Insumos.R1(3), Path.Combine(carpetaAse3, "R4-BalanceSubsidioyContribuciones_sin_total_general.xlsx"));

        var procesador = CrearProcesador();
        var ex = Assert.Throws<CalculoInvalidoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.Periodo(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.Plantilla,
                RutaSalida = salida
            }));

        Assert.Contains("ASE 3", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Total General", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(salida), "No debe existir salida certificada ante fallo.");
    }

    [Fact]
    public void Ejecutar_FaltaFuenteBalanceEnAse2_FallaNombrandoAseYSinSalida()
    {
        // HU-10 (2.4, V8): si falta el R4-BalanceSubsidioyContribuciones de un ASE, fail-fast
        // nombra el ASE y NO hay salida certificada.
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-bcerf-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-bcerf-carpeta-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodo, carpetaPeriodo);

        var carpetaAse2 = Path.Combine(carpetaPeriodo, "2-Lime");
        var balanceAse2 = Directory.EnumerateFiles(carpetaAse2, "*.xlsx", SearchOption.TopDirectoryOnly)
            .First(f => Path.GetFileNameWithoutExtension(f).StartsWith("R4-BalanceSubsidioyContribuciones", StringComparison.OrdinalIgnoreCase));
        File.Delete(balanceAse2);

        var procesador = CrearProcesador();
        var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.Periodo(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.Plantilla,
                RutaSalida = salida
            }));

        Assert.Contains("R4-BalanceSubsidioyContribuciones", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("ASE 2", ex.Message, StringComparison.OrdinalIgnoreCase);
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

    // ── HU-11 (2.5) / HU-12 (2.6 ampliada): Q2 ─────────────────────────────────────────────────

    [Fact]
    public void Ejecutar_Periodo2026072_SalidaCertificada5De5_ConDetRetri()
    {
        // HU-12 (2.6 ampliada, §4.4): el dispatch Q2 del reader (mapa <see cref="WorkbookLeafCellMapQ2"/>
        // con variante ASE5 de 2 filas Mes/Total, V0.3) LEVANTA el recorte T0-0.6 de HU-11: el
        // procesador Q2 es certificable 5/5 end-to-end (R1/R2/R4-Q2 por ASE + DetRetri-Q2 en la
        // misma pasada). A8/M1: el path Q2 del writer se ejercita contra el canónico real.
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-q2-5de5-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.PeriodoQ2().NombreArchivo);

        var procesador = CrearProcesador();
        var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.PeriodoQ2(),
            CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
            RutaPlantilla = Insumos.PlantillaQ2,
            RutaSalida = salida
        });

        Assert.True(File.Exists(salida), "Debe existir la salida certificada Q2.");
        Assert.Equal(5, resultado.Resultado.Consolidados.Count);
        Assert.Equal(5, resultado.Leafs.Count);
        Assert.All(resultado.Leafs, l => Assert.NotNull(l.DetRetriQ2));
        Assert.All(resultado.Leafs, l => Assert.NotNull(l.AjustesSfT));
        Assert.All(resultado.Leafs, l => Assert.Empty(l.Conciliacion)); // recorte 1 HU-11 (V0.6)
    }

    [Fact]
    public void Ejecutar_Q2_FaltaSaldosNotasEnAse3_FallaNombrandoAseYSinSalida()
    {
        // HU-11 Requirement 2 / §4.3: si falta la fuente SALDOS POR NOTA de un ASE en Q2,
        // fail-fast nombra el ASE y NO hay salida certificada.
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-q2-sn-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.PeriodoQ2().NombreArchivo);

        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-q2-sn-carpeta-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodoQ2, carpetaPeriodo);

        var carpetaAse3 = Path.Combine(carpetaPeriodo, "3-Ciudad Limpia");
        var saldosNotasAse3 = Directory.EnumerateFiles(carpetaAse3, "*.xlsx", SearchOption.TopDirectoryOnly)
            .First(f => Path.GetFileNameWithoutExtension(f).StartsWith("SaldosaFavorAplicadosPorNotas", StringComparison.OrdinalIgnoreCase));
        File.Delete(saldosNotasAse3);

        var procesador = CrearProcesador();
        var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.PlantillaQ2,
                RutaSalida = salida
            }));

        Assert.Contains("SaldosaFavorAplicadosPorNotas", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("ASE 3", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(salida), "No debe existir salida certificada ante fallo.");
    }

    [Fact]
    public void Ejecutar_Q2_FaltaRetribucionNegativaEnAse4_FallaNombrandoAseYSinSalida()
    {
        // HU-11 Requirement 2 / D4: si falta la fuente RETRIBUCION NEGATIVA de un ASE en Q2,
        // fail-fast nombra el ASE (el matcher es agnóstico a diacríticos: el archivo real se
        // llama RetribuciónNegativa_…).
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-q2-rn-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.PeriodoQ2().NombreArchivo);

        var carpetaPeriodo = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-q2-rn-carpeta-" + Guid.NewGuid().ToString("N"));
        CopiarArbol(Insumos.CarpetaPeriodoQ2, carpetaPeriodo);

        var carpetaAse4 = Path.Combine(carpetaPeriodo, "4-Bogota Limpia");
        var retribucionAse4 = Directory.EnumerateFiles(carpetaAse4, "*.xlsx", SearchOption.TopDirectoryOnly)
            .First(f => Path.GetFileNameWithoutExtension(f).StartsWith("Retribuci", StringComparison.OrdinalIgnoreCase));
        File.Delete(retribucionAse4);

        var procesador = CrearProcesador();
        var ex = Assert.Throws<ArchivoFuenteNoEncontradoException>(() =>
            procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = Insumos.PeriodoQ2(),
                CarpetaPeriodo = carpetaPeriodo,
                RutaPlantilla = Insumos.PlantillaQ2,
                RutaSalida = salida
            }));

        Assert.Contains("RetribuciónNegativa", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("ASE 4", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(salida), "No debe existir salida certificada ante fallo.");
    }

    [Fact]
    public void Ejecutar_Periodo2026071_ConLectorOracle_VeredictosValidacionesNoVacios()
    {
        // HU-13 (2.7, §4 Fase 3 — Unidad 3): con lector-oráculo, el procesador lee el snapshot de
        // la salida (read-only) y evalúa los gates 2.7. Q1 plantilla == canónico con caché real →
        // veredictos por ASE sin errores.
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-oracle-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.Periodo().NombreArchivo);

        var procesador = CrearProcesadorConOracle();
        var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.Periodo(),
            CarpetaPeriodo = Insumos.CarpetaPeriodo,
            RutaPlantilla = Insumos.Plantilla,
            RutaSalida = salida
        });

        Assert.True(File.Exists(salida), "Debe existir la salida certificada.");
        Assert.NotEmpty(resultado.Validaciones);
        Assert.Contains(resultado.Validaciones, l => l.Contains("ASE 1 VALIDACIONES", StringComparison.Ordinal));
    }

    [Fact]
    public void Ejecutar_Periodo2026072_ConLectorOracle_VeredictosValidacionesNoVacios()
    {
        // HU-13 (2.7): Q2 end-to-end con lector-oráculo (plantilla canónica "8 agos" con caché 0).
        // Los gates 2.7 leen caché visible (ceros legítimos de plantilla); veredictos por ASE.
        var salidaDir = Path.Combine(Path.GetTempPath(), "remuneracion-periodo-oracle-q2-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        var salida = Path.Combine(salidaDir, Insumos.PeriodoQ2().NombreArchivo);

        var procesador = CrearProcesadorConOracle();
        var resultado = procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = Insumos.PeriodoQ2(),
            CarpetaPeriodo = Insumos.CarpetaPeriodoQ2,
            RutaPlantilla = Insumos.PlantillaQ2,
            RutaSalida = salida
        });

        Assert.True(File.Exists(salida), "Debe existir la salida certificada.");
        Assert.NotEmpty(resultado.Validaciones);
        Assert.Equal(5, resultado.Validaciones.Count(l => l.Contains("VALIDACIONES:", StringComparison.Ordinal)));
    }

    private static ProcesadorPeriodo CrearProcesador() =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator());

    private static ProcesadorPeriodo CrearProcesadorConOracle() =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ValidacionOracleReader());
}
