using System.IO;
using System.Text;
using Remuneracion.Cli;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Serilog;
using Serilog.Core;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// PLAN 61 — H3/F3 (política de entrega con novedades) + H4/F4 (reporte TXT y superficies).
///
/// <para><b>Requisito del Ingeniero (literal).</b> «Ante una novedad, SIEMPRE generar el Excel para
/// poder revisar y comparar; el que use la App puede tomar la salida y ajustar si es un detalle menor
/// mientras la App se ajusta e integra. Generar el Excel más un TXT fácil de leer y entender de que
/// hubo novedades y que se deben revisar.»</para>
///
/// <para><b>Clases del §2-D-H3.</b> Divergencia de DATOS con workbook escribible = novedad (Excel
/// marcado <c>(sin validar)</c> + TXT + estado propio + código 7). Imposibilidad de construir el
/// artefacto = fallo duro (comportamiento anterior).</para>
/// </summary>
public sealed class EntregaConNovedadesH3Plan61Tests : IDisposable
{
    private readonly StringWriter _stdout = new();
    private readonly StringWriter _stderr = new();

    public void Dispose()
    {
        Log.CloseAndFlush();
        Log.Logger = Logger.None;
        _stdout.Dispose();
        _stderr.Dispose();
    }

    private static string Raiz => RaizRepositorio();

    private static string RutaPlantillaV4 => Path.Combine(
        Raiz, "Docs", "Plantilla estandarizada", "v4", "Formato Remuneracion 2026xx-x Total v4.xlsx");

    private static string CarpetaSeptiembre1 => Path.Combine(Raiz, "Docs", "Prueba Septiembre 1", "Insumos");
    private static string CarpetaSeptiembre2 => Path.Combine(Raiz, "Docs", "Prueba Septiembre-2", "Insumos");

    private static Periodo PeriodoSeptiembre1 => new() { CodigoAAAAMM = "202609", NumeroQuincena = 1 };
    private static Periodo PeriodoSeptiembre2 => new() { CodigoAAAAMM = "202609", NumeroQuincena = 2 };

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // H3.5 / H4.4 — corridas reales
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 2026091 (5 ASE): con el R10 del período alineado con la composición del dominio y el pegado
    /// del Plan 62 (N1), la corrida cierra CERTIFICADA — nombre limpio, sin reporte de novedades y
    /// sin la ficha R10 <c>-1</c>. Reescritura formal del rojo de Plan 61 H3.5: el R10 vigente
    /// coincide per-ASE con el cálculo y, por T3.2 (D-61R10), la novedad de TOTAL-R10 dejó de
    /// emitirse. CA-3: la corrida de septiembre-1 no reporta diferencias del asiento.
    /// </summary>
    [Fact]
    public void H35_2026091_SinNovedadR10_CierraCertificada()
    {
        var resultado = EspecialesR4H2Plan61Tests.Generacion2026091.Value.Resultado;

        Assert.Equal(EstadoEntrega.Certificada, resultado.EstadoEntrega);
        Assert.Empty(resultado.Novedades);
        Assert.Equal(string.Empty, resultado.RutaReporteNovedades);
        Assert.True(File.Exists(resultado.RutaSalida));

        // NO queda artefacto marcado: no hay nada que revisar.
        Assert.DoesNotContain(NombresSalida.SufijoSinValidar, Path.GetFileName(resultado.RutaSalida), StringComparison.OrdinalIgnoreCase);

        // CA-3 explícito: ni la ficha per-ASE ni la de TOTAL del R10 (el "-1" de composición).
        Assert.DoesNotContain(resultado.Novedades, n => n.Delta == -1m);
        Assert.DoesNotContain(resultado.Novedades, n => n.ConceptoOCelda == "Total del período");
        Assert.DoesNotContain(resultado.Novedades, n => n.Hoja.Contains("DetRetri2026091", StringComparison.Ordinal));

        var directorio = Path.GetDirectoryName(resultado.RutaSalida)!;
        var nombreBase = Path.GetFileNameWithoutExtension(resultado.RutaSalida);
        Assert.False(
            File.Exists(Path.Combine(directorio, nombreBase + ReporteNovedadesWriter.SufijoNovedades)),
            "Una corrida certificada NO debe dejar reporte de novedades.");
    }

    /// <summary>
    /// 2026092 (5 ASE): la corrida sigue CERTIFICADA — nombre limpio, sin reporte de novedades.
    /// Es el control que demuestra que la política nueva no tocó el cierre bueno.
    /// </summary>
    [Fact]
    public void H35_2026092_SigueCertificada_SinReporteDeNovedades()
    {
        var resultado = EspecialesR4H2Plan61Tests.Generacion2026092.Value;

        Assert.Equal(EstadoEntrega.Certificada, resultado.EstadoEntrega);
        Assert.Empty(resultado.Novedades);
        Assert.Equal(string.Empty, resultado.RutaReporteNovedades);
        Assert.True(File.Exists(resultado.RutaSalida));
        Assert.DoesNotContain(NombresSalida.SufijoSinValidar, Path.GetFileName(resultado.RutaSalida), StringComparison.OrdinalIgnoreCase);

        var directorio = Path.GetDirectoryName(resultado.RutaSalida)!;
        var nombreBase = Path.GetFileNameWithoutExtension(resultado.RutaSalida);
        Assert.False(
            File.Exists(Path.Combine(directorio, nombreBase + ReporteNovedadesWriter.SufijoNovedades)),
            "Una corrida certificada NO debe dejar reporte de novedades.");
    }

    /// <summary>
    /// H4.2 / CA-C4: la superficie CLI del cierre con novedades. Código 7 (distinto de 0 y de 6),
    /// línea grepable con las dos rutas y NUNCA <c>RESULTADO OK</c>.
    /// </summary>
    [Fact]
    public void H42_Cli_EntregaConNovedades_Codigo7SinResultadoOk()
    {
        var salidaDir = Path.Combine(Path.GetTempPath(), "p61-h3-cli-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(salidaDir);
        try
        {
            // Plan 62 (H3/T3.2): septiembre-1 ya no produce novedades propias (R10 alineado + ficha
            // de TOTAL retirada); el contrato de la superficie CLI con novedades se ejerce con una
            // divergencia de coherencia inyectada por la costura de test de <c>EjecutorCli</c>
            // (misma composición de producción para todo lo demás).
            var codigo = EjecutorCli.Ejecutar(
                ["--periodo", "2026091", "--carpeta", CarpetaSeptiembre1, "--plantilla", RutaPlantillaV4,
                    "--salida", salidaDir, "--cinco-ase", "--sobrescribir"],
                _stdout,
                _stderr,
                validador: new ValidadorConNovedadForzada(),
                lectorLeaf: null);

            Assert.Equal(7, codigo);
            Assert.Equal(CodigosSalida.NovedadesEntrega, codigo);
            Assert.NotEqual(CodigosSalida.Ok, codigo);
            Assert.NotEqual(CodigosSalida.DivergenciaDeclaradaReconocida, codigo);

            var stdout = _stdout.ToString();
            var salidaMarcada = Path.Combine(salidaDir, "Remuneración 202609-1 Total (sin validar).xlsx");
            var reporte = Path.Combine(salidaDir, "Remuneración 202609-1 Total - Novedades de la corrida.txt");

            Assert.Contains("RESULTADO CON NOVEDADES codigo=7", stdout, StringComparison.Ordinal);
            Assert.Contains($"salida={salidaMarcada}", stdout, StringComparison.Ordinal);
            Assert.Contains($"reporte={reporte}", stdout, StringComparison.Ordinal);
            Assert.Contains("runId=", stdout, StringComparison.Ordinal);
            Assert.DoesNotContain("RESULTADO OK", stdout, StringComparison.Ordinal);

            Assert.True(File.Exists(salidaMarcada));
            Assert.False(File.Exists(Path.Combine(salidaDir, "Remuneración 202609-1 Total.xlsx")));
            Assert.True(File.Exists(reporte));
        }
        finally
        {
            if (Directory.Exists(salidaDir))
            {
                Directory.Delete(salidaDir, recursive: true);
            }
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // H3.4 / CA-C3 — anti-falso-verde (tres injertos, los tres tienen que doler)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Injerto 1 — «novedad SIN TXT ⇒ rojo». Dos formas: (a) composición sin emisor; (b) emisor que
    /// no deja archivo. En ambos casos la corrida NO puede devolver un resultado entregado.
    ///
    /// <para><b>Reescritura Plan 62 (H3/T3.2):</b> septiembre-1 ya no produce novedades propias (el
    /// R10 vigente cuadra per-ASE y la ficha de TOTAL se retiró), así que el contrato anti-falso-verde
    /// se ejerce con UNA divergencia de coherencia inyectada por la costura pública
    /// <see cref="IValidador"/>: la novedad existe de verdad y la ausencia de reporte debe doler.</para>
    /// </summary>
    [Fact]
    public void H34_NovedadSinTxt_EsRojo()
    {
        // (a) Sin emisor configurado: fail-closed, nunca un Excel marcado huérfano de explicación.
        var sinEmisor = CrearProcesadorSinEmisorDeNovedades(novedadForzada: true);
        var dirA = Path.Combine(Path.GetTempPath(), "p61-h3-sintxt-a-" + Guid.NewGuid().ToString("N"));
        var salidaA = Path.Combine(dirA, "Remuneración 202609-1 Total.xlsx");
        try
        {
            var ex = Assert.Throws<CalculoInvalidoException>(() => sinEmisor.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = PeriodoSeptiembre1,
                CarpetaPeriodo = CarpetaSeptiembre1,
                RutaPlantilla = RutaPlantillaV4,
                RutaSalida = salidaA
            }));

            Assert.Equal(CodigoError.Inesperado, ex.Codigo);
            Assert.Contains("reporte de novedades", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(salidaA), "Sin reporte no se promueve al nombre limpio.");
        }
        finally
        {
            Borrar(dirA);
        }

        // (b) Emisor injertado que miente (devuelve una ruta que no escribe).
        var dirB = Path.Combine(Path.GetTempPath(), "p61-h3-sintxt-b-" + Guid.NewGuid().ToString("N"));
        var salidaB = Path.Combine(dirB, "Remuneración 202609-1 Total.xlsx");
        try
        {
            var ex = Assert.Throws<CalculoInvalidoException>(() => CrearProcesador(new EmisorQueNoEscribe(), novedadForzada: true).Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = PeriodoSeptiembre1,
                CarpetaPeriodo = CarpetaSeptiembre1,
                RutaPlantilla = RutaPlantillaV4,
                RutaSalida = salidaB
            }));

            Assert.Equal(CodigoError.Inesperado, ex.Codigo);
            Assert.Contains("no existe el reporte de novedades", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(File.Exists(salidaB), "Sin reporte no se promueve al nombre limpio.");
        }
        finally
        {
            Borrar(dirB);
        }
    }

    /// <summary>
    /// Injerto 2 — «novedad con nombre limpio ⇒ rojo». La guarda de entrega rechaza un resultado con
    /// novedades cuya ruta no lleva la marca: es imposible publicar un archivo limpio con deuda.
    /// </summary>
    [Fact]
    public void H34_NovedadConNombreLimpio_EsRojo()
    {
        var resultado = ResultadoInyectado(EstadoEntrega.EntregadaConNovedades, rutaSalida: @"C:\salida\Remuneración 202609-1 Total.xlsx");

        var ex = Assert.Throws<CalculoInvalidoException>(() => InvarianteEntrega.Garantizar(resultado));
        Assert.Equal(CodigoError.Inesperado, ex.Codigo);
        Assert.Contains(NombresSalida.SufijoSinValidar, ex.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// Injerto 3 — «novedad reportada como OK/0 ⇒ rojo». Un resultado con novedades declarado
    /// <see cref="EstadoEntrega.Certificada"/> es una entrega inconsistente, y la superficie CLI
    /// mapea las novedades a 7 (nunca a 0) sin imprimir <c>RESULTADO OK</c>.
    /// </summary>
    [Fact]
    public void H34_NovedadReportadaComoOk_EsRojo()
    {
        var comoCertificada = ResultadoInyectado(EstadoEntrega.Certificada, rutaSalida: @"C:\salida\Remuneración 202609-1 Total (sin validar).xlsx");

        var ex = Assert.Throws<CalculoInvalidoException>(() => InvarianteEntrega.Garantizar(comoCertificada));
        Assert.Equal(CodigoError.Inesperado, ex.Codigo);
        Assert.Contains("NUNCA se certifica", ex.Message, StringComparison.Ordinal);

        // El contrato de superficie: con novedades el código es 7 y jamás 0.
        Assert.NotEqual(CodigosSalida.Ok, CodigosSalida.NovedadesEntrega);
        Assert.NotEqual(CodigosSalida.DivergenciaDeclaradaReconocida, CodigosSalida.NovedadesEntrega);
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // D10 (ratificada) — qué fallo duro conserva el "(sin validar)"
    // ════════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// D10: en fallo duro POSTERIOR a la escritura, el <c>(sin validar)</c> se conserva SOLO en los
    /// fallos de LECTURA-DE-VUELTA de valores (<c>VerificarDetalleR2R4</c>,
    /// <c>VerificarIntegridadPegadoR10</c>) y se sigue borrando en <c>ERR-PLANTILLA</c>/
    /// <c>ERR-INTEGRIDAD</c>.
    /// </summary>
    [Fact]
    public void H3_D10_ParcialConservadoSoloEnLecturaDeVuelta()
    {
        // (a) Lectura-de-vuelta: ERR-ESCRITURA del gate de valores (sin excepción interna) ⇒ conserva.
        Assert.True(OpenXmlPlantillaWriter.ConservaParcialAlFallar(
            new CalculoInvalidoException(CodigoError.Escritura, "ERR-ESCRITURA: el desglose-detalle no quedó escrito…")));

        // (a2) La firma REAL: el gate de verdad, sobre el workbook de verdad. Se le pasa un bloque del
        // R10 desviado en 1 peso sobre un archivo ya escrito ⇒ el read-back del pegado lanza su
        // ERR-ESCRITURA y esa excepción —la que el writer atrapa— es la que decide CONSERVAR el
        // archivo. Reclasificado en el Plan 62 (T2.5, D-59H4): el gate de valores del asiento DetRetri
        // del Plan 59 se retiró; su sucesor es la verificación de integridad del pegado (T2.3).
        var resultado211 = EspecialesR4H2Plan61Tests.Generacion2026091.Value.Resultado;
        var hojaEscrita = resultado211.RutaSalida;
        Assert.True(File.Exists(hojaEscrita), "El fixture H3 debía dejar el workbook marcado en disco.");
        var periodo211 = resultado211.Resultado.Periodo;
        var lectorR10 = new ExcelDataReaderDetRetriR10Reader();
        var rutaR10211 = Path.Combine(
            Raiz, "Docs", "Prueba Septiembre 1", "Insumos", "R10_Remuneracion_2026091.xlsx");
        var bloqueDetReal = lectorR10.LeerBloqueDetRetri(periodo211, rutaR10211);
        var bloqueValiReal = lectorR10.LeerBloqueDetValiRetri(periodo211, rutaR10211);
        var bloqueDetDesviado = DesviarUnaCelda(bloqueDetReal, fila: 9, columna: "D", delta: 1m);

        CalculoInvalidoException exReal;
        using (var documento = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(hojaEscrita, false))
        {
            exReal = Assert.Throws<CalculoInvalidoException>(() =>
                OpenXmlPlantillaWriter.VerificarIntegridadPegadoR10(
                    documento.WorkbookPart!, periodo211, bloqueDetDesviado, bloqueValiReal));
        }

        Assert.Equal(CodigoError.Escritura, exReal.Codigo);
        Assert.Null(exReal.InnerException);
        Assert.True(OpenXmlPlantillaWriter.ConservaParcialAlFallar(exReal),
            "La lectura-de-vuelta de valores conserva el '(sin validar)': el artefacto está completo y es evidencia útil.");

        // (b) Estructura rota ⇒ borra.
        Assert.False(OpenXmlPlantillaWriter.ConservaParcialAlFallar(
            new CalculoInvalidoException(CodigoError.Plantilla, "ERR-PLANTILLA: destino con fórmula…")));
        Assert.False(OpenXmlPlantillaWriter.ConservaParcialAlFallar(
            new CalculoInvalidoException(CodigoError.Integridad, "ERR-INTEGRIDAD: el paquete no es válido…")));

        // (c) I/O envuelto (ERR-ESCRITURA CON inner) ⇒ borra: un archivo a medio copiar no es evidencia.
        Assert.False(OpenXmlPlantillaWriter.ConservaParcialAlFallar(
            new CalculoInvalidoException(CodigoError.Escritura, "no se pudo escribir", new IOException("sharing violation"))));
    }

    /// <summary>
    /// D10 end-to-end (mitad ERR-PLANTILLA): una plantilla alterada aborta y NO deja ni el marcado.
    /// Es la contraparte observable del caso (a): estructura rota no se entrega.
    /// </summary>
    [Fact]
    public void H3_D10_ErrPlantilla_NoConservaNada()
    {
        var dir = Path.Combine(Path.GetTempPath(), "p61-h3-d10-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var plantillaRota = Path.Combine(dir, "plantilla-rota.xlsx");
            File.Copy(RutaPlantillaV4, plantillaRota, overwrite: true);
            ConvertirFormulaEnValorFijo(plantillaRota, "CONSOLIDADO_TOTAL RECAUDO", "D9");

            var marcada = NombresSalida.RutaSinValidar(Path.Combine(dir, "Remuneración 202609-1 Total.xlsx"));
            var procesador = CrearProcesador(new ReporteNovedadesWriter());

            var ex = Assert.Throws<CalculoInvalidoException>(() => procesador.Ejecutar(new SolicitudProcesoPeriodo
            {
                Periodo = PeriodoSeptiembre1,
                CarpetaPeriodo = CarpetaSeptiembre1,
                RutaPlantilla = plantillaRota,
                RutaSalida = Path.Combine(dir, "Remuneración 202609-1 Total.xlsx")
            }));

            Assert.Equal(CodigoError.Plantilla, ex.Codigo);
            Assert.False(File.Exists(marcada), "ERR-PLANTILLA no deja ni el archivo marcado: una estructura rota no es evidencia.");
        }
        finally
        {
            Borrar(dir);
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // H4.1 — el reporte es determinista y compatible con el nombre anterior (D-TXT)
    // ════════════════════════════════════════════════════════════════════════════════════════════

    [Fact]
    public void H41_ReporteDeNovedades_DeterministaYConSeccionDeDeclaradas()
    {
        var dir = Path.Combine(Path.GetTempPath(), "p61-h4-txt-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var rutaLimpia = Path.Combine(dir, "Remuneración 202609-1 Total.xlsx");
            var excelEntregado = NombresSalida.RutaSinValidar(rutaLimpia);
            File.WriteAllText(excelEntregado, "placeholder");
            var writer = new ReporteNovedadesWriter();
            var novedades = new[]
            {
                new NovedadEntrega(
                    "R10 del período (R10_Remuneracion_2026091.xlsx)",
                    "DetRetri2026091",
                    3,
                    "Retribución total del ASE (asiento del período)",
                    10891338739m,
                    10891338740m,
                    -1m,
                    "Área administrativa (elabora el reporte oficial)",
                    "Compare por la etiqueta del ASE.",
                    "La cadena de la plantilla da 10.891.338.739,48 y el asiento se redondea a 10.891.338.739.")
            };
            var fecha = new DateTime(2026, 10, 28, 9, 30, 0, DateTimeKind.Unspecified);

            var ruta1 = writer.Emitir(rutaLimpia, excelEntregado, PeriodoSeptiembre1, Guid.Parse("11111111-1111-1111-1111-111111111111"),
                novedades, declaradas: null, usuario: "prueba", canal: CanalCli, fechaHora: fecha);
            var texto1 = File.ReadAllText(ruta1, Encoding.UTF8);
            var ruta2 = writer.Emitir(rutaLimpia, excelEntregado, PeriodoSeptiembre1, Guid.Parse("11111111-1111-1111-1111-111111111111"),
                novedades, declaradas: null, usuario: "prueba", canal: CanalCli, fechaHora: fecha);
            var texto2 = File.ReadAllText(ruta2, Encoding.UTF8);

            // Misma entrada → misma ruta y mismo contenido (determinismo).
            Assert.Equal(ruta1, ruta2);
            Assert.Equal(texto1, texto2);
            Assert.Equal("Remuneración 202609-1 Total" + ReporteNovedadesWriter.SufijoNovedades, Path.GetFileName(ruta1));

            // Contrato escueto (D-UX-TXT 2026-10-28): veredicto, ficha y guía. El bloque de datos de
            // corrida, la explicación general y las reglas de comparación viven en el manual.
            Assert.Contains("ENTREGADO CON NOVEDADES", texto1, StringComparison.Ordinal);
            Assert.Contains("Detalle de las novedades", texto1, StringComparison.Ordinal);
            Assert.Contains("Cómo revisar", texto1, StringComparison.Ordinal);
            Assert.DoesNotContain("Datos de la corrida", texto1, StringComparison.Ordinal);
            Assert.DoesNotContain("Qué significa", texto1, StringComparison.Ordinal);
            Assert.DoesNotContain("MANDA LA PLANTILLA", texto1, StringComparison.Ordinal);
            Assert.DoesNotContain("Regla general", texto1, StringComparison.Ordinal);
            Assert.DoesNotContain("no es un histórico", texto1, StringComparison.Ordinal);
            Assert.DoesNotContain("11111111-1111-1111-1111-111111111111", texto1, StringComparison.Ordinal);
            Assert.DoesNotContain("Divergencias declaradas presentes", texto1, StringComparison.Ordinal);

            // Con declaradas: sección final en el MISMO reporte (un solo TXT, R-DOBLE-REPORTE).
            var clasificacion = new ClasificacionDivergenciasDeclaradas
            {
                Veredictos =
                [
                    new VeredictoDivergenciaValidacion
                    {
                        AseId = 2,
                        Celda = "C15",
                        Declarada = true,
                        EntradaId = "OCCIDENTE-ASE5",
                        MontoObservado = 100m,
                        MontoEsperado = 101m
                    }
                ]
            };
            var rutaConDeclaradas = writer.Emitir(rutaLimpia, excelEntregado, PeriodoSeptiembre1,
                Guid.Parse("22222222-2222-2222-2222-222222222222"), novedades, clasificacion, fechaHora: fecha);
            var textoConDeclaradas = File.ReadAllText(rutaConDeclaradas, Encoding.UTF8);
            Assert.Contains("Divergencias declaradas presentes en esta corrida", textoConDeclaradas, StringComparison.Ordinal);
            Assert.Contains("OCCIDENTE-ASE5", textoConDeclaradas, StringComparison.Ordinal);

            // Compatibilidad D-TXT: solo declaradas conserva el nombre anterior (el otro writer).
            var rutaDeclaradas = new ReporteDivergenciasDeclaradasWriter().Emitir(rutaLimpia, PeriodoSeptiembre1,
                Guid.NewGuid(), clasificacion);
            Assert.Equal("Remuneración 202609-1 Total - Divergencias declaradas.txt", Path.GetFileName(rutaDeclaradas));
        }
        finally
        {
            Borrar(dir);
        }
    }

    // ════════════════════════════════════════════════════════════════════════════════════════════
    // Infraestructura
    // ════════════════════════════════════════════════════════════════════════════════════════════

    private const string CanalCli = HuellaAceptacion.CanalCli;

    /// <summary>
    /// Composición de producción. <paramref name="novedadForzada"/> inyecta el validador decorador
    /// de test (una divergencia de coherencia en el ASE 3) por la costura pública
    /// <see cref="IValidador"/>; el resto del flujo es el de producción.
    /// </summary>
    private static ProcesadorPeriodo CrearProcesador(IReporteNovedades? novedades, bool novedadForzada = false) =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            novedadForzada ? new ValidadorConNovedadForzada() : new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader(),
            new ValidacionOracleReader(),
            new ReporteDivergenciasDeclaradasWriter(),
            novedades);

    /// <summary>Composición de regresión: sin emisor del reporte de novedades.</summary>
    private static ProcesadorPeriodo CrearProcesadorSinEmisorDeNovedades(bool novedadForzada = false) =>
        CrearProcesador(novedades: null, novedadForzada);

    /// <summary>
    /// Plan 62 (H3): validador de producción con UNA divergencia de coherencia inyectada en el ASE 3
    /// (el resto delega en <see cref="ValidadorBasico"/>). Se inyecta por la costura pública de
    /// <see cref="IValidador"/> — sin fixture de hoja rota ni rama de producción por ASE — para
    /// ejercer el contrato «novedad sin TXT ⇒ rojo» ahora que septiembre-1 no produce novedades.
    /// </summary>
    private sealed class ValidadorConNovedadForzada : IValidador
    {
        private readonly ValidadorBasico _interno = new();

        public List<string> Validar(ResultadoRemuneracion resultado) => _interno.Validar(resultado);

        public List<string> Validar(ResultadoRemuneracion resultado, WorkbookLeafInputs leaf) => _interno.Validar(resultado, leaf);

        public List<string> Validar(ResultadoRemuneracion resultado, IReadOnlyList<WorkbookLeafInputs> leafs)
        {
            var errores = _interno.Validar(resultado, leafs);
            errores.Add("[ERR-VALIDACION] ASE 3: la hoja R2 TotalOportunoEsperado (100) no coincide con R2TotalOportuno del consolidado (200).");
            return errores;
        }

        public List<string> Validar(
            ResultadoRemuneracion resultado,
            IReadOnlyList<WorkbookLeafInputs> leafs,
            IReadOnlyList<ValidacionCruzadaSnapshot> snapshots) => _interno.Validar(resultado, leafs, snapshots);

        public ClasificacionDivergenciasDeclaradas ClasificarDivergenciasDeclaradas(
            ResultadoRemuneracion resultado,
            IReadOnlyList<WorkbookLeafInputs> leafs,
            IReadOnlyList<ValidacionCruzadaSnapshot> snapshots) =>
            _interno.ClasificarDivergenciasDeclaradas(resultado, leafs, snapshots);

        public List<string> ValidarDetRetriContraR10(IReadOnlyList<WorkbookLeafInputs> leafs, DetRetriInputs r10) =>
            _interno.ValidarDetRetriContraR10(leafs, r10);
    }

    private sealed class EmisorQueNoEscribe : IReporteNovedades
    {
        public string Emitir(
            string rutaSalida,
            string rutaExcelEntregado,
            Periodo periodo,
            Guid runId,
            IReadOnlyList<NovedadEntrega> novedades,
            ClasificacionDivergenciasDeclaradas? declaradas = null,
            string? usuario = null,
            string? canal = null,
            DateTime? fechaHora = null) =>
            Path.Combine(Path.GetDirectoryName(Path.GetFullPath(rutaSalida)) ?? ".", "reporte-que-no-existe.txt");
    }

    private static ResultadoProcesoPeriodo ResultadoInyectado(EstadoEntrega estado, string rutaSalida) =>
        new()
        {
            EstadoEntrega = estado,
            RutaSalida = rutaSalida,
            Novedades =
            [
                new NovedadEntrega(
                    "R10 del período (R10_Remuneracion_2026091.xlsx)",
                    "DetRetri2026091",
                    3,
                    "Retribución total del ASE",
                    1m,
                    2m,
                    -1m,
                    "Área administrativa",
                    "Compare por etiqueta.",
                    "Injerto de prueba.")
            ],
            RutaReporteNovedades = Path.Combine(Path.GetTempPath(), "injerto-que-no-existe.txt")
        };

    private static void ConvertirFormulaEnValorFijo(string ruta, string hoja, string celda)
    {
        using var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(ruta, true);
        var workbookPart = workbook.WorkbookPart ?? throw new InvalidOperationException("Sin WorkbookPart.");
        var workbookXml = workbookPart.Workbook ?? throw new InvalidOperationException("Sin Workbook.");
        var sheet = workbookXml.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .First(s => string.Equals(s.Name?.Value, hoja, StringComparison.OrdinalIgnoreCase));
        var worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Sin Worksheet.");
        var target = worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
            .First(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        target.CellFormula = null;
        target.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
        target.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("123");
        worksheet.Save();
    }

    /// <summary>
    /// Plan 62 (T2.5): devuelve el bloque del R10 con UNA celda desviada en <paramref name="delta"/>
    /// (simula un pegado que no coincide con la fuente, sin tocar el archivo).
    /// </summary>
    private static BloqueHojaR10 DesviarUnaCelda(BloqueHojaR10 bloque, int fila, string columna, decimal delta) =>
        new(
            bloque.Hoja,
            bloque.Periodo,
            bloque.Celdas
                .Select(c => c.Fila == fila && string.Equals(c.Columna, columna, StringComparison.OrdinalIgnoreCase)
                    ? new CeldaBloqueR10(c.Fila, c.Columna, null, (c.Numerico ?? 0m) + delta)
                    : c)
                .ToList());

    private static void Borrar(string dir)
    {
        try
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, recursive: true);
            }
        }
        catch
        {
            // best-effort
        }
    }

    private static string RaizRepositorio()
    {
        var actual = new DirectoryInfo(AppContext.BaseDirectory);
        while (actual is not null && !File.Exists(Path.Combine(actual.FullName, "AGENTS.md")))
        {
            actual = actual.Parent;
        }

        return actual?.FullName ?? throw new InvalidOperationException("No se encontró la raíz del repositorio (AGENTS.md).");
    }
}
