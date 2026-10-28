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
    /// 2026091 (5 ASE): el único gate que enrojece es la errata del R10 en el ASE 3 (739,48 → 739
    /// contra 740). Con la política nueva el período ENTREGA: Excel marcado, novedad tipada y TXT.
    /// </summary>
    [Fact]
    public void H35_2026091_EntregaConNovedades_ExcelMarcadoMasReporteLegible()
    {
        var resultado = EspecialesR4H2Plan61Tests.Generacion2026091.Value.Resultado;

        // Estado inequívoco y ruta marcada (invariante de entrega).
        Assert.Equal(EstadoEntrega.EntregadaConNovedades, resultado.EstadoEntrega);
        Assert.NotEmpty(resultado.Novedades);
        Assert.EndsWith(NombresSalida.SufijoSinValidar + ".xlsx", resultado.RutaSalida, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(resultado.RutaSalida), "El Excel marcado debe existir: sin artefacto no hay nada que revisar.");

        // NO se promueve: el nombre limpio no existe.
        var nombreLimpio = resultado.RutaSalida.Replace(NombresSalida.SufijoSinValidar, string.Empty, StringComparison.Ordinal);
        Assert.False(File.Exists(nombreLimpio), "Una corrida con novedades NUNCA toma el nombre limpio.");

        // La novedad del ASE 3 es la errata del oráculo: 739 calculado (de 739,48) vs 740 del R10.
        var novedadAse3 = Assert.Single(resultado.Novedades, n => n.AseId == 3);
        Assert.Equal(10891338739m, novedadAse3.ValorObservado);
        Assert.Equal(10891338740m, novedadAse3.ValorEsperado);
        Assert.Equal(-1m, novedadAse3.Delta);
        Assert.Contains("DetRetri2026091", novedadAse3.Hoja, StringComparison.Ordinal);
        Assert.Contains("10891338739", novedadAse3.Explicacion!.Replace(".", string.Empty, StringComparison.Ordinal), StringComparison.Ordinal);

        // El TXT existe, es el único reporte de la corrida y está en UTF-8 SIN BOM.
        Assert.True(File.Exists(resultado.RutaReporteNovedades), "Sin reporte no hay entrega (nunca un cierre silencioso).");
        var nombreBaseEntregado = Path.GetFileNameWithoutExtension(
            resultado.RutaSalida.Replace(NombresSalida.SufijoSinValidar, string.Empty, StringComparison.Ordinal));
        Assert.Equal(nombreBaseEntregado + ReporteNovedadesWriter.SufijoNovedades, Path.GetFileName(resultado.RutaReporteNovedades));
        var bytes = File.ReadAllBytes(resultado.RutaReporteNovedades);
        Assert.False(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF, "El reporte va en UTF-8 SIN BOM (precedente Plan 42).");

        var texto = File.ReadAllText(resultado.RutaReporteNovedades, Encoding.UTF8);

        // Sección 1: veredicto inequívoco.
        Assert.StartsWith("ENTREGADO CON NOVEDADES — NO CERTIFICADO", texto, StringComparison.Ordinal);

        // Sección 3: la ficha del ASE 3 con la cadena completa (739,48 → 739 vs 740).
        Assert.Contains("ASE 3", texto, StringComparison.Ordinal);
        Assert.Contains("10.891.338.739,48", texto, StringComparison.Ordinal);
        Assert.Contains("10.891.338.740", texto, StringComparison.Ordinal);

        // Sección 4: guía por ETIQUETA (nunca por fila).
        Assert.Contains("por etiqueta", texto, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("nunca por número de fila", texto, StringComparison.OrdinalIgnoreCase);

        // Lenguaje de NEGOCIO: prohibida la jerga de código (R-REPORTE-JERGA / G-TRAZA).
        Assert.DoesNotContain("ERR-", texto, StringComparison.Ordinal);
        Assert.DoesNotContain("<f>", texto, StringComparison.Ordinal);
        Assert.DoesNotContain("CalculoInvalidoException", texto, StringComparison.Ordinal);
        Assert.DoesNotContain("WorkbookLeafInputs", texto, StringComparison.Ordinal);
        Assert.DoesNotContain("VALIDACION_TOTAL", texto, StringComparison.Ordinal); // no aplica: la novedad es del R10
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
            var codigo = EjecutorCli.Ejecutar(
                ["--periodo", "2026091", "--carpeta", CarpetaSeptiembre1, "--plantilla", RutaPlantillaV4,
                    "--salida", salidaDir, "--cinco-ase", "--sobrescribir"],
                _stdout,
                _stderr);

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
    /// </summary>
    [Fact]
    public void H34_NovedadSinTxt_EsRojo()
    {
        // (a) Sin emisor configurado: fail-closed, nunca un Excel marcado huérfano de explicación.
        var sinEmisor = CrearProcesadorSinEmisorDeNovedades();
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
            var ex = Assert.Throws<CalculoInvalidoException>(() => CrearProcesador(new EmisorQueNoEscribe()).Ejecutar(new SolicitudProcesoPeriodo
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
    /// <c>VerificarAsientoDetRetri</c>) y se sigue borrando en <c>ERR-PLANTILLA</c>/
    /// <c>ERR-INTEGRIDAD</c>.
    /// </summary>
    [Fact]
    public void H3_D10_ParcialConservadoSoloEnLecturaDeVuelta()
    {
        // (a) Lectura-de-vuelta: ERR-ESCRITURA del gate de valores (sin excepción interna) ⇒ conserva.
        Assert.True(OpenXmlPlantillaWriter.ConservaParcialAlFallar(
            new CalculoInvalidoException(CodigoError.Escritura, "ERR-ESCRITURA: el desglose-detalle no quedó escrito…")));

        // (a2) La firma REAL: el gate de verdad, sobre el workbook de verdad. Se le pasa el dominio
        // desviado en 1 peso sobre un archivo ya escrito ⇒ el gate lanza su ERR-ESCRITURA y esa
        // excepción —la que el writer atrapa— es la que decide CONSERVAR el archivo.
        var resultado211 = EspecialesR4H2Plan61Tests.Generacion2026091.Value.Resultado;
        var hojaEscrita = resultado211.RutaSalida;
        Assert.True(File.Exists(hojaEscrita), "El fixture H3 debía dejar el workbook marcado en disco.");
        var referencia = resultado211.Leafs[0];
        var dominioDesviado = new WorkbookLeafInputs
        {
            Ase = referencia.Ase,
            Periodo = referencia.Periodo,
            DetRetriQ2 = new DetRetriQ2Inputs
            {
                Ase = referencia.Ase,
                TotalD104 = referencia.DetRetriQ2!.TotalD104 + 1m
            }
        };

        CalculoInvalidoException exReal;
        using (var documento = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(hojaEscrita, false))
        {
            exReal = Assert.Throws<CalculoInvalidoException>(() =>
                OpenXmlPlantillaWriter.VerificarAsientoDetRetri(documento.WorkbookPart!, [dominioDesviado]));
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

    private static ProcesadorPeriodo CrearProcesador(IReporteNovedades? novedades) =>
        new(
            new ExcelDataReaderRecaudoReader(),
            new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader(),
            new ValidacionOracleReader(),
            new ReporteDivergenciasDeclaradasWriter(),
            novedades);

    /// <summary>Composición de regresión: sin emisor del reporte de novedades.</summary>
    private static ProcesadorPeriodo CrearProcesadorSinEmisorDeNovedades() => CrearProcesador(novedades: null);

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
