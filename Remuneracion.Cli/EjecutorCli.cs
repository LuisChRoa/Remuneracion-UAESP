using System.IO;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;
using Serilog;
using Serilog.Context;

namespace Remuneracion.Cli;

/// <summary>
/// HU-15 (3.3, D3): ejecuta el flujo CLI separable de <c>Main</c> — parsea, compone los mismos
/// servicios que <c>Program.cs</c> (WinForms), ejecuta y devuelve el código 0-5 sin
/// <c>Environment.Exit</c> (tests in-process). El exit vive SOLO en <c>Program.Main</c>.
/// Misma ejecución que la UI (mismos valores), otro frontend (D7: sin runner compartido).
/// </summary>
public static class EjecutorCli
{
    /// <summary>
    /// Ejecuta el CLI con los argumentos dados. <paramref name="salidaStdout"/>/<paramref name="salidaStderr"/>
    /// son inyectables para asserts in-process; por defecto usan la consola.
    /// </summary>
    /// <returns>Código del contrato <see cref="CodigosSalida"/> (0-5).</returns>
    public static int Ejecutar(string[] args, TextWriter? salidaStdout = null, TextWriter? salidaStderr = null) =>
        Ejecutar(args, salidaStdout, salidaStderr, validador: null, lectorLeaf: null);

    /// <summary>
    /// Costura de test (Plan 46 H2 — decisión del Ingeniero 2026-10-17): permite inyectar un
    /// <see cref="IValidador"/> y un <see cref="IWorkbookLeafInputReader"/> (por defecto los de
    /// producción). La usa el FIXTURE SINTÉTICO de divergencia declarada para conservar la cobertura
    /// del mecanismo (bloqueo sin flag → cierre con flag + reporte) ahora que la divergencia real
    /// OCCIDENTE-ASE5 quedó resuelta. NO altera el mecanismo ni los defaults de producción.
    /// </summary>
    internal static int Ejecutar(
        string[] args,
        TextWriter? salidaStdout,
        TextWriter? salidaStderr,
        IValidador? validador,
        IWorkbookLeafInputReader? lectorLeaf)
    {
        var stdout = salidaStdout ?? Console.Out;
        var stderr = salidaStderr ?? Console.Error;

        OpcionesCli opciones;
        try
        {
            opciones = OpcionesCli.Parse(args);
        }
        catch (UsoCliException ex)
        {
            // G6: errores de uso → stderr + ayuda + salida 4 (sin códigos nuevos).
            stderr.WriteLine(ex.Message);
            stderr.WriteLine(OpcionesCli.Uso());
            stderr.WriteLine("Use --help para ver la ayuda completa.");
            return CodigosSalida.Inesperado;
        }

        if (opciones.Ayuda)
        {
            stdout.WriteLine(OpcionesCli.Uso());
            return CodigosSalida.Ok;
        }

        // W-2.1 (D4): el CLI genera UN RunId al arrancar y lo inyecta en la solicitud; el
        // procesador lo respeta (solicitud.RunId ?? Guid.NewGuid()) → un solo Guid correlaciona
        // todos los eventos (CA-6). LogContext lo propaga a cada evento del proceso.
        var runId = Guid.NewGuid();
        Program.ConfigurarSerilog(); // D8: canónico en Form1 (WinForms)
        using var _runIdScope = LogContext.PushProperty("RunId", runId);
        using var _periodoScope = LogContext.PushProperty("Periodo", opciones.PeriodoCodigoCompleto);
        using var _modoScope = LogContext.PushProperty("Modo", opciones.ModoTexto);

        // G5: --salida es SIEMPRE carpeta; el archivo se nombra por el período (igual que Form1).
        var salida = Path.Combine(opciones.Salida, opciones.Periodo.NombreArchivo);

        if (string.Equals(Path.GetFullPath(opciones.Plantilla), Path.GetFullPath(salida), StringComparison.OrdinalIgnoreCase))
        {
            // W-2.4: denegación salida==plantilla con Log.Error + guía del catálogo a stderr.
            var detalle = $"La ruta de salida coincide con la plantilla: {salida}.";
            var (_, guia) = CatalogoErrores.Para(CodigoError.Plantilla, new CalculoInvalidoException(CodigoError.Plantilla, detalle));
            Log.Error("[{Codigo}] {Mensaje}", CodigoError.Plantilla, guia);
            stderr.WriteLine($"[{CodigoError.Plantilla}] {guia}");
            ResultadoError(stdout, CodigoError.Plantilla, runId);
            return CodigosSalida.FuenteOPlantilla;
        }

        if (File.Exists(salida) && !opciones.Sobrescribir)
        {
            // Req 1: desatendido honesto — sin prompt; la negativa a sobrescribir es salida 5.
            Log.Warning("Proceso cancelado: salida ya existe y no se acepta sobreescritura. [{Codigo}]", CodigoError.CanceladoPorUsuario);
            stderr.WriteLine($"[{CodigoError.CanceladoPorUsuario}] El archivo de salida ya existe: {salida}. Use --sobrescribir para reemplazarlo.");
            ResultadoError(stdout, CodigoError.CanceladoPorUsuario, runId);
            return CodigosSalida.CanceladoPorUsuario;
        }

        try
        {
            ResultadoProcesoPeriodo? resultadoProceso = null;
            if (opciones.CincoAse)
            {
                resultadoProceso = EjecutarPeriodo(opciones, salida, runId, stdout, validador, lectorLeaf);
            }
            else
            {
                EjecutarUnAse(opciones, salida, runId, stdout, validador, lectorLeaf);
            }

            // Plan 61 (H4.2 / S3.4, D-CODIGO): ENTREGA CON NOVEDADES. El Excel existe (marcado "sin
            // validar") y el reporte explica qué revisar. Código 7 — DISTINTO de 0 y de 6 — y
            // NUNCA la línea "RESULTADO OK": es un estado de entrega, no un error ni un cierre limpio.
            if (resultadoProceso?.EstadoEntrega == EstadoEntrega.EntregadaConNovedades)
            {
                EmitirCierreConNovedades(stdout, resultadoProceso, runId);
                return CodigosSalida.NovedadesEntrega;
            }

            // Plan 42 (T3, D-G): cierre con divergencia declarada reconocida → código DISTINTO de
            // Ok (0). La salida textual enumera lo mismo que la UI (celda, Δ, dueño, reporte emitido).
            var clasificacion = resultadoProceso?.DivergenciasDeclaradas;
            if (clasificacion is not null && clasificacion.SoloDeclaradas)
            {
                EmitirCierreDeclarado(stdout, clasificacion, resultadoProceso!.ReporteDivergenciasDeclaradas, salida, runId);
                return CodigosSalida.DivergenciaDeclaradaReconocida;
            }

            stdout.WriteLine($"RESULTADO OK codigo={CodigosSalida.Ok} salida={salida} runId={runId}");
            return CodigosSalida.Ok;
        }
        catch (CalculoInvalidoException ex) when (ex.Clasificacion is not null && ex.Clasificacion.SoloDeclaradas)
        {
            // Plan 42 (T3, S3/D-G): sin el flag, la declarada también bloquea (FAIL Validacion,
            // exit 1). El CLI enumera la divergencia (celda, Δ, dueño) y cómo reconocerla.
            Log.Warning("Cierre bloqueado por divergencia declarada sin reconocimiento explícito: {Mensaje}", ex.Message);
            stderr.WriteLine($"[{CodigoError.Validacion}] La corrida no cerró: hay una divergencia declarada pendiente de reconocimiento explícito en esta corrida.");
            foreach (var veredicto in ex.Clasificacion.Declaradas
                         .GroupBy(v => (v.EntradaId, v.Celda))
                         .Select(g => g.First()))
            {
                var entrada = DivergenciasDeclaradasCatalogo.Entradas
                    .FirstOrDefault(e => string.Equals(e.Id, veredicto.EntradaId, StringComparison.OrdinalIgnoreCase));
                stderr.WriteLine(
                    $"  Diferencia declarada {veredicto.EntradaId}: hoja VALIDACION_TOTAL, celda {veredicto.Celda}; " +
                    $"Δ {veredicto.Delta:0.##}; responsable {entrada?.Dueno ?? "por confirmar"}.");
            }

            stderr.WriteLine(
                $"Para cerrar esta corrida reconozca la divergencia con --aceptar-divergencia-declarada <id> (p. ej. {string.Join(", ", ex.Clasificacion.EntradasDeclaradas)}).");
            ResultadoError(stdout, CodigoError.Validacion, runId);
            return CodigosSalida.Validacion;
        }
        catch (Exception ex)
        {
            // D6: código del catálogo vía CodigoDe (una sola fuente, compartida con Form1).
            var codigo = CatalogoErrores.CodigoDe(ex);
            Log.Error(ex, "[{Codigo}] Error en la ejecución del proceso: {Mensaje}", codigo, ex.Message);
            var (_, guia) = CatalogoErrores.Para(codigo, ex);
            ResultadoError(stdout, codigo, runId);
            stderr.WriteLine($"[{codigo}] {guia}");
            return CatalogoErrores.CodigoSalidaPara(codigo);
        }
    }

    /// <summary>
    /// Modo período (5 ASE): composición idéntica a <c>Program.cs</c> del WinForms (V8), incluidos
    /// el oráculo R10 obligatorio (HU-20/G3) y el oráculo de lectura de validaciones cruzadas.
    /// Plan 42 (T2/T3): además, el emisor del reporte durable de divergencias declaradas y los ids
    /// reconocidos por el usuario en esta corrida.
    /// </summary>
    private static ResultadoProcesoPeriodo EjecutarPeriodo(
        OpcionesCli opciones,
        string salida,
        Guid runId,
        TextWriter stdout,
        IValidador? validador,
        IWorkbookLeafInputReader? lectorLeaf)
    {
        IProcesadorPeriodo procesador = new ProcesadorPeriodo(
            new ExcelDataReaderRecaudoReader(),
            lectorLeaf ?? new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            validador ?? new ValidadorBasico(),
            new OpenXmlPlantillaWriter(),
            new ArchivoFuenteLocator(),
            new ExcelDataReaderDetRetriR10Reader(),
            new ValidacionOracleReader(),
            new ReporteDivergenciasDeclaradasWriter(),
            // Plan 61 (H4.1): emisor del reporte de novedades. Sin él, una corrida con novedades
            // FALLA (sin reporte no hay entrega).
            new ReporteNovedadesWriter());

        return procesador.Ejecutar(new SolicitudProcesoPeriodo
        {
            Periodo = opciones.Periodo,
            CarpetaPeriodo = opciones.Carpeta,
            RutaPlantilla = opciones.Plantilla,
            RutaSalida = salida,
            RunId = runId, // W-2.1
            DivergenciasDeclaradasAceptadas = opciones.DivergenciasDeclaradasAceptadas,
            // Plan 43 (O-3): el reporte/log registran el canal (CLI) por el que se aceptó.
            Canal = HuellaAceptacion.CanalCli
        }, new ProgresoConsola(stdout));
    }

    /// <summary>
    /// Plan 61 (H4.2 / S3.4, CA-C4): salida textual del cierre ENTREGADO CON NOVEDADES. Enumera lo
    /// mismo que la UI (cuántas novedades, dónde está el Excel marcado y dónde el reporte) y deja la
    /// línea final grepable del contrato. NUNCA se imprime <c>RESULTADO OK</c>.
    /// </summary>
    private static void EmitirCierreConNovedades(
        TextWriter stdout,
        ResultadoProcesoPeriodo resultado,
        Guid runId)
    {
        stdout.WriteLine("CIERRE ENTREGADO CON NOVEDADES — NO CERTIFICADO");
        stdout.WriteLine("El Excel del período está completo y se entrega marcado como \"(sin validar)\" para revisión;");
        stdout.WriteLine("la corrida NO se certificó porque hay diferencias que deben mirarse (ver el reporte).");

        foreach (var novedad in resultado.Novedades)
        {
            var ase = novedad.AseId == 0 ? "Período completo" : $"ASE {novedad.AseId}";
            var monto = novedad.ValorObservado is null || novedad.ValorEsperado is null
                ? string.Empty
                : $"; observado {novedad.ValorObservado:0.##}; esperado {novedad.ValorEsperado:0.##}; Δ {novedad.Delta:0.##}";
            stdout.WriteLine($"  {ase} · {novedad.Hoja} · {novedad.ConceptoOCelda}{monto}; responsable {novedad.Dueno}.");
        }

        stdout.WriteLine($"Reporte de novedades: {resultado.RutaReporteNovedades}");
        stdout.WriteLine($"RESULTADO CON NOVEDADES codigo={CodigosSalida.NovedadesEntrega} salida={resultado.RutaSalida} reporte={resultado.RutaReporteNovedades} runId={runId}");
    }

    /// <summary>
    /// Plan 42 (T3, R-CLI-1/D-G): salida textual del cierre con divergencia declarada reconocida.
    /// Enumera lo mismo que la UI (celda, observado vs esperado, Δ, dueño, reporte emitido) en
    /// lenguaje de negocio (sin jerga de código) y deja la línea final grepable del contrato.
    /// </summary>
    private static void EmitirCierreDeclarado(
        TextWriter stdout,
        ClasificacionDivergenciasDeclaradas clasificacion,
        string rutaReporte,
        string salida,
        Guid runId)
    {
        stdout.WriteLine("CIERRE CON DIVERGENCIA DECLARADA RECONOCIDA");
        stdout.WriteLine("La corrida cerró con una diferencia conocida y reconocida para esta corrida (no es un error del cálculo).");

        foreach (var veredicto in clasificacion.Declaradas
                     .GroupBy(v => (v.EntradaId, v.Celda))
                     .Select(g => g.First()))
        {
            var entrada = DivergenciasDeclaradasCatalogo.Entradas
                .FirstOrDefault(e => string.Equals(e.Id, veredicto.EntradaId, StringComparison.OrdinalIgnoreCase));
            stdout.WriteLine(
                $"  Diferencia declarada {veredicto.EntradaId}: hoja VALIDACION_TOTAL, celda {veredicto.Celda}; " +
                $"observado {veredicto.MontoObservado:0.##}; esperado {veredicto.MontoEsperado:0.##}; Δ {veredicto.Delta:0.##}; " +
                $"responsable {entrada?.Dueno ?? "por confirmar"}.");
        }

        stdout.WriteLine($"Reporte emitido: {rutaReporte}");
        stdout.WriteLine($"RESULTADO OK codigo={CodigosSalida.DivergenciaDeclaradaReconocida} salida={salida} reporte={rutaReporte} runId={runId}");
    }

    /// <summary>
    /// Modo 1-ASE: resuelve carpeta + R1/R2/R4 con <see cref="ArchivoFuenteLocator"/> exactamente
    /// como <c>Form1.ObtenerCarpetaAse</c> + <c>EjecutarModoUnAse</c> (mismo orden de prefijos,
    /// mismo fallback Reversión/Reversion) y ejecuta <see cref="ProcesadorRemuneracion"/>.
    /// </summary>
    private static void EjecutarUnAse(
        OpcionesCli opciones,
        string salida,
        Guid runId,
        TextWriter stdout,
        IValidador? validador,
        IWorkbookLeafInputReader? lectorLeaf)
    {
        var ase = AseFactory.DesdeId(opciones.AseId!.Value);
        var locator = new ArchivoFuenteLocator();
        var carpetaAse = locator.ObtenerCarpetasAse(opciones.Carpeta)
            .FirstOrDefault(c => Path.GetFileName(c).StartsWith(ase.Id.ToString(), StringComparison.OrdinalIgnoreCase))
            ?? throw new ArchivoFuenteNoEncontradoException(
                CodigoError.FuenteNoEncontrada,
                $"No se encontró la carpeta del ASE {ase.Id} ({CarpetasAse.Prefijos[ase.Id - 1]}) en '{opciones.Carpeta}'.");

        var rutaR1 = locator.BuscarArchivo(carpetaAse, "Recaudoporcomponente")
            ?? throw new ArchivoFuenteNoEncontradoException(CodigoError.FuenteNoEncontrada, $"No se encontró R1 del ASE {ase.Id} en {carpetaAse}.");
        var rutaR2 = locator.BuscarArchivo(carpetaAse, "RerpoteDetalleSaldosaFavor")
            ?? throw new ArchivoFuenteNoEncontradoException(CodigoError.FuenteNoEncontrada, $"No se encontró R2 del ASE {ase.Id} en {carpetaAse}.");
        var rutaR4 = locator.BuscarArchivo(carpetaAse, "ReversiónPorComponente")
            ?? locator.BuscarArchivo(carpetaAse, "ReversionPorComponente")
            ?? throw new ArchivoFuenteNoEncontradoException(CodigoError.FuenteNoEncontrada, $"No se encontró R4 del ASE {ase.Id} en {carpetaAse}.");

        IProcesadorRemuneracion procesador = new ProcesadorRemuneracion(
            new ExcelDataReaderRecaudoReader(),
            lectorLeaf ?? new ExcelDataReaderWorkbookLeafInputReader(),
            new CalculoRemuneracion(),
            validador ?? new ValidadorBasico(),
            new OpenXmlPlantillaWriter());

        procesador.Ejecutar(new SolicitudProcesoAse
        {
            Ase = ase,
            Periodo = opciones.Periodo,
            RutaR1 = rutaR1,
            RutaR2 = rutaR2,
            RutaR4 = rutaR4,
            RutaPlantilla = opciones.Plantilla,
            RutaSalida = salida,
            RunId = runId // W-2.1
        }, new ProgresoConsola(stdout));
    }

    /// <summary>
    /// Línea final grepable del contrato (§2.1): RESULTADO ERROR codigo=&lt;ERR-…&gt; salida=&lt;N&gt; runId=&lt;guid&gt;.
    /// </summary>
    private static void ResultadoError(TextWriter stdout, string codigo, Guid runId)
    {
        stdout.WriteLine($"RESULTADO ERROR codigo={codigo} salida={CatalogoErrores.CodigoSalidaPara(codigo)} runId={runId}");
    }

    /// <summary>
    /// Hitos del procesador a stdout (misma doctrina HU-14 D4: hitos a stdout/UI, detalle al log).
    /// Síncrono (consola, sin SynchronizationContext) para no entremezclar líneas.
    /// </summary>
    private sealed class ProgresoConsola : IProgress<string>
    {
        private readonly TextWriter _salida;

        public ProgresoConsola(TextWriter salida) => _salida = salida;

        public void Report(string value) => _salida.WriteLine(value);
    }
}
