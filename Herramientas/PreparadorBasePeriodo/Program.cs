using System.Text.RegularExpressions;
using Remuneracion.Core.Models;
using Remuneracion.Infrastructure.Excel;

namespace Herramientas.PreparadorBasePeriodo;

/// <summary>
/// Wrapper CLI DELGADO (Plan 39, T1/D-C) del componente compartido
/// <see cref="PreparadorBasePipeline"/>: convierte la plantilla estandarizada única en la base
/// canónica de un período (renombra las 3 hojas de período, reescribe las 11 <c>&lt;f&gt;</c> del
/// token de período, sella <c>CONSOLIDADO!N3</c>, selecciona la columna <c>L9:L13</c> por quincena
/// —R-INTERV-Q—, actualiza <c>docProps/app.xml</c>/defined names y escribe de forma atómica).
///
/// CERO lógica propia: el cuerpo mecánico vive en <see cref="PreparadorBasePipeline"/> (BCL puro
/// <c>ZipArchive</c> + <c>Regex</c>; sin OpenXML/ExcelDataReader/Excel/COM). El mismo componente es
/// la Phase 0 automática del pipeline (el writer lo invoca sobre la copia de trabajo).
///
/// El modo legacy (por período desde un manual, Plan 30 T3) se retiró: sus entradas
/// (<c>Docs/Prueba2/...</c>) ya no existen y ningún consumidor lo invocaba.
/// </summary>
internal static class Program
{
    private static int Main(string[] args)
    {
        try
        {
            Ejecutar(args);
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine("FALLO: " + ex.Message);
            return 1;
        }
    }

    private static void Ejecutar(string[] args)
    {
        var plantillaUnica = Argumento(args, "--plantilla-unica");
        if (string.IsNullOrWhiteSpace(plantillaUnica))
        {
            throw new InvalidOperationException(
                "Falta --plantilla-unica. Uso: --plantilla-unica <ruta> --periodo <AAAAMMQ> [--salida <ruta>] [--base <AAAAMMQ>].");
        }

        var periodoDestino = ExigirPeriodo(Argumento(args, "--periodo"), "--periodo");
        var salidaPath = Argumento(args, "--salida")
            ?? Path.Combine(
                Path.GetDirectoryName(Path.GetFullPath(plantillaUnica)) ?? ".",
                $"Plantilla_Remuneracion_{periodoDestino}.xlsx");

        if (string.Equals(Path.GetFullPath(plantillaUnica), Path.GetFullPath(salidaPath), StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("La salida no puede ser la misma ruta que la plantilla única (el origen nunca se modifica).");
        }

        Console.WriteLine("=== Preparador de base canónica desde la plantilla estandarizada única (Plan 39, T1) ===");

        var baseArg = Argumento(args, "--base");
        PreparadorBasePipeline.Ejecutar(
            plantillaUnica,
            Periodo.Parse(periodoDestino),
            salidaPath,
            string.IsNullOrWhiteSpace(baseArg) ? null : ExigirPeriodo(baseArg, "--base"),
            Console.WriteLine);
    }

    private static string ExigirPeriodo(string? valor, string argumento)
    {
        if (string.IsNullOrWhiteSpace(valor) || !Regex.IsMatch(valor, "^\\d{7}$"))
        {
            throw new InvalidOperationException($"El argumento {argumento} debe ser AAAAMMQ (7 dígitos); se recibió '{valor}'.");
        }

        return valor;
    }

    private static string? Argumento(string[] args, string nombre)
    {
        for (var i = 0; i < args.Length - 1; i++)
        {
            if (string.Equals(args[i], nombre, StringComparison.OrdinalIgnoreCase))
            {
                return args[i + 1];
            }
        }

        return null;
    }
}
