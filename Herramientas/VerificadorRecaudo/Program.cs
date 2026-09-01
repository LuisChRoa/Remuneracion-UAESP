using System.Globalization;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Infrastructure.Excel;

namespace Herramientas.VerificadorRecaudo;

/// <summary>
/// Harness de verificación para HU-02 — Lee archivos reales R1/R2/R4 y compara con valores esperados.
/// Tolerancia: ±0.5 (regla de negocio #9).
/// </summary>
internal static class Program
{
    private const decimal Tolerancia = 0.5m;

    // Rutas relativas al directorio de ejecución (dotnet run usa la raíz de la solución)
    private static readonly string BasePath = Path.Combine(
        "Docs", "Insumos", "REMUNERACION 2026071", "1-Promoambiental");

    private static readonly string ArchivoR1 = Path.Combine(BasePath,
        "Recaudoporcomponente_to_date01072026ddMMyyyy_to_date15072026ddMMyyyy___20267161653925.xlsx");

    private static readonly string ArchivoR2 = Path.Combine(BasePath,
        "RerpoteDetalleSaldosaFavor_to_date01072026ddMMyyyy_to_date15072026ddMMyyyy___202671616325453.xlsx");

    private static readonly string ArchivoR4 = Path.Combine(BasePath,
        "ReversiónPorComponente_to_date01072026ddMMyyyy_to_date15072026ddMMyyyy___2026716163221489.xlsx");

    // Valores de referencia verificados contra archivos físicos (1-Promoambiental, 202607-1)
    private const decimal EsperadoR1TotalOportuno = 19556118465.99m;
    private const decimal EsperadoR1Extemporaneo = 19549786950.62m;
    private const decimal EsperadoR2GrandTotal = 54273647.38m;
    // Nota: El archivo R2 de Promoambiental SÍ tiene columna "Especiales" (index 10, valor 57261.70).
    // El plan T4 decía 0, pero la fuente real tiene la columna. TotalOportuno = GrandTotal - ServEspK = 54216385.68.
    private const decimal EsperadoR2ServEspK = 57261.70m;
    private const decimal EsperadoR4TotalReversiones = -12054255.65m;

    private static int _passed;
    private static int _failed;
    private static int _total;

    internal static void Main(string[] args)
    {
        Console.WriteLine("=== Verificador de Recaudo — HU-02 ===");
        Console.WriteLine($"Tolerancia: ±{Tolerancia}");
        Console.WriteLine();

        IRecaudoReader reader = new ExcelDataReaderRecaudoReader();

        // Casos positivos (T1-T5)
        EjecutarCaso("T1", "R1 TotalOportuno (A=Componente∧B=Total→colF)",
            () => Verificar(reader.LeerR1(ArchivoR1).TotalOportuno, EsperadoR1TotalOportuno));

        EjecutarCaso("T2", "R1 Extemporaneo (B=Mes→colF)",
            () => Verificar(reader.LeerR1(ArchivoR1).Extemporaneo, EsperadoR1Extemporaneo));

        EjecutarCaso("T3", "R2 GrandTotal (A=Total∧B=vacío→colE)",
            () => Verificar(reader.LeerR2(ArchivoR2).GrandTotal, EsperadoR2GrandTotal));

        EjecutarCaso("T4", "R2 ServEspK (columna Especiales detectada dinámicamente)",
            () =>
            {
                var r2 = reader.LeerR2(ArchivoR2);
                var ok = Verificar(r2.ServEspK, EsperadoR2ServEspK);
                Console.WriteLine($"    TieneColumnaEspeciales = {r2.TieneColumnaEspeciales}");
                Console.WriteLine($"    TotalOportuno (computed) = {r2.TotalOportuno}");
                return ok && r2.TieneColumnaEspeciales;
            });

        EjecutarCaso("T5", "R4 TotalReversiones (A=Total∧B=vacío→colD, negativo)",
            () => Verificar(reader.LeerR4(ArchivoR4).TotalReversiones, EsperadoR4TotalReversiones));

        Console.WriteLine();

        // Casos negativos
        Console.WriteLine("--- Casos negativos ---");
        Console.WriteLine();

        // Caso negativo 1: archivo inexistente → ArchivoFuenteNoEncontradoException
        EjecutarCasoNegativo("N1", "Archivo inexistente → ArchivoFuenteNoEncontradoException",
            () =>
            {
                try
                {
                    reader.LeerR1(Path.Combine(BasePath, "archivo_inexistente.xlsx"));
                    Console.WriteLine("    ERROR: No se lanzó excepción");
                    return false;
                }
                catch (ArchivoFuenteNoEncontradoException ex)
                {
                    Console.WriteLine($"    Excepción lanzada: {ex.GetType().Name}");
                    Console.WriteLine($"    Mensaje: {ex.Message}");
                    return true;
                }
            });

        // Caso negativo 2: archivo con estructura inesperada → CalculoInvalidoException
        // Usamos el archivo R4 como entrada a LeerR1: R4 no tiene la fila "Componente"/"Total" que R1 busca
        EjecutarCasoNegativo("N2", "Archivo R4 leído como R1 (sin fila Componente/Total) → CalculoInvalidoException",
            () =>
            {
                try
                {
                    reader.LeerR1(ArchivoR4);
                    Console.WriteLine("    ERROR: No se lanzó excepción");
                    return false;
                }
                catch (CalculoInvalidoException ex)
                {
                    Console.WriteLine($"    Excepción lanzada: {ex.GetType().Name}");
                    Console.WriteLine($"    Mensaje: {ex.Message}");
                    return true;
                }
            });

        Console.WriteLine();
        Console.WriteLine("=== Resumen ===");
        Console.WriteLine($"Total: {_total} | PASS: {_passed} | FAIL: {_failed}");
        Console.WriteLine();

        if (_failed > 0)
        {
            Console.WriteLine("❌ HAY FALLOS — revise la salida anterior.");
            Environment.Exit(1);
        }
        else
        {
            Console.WriteLine("✅ TODOS LOS CASOS PASARON.");
            Environment.Exit(0);
        }
    }

    private static bool Verificar(decimal real, decimal esperado)
    {
        var diferencia = Math.Abs(esperado - real);
        var ok = diferencia <= Tolerancia;

        Console.WriteLine($"    Esperado:  {esperado.ToString("N2", CultureInfo.InvariantCulture)}");
        Console.WriteLine($"    Real:      {real.ToString("N2", CultureInfo.InvariantCulture)}");
        Console.WriteLine($"    Diferencia: {diferencia.ToString("N2", CultureInfo.InvariantCulture)}");
        Console.WriteLine($"    Resultado: {(ok ? "PASS" : "FAIL")}");

        return ok;
    }

    private static void EjecutarCaso(string id, string descripcion, Func<bool> accion)
    {
        _total++;
        Console.Write($"[{id}] {descripcion}... ");

        try
        {
            var ok = accion();
            if (ok)
            {
                _passed++;
                Console.WriteLine("✅ PASS");
            }
            else
            {
                _failed++;
                Console.WriteLine("❌ FAIL");
            }
        }
        catch (Exception ex)
        {
            _failed++;
            Console.WriteLine($"❌ FAIL (excepción inesperada: {ex.GetType().Name}: {ex.Message})");
        }

        Console.WriteLine();
    }

    private static void EjecutarCasoNegativo(string id, string descripcion, Func<bool> accion)
    {
        _total++;
        Console.Write($"[{id}] {descripcion}... ");

        try
        {
            var ok = accion();
            if (ok)
            {
                _passed++;
                Console.WriteLine("✅ PASS");
            }
            else
            {
                _failed++;
                Console.WriteLine("❌ FAIL");
            }
        }
        catch (Exception ex)
        {
            _failed++;
            Console.WriteLine($"❌ FAIL (excepción inesperada: {ex.GetType().Name}: {ex.Message})");
        }

        Console.WriteLine();
    }

}
