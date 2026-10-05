namespace Remuneracion.Core.Models;

/// <summary>
/// Plan 26 (T1, §2.1/D-E): describe UN insumo faltante del período en lenguaje administrativo.
/// El modelo es puro (sin I/O); el sitio único que compone los textos
/// <see cref="QueFalta"/>/<see cref="DondeDebeIr"/>/<see cref="QueHacer"/> es
/// <c>FormateadorInsumosFaltantes</c>, y <c>ValidadorInsumosPeriodo</c> solo enumera los hechos
/// (qué finder devolvió <c>null</c>). Nombres de campo a elección del implementador (§2.1).
/// </summary>
public sealed class InsumoFaltante
{
    /// <summary>
    /// Ámbito del faltante (p. ej. "ASE 3", "Período"). Traza el orden estable del mensaje
    /// (ASE 1..5, luego conciliaciones, luego R10).
    /// </summary>
    public string Alcance { get; init; } = string.Empty;

    /// <summary>
    /// Qué falta, en el nombre que el usuario reconoce: reporte o carpeta, con el inicio de
    /// nombre de archivo esperado cuando aplica (D-E).
    /// </summary>
    public string QueFalta { get; init; } = string.Empty;

    /// <summary>
    /// Dónde debe ir: carpeta o ruta esperada dentro de la carpeta del período (D-E).
    /// </summary>
    public string DondeDebeIr { get; init; } = string.Empty;

    /// <summary>
    /// Qué hacer: pedir/generar el archivo, colocarlo y volver a ejecutar (D-E).
    /// </summary>
    public string QueHacer { get; init; } = string.Empty;
}
