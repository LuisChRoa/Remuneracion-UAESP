using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// HU-20 (G3): contrato del lector del oráculo <c>R10</c> por período
/// (<c>R10_Remuneracion_{AAAAMMQ}.xlsx</c>, hoja <c>DetRetri{AAAAMMQ}</c>).
///
/// El R10 es un ORÁCULO DE VALIDACIÓN (G3-D1): se contrasta contra el DetRetri calculado con
/// tolerancia ±0.5 post-redondeo; NUNCA es fuente de escritura del template. El reader jamás
/// escribe (read-only) y falla nombrando período + archivo si la hoja/celdas esperadas no
/// existen (nunca 0 silencioso).
/// </summary>
public interface IDetRetriR10Reader
{
    /// <summary>
    /// Lee el DetRetri del período desde la hoja <c>DetRetri{CodigoCompleto}</c> del R10.
    /// </summary>
    /// <param name="periodo">Período de la liquidación (determina el sufijo de la hoja).</param>
    /// <param name="rutaR10">Ruta del archivo <c>R10_Remuneracion_*.xlsx</c>.</param>
    /// <returns>DetRetri por ASE + total + rango del período.</returns>
    DetRetriInputs LeerDetRetri(Periodo periodo, string rutaR10);
}
