using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// HU-20 (G3): contrato del lector del insumo <c>R10</c> por período
/// (<c>R10_Remuneracion_{AAAAMMQ}.xlsx</c>, hojas <c>DetRetri{AAAAMMQ}</c> /
/// <c>DetValiRetri{AAAAMMQ}</c>).
///
/// <para>
/// <b>Dos superficies, dos usos:</b>
/// </para>
/// <list type="bullet">
/// <item><see cref="LeerDetRetri"/> (HU-20/G3): lee el DetRetri por ASE (<c>D9:D13</c>) y su total
/// (<c>D14</c>) para contrastarlo con el DetRetri calculado (tolerancia ±0.5 post-redondeo).</item>
/// <item><see cref="LeerBloqueDetRetri"/> / <see cref="LeerBloqueDetValiRetri"/> (Plan 62, N1): leen
/// el bloque COMPLETO del contrato T0 (<c>A3:O19</c> y <c>A3:O14</c>) que la app copia a las hojas
/// homónimas de la plantilla. Desde el Plan 62 el R10 es la <b>FUENTE</b> de esas hojas en TODOS los
/// períodos y AMBAS quincenas (reemplaza a G3-D1 para la hoja; el gate total calculado-vs-R10 queda
/// obsoleto por tautología).</item>
/// </list>
///
/// <para>
/// El reader jamás escribe (read-only) y falla nombrando período + archivo + hoja (y la celda cuando
/// el problema es de contenido): nunca 0 silencioso ni bloque parcial.
/// </para>
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

    /// <summary>
    /// Plan 62 (T1.1): lee el bloque completo de <c>DetRetri{CodigoCompleto}</c> — rango
    /// <c>A3:O19</c>, incluidos los encabezados de las filas 3–8 y el pie de las filas 15–19.
    /// </summary>
    /// <param name="periodo">Período del dominio (determina la hoja por <see cref="NombresHojaPeriodo"/>).</param>
    /// <param name="rutaR10">Ruta del archivo <c>R10_Remuneracion_*.xlsx</c>.</param>
    /// <returns>Bloque del período con todas las celdas del rango (las vacías, explícitas).</returns>
    BloqueHojaR10 LeerBloqueDetRetri(Periodo periodo, string rutaR10);

    /// <summary>
    /// Plan 62 (T1.1): lee el bloque completo de <c>DetValiRetri{CodigoCompleto}</c> — rango
    /// <c>A3:O14</c>, incluidos los encabezados de las filas 3–8.
    /// </summary>
    /// <param name="periodo">Período del dominio (determina la hoja por <see cref="NombresHojaPeriodo"/>).</param>
    /// <param name="rutaR10">Ruta del archivo <c>R10_Remuneracion_*.xlsx</c>.</param>
    /// <returns>Bloque del período con todas las celdas del rango (las vacías, explícitas).</returns>
    BloqueHojaR10 LeerBloqueDetValiRetri(Periodo periodo, string rutaR10);
}
