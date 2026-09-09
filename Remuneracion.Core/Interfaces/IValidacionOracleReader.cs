using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// HU-13 (2.7): contrato del oráculo de LECTURA de validaciones cruzadas (D1).
/// Lee en read-only (OpenXML sin recalcular) el caché de las hojas de validación protegidas de un
/// workbook y produce snapshots por ASE. El validador de dominio NUNCA abre .xlsx: solo compara
/// los números de <see cref="ValidacionCruzadaSnapshot"/> (DIP, testeabilidad in-memory).
///
/// W2 (D7): cada hoja/celda-oráculo debe ASSERTAR existencia (hoja existe, celda existe, fórmula
/// presente donde T0 lo exige) antes de devolver el snapshot. Lectura que no encuentra una celda
/// → fallo que NOMBRA hoja+celda, nunca 0 silencioso.
///
/// El reader jamás escribe (A4): abrir el workbook en read-only y no mutarlo.
/// </summary>
public interface IValidacionOracleReader
{
    /// <summary>
    /// Lee el snapshot-oráculo de validaciones cruzadas de los 5 ASE para el período indicado.
    /// </summary>
    /// <param name="rutaWorkbook">Ruta del workbook (salida o golden) a leer read-only.</param>
    /// <param name="periodo">Período que determina el sufijo de hojas DetRetri/DetValiRetri
    /// (2026071 vs 2026072) y las referencias por período.</param>
    /// <returns>Lista de 5 snapshots, uno por ASE (1..5), con matcheo estricto por Ase.Id.</returns>
    IReadOnlyList<ValidacionCruzadaSnapshot> LeerSnapshots(string rutaWorkbook, Periodo periodo);
}
