using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Define la validación de un resultado de remuneración.
/// </summary>
public interface IValidador
{
    /// <summary>
    /// Valida el resultado de la remuneración y retorna la lista de advertencias o errores.
    /// </summary>
    /// <param name="resultado">Resultado a validar.</param>
    /// <returns>Lista de mensajes de validación (vacía si no hay problemas).</returns>
    List<string> Validar(ResultadoRemuneracion resultado);

    /// <summary>
    /// Valida el resultado utilizando además los inputs leaf que alimentan las fórmulas del workbook.
    /// </summary>
    /// <param name="resultado">Resultado a validar.</param>
    /// <param name="leaf">Inputs leaf asociados a la salida que se quiere certificar.</param>
    /// <returns>Lista de mensajes de validación (vacía si no hay problemas).</returns>
    List<string> Validar(ResultadoRemuneracion resultado, WorkbookLeafInputs leaf);

    /// <summary>
    /// Valida un resultado multi-ASE (modo período) contra la lista completa de inputs leaf,
    /// con matcheo estricto por <see cref="Ase.Id"/> y gates por bloque.
    /// </summary>
    /// <param name="resultado">Resultado con los consolidados por ASE.</param>
    /// <param name="leafs">Inputs leaf, uno por ASE (5 en modo período).</param>
    /// <returns>Lista de mensajes de validación (vacía si no hay problemas).</returns>
    /// <remarks>
    /// HU-09 (2.3, §2.5): cuando algún leaf trae <see cref="WorkbookLeafInputs.ReporteBanco"/>
    /// no nulo, además de los gates HU-07/HU-08 se evalúan los gates D5:
    /// (i) bloque ASE = resumen fuente ±0.5 por empresa (<see cref="ReporteBancoEmpresaInputs.Total"/>
    /// vs <see cref="ReporteBancoEmpresaInputs.TotalFuente"/>),
    /// (ii) consolidado 1–7 = Σ bloques por empresa ±0.5 (si el dominio trae
    /// <see cref="ReporteBancoInputs.Consolidado"/>; en el desenlace T0 D2(b) es null y solo
    /// se verifica Σ contra el caché golden en Capa A),
    /// (iii) C59 = <see cref="Periodo.NumeroQuincena"/>.
    /// La diferencia de las filas 59–80 es informativa (V9): NUNCA falla por ella.
    /// HU-10 (2.4, §2.5): cuando algún leaf trae <see cref="WorkbookLeafInputs.BalanceSc"/>
    /// no nulo se evalúan los gates D5 2.4: (i) Total BSC = Total General fuente ±0.5 por ASE,
    /// (ii) Σ Total BSC = Σ TotalFuente ±0.5 (sumas fila 11 en dominio),
    /// (iii) H≈ROUND(F,0) solo en desenlace D2(a) (<see cref="BalanceScAseInputs.Sistema"/>);
    /// en Q1 H es fórmula (D2(b)) y H≈F se verifica contra el caché golden en Capa A.
    /// El validador NO abre .xlsx (J/K/M downstream y bloque 18–24 van a Capa B).
    /// </remarks>
    List<string> Validar(ResultadoRemuneracion resultado, IReadOnlyList<WorkbookLeafInputs> leafs);

    /// <summary>
    /// HU-13 (2.7): valida un resultado multi-ASE contra los snapshots-oráculo de validaciones
    /// cruzadas (gates 2.7 aditivos por quincena, D3). Snapshot ausente/vacío para un ASE =
    /// comportamiento HU-12 puro (sin gates nuevos para ese ASE).
    /// </summary>
    /// <param name="resultado">Resultado con los consolidados por ASE.</param>
    /// <param name="leafs">Inputs leaf, uno por ASE (5 en modo período).</param>
    /// <param name="snapshots">Snapshots-oráculo leídos read-only del workbook (D1), uno por ASE
    /// (5 en modo período). El validador NO abre .xlsx.</param>
    /// <returns>Lista de mensajes de validación (vacía si no hay problemas).</returns>
    /// <remarks>
    /// Gates 2.7 (§2.5 del Plan 13, congelados por T0 en ambos canónicos):
    /// (i) por empresa y ASE: columna O (H−N de Recaudo * vs REMUNERACION_*) == 0 ±0.5 y
    ///     columna P (INT(O)=0) == true exacto — error que nombra ASE + empresa;
    /// (ii) DetValiRetri D16..D20 == 0 ±0.5 por ASE y D24..D28 == true exacto — error que nombra
    ///     ASE + celda; la fila Total D21 y D9:D14/J9:J14 ≠ ROUND quedan EXCLUIDAS (D6);
    /// (iii) VALIDACION_TOTAL O9 == 0 ±0.5 y P9 == true exacto;
    /// (iv) controles Valida -* = informativos (T0: sin celda numérica estable; sin gate).
    /// El resto de la validación HU-01..HU-12 queda intacta; sin snapshots = HU-12 puro.
    /// </remarks>
    List<string> Validar(ResultadoRemuneracion resultado, IReadOnlyList<WorkbookLeafInputs> leafs, IReadOnlyList<ValidacionCruzadaSnapshot> snapshots);
}
