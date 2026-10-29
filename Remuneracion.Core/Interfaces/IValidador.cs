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

    /// <summary>
    /// Plan 42 (T1, R-CLS-1/R-CLS-2/R-CLS-3, D-B/D-F): clasifica los sub-bloques FALSE de
    /// <c>VALIDACION_TOTAL</c> en declarados (entrada vigente de
    /// <see cref="DivergenciasDeclaradasCatalogo"/>) vs NO declarados (defecto interno), y devuelve
    /// el resto de errores de la validación cruzada que no son sub-bloques declarados. Es la ÚNICA
    /// decisión de clasificación (Core): UI y CLI solo presentan el resultado.
    ///
    /// D-F (la declarada es la ÚLTIMA puerta): si <see cref="ClasificacionDivergenciasDeclaradas.HayNoDeclaradas"/>
    /// es verdadero, el run debe fallar en duro aunque la declarada esté reconocida.
    /// </summary>
    /// <param name="resultado">Resultado con los consolidados por ASE.</param>
    /// <param name="leafs">Inputs leaf, uno por ASE (5 en modo período).</param>
    /// <param name="snapshots">Snapshots-oráculo leídos read-only del workbook, uno por ASE.</param>
    /// <returns>Clasificación con los veredictos por ASE·celda y los errores no declarados.</returns>
    ClasificacionDivergenciasDeclaradas ClasificarDivergenciasDeclaradas(
        ResultadoRemuneracion resultado,
        IReadOnlyList<WorkbookLeafInputs> leafs,
        IReadOnlyList<ValidacionCruzadaSnapshot> snapshots);

    /// <summary>
    /// HU-20 (G3): valida el DetRetri CALCULADO por ASE (<see cref="WorkbookLeafInputs.DetRetriQ2"/>,
    /// ROUND(D104:D108,0)) contra el oráculo <c>R10</c> del período, tolerancia ±0.5 tras el
    /// redondeo a entero. El proceso puebla el detalle en AMBAS quincenas, por lo que el gate corre
    /// en Q1 y Q2; los leafs sin detalle (construcciones manuales de test) se omiten. Divergencia →
    /// mensaje con ASE + ambos valores; el procesador lo eleva a novedad de entrega.
    ///
    /// <para>Plan 62 (T3.2, D-61R10): solo se evalúa el gate PER-ASE. El gate de TOTAL
    /// calculado-vs-R10 se retiró: con el pegado (N1) la hoja DetRetri/DetValiRetri es el R10 por
    /// construcción y su integridad la verifica <c>VerificarIntegridadPegadoR10</c>.</para>
    /// </summary>
    /// <param name="leafs">Inputs leaf, uno por ASE (5 en modo período).</param>
    /// <param name="r10">DetRetri leído del R10 del período (oráculo).</param>
    /// <returns>Lista de mensajes de validación (vacía si no hay problemas).</returns>
    List<string> ValidarDetRetriContraR10(IReadOnlyList<WorkbookLeafInputs> leafs, DetRetriInputs r10);
}
