using Remuneracion.Core.Models;

namespace Remuneracion.Core.Interfaces;

/// <summary>
/// Contrato para poblar los inputs leaf reales del workbook a partir de las fuentes R1/R2/R4.
/// </summary>
public interface IWorkbookLeafInputReader
{
    /// <summary>
    /// Lee los inputs leaf editables del workbook para el ASE y periodo indicados.
    /// </summary>
    /// <param name="ase">ASE asociado al cálculo.</param>
    /// <param name="periodo">Periodo de la liquidación.</param>
    /// <param name="rutaR1">Ruta del archivo fuente R1.</param>
    /// <param name="rutaR2">Ruta del archivo fuente R2.</param>
    /// <param name="rutaR4">Ruta del archivo fuente R4.</param>
    /// <returns>Modelo <see cref="WorkbookLeafInputs"/> listo para la escritura real del workbook.</returns>
    WorkbookLeafInputs LeerLeafInputs(Ase ase, Periodo periodo, string rutaR1, string rutaR2, string rutaR4);

    /// <summary>
    /// HU-08 (2.2): lee la conciliación por empresa del ASE indicado a partir de las fuentes
    /// R1/R2/R4 (desglose por empresa, T0-0.4). Fail-fast: si falta un label de empresa en la
    /// fuente, lanza <c>CalculoInvalidoException</c> que nombra ASE y empresa.
    /// </summary>
    /// <param name="ase">ASE asociado.</param>
    /// <param name="periodo">Período de la liquidación.</param>
    /// <param name="rutaR1">Ruta del archivo fuente R1.</param>
    /// <param name="rutaR2">Ruta del archivo fuente R2.</param>
    /// <param name="rutaR4">Ruta del archivo fuente R4.</param>
    /// <returns>Inputs de conciliación por empresa (5 empresas × operandos R1/R2/R4).</returns>
    IReadOnlyList<ConciliacionEmpresaInputs> LeerConciliacionEmpresas(Ase ase, Periodo periodo, string rutaR1, string rutaR2, string rutaR4);

    /// <summary>
    /// HU-08 (2.2): lee las 5 hojas <c>Recaudo *</c> desde los archivos de conciliación
    /// <c>Consolidado/Conciliaciones/Conjunta {prefijo}*.xlsx</c> (T0-0.6). Fail-fast: si falta
    /// el archivo de conciliación de una empresa, lanza <c>ArchivoFuenteNoEncontradoException</c>
    /// que nombra la empresa.
    /// </summary>
    /// <param name="rutaConciliacionPorEmpresa">Función que resuelve la ruta del archivo de
    /// conciliación para una empresa (devuelve <c>null</c> si no existe).</param>
    /// <returns>Inputs de las hojas <c>Recaudo *</c> (5 empresas).</returns>
    IReadOnlyList<RecaudoEmpresaInputs> LeerRecaudosEmpresa(Func<EmpresaFacturacion, string?> rutaConciliacionPorEmpresa);

    /// <summary>
    /// HU-09 (2.3): lee el "Resumen Recaudo Aplicado Por Servicio" del final del
    /// <c>ReportePagosxBanco_*.xlsx</c> del ASE indicado (búsqueda dinámica desde el final,
    /// match de conceptos por prefijo normalizado — D1). Fail-fast: si falta la etiqueta o
    /// una empresa-columna esperada del mapa T0-0.7, lanza <c>CalculoInvalidoException</c>
    /// que nombra ASE y empresa-columna (nunca valor inventado). La quincena se toma de
    /// <see cref="Periodo.NumeroQuincena"/> (C59 = dominio, nunca fuente; Requirement 4).
    /// </summary>
    /// <param name="ase">ASE asociado.</param>
    /// <param name="periodo">Período de la liquidación.</param>
    /// <param name="rutaReportePagosxBanco">Ruta del archivo <c>ReportePagosxBanco_*</c>.</param>
    /// <returns>Inputs del bloque banco del ASE (conceptos por empresa + TotalFuente + C59).</returns>
    ReporteBancoInputs LeerReporteBanco(Ase ase, Periodo periodo, string rutaReportePagosxBanco);

    /// <summary>
    /// HU-10 (2.4): lee la fila "Total General" del final del <c>R4-BalanceSubsidioyContribuciones_*.xlsx</c>
    /// del ASE indicado (búsqueda dinámica desde el final, match normalizado — D1). Fail-fast:
    /// si falta la etiqueta o la coherencia E+F vs columna G no cierra ±0.5, lanza
    /// <c>CalculoInvalidoException</c> que nombra el ASE (nunca valor inventado). La asignación
    /// D/E aplicada es la del veredicto T0-0.2 (hipótesis líder probada): template-D
    /// (CONTRIBUCION) ← columna F-fuente; template-E (SUBSIDIO) ← columna E-fuente.
    /// Columna H (SISTEMA) = D2(b): es fórmula en el template → <see cref="BalanceScAseInputs.Sistema"/>
    /// queda null y H entra al mapa de fórmulas protegidas.
    /// </summary>
    /// <param name="ase">ASE asociado.</param>
    /// <param name="periodo">Período de la liquidación (sin uso directo; Q2 sigue bloqueada aguas arriba).</param>
    /// <param name="rutaBalance">Ruta del archivo <c>R4-BalanceSubsidioyContribuciones_*</c>.</param>
    /// <returns>Inputs del balance del ASE (Subsidio + Contribucion + TotalFuente).</returns>
    BalanceScInputs LeerBalanceSc(Ase ase, Periodo periodo, string rutaBalance);

    /// <summary>
    /// HU-11 (2.5, D3): lee el bloque SALDOS POR NOTA del ASE desde
    /// <c>SaldosaFavorAplicadosPorNotas_*.xlsx</c> (búsqueda header-driven por títulos, patrón R2;
    /// columna "Especiales" opcional — ausente → 0). Fail-fast: si falta un header esperado del
    /// mapa T0-0.7, lanza <c>CalculoInvalidoException</c> que nombra el ASE (nunca valor inventado).
    /// La aritmética de dominio <see cref="SaldosNotasAseInputs.TotalSaldosNotas"/> replica el
    /// visible Cn-In del template (composición T0-0.3).
    /// </summary>
    /// <param name="ase">ASE asociado.</param>
    /// <param name="rutaSaldosNotas">Ruta del archivo <c>SaldosaFavorAplicadosPorNotas_*</c>.</param>
    /// <returns>Inputs del bloque SALDOS POR NOTA del ASE.</returns>
    SaldosNotasAseInputs LeerSaldosNotas(Ase ase, string rutaSaldosNotas);

    /// <summary>
    /// HU-11 (2.5, D3): lee el bloque RETRIBUCION NEGATIVA del ASE desde
    /// <c>RetribuciónNegativa_*.xlsx</c> (búsqueda header-driven por títulos; valores negativos
    /// por componente). Distingue "fuente vacía = 0 legítimo" (Q2 trae archivos solo con el rango
    /// de fechas) de "header ausente = fallo que nombra ASE + reporte" (Riesgo 6).
    /// </summary>
    /// <param name="ase">ASE asociado.</param>
    /// <param name="rutaRetribucionNegativa">Ruta del archivo <c>RetribuciónNegativa_*</c>.</param>
    /// <returns>Inputs del bloque RETRIBUCION NEGATIVA del ASE.</returns>
    RetribucionNegativaAseInputs LeerRetribucionNegativa(Ase ase, string rutaRetribucionNegativa);

    /// <summary>
    /// HU-16 (D3a): lee las L-Especiales menores del <c>Reporte Componentes R1</c> del ASE desde
    /// la misma fuente <c>Recaudoporcomponente_*</c> (columna SERVICIO ESPECIALES por rol de fila
    /// congelado en T0-0.5, dispatch por período). Fail-fast: si la fila del rol/ocurrencia no
    /// existe en la fuente, lanza <c>CalculoInvalidoException</c> que nombra ASE + hoja + celda
    /// (doctrina "slot ausente" ≠ "leído 0"; nunca 0 silencioso en operando).
    /// </summary>
    /// <param name="ase">ASE asociado.</param>
    /// <param name="periodo">Período de la liquidación (dispatch Q1/Q2 del mapa).</param>
    /// <param name="rutaR1">Ruta del archivo fuente R1 del ASE.</param>
    /// <returns>Inputs L-menores del ASE (celdas del template → valor ESP de la fuente).</returns>
    LEspecialesMenoresAseInputs LeerLEspecialesMenores(Ase ase, Periodo periodo, string rutaR1);
}
