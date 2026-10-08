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
    /// HU-08 (2.2) / HU-20: lee las 5 hojas <c>Recaudo *</c> desde los archivos de conciliación
    /// <c>{periodo}/Conciliaciones/Conjunta {prefijo}*.xlsx</c>. La quincena la gobierna
    /// <see cref="Periodo.NumeroQuincena"/> (G2-D2: dominio, nunca fuente): Q1 lee D/E (VALOR 1°Q),
    /// Q2 lee F/G (VALOR 2°Q). Fail-fast: si falta el archivo de conciliación de una empresa,
    /// lanza <c>ArchivoFuenteNoEncontradoException</c> que nombra la empresa y el prefijo.
    /// </summary>
    /// <param name="periodo">Período de la liquidación (dispatch Q1/Q2 del par de columnas).</param>
    /// <param name="rutaConciliacionPorEmpresa">Función que resuelve la ruta del archivo de
    /// conciliación para una empresa (devuelve <c>null</c> si no existe).</param>
    /// <returns>Inputs de las hojas <c>Recaudo *</c> (5 empresas).</returns>
    IReadOnlyList<RecaudoEmpresaInputs> LeerRecaudosEmpresa(Periodo periodo, Func<EmpresaFacturacion, string?> rutaConciliacionPorEmpresa);

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
    /// D/E aplicada es la del veredicto Plan 29 (T1, Unidad B — «header-manda»): template-D
    /// (SUBSIDIO) ← columna E-fuente; template-E (CONTRIBUCION) ← columna F-fuente. El veredicto
    /// T0-0.2 del Plan 10 («D=CONTRIBUCION») quedó refutado por V4: el header de la plantilla
    /// (D2=SUBSIDIO / E2=CONTRIBUCION), el manual y la fuente lo contradicen, y como F=D+E es
    /// conmutativa, el golden ±0.5 no podía detectar el swap.
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
    /// Plan 21 (T2, R-E-1/R-E-6): lee la fuente R1 (<c>Recaudoporcomponente_*</c>, hoja
    /// <c>Sheet1</c>) como BLOQUE ESPEJO del ASE: la secuencia ordenada de filas tipadas
    /// (<see cref="FilaEspejoR1"/>) con sus valores por ENCABEZADO de columna (detección
    /// dinámica, nunca por índice fijo). No exige cardinalidades ni ocurrencias (la secuencia
    /// observada ES la especificación del período — D-B/R-E-1); la columna
    /// <c>SERVICIO ESPECIALES</c> es opcional (ausente → 0, R-E-6). Delimita la zona por
    /// etiquetas y exige las invariantes duras de cierre T0e (<c>Componente/Total</c>,
    /// <c>Subs/Cont/Total</c> y <c>Total</c> final); si falta una, lanza
    /// <c>CalculoInvalidoException</c> nombrando ASE + reporte + fila esperada.
    ///
    /// Plan 21 (T5): esta es la ÚNICA vía de lectura/gobernanza de la columna L-menores del R1;
    /// el path legado rol/ocurrencia (<c>LeerLEspecialesMenores</c> + mapa T0-0.5) se retiró al
    /// probar la absorción 100% en los 15 escenarios (5 ASE × Q1/Q2/agosto).
    /// </summary>
    /// <param name="ase">ASE asociado.</param>
    /// <param name="rutaR1">Ruta del archivo fuente R1 del ASE.</param>
    /// <returns>Bloque espejo del ASE (secuencia de filas + encabezados detectados).</returns>
    BloqueEspejoAseInputs LeerEspejoR1(Ase ase, string rutaR1);

    /// <summary>
    /// Plan 29 (T2, Unidad R — SOLO LECTURA): lee la matriz de detalle por componente de la fuente
    /// R2 (<c>RerpoteDetalleSaldosaFavor_*</c>, hoja <c>Sheet1</c>). Localiza las filas por LABEL
    /// (A–D) y las columnas por ENCABEZADO de componente (col E..), tolerante a la deriva de la
    /// malla entre períodos (julio E..P sin <c>Especiales</c>; agosto E..Q con <c>Especiales</c> —
    /// T0b §2.1). Fail-fast con archivo+fila si falta la primera <c>Vlr Servicio</c>, la fila de
    /// encabezados de componente o la fila de cierre <c>Total</c>. La columna <c>Especiales</c>
    /// ausente se emite como <c>0</c> explícito (nunca inventada). Esta lectura NO escribe nada.
    /// </summary>
    /// <param name="ase">ASE asociado.</param>
    /// <param name="rutaR2">Ruta del archivo fuente R2 del ASE.</param>
    /// <returns>Matriz de detalle R2 (filas tipadas + componentes por encabezado).</returns>
    DetalleR2AseInputs LeerDetalleR2(Ase ase, string rutaR2);

    /// <summary>
    /// Plan 29 (T2, Unidad R — SOLO LECTURA): lee la matriz de detalle por componente de la fuente
    /// R4 (<c>ReversiónPorComponente_*</c>, hoja <c>Sheet1</c>). Localiza las filas por LABEL (A–C)
    /// y las columnas por ENCABEZADO de componente (col D..), tolerante a la deriva de la malla
    /// entre períodos (agosto reduce las filas de PROMO/CIUDAD/BOGOTA — T0b §2.3). Fail-fast con
    /// archivo+fila si falta la primera <c>Vlr Servicio</c>, la fila de encabezados o la fila de
    /// cierre <c>Total</c>. Esta lectura NO escribe nada.
    /// </summary>
    /// <param name="ase">ASE asociado.</param>
    /// <param name="rutaR4">Ruta del archivo fuente R4 del ASE.</param>
    /// <returns>Matriz de detalle R4 (filas tipadas + componentes por encabezado).</returns>
    DetalleR4AseInputs LeerDetalleR4(Ase ase, string rutaR4);
}
