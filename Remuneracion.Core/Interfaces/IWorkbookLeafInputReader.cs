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
}
