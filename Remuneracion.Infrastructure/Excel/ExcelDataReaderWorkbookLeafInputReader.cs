using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Reader dedicado a poblar los inputs leaf del workbook a partir de los archivos R1/R2/R4 reales.
/// Localiza filas por labels/headers; no usa números de fila fijos de las fuentes.
///
/// Mapeo verificado contra <c>Docs/Insumos/Remuneracion 202607-1 Total.xlsx</c> (Promoambiental):
/// R1 F25 = primera fila Mes/Total col F (Extemporáneo HU-02);
/// R1 F41 = segunda fila Mes/Total col F (Subs/Cont de Mes);
/// R1 L25 = primera fila Mes/Total col L (SERVICIO ESPECIALES);
/// R1 F30 = primera fila Subsidio(-)/Contribucion(+) col F (extemporáneo);
/// R1 F10 = fila Aplicacion nuevos x reversion col F (extemporáneo);
/// R1 L10 = 0 en la plantilla de referencia (sin especiales en esa sección);
/// R2 E15 = fila Componente/Total col E;
/// R2 E26 = fila Subs/Cont/Total col E;
/// R2 K15 = Especiales de Componente/Total (0 si no hay columna);
/// R4 D9 = fila Total con B vacío col D;
/// R4 P9 = 0 (plantilla de referencia).
/// </summary>
public sealed class ExcelDataReaderWorkbookLeafInputReader : IWorkbookLeafInputReader
{
    public WorkbookLeafInputs LeerLeafInputs(Ase ase, Periodo periodo, string rutaR1, string rutaR2, string rutaR4)
    {
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(rutaR1);
        ArgumentNullException.ThrowIfNull(rutaR2);
        ArgumentNullException.ThrowIfNull(rutaR4);

        var recaudoReader = new ExcelDataReaderRecaudoReader();
        var r1 = recaudoReader.LeerR1(rutaR1);
        var r2 = recaudoReader.LeerR2(rutaR2);
        var r4 = recaudoReader.LeerR4(rutaR4);

        var filasR1 = ExcelWorksheetNavigator.LeerFilas(rutaR1);
        var filasR2 = ExcelWorksheetNavigator.LeerFilas(rutaR2);
        var filasR4 = ExcelWorksheetNavigator.LeerFilas(rutaR4);

        var leaf = new WorkbookLeafInputs
        {
            Ase = ase,
            Periodo = periodo,
            R1 = MapearR1(filasR1),
            R2 = MapearR2(filasR2),
            R4 = MapearR4(filasR4)
        };

        WorkbookLeafCoherence.ValidarContraFuentes(leaf, r1, r2, r4);
        return leaf;
    }

    private static WorkbookLeafInputsR1 MapearR1(List<object?[]> filas)
    {
        var indiceEspeciales = ExcelWorksheetNavigator.IndiceColumnaPorEncabezado(filas, "SERVICIO ESPECIALES");
        var filasMes = filas
            .Where(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1)).Equals("Mes", StringComparison.OrdinalIgnoreCase)
                && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(2)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (filasMes.Count < 2)
        {
            throw new CalculoInvalidoException(
                "No se encontraron las dos filas 'Mes/Total' del R1 necesarias para F25 (extemporáneo) y F41 (subsidio/contribución).");
        }

        var filaExtemp = filasMes[0];
        var filaSubsMes = filasMes[1];

        var filaAplicacion = filas.FirstOrDefault(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1))
                .Contains("Aplicacion nuevos x reversion", StringComparison.OrdinalIgnoreCase)
            && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(2)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException("No se encontró la fila 'Aplicacion nuevos x reversion' del R1 para F10.");

        var filasSubsidio = filas
            .Where(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(4))
                .Contains("Subsidio(-)/Contribucion(+)", StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (filasSubsidio.Count == 0)
        {
            throw new CalculoInvalidoException("No se encontró ninguna fila 'Subsidio(-)/Contribucion(+)' en R1 para F30.");
        }

        return new WorkbookLeafInputsR1
        {
            F25 = ExcelWorksheetNavigator.CeldaNumero(filaExtemp.ElementAtOrDefault(5)),
            F41 = ExcelWorksheetNavigator.CeldaNumero(filaSubsMes.ElementAtOrDefault(5)),
            L25 = indiceEspeciales >= 0
                ? ExcelWorksheetNavigator.CeldaNumero(filaExtemp.ElementAtOrDefault(indiceEspeciales))
                : 0m,
            F30 = ExcelWorksheetNavigator.CeldaNumero(filasSubsidio[0].ElementAtOrDefault(5)),
            F10 = ExcelWorksheetNavigator.CeldaNumero(filaAplicacion.ElementAtOrDefault(5)),
            L10 = 0m
        };
    }

    private static WorkbookLeafInputsR2 MapearR2(List<object?[]> filas)
    {
        var indiceEspeciales = ExcelWorksheetNavigator.IndiceColumnaPorEncabezado(filas, "Especiales");

        var filaComponente = filas.FirstOrDefault(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(0)).Equals("Componente", StringComparison.OrdinalIgnoreCase)
            && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException("No se encontró la fila Componente/Total del R2 para E15.");

        var filaSubsCont = filas.FirstOrDefault(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(0)).Equals("Subs/Cont", StringComparison.OrdinalIgnoreCase)
            && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException("No se encontró la fila Subs/Cont/Total del R2 para E26.");

        return new WorkbookLeafInputsR2
        {
            E15 = ExcelWorksheetNavigator.CeldaNumero(filaComponente.ElementAtOrDefault(4)),
            E26 = ExcelWorksheetNavigator.CeldaNumero(filaSubsCont.ElementAtOrDefault(4)),
            K15 = indiceEspeciales >= 0
                ? ExcelWorksheetNavigator.CeldaNumero(filaComponente.ElementAtOrDefault(indiceEspeciales))
                : 0m
        };
    }

    private static WorkbookLeafInputsR4 MapearR4(List<object?[]> filas)
    {
        var filaTotal = filas.FirstOrDefault(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(0)).Equals("Total", StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrWhiteSpace(ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1))))
            ?? throw new CalculoInvalidoException("No se encontró la fila Total (B vacío) del R4 para D9.");

        return new WorkbookLeafInputsR4
        {
            D9 = ExcelWorksheetNavigator.CeldaNumero(filaTotal.ElementAtOrDefault(3)),
            P9 = 0m
        };
    }
}
