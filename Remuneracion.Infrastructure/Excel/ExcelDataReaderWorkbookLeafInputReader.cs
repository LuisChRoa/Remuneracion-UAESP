using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Reader dedicado a poblar los inputs leaf del workbook a partir de los archivos R1/R2/R4 reales.
/// Localiza filas por labels/headers; no usa números de fila fijos de las fuentes.
///
/// HU-07 (multi-ASE): el mapeo es POR BLOQUE, congelado por T0 en
/// <c>plans/07 - HU-07 T0 Evidencia.md</c>. Los labels de las fuentes ASE2..5 difieren de ASE1:
/// - ASE2/ASE5 NO tienen fila 'Aplicacion nuevos x reversion' (su EXTEMP es valor estático 0);
/// - ASE4 usa Aplicacion[1] como primer operando EXTEMP (no Subsidio[0]);
/// - TOT_OPT es fórmula de 5 términos para ASE2..5 (no 3 como ASE1).
/// Las propiedades ASE1 (F25/F41/L25/F30/F10/L10) se conservan con semántica por bloque:
/// F25 = primera fila Mes/Total col F, F41 = segunda, L25 = especiales de la primera,
/// F30/F10 = operandos EXTEMP del bloque.
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
            R1 = MapearR1(ase, filasR1),
            R2 = MapearR2(ase, filasR2),
            R4 = MapearR4(ase, filasR4)
        };

        WorkbookLeafCoherence.ValidarContraFuentes(leaf, r1, r2, r4);
        return leaf;
    }

    private static WorkbookLeafInputsR1 MapearR1(Ase ase, List<object?[]> filas)
    {
        var indiceEspeciales = ExcelWorksheetNavigator.IndiceColumnaPorEncabezado(filas, "SERVICIO ESPECIALES");
        var filasMes = filas
            .Where(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1)).Equals("Mes", StringComparison.OrdinalIgnoreCase)
                && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(2)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var filasAplicacion = filas
            .Where(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1))
                .Contains("Aplicacion nuevos x reversion", StringComparison.OrdinalIgnoreCase)
                && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(2)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var filasSubsidio = filas
            .Where(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(4))
                .Contains("Subsidio(-)/Contribucion(+)", StringComparison.OrdinalIgnoreCase))
            .ToList();

        decimal Esp(object?[]? fila) =>
            indiceEspeciales >= 0
                ? ExcelWorksheetNavigator.CeldaNumero(fila?.ElementAtOrDefault(indiceEspeciales))
                : 0m;

        decimal F(object?[]? fila) => ExcelWorksheetNavigator.CeldaNumero(fila?.ElementAtOrDefault(5));

        // --- TOT_OPT: 3 términos para ASE1; 5 términos para ASE2..5 ---
        decimal totOpt;
        decimal extemp;
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        var filaMes0 = filasMes.ElementAtOrDefault(0);
        var filaMes1 = filasMes.ElementAtOrDefault(1);
        var filaMes2 = filasMes.ElementAtOrDefault(2);

        switch (ase.Id)
        {
            case 1:
                if (filasMes.Count < 2)
                {
                    throw new CalculoInvalidoException(
                        $"ASE {ase.Id}: no se encontraron las dos filas 'Mes/Total' del R1 necesarias para F25 (extemporáneo) y F41 (subsidio/contribución).");
                }

                var filaExtemp = filaMes0;
                var filaSubsMes = filaMes1;

                var filaAplicacion = filasAplicacion.FirstOrDefault()
                    ?? throw new CalculoInvalidoException($"ASE {ase.Id}: no se encontró la fila 'Aplicacion nuevos x reversion' del R1 para F10.");

                if (filasSubsidio.Count == 0)
                {
                    throw new CalculoInvalidoException($"ASE {ase.Id}: no se encontró ninguna fila 'Subsidio(-)/Contribucion(+)' en R1 para F30.");
                }

                celdas["F25"] = F(filaExtemp);
                celdas["F41"] = F(filaSubsMes);
                celdas["L25"] = Esp(filaExtemp);
                celdas["F30"] = F(filasSubsidio[0]);
                celdas["F10"] = F(filaAplicacion);
                celdas["L10"] = 0m;
                totOpt = celdas["F25"] + celdas["F41"] - celdas["L25"];
                extemp = celdas["F30"] + celdas["F10"] - celdas["L10"];
                break;

            case 2:
                // F176 = F113+F130-L113+F90-L90 ; EXTEMP F178 es valor estático 0 (T0-0.2).
                if (filasMes.Count < 3)
                {
                    throw new CalculoInvalidoException(
                        $"ASE {ase.Id}: se requieren tres filas 'Mes/Total' del R1 para la fórmula TOT_OPT de 5 términos.");
                }

                celdas["F113"] = F(filaMes1);
                celdas["F130"] = F(filaMes2);
                celdas["L113"] = Esp(filaMes1);
                celdas["F90"] = F(filaMes0);
                celdas["L90"] = Esp(filaMes0);
                totOpt = celdas["F113"] + celdas["F130"] - celdas["L113"] + celdas["F90"] - celdas["L90"];
                extemp = 0m;
                break;

            case 3:
                // F316 = F238+F254+F217-L217-L238 ; F318 = F243+F223-L223.
                if (filasMes.Count < 3)
                {
                    throw new CalculoInvalidoException(
                        $"ASE {ase.Id}: se requieren tres filas 'Mes/Total' del R1 para la fórmula TOT_OPT de 5 términos.");
                }

                var filaAplicacion3 = filasAplicacion.FirstOrDefault()
                    ?? throw new CalculoInvalidoException($"ASE {ase.Id}: no se encontró la fila 'Aplicacion nuevos x reversion' del R1 para F223.");

                if (filasSubsidio.Count == 0)
                {
                    throw new CalculoInvalidoException($"ASE {ase.Id}: no se encontró ninguna fila 'Subsidio(-)/Contribucion(+)' en R1 para F243.");
                }

                celdas["F238"] = F(filaMes1);
                celdas["F254"] = F(filaMes2);
                celdas["F217"] = F(filaMes0);
                celdas["L217"] = Esp(filaMes0);
                celdas["L238"] = Esp(filaMes1);
                celdas["F243"] = F(filasSubsidio[0]);
                celdas["F223"] = F(filaAplicacion3);
                celdas["L223"] = 0m;
                totOpt = celdas["F238"] + celdas["F254"] + celdas["F217"] - celdas["L217"] - celdas["L238"];
                extemp = celdas["F243"] + celdas["F223"] - celdas["L223"];
                break;

            case 4:
                // F437 = F391+F417+F357-L357-L391 ; F439 = F401+F369-L369.
                // ASE4: el primer operando EXTEMP es Aplicacion[1], el segundo Aplicacion[0] (T0-0.3).
                if (filasMes.Count < 3)
                {
                    throw new CalculoInvalidoException(
                        $"ASE {ase.Id}: se requieren tres filas 'Mes/Total' del R1 para la fórmula TOT_OPT de 5 términos.");
                }

                if (filasAplicacion.Count < 2)
                {
                    throw new CalculoInvalidoException(
                        $"ASE {ase.Id}: se requieren dos filas 'Aplicacion nuevos x reversion' del R1 para F401 y F369.");
                }

                celdas["F391"] = F(filaMes1);
                celdas["F417"] = F(filaMes2);
                celdas["F357"] = F(filaMes0);
                celdas["L357"] = Esp(filaMes0);
                celdas["L391"] = Esp(filaMes1);
                celdas["F401"] = F(filasAplicacion[1]);
                celdas["F369"] = F(filasAplicacion[0]);
                celdas["L369"] = 0m;
                totOpt = celdas["F391"] + celdas["F417"] + celdas["F357"] - celdas["L357"] - celdas["L391"];
                extemp = celdas["F401"] + celdas["F369"] - celdas["L369"];
                break;

            case 5:
                // F519 = F513+F498+F478-L478-L498 ; EXTEMP F521 es valor estático 0 (T0-0.2).
                if (filasMes.Count < 3)
                {
                    throw new CalculoInvalidoException(
                        $"ASE {ase.Id}: se requieren tres filas 'Mes/Total' del R1 para la fórmula TOT_OPT de 5 términos.");
                }

                celdas["F513"] = F(filaMes2);
                celdas["F498"] = F(filaMes1);
                celdas["F478"] = F(filaMes0);
                celdas["L478"] = Esp(filaMes0);
                celdas["L498"] = Esp(filaMes1);
                totOpt = celdas["F513"] + celdas["F498"] + celdas["F478"] - celdas["L478"] - celdas["L498"];
                extemp = 0m;
                break;

            default:
                throw new CalculoInvalidoException($"ASE {ase.Id} no está soportado por el lector leaf.");
        }

        return new WorkbookLeafInputsR1
        {
            F25 = F(filaMes0),
            F41 = F(filaMes1),
            L25 = Esp(filaMes0),
            F30 = ObtenerF30(ase, filasSubsidio, filasAplicacion),
            F10 = ObtenerF10(ase, filasAplicacion),
            L10 = 0m,
            CeldasPorAse = celdas,
            TotalOportunoEsperadoPorAse = totOpt,
            ExtemporaneoEsperadoPorAse = extemp
        };
    }

    private static WorkbookLeafInputsR2 MapearR2(Ase ase, List<object?[]> filas)
    {
        var indiceEspeciales = ExcelWorksheetNavigator.IndiceColumnaPorEncabezado(filas, "Especiales");

        var filaComponente = filas.FirstOrDefault(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(0)).Equals("Componente", StringComparison.OrdinalIgnoreCase)
            && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException($"ASE {ase.Id}: no se encontró la fila Componente/Total del R2 para E15.");

        var filaSubsCont = filas.FirstOrDefault(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(0)).Equals("Subs/Cont", StringComparison.OrdinalIgnoreCase)
            && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException($"ASE {ase.Id}: no se encontró la fila Subs/Cont/Total del R2 para E26.");

        var e15 = ExcelWorksheetNavigator.CeldaNumero(filaComponente.ElementAtOrDefault(4));
        var e26 = ExcelWorksheetNavigator.CeldaNumero(filaSubsCont.ElementAtOrDefault(4));
        var k15 = indiceEspeciales >= 0
            ? ExcelWorksheetNavigator.CeldaNumero(filaComponente.ElementAtOrDefault(indiceEspeciales))
            : 0m;

        var celdas = ase.Id switch
        {
            1 => new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
            {
                ["E15"] = e15, ["E26"] = e26, ["K15"] = k15
            },
            2 => new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
            {
                ["E85"] = e15, ["E103"] = e26, ["K85"] = k15
            },
            3 => new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
            {
                ["E167"] = e15, ["E178"] = e26, ["K167"] = k15
            },
            4 => new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
            {
                ["E290"] = e15, ["E308"] = e26, ["K290"] = k15
            },
            5 => new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
            {
                // T0: template E374 = Componente (e15), E385 = Subs/Cont (e26); E413 = E385+E374-K374.
                ["E374"] = e15, ["E385"] = e26, ["K374"] = k15
            },
            _ => throw new CalculoInvalidoException($"ASE {ase.Id} no está soportado por el lector leaf.")
        };

        return new WorkbookLeafInputsR2
        {
            E15 = e15,
            E26 = e26,
            K15 = k15,
            CeldasPorAse = celdas
        };
    }

    private static WorkbookLeafInputsR4 MapearR4(Ase ase, List<object?[]> filas)
    {
        var filaTotal = filas.FirstOrDefault(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(0)).Equals("Total", StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrWhiteSpace(ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1))))
            ?? throw new CalculoInvalidoException($"ASE {ase.Id}: no se encontró la fila Total (B vacío) del R4 para D9.");

        var d9 = ExcelWorksheetNavigator.CeldaNumero(filaTotal.ElementAtOrDefault(3));
        var p9 = 0m;

        var celdas = ase.Id switch
        {
            1 => new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase) { ["D9"] = d9, ["P9"] = p9 },
            2 => new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase) { ["D98"] = d9, ["P98"] = p9 },
            3 => new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase) { ["D193"] = d9, ["P193"] = p9 },
            4 => new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase) { ["D236"] = d9, ["P236"] = p9 },
            5 => new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase) { ["D344"] = d9, ["P344"] = p9 },
            _ => throw new CalculoInvalidoException($"ASE {ase.Id} no está soportado por el lector leaf.")
        };

        return new WorkbookLeafInputsR4
        {
            D9 = d9,
            P9 = p9,
            CeldasPorAse = celdas
        };
    }

    private static decimal ObtenerF30(Ase ase, List<object?[]> filasSubsidio, List<object?[]> filasAplicacion)
    {
        if (ase.Id is 1 or 3)
        {
            return filasSubsidio.Count > 0
                ? ExcelWorksheetNavigator.CeldaNumero(filasSubsidio[0].ElementAtOrDefault(5))
                : 0m;
        }

        if (ase.Id == 4)
        {
            return filasAplicacion.Count > 1
                ? ExcelWorksheetNavigator.CeldaNumero(filasAplicacion[1].ElementAtOrDefault(5))
                : 0m;
        }

        return 0m;
    }

    private static decimal ObtenerF10(Ase ase, List<object?[]> filasAplicacion)
    {
        if (ase.Id is 1 or 3 or 4)
        {
            return filasAplicacion.Count > 0
                ? ExcelWorksheetNavigator.CeldaNumero(filasAplicacion[0].ElementAtOrDefault(5))
                : 0m;
        }

        return 0m;
    }
}