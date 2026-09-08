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

    /// <inheritdoc />
    public IReadOnlyList<ConciliacionEmpresaInputs> LeerConciliacionEmpresas(
        Ase ase, Periodo periodo, string rutaR1, string rutaR2, string rutaR4)
    {
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(rutaR1);
        ArgumentNullException.ThrowIfNull(rutaR2);
        ArgumentNullException.ThrowIfNull(rutaR4);

        var filasR1 = ExcelWorksheetNavigator.LeerFilas(rutaR1);
        var filasR2 = ExcelWorksheetNavigator.LeerFilas(rutaR2);
        var filasR4 = ExcelWorksheetNavigator.LeerFilas(rutaR4);

        var resultado = new List<ConciliacionEmpresaInputs>();
        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            var inputs = new ConciliacionEmpresaInputs
            {
                Empresa = empresa,
                Ase = ase,
                CeldasR1 = ExtraerCeldasR1(ase, empresa, filasR1),
                CeldasR2 = ExtraerCeldasR2(ase, empresa, filasR2),
                CeldasR4 = ExtraerCeldasR4(ase, empresa, filasR4)
            };

            inputs.VisibleR1 = CalcularVisibleR1(ase.Id, empresa.Id, inputs.CeldasR1);
            inputs.VisibleR2 = CalcularVisibleR2(inputs.CeldasR2);
            inputs.VisibleR4 = CalcularVisibleR4(inputs.CeldasR4);

            resultado.Add(inputs);
        }

        return resultado;
    }

    /// <inheritdoc />
    public IReadOnlyList<RecaudoEmpresaInputs> LeerRecaudosEmpresa(
        Func<EmpresaFacturacion, string?> rutaConciliacionPorEmpresa)
    {
        ArgumentNullException.ThrowIfNull(rutaConciliacionPorEmpresa);

        var resultado = new List<RecaudoEmpresaInputs>();
        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            var ruta = rutaConciliacionPorEmpresa(empresa)
                ?? throw new ArchivoFuenteNoEncontradoException(
                    $"No se encontró el archivo de conciliación de {empresa.Nombre} (prefijo '{empresa.PrefijoConciliacion}') en Consolidado/Conciliaciones.");

            resultado.Add(LeerRecaudoEmpresa(empresa, ruta));
        }

        return resultado;
    }

    /// <inheritdoc />
    public ReporteBancoInputs LeerReporteBanco(Ase ase, Periodo periodo, string rutaReportePagosxBanco)
    {
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(rutaReportePagosxBanco);

        // HU-09 (2.3, D1): búsqueda dinámica del Resumen desde el FINAL del Sheet1. Prohibida
        // fila fija: el conteo de filas diarias varía por ASE (V4/V6).
        var filas = ExcelWorksheetNavigator.LeerFilas(rutaReportePagosxBanco);
        var indiceResumen = BuscarEtiquetaDesdeElFinal(filas, WorkbookLeafCellMapReporteBanco.PrefijoEtiquetaResumen);
        if (indiceResumen < 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: no se encontró la etiqueta 'Resumen Recaudo Aplicado Por Servicio' en el ReportePagosxBanco.");
        }

        var filaHeaders = filas.ElementAtOrDefault(indiceResumen + 1)
            ?? throw new CalculoInvalidoException(
                $"ASE {ase.Id}: no se encontró la fila de encabezados de empresas después del Resumen en el ReportePagosxBanco.");

        // Mapa de columnas fuente por ASE (T0-0.3, V10): empresa → índice de columna (1-based).
        // Fail-fast: si la columna esperada no trae el header exacto, falla nombrando ASE+empresa-columna.
        var columnas = WorkbookLeafCellMapReporteBanco.ColumnasFuentePorAse[ase.Id];
        var indicesEmpresa = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var (empresa, indice) in columnas)
        {
            var header = ExcelWorksheetNavigator.CeldaTexto(filaHeaders.ElementAtOrDefault(indice));
            if (!header.Equals(empresa, StringComparison.OrdinalIgnoreCase))
            {
                throw new CalculoInvalidoException(
                    $"ASE {ase.Id}: la columna fuente {indice} del Resumen debería ser '{empresa}' pero dice '{header}'. Mapa T0-0.3 no coincide con la fuente.");
            }

            indicesEmpresa[empresa] = indice;
        }

        // Filas de concepto (1/2/3/7) + fila Total, por prefijo normalizado (T0-0.4).
        // Fila ausente = 0 demostrable: el Total de la fuente lo confirma (Riesgo 3 mitigado).
        var conceptosPorEmpresa = indicesEmpresa.Keys.ToDictionary(e => e, _ => new decimal[4], StringComparer.OrdinalIgnoreCase);
        var totalesFuente = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        for (var i = indiceResumen + 2; i < filas.Count; i++)
        {
            var fila = filas[i];
            if (fila is null)
            {
                continue;
            }

            var etiqueta = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(0));
            if (string.IsNullOrWhiteSpace(etiqueta))
            {
                continue;
            }

            var normalizada = NormalizarEtiqueta(etiqueta);

            // Fila "Total" del Resumen: cierra los conceptos y aporta el Total por empresa
            // (gate D5-i: Σ conceptos == Total fuente).
            if (normalizada == "total")
            {
                foreach (var (empresa, indice) in indicesEmpresa)
                {
                    totalesFuente[empresa] = ExcelWorksheetNavigator.CeldaNumero(fila.ElementAtOrDefault(indice));
                }

                break;
            }

            foreach (var (indiceConcepto, prefijo) in WorkbookLeafCellMapReporteBanco.Conceptos)
            {
                if (!normalizada.StartsWith(prefijo, StringComparison.Ordinal))
                {
                    continue;
                }

                foreach (var (empresa, indice) in indicesEmpresa)
                {
                    conceptosPorEmpresa[empresa][indiceConcepto] = ExcelWorksheetNavigator.CeldaNumero(fila.ElementAtOrDefault(indice));
                }

                break;
            }
        }

        // Los conceptos ausentes quedaron en 0 (fila ausente = 0 demostrable).
        var empresas = indicesEmpresa.Keys
            .OrderBy(e => indicesEmpresa[e])
            .Select(empresa => new ReporteBancoEmpresaInputs
            {
                Empresa = empresa,
                AplicadosFacturacion = conceptosPorEmpresa[empresa][0],
                SaldosFavorGenerados = conceptosPorEmpresa[empresa][1],
                FinanciacionesNuevas = conceptosPorEmpresa[empresa][2],
                RecibosServEspeciales = conceptosPorEmpresa[empresa][3],
                TotalFuente = totalesFuente.GetValueOrDefault(empresa)
            })
            .ToList();

        // Gate D5-i (coherencia): bloque == resumen fuente antes de devolver (fail-fast ASE+empresa).
        WorkbookLeafCoherence.ValidarContraFuentesBanco(new ReporteBancoAseInputs { Ase = ase, Empresas = empresas }, empresas);

        return new ReporteBancoInputs
        {
            Ases = [new ReporteBancoAseInputs { Ase = ase, Empresas = empresas }],
            Quincena = periodo.NumeroQuincena,
            Consolidado = null // D2(b): filas 1–7 son fórmulas (T0-0.1); se protegen, no se escriben.
        };
    }

    /// <summary>
    /// HU-09 (2.3): busca la fila del Resumen desde el final del <c>Sheet1</c> por prefijo
    /// normalizado (case-insensitive, sin espacios). Nunca fila fija (V4/V6).
    /// </summary>
    private static int BuscarEtiquetaDesdeElFinal(List<object?[]> filas, string prefijoNormalizado)
    {
        for (var i = filas.Count - 1; i >= 0; i--)
        {
            var fila = filas[i];
            if (fila is null)
            {
                continue;
            }

            for (var j = 0; j < fila.Length; j++)
            {
                var texto = ExcelWorksheetNavigator.CeldaTexto(fila[j]);
                if (!string.IsNullOrWhiteSpace(texto)
                    && NormalizarEtiqueta(texto).StartsWith(prefijoNormalizado, StringComparison.Ordinal))
                {
                    return i;
                }
            }
        }

        return -1;
    }

    /// <summary>
    /// HU-09 (2.3): normaliza una etiqueta para el match por prefijo (T0-0.4): minúsculas,
    /// sin espacios, sin signos de puntuación. "3-APLICADOS A FINANCIACIONES NUEVAS" y
    /// "3-APLICADOS A FINANCIACIONES" matchean el mismo prefijo.
    /// </summary>
    private static string NormalizarEtiqueta(string texto)
    {
        var sb = new System.Text.StringBuilder(texto.Length);
        foreach (var ch in texto)
        {
            if (char.IsLetterOrDigit(ch) || ch == '-')
            {
                sb.Append(char.ToLowerInvariant(ch));
            }
        }

        return sb.ToString();
    }

    /// <inheritdoc />
    public BalanceScInputs LeerBalanceSc(Ase ase, Periodo periodo, string rutaBalance)
    {
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(rutaBalance);

        // HU-10 (2.4, D1): búsqueda dinámica de la etiqueta "Total General" desde el FINAL del
        // Sheet1. Prohibida fila fija: la posición varía por ASE (V6: R118/R110/R38/R42/R28).
        var filas = ExcelWorksheetNavigator.LeerFilas(rutaBalance);
        var indiceTotalGeneral = BuscarTotalGeneralDesdeElFinal(filas);
        if (indiceTotalGeneral < 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: no se encontró la etiqueta 'Total General' en el Balance de subsidios y contribuciones ({Path.GetFileName(rutaBalance)}).");
        }

        var fila = filas[indiceTotalGeneral];

        // Veredicto T0-0.2 (hipótesis líder PROBADA en los 5 ASE Q1): template-D (CONTRIBUCION,
        // positivo) ← columna F-fuente; template-E (SUBSIDIO, negativo) ← columna E-fuente.
        // Columna G-fuente (Valor) = contraparte del gate D5(i) BCE = fuente.
        var subsidio = ExcelWorksheetNavigator.CeldaNumero(fila.ElementAtOrDefault(WorkbookLeafCellMapBalanceSc.ColumnaSubsidioFuente));
        var contribucion = ExcelWorksheetNavigator.CeldaNumero(fila.ElementAtOrDefault(WorkbookLeafCellMapBalanceSc.ColumnaContribucionFuente));
        var totalFuente = ExcelWorksheetNavigator.CeldaNumero(fila.ElementAtOrDefault(WorkbookLeafCellMapBalanceSc.ColumnaTotalFuente));

        var bloque = new BalanceScAseInputs
        {
            Ase = ase,
            Contribucion = contribucion,
            Subsidio = subsidio,
            TotalFuente = totalFuente,
            Sistema = null // D2(b): H es fórmula en el template (T0-0.6); se protege, no se escribe.
        };

        // Gate D5-i (coherencia): TotalBsc (E+F) == Total General fuente (col G) antes de
        // devolver (fail-fast nombra ASE). Mismo patrón que ValidarContraFuentesBanco (HU-09).
        WorkbookLeafCoherence.ValidarContraFuentesBalanceSc(bloque);

        return new BalanceScInputs
        {
            Ases = [bloque]
        };
    }

    /// <summary>
    /// HU-10 (2.4): busca la fila "Total General" desde el final del <c>Sheet1</c> por match
    /// NORMALIZADO EXACTO (minúsculas, sin espacios — T0-0.4). Nunca fila fija (V6). Las filas
    /// "Total" de cada localidad NO matchean (match exacto, no por prefijo).
    /// </summary>
    private static int BuscarTotalGeneralDesdeElFinal(List<object?[]> filas)
    {
        for (var i = filas.Count - 1; i >= 0; i--)
        {
            var fila = filas[i];
            if (fila is null)
            {
                continue;
            }

            for (var j = 0; j < fila.Length; j++)
            {
                var texto = ExcelWorksheetNavigator.CeldaTexto(fila[j]);
                if (!string.IsNullOrWhiteSpace(texto)
                    && string.Equals(NormalizarEtiqueta(texto), WorkbookLeafCellMapBalanceSc.EtiquetaTotalGeneral, StringComparison.Ordinal))
                {
                    return i;
                }
            }
        }

        return -1;
    }

    private static RecaudoEmpresaInputs LeerRecaudoEmpresa(EmpresaFacturacion empresa, string rutaConciliacion)
    {
        // T0-0.6: los archivos Conjunta * tienen UNA sola hoja (RESUMEN MES); LeerFilas lee la primera.
        // La estructura de bloques es uniforme (ASE1..5 + X + total) aunque los encabezados varíen
        // ("OPORTUNO"/"EXTEMP."/"TOTAL" en ENEL vs "Ciudad Limpia - Prestador"/"EAAB - Prestador"
        // en Otros). Se detectan los bloques por la corrida de filas con ASE 1..5 (col C).
        var filas = ExcelWorksheetNavigator.LeerFilas(rutaConciliacion);
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        bool EsAse(object?[]? fila, int esperado)
        {
            if (fila is null)
            {
                return false;
            }

            var texto = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(2));
            return decimal.TryParse(texto, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var valor)
                && Math.Abs(valor - esperado) < 0.01m;
        }

        var inicioBloques = new List<int>();
        for (var i = 0; i + 4 < filas.Count; i++)
        {
            if (EsAse(filas[i], 1) && EsAse(filas[i + 1], 2) && EsAse(filas[i + 2], 3)
                && EsAse(filas[i + 3], 4) && EsAse(filas[i + 4], 5))
            {
                inicioBloques.Add(i);
                i += 6; // salta los 5 ASE + la fila X
            }
        }

        if (inicioBloques.Count < 3)
        {
            throw new CalculoInvalidoException(
                $"No se encontraron los 3 bloques (ASE 1..5) en el RESUMEN MES de {empresa.Nombre} ({rutaConciliacion}).");
        }

        // Bloques → Recaudo: OPORTUNO D3:D9, EXTEMP D12:D18, TOTAL D21:D27 (D = valor, E = n° reg).
        var filasInicio = new[] { 3, 12, 21 };
        for (var b = 0; b < 3; b++)
        {
            for (var i = 0; i < 7; i++)
            {
                var fila = filas.ElementAtOrDefault(inicioBloques[b] + i);
                var valor = fila?.ElementAtOrDefault(3);
                var registros = fila?.ElementAtOrDefault(4);
                if (valor is null && registros is null)
                {
                    continue; // fila "X" sin datos
                }

                celdas[$"D{filasInicio[b] + i}"] = ExcelWorksheetNavigator.CeldaNumero(valor);
                celdas[$"E{filasInicio[b] + i}"] = ExcelWorksheetNavigator.CeldaNumero(registros);
            }
        }

        return new RecaudoEmpresaInputs
        {
            Empresa = empresa,
            HojaRecaudo = empresa.HojaRecaudo,
            Celdas = celdas,
            TotalOportuno = celdas.GetValueOrDefault("D9"),
            TotalExtemporaneo = celdas.GetValueOrDefault("D18"),
            Total = celdas.GetValueOrDefault("D27")
        };
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

    /// <summary>
    /// Extrae los operandos R1 por empresa según <see cref="WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa"/>.
    /// Fail-fast: si la fuente no trae la fila Total/Subs de una empresa con operando declarado,
    /// lanza <c>CalculoInvalidoException</c> que nombra ASE y empresa (nunca valor inventado).
    /// </summary>
    private static IReadOnlyDictionary<string, decimal> ExtraerCeldasR1(
        Ase ase, EmpresaFacturacion empresa, List<object?[]> filas)
    {
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        if (!WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.TryGetValue((empresa.Id, ase.Id), out var mapa))
        {
            return celdas; // empresa sin detalle no-estático en este ASE (visible 0)
        }

        var label = ObtenerLabelFuente(empresa.Id, ase.Id);

        // Zonas de la fuente R1: la fila "Componente" (col A) separa el bloque de Componente
        // (filas Total por empresa) del bloque de Subsidio(-)/Contribucion(+).
        var indiceComponente = filas.FindIndex(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(0)).Equals("Componente", StringComparison.OrdinalIgnoreCase));
        if (indiceComponente < 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: no se encontró la fila 'Componente' en la fuente R1 para localizar los bloques por empresa.");
        }

        bool EsFilaEmpresa(object?[] f) =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(2)).Equals(label, StringComparison.OrdinalIgnoreCase)
            && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(3)).Equals("Total", StringComparison.OrdinalIgnoreCase);

        var filasTotal = filas
            .Select((fila, indice) => (fila, indice))
            .Where(t => t.indice < indiceComponente && EsFilaEmpresa(t.fila))
            .Select(t => t.fila)
            .OrderByDescending(f => Math.Abs(ExcelWorksheetNavigator.CeldaNumero(f.ElementAtOrDefault(5))))
            .ToList();

        var filasSubs = filas
            .Select((fila, indice) => (fila, indice))
            .Where(t => t.indice > indiceComponente && EsFilaEmpresa(t.fila))
            .Select(t => t.fila)
            .OrderByDescending(f => Math.Abs(ExcelWorksheetNavigator.CeldaNumero(f.ElementAtOrDefault(5))))
            .ToList();

        decimal Bloque5Total() =>
            filas
                .Where(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(3)).Equals("5", StringComparison.OrdinalIgnoreCase)
                    && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(4)).Equals("Total", StringComparison.OrdinalIgnoreCase))
                .Select(f => ExcelWorksheetNavigator.CeldaNumero(f.ElementAtOrDefault(5)))
                .FirstOrDefault();

        foreach (var (celda, fuente) in mapa)
        {
            decimal valor = fuente switch
            {
                WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMain => ValorDe(filasTotal, 0, ase.Id, empresa.Nombre, "Total"),
                WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMain => EspecialesDe(filasTotal, 0, ase.Id, empresa.Nombre),
                WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMain => ValorDe(filasSubs, 0, ase.Id, empresa.Nombre, "Subsidio/Contribución"),
                WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor => ValorDe(filasTotal, filasTotal.Count - 1, ase.Id, empresa.Nombre, "Total (bloque menor)"),
                WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMenor => EspecialesDe(filasTotal, filasTotal.Count - 1, ase.Id, empresa.Nombre),
                WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMenor => ValorDe(filasSubs, filasSubs.Count - 1, ase.Id, empresa.Nombre, "Subsidio/Contribución (bloque menor)"),
                WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMedio => ValorDe(filasTotal, 1, ase.Id, empresa.Nombre, "Total (bloque medio)"),
                WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMedio => EspecialesDe(filasTotal, 1, ase.Id, empresa.Nombre),
                WorkbookLeafCellMapPorEmpresa.FuenteR1.Block5Total => Bloque5Total(),
                _ => 0m
            };

            celdas[celda] = valor;
        }

        return celdas;
    }

    private static IReadOnlyDictionary<string, decimal> ExtraerCeldasR2(
        Ase ase, EmpresaFacturacion empresa, List<object?[]> filas)
    {
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        if (!WorkbookLeafCellMapPorEmpresa.EditablesR2PorEmpresa.TryGetValue((empresa.Id, ase.Id), out var mapa))
        {
            return celdas;
        }

        // R2 fuente: filas con col B = label y col C = "Total". Total = mayor |E|; SubsCont = menor |E|.
        // La columna de Especiales se localiza por encabezado; si no existe (ASE2/ASE4, análogo a
        // ASE4 sin Especiales), el valor es 0 (ServEspK).
        var indiceEspeciales = ExcelWorksheetNavigator.IndiceColumnaPorEncabezado(filas, "Especiales");
        var filasEmpresa = filas
            .Where(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1)).Equals(ObtenerLabelFuente(empresa.Id, ase.Id), StringComparison.OrdinalIgnoreCase)
                && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(2)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(f => Math.Abs(ExcelWorksheetNavigator.CeldaNumero(f.ElementAtOrDefault(4))))
            .ToList();

        foreach (var (celda, fuente) in mapa)
        {
            decimal valor = fuente switch
            {
                WorkbookLeafCellMapPorEmpresa.FuenteR2.Total => ValorDe(filasEmpresa, 0, ase.Id, empresa.Nombre, "Total", columna: 4),
                WorkbookLeafCellMapPorEmpresa.FuenteR2.SubsCont => ValorDe(filasEmpresa, filasEmpresa.Count - 1, ase.Id, empresa.Nombre, "Subs/Cont", columna: 4),
                WorkbookLeafCellMapPorEmpresa.FuenteR2.Especiales => indiceEspeciales >= 0
                    ? EspecialesDe(filasEmpresa, 0, ase.Id, empresa.Nombre, columna: indiceEspeciales)
                    : 0m,
                _ => 0m
            };

            celdas[celda] = valor;
        }

        return celdas;
    }

    private static IReadOnlyDictionary<string, decimal> ExtraerCeldasR4(
        Ase ase, EmpresaFacturacion empresa, List<object?[]> filas)
    {
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        if (!WorkbookLeafCellMapPorEmpresa.EditablesR4PorEmpresa.TryGetValue((empresa.Id, ase.Id), out var mapa))
        {
            return celdas;
        }

        // R4 fuente: fila con col A = label y col B = "Total"; valor en col D (negativo).
        var filaTotal = filas.FirstOrDefault(f =>
                ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(0)).Equals(ObtenerLabelFuente(empresa.Id, ase.Id), StringComparison.OrdinalIgnoreCase)
                && ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(
                $"ASE {ase.Id}: no se encontró la fila Total de la empresa {empresa.Nombre} en la fuente R4.");

        var total = ExcelWorksheetNavigator.CeldaNumero(filaTotal.ElementAtOrDefault(3));
        foreach (var (celda, fuente) in mapa)
        {
            celdas[celda] = fuente switch
            {
                WorkbookLeafCellMapPorEmpresa.FuenteR4.Total => total,
                WorkbookLeafCellMapPorEmpresa.FuenteR4.P => 0m,
                _ => 0m
            };
        }

        return celdas;
    }

    /// <summary>
    /// Label de la empresa en las fuentes R1/R2/R4. RECIPROCIDAD usa "NUEVO ESQUEMA" en ASE2/ASE4 (T0-0.1/0.4).
    /// </summary>
    private static string ObtenerLabelFuente(int empresaId, int aseId) =>
        empresaId == 1 && aseId is 2 or 4
            ? "NUEVO ESQUEMA"
            : EmpresaFacturacion.Obtener(empresaId).LabelTemplate;

    private static decimal ValorDe(List<object?[]> filas, int indice, int aseId, string empresa, string concepto, int columna = 5)
    {
        var fila = filas.ElementAtOrDefault(indice)
            ?? throw new CalculoInvalidoException(
                $"ASE {aseId}: no se encontró la fila {concepto} de la empresa {empresa} en la fuente.");
        return ExcelWorksheetNavigator.CeldaNumero(fila.ElementAtOrDefault(columna));
    }

    private static decimal EspecialesDe(List<object?[]> filas, int indice, int aseId, string empresa, int columna = 11)
    {
        var fila = filas.ElementAtOrDefault(indice)
            ?? throw new CalculoInvalidoException(
                $"ASE {aseId}: no se encontró la fila Total de la empresa {empresa} en la fuente para sus Especiales.");
        return ExcelWorksheetNavigator.CeldaNumero(fila.ElementAtOrDefault(columna));
    }

    /// <summary>
    /// Visible R1 por empresa según la cadena T0 por ASE (OPORTUNO; los operandos EXTEMP no participan del gate Σ).
    /// </summary>
    private static decimal CalcularVisibleR1(int aseId, int empresaId, IReadOnlyDictionary<string, decimal> celdas)
    {
        if (celdas.Count == 0)
        {
            return 0m;
        }

        decimal totalMain = 0m, especialesMain = 0m, subsMain = 0m, totalMenor = 0m, block5 = 0m;
        foreach (var (celda, fuente) in WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa[(empresaId, aseId)])
        {
            var valor = celdas.GetValueOrDefault(celda);
            switch (fuente)
            {
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMain: totalMain = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.EspecialesMain: especialesMain = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.SubsMain: subsMain = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.TotalMenor: totalMenor = valor; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR1.Block5Total: block5 = valor; break;
            }
        }

        // Cadena T0 por ASE: ASE1/ASE4-RECIP = 3 términos; resto = 5 términos; ASE2-RECIP suma bloque "5".
        return aseId switch
        {
            1 => totalMain + subsMain - especialesMain,
            2 when empresaId == 1 => totalMain + subsMain + block5,
            2 => totalMain + subsMain + totalMenor - especialesMain,
            3 => totalMain + subsMain + totalMenor - especialesMain,
            4 when empresaId == 1 => totalMain + subsMain - especialesMain,
            4 => totalMain + subsMain + totalMenor - especialesMain,
            5 => totalMain + subsMain + totalMenor - especialesMain,
            _ => 0m
        };
    }

    private static decimal CalcularVisibleR2(IReadOnlyDictionary<string, decimal> celdas)
    {
        if (celdas.Count == 0)
        {
            return 0m;
        }

        // Uniforme en los 5 bloques: visible = Total + SubsCont − Especiales.
        decimal total = 0m, subs = 0m, especiales = 0m;
        foreach (var (celda, fuente) in WorkbookLeafCellMapPorEmpresa.EditablesR2PorEmpresa.SelectMany(kv => kv.Value))
        {
            if (!celdas.ContainsKey(celda))
            {
                continue;
            }

            switch (fuente)
            {
                case WorkbookLeafCellMapPorEmpresa.FuenteR2.Total: total = celdas[celda]; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR2.SubsCont: subs = celdas[celda]; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR2.Especiales: especiales = celdas[celda]; break;
            }
        }

        return total + subs - especiales;
    }

    private static decimal CalcularVisibleR4(IReadOnlyDictionary<string, decimal> celdas)
    {
        if (celdas.Count == 0)
        {
            return 0m;
        }

        decimal total = 0m, p = 0m;
        foreach (var (celda, fuente) in WorkbookLeafCellMapPorEmpresa.EditablesR4PorEmpresa.SelectMany(kv => kv.Value))
        {
            if (!celdas.ContainsKey(celda))
            {
                continue;
            }

            switch (fuente)
            {
                case WorkbookLeafCellMapPorEmpresa.FuenteR4.Total: total = celdas[celda]; break;
                case WorkbookLeafCellMapPorEmpresa.FuenteR4.P: p = celdas[celda]; break;
            }
        }

        return total - p;
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