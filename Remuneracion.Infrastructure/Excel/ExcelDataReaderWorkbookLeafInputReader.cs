using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Serilog;

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

        // Plan 54 (H1.1/H1.4, D1/D4): resolución ÚNICA para ambas quincenas. La quincena ya NO
        // parametriza el mapeo de R1/R2/R4 — solo gobierna fuentes y declaración de visibles (D7).
        // El mapa productivo es la geometría v2 (<see cref="WorkbookLeafCellMapQ2"/>).
        var leaf = new WorkbookLeafInputs
        {
            Ase = ase,
            Periodo = periodo,
            R1 = MapearR1(ase, filasR1),
            R2 = MapearR2(ase, filasR2),
            R4 = MapearR4(ase, filasR4)
        };

        // Plan 29 (T2, Unidad R — SOLO LECTURA): el desglose-detalle por componente de R2/R4 se
        // construye desde la MISMA observación ya cargada (no re-lee el archivo; R-DOBLE-FUENTE).
        // Viaja dentro del leaf para que la ESCRITURA lo consuma; aquí NO se escribe.
        //
        // Plan 59 (§11-bis / §0.2.2 — rama por quincena ELIMINADA): hasta aquí el detalle solo se
        // construía en la 2.ª quincena. El oráculo administrativo de 2026101 y de 2026071 (ambos
        // primera quincena) trae los tres niveles del bloque llenos —detalle por canal/concepto,
        // total por canal y cierre de sección— y sin ellos el bloque queda en los literales 0 de la
        // plantilla, E57 (que los suma) da 0 y la línea de anticipos de CONSOLIDADO!D28:D32 se cae.
        // La quincena NO decide si hay detalle: el detalle es del período. Lo único que la quincena
        // parametriza de verdad es el par de columnas del reporte de conciliación por empresa
        // (D/E en la 1.ª, F/G en la 2.ª), y esa selección vive donde corresponde, no aquí.
        leaf.DetalleR2 = ConstruirDetalleR2(ase, rutaR2, filasR2);
        leaf.DetalleR4 = ConstruirDetalleR4(ase, rutaR4, filasR4);

        WorkbookLeafCoherence.ValidarContraFuentes(leaf, r1, r2, r4);
        return leaf;
    }

    /// <inheritdoc />
    public DetalleR2AseInputs LeerDetalleR2(Ase ase, string rutaR2)
    {
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(rutaR2);

        var filas = ExcelWorksheetNavigator.LeerFilas(rutaR2);
        return ConstruirDetalleR2(ase, rutaR2, filas);
    }

    /// <inheritdoc />
    public DetalleR4AseInputs LeerDetalleR4(Ase ase, string rutaR4)
    {
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(rutaR4);

        var filas = ExcelWorksheetNavigator.LeerFilas(rutaR4);
        return ConstruirDetalleR4(ase, rutaR4, filas);
    }

    /// <inheritdoc />
    public string? RutaPlantilla { get; set; }

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

        // Plan 55 (H1.1/H1.3, D4): la resolución por empresa es por CRUCE DE ETIQUETAS contra la
        // plantilla (autoridad de los destinos) y el agregado se deriva de la FUENTE sobre el mismo
        // conjunto de roles del visible de bloque (anti-tautología). La tabla congelada por
        // (empresa, ASE) quedó retirada del camino productivo (LEGACY).
        var secuencia = ConstruirSecuenciaEspejoR1(ase, filasR1, nombreArchivo: null);
        var resultado = ResolvedorConciliacionEmpresa
            .Resolver(ase, secuencia.Filas, filasR2, filasR4, RutaPlantilla)
            .ToList();

        // Plan 50 (A3, §2.5): evidencia de etiquetas observadas fuera del vocabulario de fuentes.
        var etiquetasNoReconocidas = DetectarEtiquetasNoReconocidas(ase, filasR1, filasR2, filasR4);
        foreach (var inputs in resultado)
        {
            inputs.EtiquetasNoReconocidas = etiquetasNoReconocidas;
        }

        return resultado;
    }

    /// <summary>
    /// Plan 50 (A3, §2.5 — control mínimo alternativo): etiquetas de fuente observadas en las filas de
    /// cierre de cada reporte que NO pertenecen al vocabulario de fuentes del negocio
    /// (<see cref="EmpresaFacturacion.Catalogo"/> + su label por ASE + la sinonimia citada del
    /// Plan 34). Es EVIDENCIA (sin juicio): no abre workbooks extra, reusa las filas ya leídas.
    /// Formato <c>{Reporte}:{label}</c>.
    /// </summary>
    private static IReadOnlyList<string> DetectarEtiquetasNoReconocidas(
        Ase ase, List<object?[]> filasR1, List<object?[]> filasR2, List<object?[]> filasR4)
    {
        var noReconocidas = new List<string>();

        // Vocabulario de FUENTES de recaudo del negocio (dato del catálogo, nunca una rama por
        // período/empresa): los rótulos de la plantilla y el rótulo vigente de la fuente por ASE.
        var vocabulario = EmpresaFacturacion.Catalogo
            .SelectMany(empresa => new[] { empresa.LabelTemplate, ObtenerLabelFuente(empresa.Id, ase.Id) })
            .Where(label => label.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        void Analizar(string reporte, IEnumerable<string> observadas)
        {
            foreach (var observada in observadas.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (vocabulario.Any(conocida => SinonimosEmpresaR1.SonMismaEmpresa(observada, conocida)))
                {
                    continue;
                }

                noReconocidas.Add($"{reporte}:{observada}");
            }
        }

        // R1: la etiqueta de la fuente vive en C y la fila de cierre es D='Total'.
        Analizar(
            "R1",
            filasR1.Where(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(3)).Equals("Total", StringComparison.OrdinalIgnoreCase))
                .Select(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(2)))
                .Where(label => label.Length > 0));

        // R2: la etiqueta vive en B y la fila de cierre es C='Total'.
        Analizar(
            "R2",
            filasR2.Where(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(2)).Equals("Total", StringComparison.OrdinalIgnoreCase))
                .Select(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1)))
                .Where(label => label.Length > 0));

        // R4: la etiqueta vive en A y la fila de cierre es B='Total'.
        Analizar(
            "R4",
            filasR4.Where(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1)).Equals("Total", StringComparison.OrdinalIgnoreCase))
                .Select(f => ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(0)))
                .Where(label => label.Length > 0));

        return noReconocidas;
    }

    /// <inheritdoc />
    public IReadOnlyList<RecaudoEmpresaInputs> LeerRecaudosEmpresa(
        Periodo periodo, Func<EmpresaFacturacion, string?> rutaConciliacionPorEmpresa)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(rutaConciliacionPorEmpresa);

        // G2-D2: la quincena la gobierna el DOMINIO (Periodo.NumeroQuincena), jamás la detección
        // de contenido del xlsx. Q1 → par de columnas D/E (VALOR 1°Q / N° REG. 1°Q); Q2 → F/G.
        var esQuincena2 = periodo.NumeroQuincena == 2;

        var resultado = new List<RecaudoEmpresaInputs>();
        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            var ruta = rutaConciliacionPorEmpresa(empresa)
                ?? throw new ArchivoFuenteNoEncontradoException(
                    $"No se encontró el archivo de conciliación de {empresa.Nombre} (prefijo '{empresa.PrefijoConciliacion}') en la carpeta Conciliaciones/ del período.");

            resultado.Add(LeerRecaudoEmpresa(empresa, ruta, esQuincena2));
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

        // Plan 51 (A1, §2.1 D-1): reconocimiento de las columnas fuente POR HEADER, nunca por
        // índice congelado. El vocabulario es el del negocio ya usado por A3 en R1/R2/R4
        // (catálogo + rótulo vigente por ASE + sinonimia citada) y lo ESPERADO por ASE lo declara
        // la plantilla (`EditablesPorAse[ase].Keys`). `Total` se excluye antes del vocabulario;
        // una celda vacía no es evidencia; una etiqueta reconocida sin celdas destino es desync
        // mapa↔plantilla (fail-fast); cero fuentes reconocidas es estructura rota (fail-fast).
        var editables = WorkbookLeafCellMapReporteBanco.ObtenerEditables(ase.Id);
        var esperadas = editables.Keys.ToList();
        var indicesEmpresa = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var etiquetasNoReconocidas = new List<string>();
        for (var indice = 1; indice < filaHeaders.Length; indice++)
        {
            var header = ExcelWorksheetNavigator.CeldaTexto(filaHeaders[indice]);
            if (string.IsNullOrWhiteSpace(header))
            {
                continue;
            }

            if (NormalizarEtiqueta(header) == "total")
            {
                continue; // "Total" no es fuente (exclusión explícita, §2.3.1).
            }

            if (!EsFuenteDelVocabulario(header, ase.Id))
            {
                etiquetasNoReconocidas.Add($"BANCO:{header}");
                continue;
            }

            var empresa = esperadas.FirstOrDefault(esperada => SinonimosEmpresaR1.SonMismaEmpresa(header, esperada));
            if (empresa is null)
            {
                throw new CalculoInvalidoException(
                    $"ASE {ase.Id} · BANCO: la fuente '{header}' viene en el Resumen pero el mapa de la plantilla no declara celdas destino para ese ASE (mapa↔plantilla desincronizados).");
            }

            if (!indicesEmpresa.TryAdd(empresa, indice))
            {
                throw new CalculoInvalidoException(
                    $"ASE {ase.Id} · BANCO: la fuente '{empresa}' aparece más de una vez en la fila de headers del Resumen (columnas {indicesEmpresa[empresa]} y {indice}); la geometría es ambigua.");
            }
        }

        if (indicesEmpresa.Count == 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id} · BANCO: no se reconoció ninguna fuente de recaudo en la fila de headers del Resumen. Estructura rota: el reporte por banco debe traer al menos una fuente del catálogo.");
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

        // Plan 51 (A2, §2.2): la fuente esperada que no vino se OMITE (no se emite fila: el writer
        // itera presentes y sus celdas quedan en el 0 de la plantilla) con 1 Warning por (ASE,
        // fuente) y huella durable `BANCO:{label}`.
        var fuentesAusentes = new List<string>();
        foreach (var esperada in esperadas.Where(e => !indicesEmpresa.ContainsKey(e)))
        {
            RegistrarFuenteAusenteBanco(ase, esperada, editables[esperada].Length);
            fuentesAusentes.Add($"BANCO:{esperada}");
        }

        var bloque = new ReporteBancoAseInputs
        {
            Ase = ase,
            Empresas = empresas,
            FuentesAusentesBanco = fuentesAusentes,
            EtiquetasBancoNoReconocidas = etiquetasNoReconocidas
        };

        // Gate D5-i (coherencia): bloque == resumen fuente antes de devolver (fail-fast ASE+empresa).
        WorkbookLeafCoherence.ValidarContraFuentesBanco(bloque, empresas);

        return new ReporteBancoInputs
        {
            Ases = [bloque],
            Quincena = periodo.NumeroQuincena,
            Consolidado = null // D2(b): filas 1–7 son fórmulas (T0-0.1); se protegen, no se escriben.
        };
    }

    /// <summary>
    /// Plan 51 (A1/D-1): verdad si el header pertenece al vocabulario de FUENTES de recaudo del
    /// negocio — los rótulos del catálogo, el rótulo vigente de la fuente por ASE (Plan 34) y la
    /// sinonimia citada (<c>RECIPROCIDAD ↔ NUEVO ESQUEMA</c>). Es evidencia de reconocimiento, sin
    /// juicio: no crea empresas ni escribe nada.
    /// </summary>
    private static bool EsFuenteDelVocabulario(string header, int aseId)
    {
        foreach (var empresa in EmpresaFacturacion.Catalogo)
        {
            if (SinonimosEmpresaR1.SonMismaEmpresa(header, empresa.LabelTemplate)
                || SinonimosEmpresaR1.SonMismaEmpresa(header, ObtenerLabelFuente(empresa.Id, aseId)))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Plan 51 (A2, §2.2): huella de una fuente de recaudo ausente en el reporte por banco — UN
    /// <c>Log.Warning</c> por (ASE, fuente), con el conteo de celdas destino que quedan en 0.
    /// Lenguaje administrativo (doctrina Plan 26 D-E): sin jerga de código.
    /// </summary>
    private static void RegistrarFuenteAusenteBanco(Ase ase, string label, int celdasEnCero)
    {
        Log.ForContext("AseId", ase.Id)
            .ForContext("Reporte", "BANCO")
            .ForContext("Fuente", label)
            .ForContext("CeldasEnCero", celdasEnCero)
            .Warning(
                "ASE {AseId}: la fuente de recaudo {Fuente} no viene en el reporte por banco del período; sus {CeldasEnCero} celdas quedan en 0 (ausencia registrada, no es error).",
                ase.Id,
                label,
                celdasEnCero);
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

        // Lectura fuente-fiel (header de la fuente, T0a): columnas E=Subsidio (negativo) y
        // F=Contribución (positivo) de la fila "TOTAL GENERAL". El DESTINO en la plantilla se
        // corrigió en Plan 29 T1 / T0-V4 (D=SUBSIDIO, E=CONTRIBUCION); este reader NO cambia
        // (el veredicto viejo del Plan 10 "template-D=CONTRIBUCION" quedó refutado por el header).
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

    /// <inheritdoc />
    public SaldosNotasAseInputs LeerSaldosNotas(Ase ase, string rutaSaldosNotas)
    {
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(rutaSaldosNotas);

        // HU-11 (2.5, D3): lectura header-driven por TÍTULO de columna (D3/G4) con la fila de
        // headers localizada dinámicamente (T0-0.5: fila 3 en las fuentes Q2; nunca fila fija).
        var filas = ExcelWorksheetNavigator.LeerFilas(rutaSaldosNotas);

        // Fuente vacía legítima (solo la fila 1 con el rango de fechas; ASE5-Q2, T0-0.5):
        // total 0 demostrable, sin inventar headers ni valores (Riesgo 6 del plan).
        var mapeoObligatorio = WorkbookLeafCellMapAjustesSfT.MapeoColumnasPorTitulo;
        var indiceHeaders = BuscarFilaHeaders(filas, mapeoObligatorio);
        if (indiceHeaders < 0)
        {
            return new SaldosNotasAseInputs
            {
                Ase = ase,
                TieneColumnaEspeciales = false,
                Total = 0m,
                ServEspK = 0m,
                Celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
            };
        }

        var filaHeaders = filas[indiceHeaders];
        var indicesColumna = mapeoObligatorio
            .Select(par => (par, indice: BuscarIndiceColumnaPorTitulo(filaHeaders, par.HeaderFuente)))
            .ToArray();

        // Fail-fast (Requirement 2): si falta un header obligatorio del mapa T0-0.5, fallo que
        // nombra el ASE y el reporte; NUNCA valor inventado. "Deb/Cred" es OPCIONAL (solo ASE2).
        var faltantes = indicesColumna
            .Where(t => t.indice < 0)
            .Select(t => t.par.HeaderFuente)
            .ToArray();
        if (faltantes.Length > 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: la fuente SALDOS POR NOTA no trae los headers esperados del mapa T0-0.5: {string.Join(", ", faltantes)}.");
        }

        var indiceDebCred = BuscarIndiceColumnaPorTitulo(filaHeaders, WorkbookLeafCellMapAjustesSfT.HeaderDebCredOpcional);

        // Conceptos por fila del template (mapa explícito por Ase.Id, D7). HU-22 (Plan 23, R-F-1/T2,
        // T0d §5.1): el flag EsOpcional del mapa distingue el concepto opcional ("Vlr Intereses")
        // del core; la ausencia del opcional se emite como 0 explícito (nunca fail-fast ni valor
        // fabricado). La decisión obligatorio/opcional vive SOLO en el mapa, no en el reader.
        var conceptos = WorkbookLeafCellMapAjustesSfT.ObtenerConceptosSaldos(ase.Id);
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        decimal total = 0m;

        foreach (var (filaTemplate, concepto, esOpcional) in conceptos)
        {
            var filaFuente = BuscarFilaConcepto(filas, concepto);
            if (filaFuente < 0)
            {
                // Opcional ausente (T0d): fila completa en 0 explícito. Se conserva el 0 de
                // plantilla en O/Deb/Cred (no se escribe) y la columna I se fija a 0 más abajo.
                if (esOpcional)
                {
                    foreach (var (par, _) in indicesColumna)
                    {
                        celdas[$"{par.ColumnaTemplate}{filaTemplate}"] = 0m;
                    }

                    continue;
                }

                throw new CalculoInvalidoException(
                    $"ASE {ase.Id}: no se encontró la fila de concepto '{concepto}' en la fuente SALDOS POR NOTA.");
            }

            foreach (var (par, indice) in indicesColumna)
            {
                var valor = ExcelWorksheetNavigator.CeldaNumero(filas[filaFuente].ElementAtOrDefault(indice));
                celdas[$"{par.ColumnaTemplate}{filaTemplate}"] = valor;

                if (string.Equals(par.HeaderFuente, "Total", StringComparison.OrdinalIgnoreCase))
                {
                    total = valor;
                }
            }

            // Columna O (Deb/Cred) del template: opcional — si la fuente la trae se escribe el
            // valor; si no, queda el 0 ya presente en la plantilla (el writer no la escribe).
            if (indiceDebCred >= 0)
            {
                celdas[$"{WorkbookLeafCellMapAjustesSfT.ColumnaTemplateDebCred}{filaTemplate}"] =
                    ExcelWorksheetNavigator.CeldaNumero(filas[filaFuente].ElementAtOrDefault(indiceDebCred));
            }
        }

        // La columna I del template ("Especiales") no existe en las fuentes Q2 (T0-0.5) → 0.
        // TotalSaldosNotas = Total − ServEspK (aritmética T0-0.3, visible Cn-In).
        foreach (var (filaTemplate, _, _) in conceptos)
        {
            celdas[$"I{filaTemplate}"] = 0m;
        }

        return new SaldosNotasAseInputs
        {
            Ase = ase,
            TieneColumnaEspeciales = false,
            Total = total,
            ServEspK = 0m,
            Celdas = celdas
        };
    }

    /// <inheritdoc />
    public RetribucionNegativaAseInputs LeerRetribucionNegativa(Ase ase, string rutaRetribucionNegativa)
    {
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(rutaRetribucionNegativa);

        // HU-11 (2.5, D3): misma lectura header-driven por TÍTULO que SALDOS POR NOTA. T0-0.4/0.5:
        // en Q2 las 5 fuentes traen SOLO la fila 1 (rango de fechas) → 0 legítimo (el golden
        // D28:D32 = 0 lo confirma). El reader distingue "fuente vacía = 0" de "header ausente
        // dentro de una fuente con datos = fallo que nombra ASE + reporte" (Riesgo 6).
        var filas = ExcelWorksheetNavigator.LeerFilas(rutaRetribucionNegativa);
        var mapeoObligatorio = WorkbookLeafCellMapAjustesSfT.MapeoColumnasPorTitulo;
        var indiceHeaders = BuscarFilaHeaders(filas, mapeoObligatorio);
        if (indiceHeaders < 0)
        {
            return new RetribucionNegativaAseInputs
            {
                Ase = ase,
                TieneColumnaEspeciales = false,
                Total = 0m,
                ServEspK = 0m,
                Celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
            };
        }

        var filaHeaders = filas[indiceHeaders];
        var indicesColumna = mapeoObligatorio
            .Select(par => (par, indice: BuscarIndiceColumnaPorTitulo(filaHeaders, par.HeaderFuente)))
            .ToArray();

        var faltantes = indicesColumna
            .Where(t => t.indice < 0)
            .Select(t => t.par.HeaderFuente)
            .ToArray();
        if (faltantes.Length > 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: la fuente RETRIBUCION NEGATIVA no trae los headers esperados del mapa T0-0.5: {string.Join(", ", faltantes)}.");
        }

        var indiceDebCred = BuscarIndiceColumnaPorTitulo(filaHeaders, WorkbookLeafCellMapAjustesSfT.HeaderDebCredOpcional);

        // HU-22 (Plan 23, R-F-2/T2, T0d §5.3): MISMA regla de opcionalidad que SALDOS POR NOTA.
        // El flag EsOpcional del mapa distingue el concepto opcional ("Vlr Intereses") del core;
        // su ausencia se emite como 0 explícito (nunca fail-fast). El defecto es latente (las
        // fuentes Q2/agosto traen solo la fila de fecha → early-return 0), pero la rama queda
        // simétrica y ejercitable si una fuente futura trae headers sin el concepto opcional.
        var conceptos = WorkbookLeafCellMapAjustesSfT.ObtenerConceptosRetribucion(ase.Id);
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        decimal total = 0m;

        foreach (var (filaTemplate, concepto, esOpcional) in conceptos)
        {
            var filaFuente = BuscarFilaConcepto(filas, concepto);
            if (filaFuente < 0)
            {
                // Opcional ausente: fila completa en 0 explícito; O/Deb/Cred conserva el 0 de
                // plantilla (no se escribe) e I{fila} se fija a 0 más abajo.
                if (esOpcional)
                {
                    foreach (var (par, _) in indicesColumna)
                    {
                        celdas[$"{par.ColumnaTemplate}{filaTemplate}"] = 0m;
                    }

                    continue;
                }

                throw new CalculoInvalidoException(
                    $"ASE {ase.Id}: no se encontró la fila de concepto '{concepto}' en la fuente RETRIBUCION NEGATIVA.");
            }

            foreach (var (par, indice) in indicesColumna)
            {
                var valor = ExcelWorksheetNavigator.CeldaNumero(filas[filaFuente].ElementAtOrDefault(indice));
                celdas[$"{par.ColumnaTemplate}{filaTemplate}"] = valor;

                if (string.Equals(par.HeaderFuente, "Total", StringComparison.OrdinalIgnoreCase))
                {
                    total = valor;
                }
            }

            if (indiceDebCred >= 0)
            {
                celdas[$"{WorkbookLeafCellMapAjustesSfT.ColumnaTemplateDebCred}{filaTemplate}"] =
                    ExcelWorksheetNavigator.CeldaNumero(filas[filaFuente].ElementAtOrDefault(indiceDebCred));
            }
        }

        foreach (var (filaTemplate, _, _) in conceptos)
        {
            celdas[$"I{filaTemplate}"] = 0m;
        }

        return new RetribucionNegativaAseInputs
        {
            Ase = ase,
            TieneColumnaEspeciales = false,
            Total = total,
            ServEspK = 0m,
            Celdas = celdas
        };
    }

    /// <inheritdoc />
    public BloqueEspejoAseInputs LeerEspejoR1(Ase ase, string rutaR1)
    {
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(rutaR1);

        // Plan 21 (T2, R-E-1): la fuente R1 trae UN bloque por archivo. Se lee la hoja Sheet1 y se
        // delimita la zona por ETIQUETAS (nunca por fila/columna fija). Sin cardinalidades ni
        // ocurrencias: la secuencia observada ES la especificación del período (D-B).
        // Plan 53 (H1.2, D3.4): la construcción de la secuencia es COMPARTIDA con la conciliación por
        // empresa — un solo scan de fronteras, prohibido reimplementarlo. Plan 55 (H1.1): el consumidor
        // de la conciliación es `ResolvedorConciliacionEmpresa` (recibe esta MISMA secuencia), que
        // reemplazó a `ExtraerCeldasR1` (retirado con la tabla congelada por (empresa, ASE) a LEGACY).
        const string hoja = "Sheet1";
        var filas = ExcelWorksheetNavigator.LeerFilas(rutaR1, hoja);
        var secuencia = ConstruirSecuenciaEspejoR1(ase, filas, Path.GetFileName(rutaR1));

        var bloque = new BloqueEspejoAseInputs
        {
            Ase = ase,
            Encabezados = secuencia.Encabezados,
            Filas = secuencia.Filas,
            TieneColumnaEspeciales = secuencia.TieneEspeciales
        };

        // Invariantes duras T0e: si falta una, fail-fast que nombra ASE + reporte + fila esperada.
        // W-7 (auditoría PR3): el chequeo vive en el modelo (FaltantesInvariantesDeCierre) para no
        // duplicar la definición con ProcesadorPeriodo.ValidarInvariantesEspejoR1.
        var faltantes = bloque.FaltantesInvariantesDeCierre();
        if (faltantes.Count > 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: la fuente R1 ({Path.GetFileName(rutaR1)}) no trae las filas invariantes de cierre T0e: {string.Join(", ", faltantes)}.");
        }

        return bloque;
    }

    /// <summary>
    /// Plan 53 (H1.2, D3.4): secuencia espejo de la hoja <c>Sheet1</c> de la fuente R1 — la MISMA
    /// ventana que consume <see cref="LeerEspejoR1"/> (primera fila <c>E='Vlr Servicio'</c> hasta la fila
    /// invariante <c>A='Total'</c> con <c>B</c> vacío), con el <b>índice crudo</b> de cada fila (para
    /// resolver sus valores por las columnas F/L del contrato) y los encabezados detectados.
    /// Punto ÚNICO de construcción: prohibido reimplementar el scan de fronteras (Plan 53 §4.5).
    /// </summary>
    private static SecuenciaEspejoR1 ConstruirSecuenciaEspejoR1(Ase ase, List<object?[]> filas, string? nombreArchivo)
    {
        var archivo = nombreArchivo is null ? "la fuente R1" : $"la fuente R1 ({nombreArchivo})";

        // Encabezados por detección dinámica: la fila de headers es la inmediatamente anterior a
        // la primera fila de datos (E='Vlr Servicio'). Así la ausencia de la columna
        // SERVICIO ESPECIALES no impide localizar el resto (R-E-6).
        var indicePrimeraFila = filas.FindIndex(f =>
            NormalizarEtiqueta(ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(4))).Equals("vlrservicio", StringComparison.Ordinal));
        if (indicePrimeraFila <= 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: {archivo} no trae la fila 'Vlr Servicio' que delimita el inicio del bloque espejo.");
        }

        var indiceHeader = indicePrimeraFila - 1;
        while (indiceHeader > 0 && FilaSinTexto(filas[indiceHeader]))
        {
            indiceHeader--;
        }

        var (encabezados, indicesPorEncabezado) = MapearEncabezados(filas[indiceHeader]);
        var tieneEspeciales = indicesPorEncabezado.Keys.Any(EsEncabezadoEspeciales);
        var indiceEspeciales = indicesPorEncabezado
            .Where(kv => EsEncabezadoEspeciales(kv.Key))
            .Select(kv => kv.Value)
            .FirstOrDefault(-1);

        // Zona: desde la primera fila de datos hasta la fila 'Total' final (A='Total', B vacío),
        // ambas inclusive. Si no aparece, la secuencia no es interpretable → fail-fast.
        var indiceTotal = -1;
        for (var i = indicePrimeraFila; i < filas.Count; i++)
        {
            var fila = filas[i];
            if (string.Equals(ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(0)), "Total", StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrWhiteSpace(ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(1))))
            {
                indiceTotal = i;
                break;
            }
        }

        if (indiceTotal < 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: {archivo} no trae la fila invariante 'Total' (A='Total', B vacío) que cierra el bloque espejo.");
        }

        var filasEspejo = new List<FilaEspejoR1>();
        var indicesCrudos = new List<int>();
        for (var i = indicePrimeraFila; i <= indiceTotal; i++)
        {
            var fila = filas[i];
            if (FilaSinTexto(fila) && FilaSinValores(fila, indicesPorEncabezado))
            {
                continue; // fila fantasma vacía: no forma parte de la secuencia observada
            }

            filasEspejo.Add(ConstruirFilaEspejo(fila, encabezados, indicesPorEncabezado, indiceEspeciales));
            indicesCrudos.Add(i);
        }

        return new SecuenciaEspejoR1(filasEspejo, indicesCrudos, encabezados, indicesPorEncabezado, tieneEspeciales, indiceEspeciales);
    }

    /// <summary>
    /// Plan 53 (H1.2): secuencia espejo construida una sola vez por hoja + el índice CRUDO de cada fila
    /// (la ventana puede saltar filas fantasma, así que el índice del espejo NO coincide con el crudo).
    /// </summary>
    private sealed record SecuenciaEspejoR1(
        List<FilaEspejoR1> Filas,
        List<int> IndicesCrudos,
        List<string> Encabezados,
        Dictionary<string, int> IndicesPorEncabezado,
        bool TieneEspeciales,
        int IndiceEspeciales);

    /// <summary>
    /// Plan 21 (T2): encabezados de columna de la fila de headers, en orden físico, con su índice
    /// 0-based. Los encabezados vacíos se omiten; los duplicados se conservan una sola vez.
    /// </summary>
    private static (List<string> Encabezados, Dictionary<string, int> Indices) MapearEncabezados(object?[] filaHeader)
    {
        var encabezados = new List<string>();
        var indices = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var j = 0; j < filaHeader.Length; j++)
        {
            var texto = ExcelWorksheetNavigator.CeldaTexto(filaHeader[j]);
            if (string.IsNullOrWhiteSpace(texto))
            {
                continue;
            }

            var encabezado = texto.Trim();
            if (indices.ContainsKey(encabezado))
            {
                continue;
            }

            indices[encabezado] = j;
            encabezados.Add(encabezado);
        }

        return (encabezados, indices);
    }

    /// <summary>
    /// Plan 21 (T2, R-E-6): verdad si el encabezado es la columna de especiales
    /// (<c>SERVICIO ESPECIALES</c> o su alias <c>Especiales</c>, normalizados).
    /// </summary>
    private static bool EsEncabezadoEspeciales(string encabezado)
    {
        var normalizado = NormalizarEtiqueta(encabezado);
        return normalizado is "servicioespeciales" or "especiales";
    }

    /// <summary>
    /// Plan 21 (T2): construye una <see cref="FilaEspejoR1"/> con su firma A–E y sus valores por
    /// encabezado. Si la columna de especiales no existe, aporta 0 (R-E-6).
    /// </summary>
    private static FilaEspejoR1 ConstruirFilaEspejo(
        object?[] fila,
        IReadOnlyList<string> encabezados,
        IReadOnlyDictionary<string, int> indicesPorEncabezado,
        int indiceEspeciales)
    {
        var valores = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);
        foreach (var encabezado in encabezados)
        {
            valores[encabezado] = CeldaNumeroONulo(fila.ElementAtOrDefault(indicesPorEncabezado[encabezado]));
        }

        if (indiceEspeciales < 0)
        {
            valores["SERVICIO ESPECIALES"] = 0m;
        }

        return new FilaEspejoR1
        {
            A = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(0)),
            B = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(1)),
            C = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(2)),
            D = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(3)),
            E = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(4)),
            ValoresPorColumna = valores
        };
    }

    /// <summary>Plan 21 (T2): verdad si la fila no tiene texto en ninguna columna.</summary>
    private static bool FilaSinTexto(object?[]? fila)
    {
        if (fila is null)
        {
            return true;
        }

        foreach (var celda in fila)
        {
            if (!string.IsNullOrWhiteSpace(ExcelWorksheetNavigator.CeldaTexto(celda)))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Plan 21 (T2): verdad si la fila no tiene valor numérico en ninguna de las columnas de
    /// encabezado detectadas.
    /// </summary>
    private static bool FilaSinValores(object?[] fila, IReadOnlyDictionary<string, int> indicesPorEncabezado)
    {
        foreach (var indice in indicesPorEncabezado.Values)
        {
            if (CeldaNumeroONulo(fila.ElementAtOrDefault(indice)) is not null)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Plan 21 (T2): convierte una celda a número nullable; <c>null</c> = celda vacía (distingue
    /// "ausente" de "cero", R-E-5). Texto no numérico → <c>null</c>.
    /// </summary>
    private static decimal? CeldaNumeroONulo(object? valor)
    {
        if (valor is null)
        {
            return null;
        }

        if (valor is string texto)
        {
            texto = texto.Trim();
            if (texto.Length == 0)
            {
                return null;
            }

            return decimal.TryParse(texto, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var numero)
                ? numero
                : null;
        }

        return ExcelWorksheetNavigator.CeldaNumero(valor);
    }

    /// <summary>
    /// HU-11 (2.5): localiza la fila de headers de la fuente (la que contiene la mayoría de los
    /// títulos del mapa T0-0.5). Devuelve -1 si la fuente no trae la estructura esperada
    /// (fuente vacía legítima = solo la fila 1 con el rango de fechas).
    /// </summary>
    private static int BuscarFilaHeaders(
        List<object?[]> filas,
        (string HeaderFuente, string ColumnaTemplate)[] mapeo)
    {
        for (var i = 0; i < filas.Count; i++)
        {
            var fila = filas[i];
            if (fila is null)
            {
                continue;
            }

            var aciertos = mapeo.Count(par =>
                BuscarIndiceColumnaPorTitulo(fila, par.HeaderFuente) >= 0);
            if (aciertos >= mapeo.Length - 1)
            {
                return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// HU-11 (2.5): índice (0-based) de la columna cuyo título (normalizado) coincide con el
    /// header esperado. -1 si no existe.
    /// </summary>
    private static int BuscarIndiceColumnaPorTitulo(object?[] fila, string headerEsperado)
    {
        var esperado = NormalizarEtiqueta(headerEsperado);
        for (var j = 0; j < fila.Length; j++)
        {
            var texto = ExcelWorksheetNavigator.CeldaTexto(fila[j]);
            if (!string.IsNullOrWhiteSpace(texto)
                && string.Equals(NormalizarEtiqueta(texto), esperado, StringComparison.Ordinal))
            {
                return j;
            }
        }

        return -1;
    }

    /// <summary>
    /// HU-11 (2.5): índice (0-based) de la fila cuyo label (col A o col B, normalizado) coincide
    /// con el concepto esperado ("Vlr Servicio", "Componente", "Subsidio(-)/Contribucion(+)",
    /// "Subs/Cont", "Total", "Vlr Intereses"). El concepto "Total" SOLO matchea la columna A:
    /// en las fuentes, la fila Total del bloque tiene A="Total" (B vacío), mientras que las filas
    /// "Componente"/"Subs/Cont" tienen B="Total" (header de la segunda columna). -1 si no existe.
    /// </summary>
    private static int BuscarFilaConcepto(List<object?[]> filas, string concepto)
    {
        var esperado = NormalizarEtiqueta(concepto);
        var esTotal = string.Equals(esperado, "total", StringComparison.Ordinal);
        for (var i = 0; i < filas.Count; i++)
        {
            var fila = filas[i];
            if (fila is null)
            {
                continue;
            }

            for (var j = 0; j < fila.Length && j < 2; j++)
            {
                if (esTotal && j == 1)
                {
                    continue; // "Total" solo en col A (fila Total del bloque).
                }

                var texto = ExcelWorksheetNavigator.CeldaTexto(fila[j]);
                if (!string.IsNullOrWhiteSpace(texto)
                    && string.Equals(NormalizarEtiqueta(texto), esperado, StringComparison.Ordinal))
                {
                    return i;
                }
            }
        }

        return -1;
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

    /// <summary>
    /// Nombre exacto de la hoja de resumen de las conciliaciones por empresa
    /// (<c>Conjunta *</c>/<c>Directa*</c>). Plan 27 (T1/D-A): la hoja se resuelve POR NOMBRE, no
    /// por posición; julio trae una sola hoja con este nombre y agosto multi-hoja (2ª/3ª posición).
    /// El match del navigator es case-insensitive.
    /// </summary>
    private const string NombreHojaResumenMes = "RESUMEN MES";

    /// <summary>
    /// Lee una hoja <c>Recaudo *</c> desde el <c>RESUMEN MES</c> de la conciliación por empresa.
    ///
    /// Plan 27 (T1, D-A): la hoja se resuelve POR NOMBRE (<see cref="NombreHojaResumenMes"/>), no
    /// "la primera a ciegas". T0-0.6/HU-20-T0a asumía UNA sola hoja; agosto trae multi-hoja
    /// (<c>Oportuno | EXTEMPORANEO | RESUMEN MES</c>) y leer la primera devolvía el detalle crudo
    /// (sin los 3 bloques). La sobrecarga por nombre falla si la hoja falta, nombrando archivo+hoja.
    /// La estructura de bloques es uniforme (ASE1..5 + X + total) aunque los encabezados varíen
    /// ("OPORTUNO"/"EXTEMP."/"TOTAL" en ENEL vs "Ciudad Limpia - Prestador"/"EAAB - Prestador" en
    /// Otros). Se detectan los bloques por la corrida de filas con ASE 1..5 (col C).
    ///
    /// G2-D2 (quincena = dominio): el par de columnas leído se parametriza por
    /// <paramref name="esQuincena2"/>: Q1 → (3,4) = D/E (VALOR 1°Q / N° REG. 1°Q), Q2 → (5,6) =
    /// F/G (VALOR 2°Q / N° REG. 2°Q). Las claves de <c>Celdas</c> reflejan la columna destino REAL
    /// (D/E en Q1; F/G en Q2) para que <c>EscribirCeldasEmpresa</c> escriba la celda correcta.
    /// </summary>
    private static RecaudoEmpresaInputs LeerRecaudoEmpresa(EmpresaFacturacion empresa, string rutaConciliacion, bool esQuincena2)
    {
        var filas = ExcelWorksheetNavigator.LeerFilas(rutaConciliacion, NombreHojaResumenMes);
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        // Q1 → D/E (índices 3/4); Q2 → F/G (índices 5/6).
        var indiceValor = esQuincena2 ? 5 : 3;
        var indiceRegistros = esQuincena2 ? 6 : 4;
        var letraValor = esQuincena2 ? "F" : "D";
        var letraRegistros = esQuincena2 ? "G" : "E";

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

        // Bloques → Recaudo: OPORTUNO 3:9, EXTEMP 12:18, TOTAL 21:27 (fila base por bloque).
        var filasInicio = new[] { 3, 12, 21 };
        for (var b = 0; b < 3; b++)
        {
            for (var i = 0; i < 7; i++)
            {
                var fila = filas.ElementAtOrDefault(inicioBloques[b] + i);
                var valor = fila?.ElementAtOrDefault(indiceValor);
                var registros = fila?.ElementAtOrDefault(indiceRegistros);
                if (valor is null && registros is null)
                {
                    continue; // fila "X" sin datos
                }

                celdas[$"{letraValor}{filasInicio[b] + i}"] = ExcelWorksheetNavigator.CeldaNumero(valor);
                celdas[$"{letraRegistros}{filasInicio[b] + i}"] = ExcelWorksheetNavigator.CeldaNumero(registros);
            }
        }

        return new RecaudoEmpresaInputs
        {
            Empresa = empresa,
            HojaRecaudo = empresa.HojaRecaudo,
            Celdas = celdas,
            TotalOportuno = celdas.GetValueOrDefault($"{letraValor}9"),
            TotalExtemporaneo = celdas.GetValueOrDefault($"{letraValor}18"),
            Total = celdas.GetValueOrDefault($"{letraValor}27")
        };
    }

    /// <summary>
    /// Plan 54 (H1.1, D1 — U-A): resolución R1 ÚNICA para ambas quincenas (absorbe a la antigua rama
    /// Q1 por <c>switch (ase.Id)</c> con ordinales congelados, eliminada). Resuelve cada celda destino
    /// por FIRMA de etiquetas (A–E) sobre la secuencia construida desde la MISMA lectura de filas
    /// (reusa, no duplica el I/O; R-DOBLE-FUENTE). La firma vive en Core (<see cref="FilaEspejoR1"/>);
    /// aquí solo se construye la secuencia y se resuelve por orden de aparición.
    ///
    /// Regla de ausencia por rol (T0e §6, enmendada 2026-10-22 — enmienda 2 del Plan 54):
    /// <c>Mes0/Mes1</c> y <c>Lmes0/Lmes1</c> son core OBLIGATORIOS (fail-fast que nombra ASE +
    /// reporte + celda, mensaje intacto); <c>Mes2/Lmes2</c> y <c>Aplic0/Aplic1/LAplic0</c> son
    /// OPCIONALES (ausente = 0 explícito: el slot que la fuente no llena queda en el cero del
    /// template). El agregado EXTEMP (D-C) = ΣF(todas las filas <c>Aplicacion</c>) − Especiales de la
    /// primera (0 si no hay); el TOT_OPT (D-C) = ΣF(todas las filas <c>Mes</c>) − Especiales de todas
    /// menos la última. Cero cambios de aritmética respecto de Plan 25.
    /// </summary>
    private static WorkbookLeafInputsR1 MapearR1(Ase ase, List<object?[]> filas)
    {
        var indiceEspeciales = ExcelWorksheetNavigator.IndiceColumnaPorEncabezado(filas, "SERVICIO ESPECIALES");
        var indiceTotal = ExcelWorksheetNavigator.IndiceColumnaPorEncabezado(filas, "Total");
        if (indiceTotal < 0)
        {
            indiceTotal = 5; // columna F (V0.3): el header 'Total' coincide en las fuentes reales
        }

        // Secuencia tipada construida desde la MISMA lectura: la firma (Core) resuelve los roles.
        var secuencia = new BloqueEspejoAseInputs
        {
            Ase = ase,
            Filas = filas.Select(f => ConstruirFilaEspejoRolQ2(f, indiceTotal, indiceEspeciales)).ToList()
        };

        var filasMes = secuencia.FilasMes();
        var filasAplicacion = secuencia.FilasAplicacion();

        decimal F(FilaEspejoR1? fila) => fila?.Valor("Total") ?? 0m;
        decimal Esp(FilaEspejoR1? fila) => fila?.Valor("SERVICIO ESPECIALES") ?? 0m;

        // Plan 59 (H2.1): el lector NO conoce celdas del template. Entrega los valores por ROL de
        // destino (concepto) y el escritor los resuelve contra la geometría que la plantilla declare.
        // El número de roles lo decide la FUENTE (subsecciones variables, §2.6 del T0); los slots que
        // la fuente no trae quedan en el cero del template (doctrina Plan 50).
        if (filasMes.Count < 2)
        {
            // Plan 54 (enmienda 2) conserva el contrato de ROLES CORE: el total oportuno del bloque
            // exige las dos primeras subsecciones 'Mes' (Mes0/Mes1). Sin ellas, los valores se
            // desplazarían en silencio al rol anterior — y eso es un reacomodo, no una lectura.
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ASE {ase.Id}: la fuente R1 (R1-Q2, Recaudoporcomponente) no trae la fila del rol obligatorio 'Mes' #2 (segunda subsección de cierre 'Mes/Total'): el total oportuno del bloque se compone de las dos primeras y sin ella los valores se desplazarían de rol. Fail-fast: nunca se reacomoda la fuente.");
        }

        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        void Rol(
            ResolvedorDestinoValor.NivelDestino nivel,
            string etiqueta,
            int ocurrencia,
            string encabezado,
            decimal valor) =>
            celdas[new ResolvedorDestinoValor.RolValor(nivel, etiqueta, ocurrencia, encabezado).Clave] = valor;

        RegistrarFamilia(ResolvedorDestinoValor.SubseccionOportuna, filasMes, ResolvedorDestinoValor.EncabezadoEspecialesR1);
        RegistrarFamilia(ResolvedorDestinoValor.SubseccionExtemporanea, filasAplicacion, ResolvedorDestinoValor.EncabezadoEspecialesR1);

        void RegistrarFamilia(string etiqueta, IReadOnlyList<FilaEspejoR1> filas, string encabezadoEspeciales)
        {
            for (var i = 0; i < filas.Count; i++)
            {
                Rol(
                    ResolvedorDestinoValor.NivelDestino.Subseccion, etiqueta, i + 1,
                    ResolvedorDestinoValor.EncabezadoTotal, F(filas[i]));

                // Los especiales son operando de la fórmula del total para TODAS las ocurrencias menos
                // la última (regla del dominio, Plan 25 D-C): el rol se emite con la misma regla, así
                // que el número de términos lo decide la fuente y no una lista congelada.
                if (i < filas.Count - 1)
                {
                    Rol(
                        ResolvedorDestinoValor.NivelDestino.Subseccion, etiqueta, i + 1,
                        encabezadoEspeciales, Esp(filas[i]));
                }
            }
        }

        // TOT_OPT: patrón visible del template resuelto por firma/observación —
        // ΣF(todas las filas Mes) − Especiales de todas MENOS la última. Es idéntico al mapa
        // congelado en ASE1-4 (3 filas) y ASE5-julio (2 filas, Lmes1=0); generaliza ASE5-agosto
        // (3 filas Mes — T0a) al mismo 5-term visible, sin depender del conteo congelado.
        var totOpt = filasMes.Sum(f => f.Valor("Total") ?? 0m)
            - (filasMes.Count > 1
                ? filasMes.Take(filasMes.Count - 1).Sum(f => f.Valor("SERVICIO ESPECIALES") ?? 0m)
                : 0m);

        // D-C (R-F-2): EXTEMP = ΣF(todas las filas Aplicacion) − Especiales de la primera; 0 si no hay.
        var extemp = filasAplicacion.Sum(f => f.Valor("Total") ?? 0m)
            - (filasAplicacion.Count > 0 ? Esp(filasAplicacion[0]) : 0m);

        // F25/F41/L25 conservan la semántica Q1 (primera/segunda fila Mes/Total) para el gate
        // R1-F25 vs Extemporáneo HU-02 (Q1 intacto por construcción, G5).
        var filaMes0 = filasMes.ElementAtOrDefault(0);
        var filaMes1 = filasMes.ElementAtOrDefault(1);

        return new WorkbookLeafInputsR1
        {
            F25 = F(filaMes0),
            F41 = F(filaMes1),
            L25 = Esp(filaMes0),
            F30 = 0m,
            F10 = 0m,
            L10 = 0m,
            ValoresPorRol = celdas,
            TotalOportunoEsperadoPorAse = totOpt,
            ExtemporaneoEsperadoPorAse = extemp
        };
    }

    /// <summary>
    /// Plan 25 (T2): construye una <see cref="FilaEspejoR1"/> de rol desde una fila cruda, con los
    /// valores por ENCABEZADO (<c>Total</c> = columna F; <c>SERVICIO ESPECIALES</c> = L o 0 si
    /// ausente). Reusa la semántica de <c>ConstruirFilaEspejo</c> sin exigir la zona del espejo
    /// (la secuencia de roles es un subconjunto de la lectura cruda ya cargada).
    /// </summary>
    private static FilaEspejoR1 ConstruirFilaEspejoRolQ2(object?[] fila, int indiceTotal, int indiceEspeciales)
    {
        var valores = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase)
        {
            ["Total"] = CeldaNumeroONulo(fila.ElementAtOrDefault(indiceTotal)),
            ["SERVICIO ESPECIALES"] = indiceEspeciales >= 0
                ? CeldaNumeroONulo(fila.ElementAtOrDefault(indiceEspeciales))
                : 0m
        };

        return new FilaEspejoR1
        {
            A = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(0)),
            B = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(1)),
            C = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(2)),
            D = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(3)),
            E = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(4)),
            ValoresPorColumna = valores
        };
    }
    /// <summary>
    /// Label de la empresa en las fuentes R1/R2/R4. RECIPROCIDAD usa "NUEVO ESQUEMA" en ASE2/ASE4 (T0-0.1/0.4).
    /// </summary>
    private static string ObtenerLabelFuente(int empresaId, int aseId) =>
        empresaId == 1 && aseId is 2 or 4
            ? "NUEVO ESQUEMA"
            : EmpresaFacturacion.Obtener(empresaId).LabelTemplate;

    /// <summary>
    /// Plan 54 (H1.4, D4): mapeo R2 ÚNICO para ambas quincenas — labels (Componente/Total,
    /// Subs/Cont/Total, Especiales opcional) + tabla destino ÚNICA (el mapa productivo v2). La
    /// caracterización T0.4 probó que la rama Q2 y esta producían valores idénticos en los 5 ASE
    /// (julio-1 + julio-2), así que el switch por ASE se retiró por redundante.
    /// </summary>
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

        // Plan 59 (H2.1): el destino se entrega por ROL (cierre de cada sección + especiales de todas
        // menos la última); el escritor resuelve la celda contra la plantilla. Cero celdas aquí.
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
        {
            [RolSeccion(SeccionComponente, ResolvedorDestinoValor.EncabezadoTotal)] = e15,
            [RolSeccion(SeccionSubsCont, ResolvedorDestinoValor.EncabezadoTotal)] = e26,
            [RolSeccion(SeccionComponente, ResolvedorDestinoValor.EncabezadoEspecialesR2)] = k15
        };

        return new WorkbookLeafInputsR2
        {
            E15 = e15,
            E26 = e26,
            K15 = k15,
            ValoresPorRol = celdas
        };
    }

    /// <summary>Rótulo de la sección de anticipos del R2 (anticipos a usuarios).</summary>
    private const string SeccionComponente = "Componente";

    /// <summary>Rótulo de la sección de subsidio/contribución del R2.</summary>
    private const string SeccionSubsCont = "Subs/Cont";

    private static string RolSeccion(string etiqueta, string encabezado) =>
        new ResolvedorDestinoValor.RolValor(
            ResolvedorDestinoValor.NivelDestino.Seccion, etiqueta, 1, encabezado).Clave;
    /// <summary>
    /// Plan 54 (H1.4, D4): mapeo R4 ÚNICO para ambas quincenas — fila <c>Total</c> con <c>B</c> vacío
    /// sobre la columna D, con la tabla destino ÚNICA del mapa productivo v2. <c>P = 0</c> (sin análogo
    /// en la fuente, patrón histórico). Valores idénticos a la rama Q2 por caracterización T0.4.
    /// </summary>
    private static WorkbookLeafInputsR4 MapearR4(Ase ase, List<object?[]> filas)
    {
        var filaTotal = filas.FirstOrDefault(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(0)).Equals("Total", StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrWhiteSpace(ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(1))))
            ?? throw new CalculoInvalidoException($"ASE {ase.Id}: no se encontró la fila Total (B vacío) del R4 para D9.");

        var d9 = ExcelWorksheetNavigator.CeldaNumero(filaTotal.ElementAtOrDefault(3));
        var p9 = 0m;

        // Plan 59 (H2.1): el destino se entrega por ROL (gran total del bloque + sus especiales); el
        // escritor resuelve la celda contra la plantilla. Cero celdas aquí.
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
        {
            [RolGranTotal(ResolvedorDestinoValor.EncabezadoTotal)] = d9,
            [RolGranTotal(ResolvedorDestinoValor.EncabezadoEspecialesR4)] = p9
        };

        return new WorkbookLeafInputsR4
        {
            D9 = d9,
            P9 = p9,
            ValoresPorRol = celdas
        };
    }

    private static string RolGranTotal(string encabezado) =>
        new ResolvedorDestinoValor.RolValor(
            ResolvedorDestinoValor.NivelDestino.GranTotal, string.Empty, 1, encabezado).Clave;

    /// <summary>
    /// Plan 29 (T2, Unidad R): construye la matriz de detalle R2 desde las filas crudas de
    /// <c>Sheet1</c>. Localiza: (a) la primera fila <c>Vlr Servicio</c> (col D) como corte superior;
    /// (b) la fila de ENCABEZADOS de componente inmediatamente superior (col E..); (c) la fila de
    /// cierre <c>Total</c> (col A). Resuelve columnas por NOMBRE de encabezado (nunca índice fijo),
    /// por lo que sirve para julio (E..P) y agosto (E..Q, con <c>Especiales</c>). Fail-fast que
    /// nombra ASE + archivo si falta alguna de las tres referencias (nunca valor inventado).
    /// </summary>
    private static DetalleR2AseInputs ConstruirDetalleR2(Ase ase, string rutaR2, List<object?[]> filas)
    {
        var indicePrimeraFila = filas.FindIndex(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(3)).Equals("Vlr Servicio", StringComparison.OrdinalIgnoreCase));
        if (indicePrimeraFila <= 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: la fuente R2 ({Path.GetFileName(rutaR2)}) no trae la fila 'Vlr Servicio' (col D) que delimita el inicio del detalle.");
        }

        var indiceHeader = indicePrimeraFila - 1;
        while (indiceHeader > 0 && !FilaTieneTextoDesde(filas[indiceHeader], 4))
        {
            indiceHeader--;
        }

        var (componentes, indicesPorEncabezado) = MapearEncabezadosDesde(filas[indiceHeader], 4);
        if (componentes.Count == 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: la fuente R2 ({Path.GetFileName(rutaR2)}) no trae encabezados de componente (col E..) en la fila {indiceHeader + 1}.");
        }

        var indiceTotal = -1;
        for (var i = indicePrimeraFila; i < filas.Count; i++)
        {
            if (ExcelWorksheetNavigator.CeldaTexto(filas[i].ElementAtOrDefault(0)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                indiceTotal = i;
                break;
            }
        }

        if (indiceTotal < 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: la fuente R2 ({Path.GetFileName(rutaR2)}) no trae la fila de cierre 'Total' (col A) del detalle.");
        }

        var tieneEspeciales = indicesPorEncabezado.Keys.Any(k => NormalizarEtiqueta(k) == "especiales");
        var filasDetalle = new List<DetalleR2Fila>();

        for (var i = indicePrimeraFila; i <= indiceTotal; i++)
        {
            var fila = filas[i];
            var valores = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);
            foreach (var componente in componentes)
            {
                valores[componente] = CeldaNumeroONulo(fila.ElementAtOrDefault(indicesPorEncabezado[componente]));
            }

            var a = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(0));
            var b = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(1));
            var c = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(2));
            var d = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(3));

            var hayTexto = !(string.IsNullOrWhiteSpace(a) && string.IsNullOrWhiteSpace(b)
                && string.IsNullOrWhiteSpace(c) && string.IsNullOrWhiteSpace(d));
            var hayValor = componentes.Any(componente => valores[componente] is not null);
            if (!hayTexto && !hayValor)
            {
                continue; // fila fantasma vacía: no forma parte de la secuencia observada
            }

            // Columna 'Especiales' ausente en la fuente (julio) → 0 explícito por fila (invariante).
            if (!tieneEspeciales)
            {
                valores["Especiales"] = 0m;
            }

            filasDetalle.Add(new DetalleR2Fila { A = a, B = b, C = c, D = d, ValoresPorComponente = valores });
        }

        return new DetalleR2AseInputs
        {
            Ase = ase,
            Componentes = componentes,
            TieneColumnaEspeciales = tieneEspeciales,
            Fuente = rutaR2,
            Filas = filasDetalle
        };
    }

    /// <summary>
    /// Plan 29 (T2, Unidad R): construye la matriz de detalle R4 desde las filas crudas de
    /// <c>Sheet1</c>. Misma mecánica que <see cref="ConstruirDetalleR2"/>, con la firma A–C
    /// (concepto en col C) y las columnas de componente a partir de la col D. Fail-fast que nombra
    /// ASE + archivo si falta la primera <c>Vlr Servicio</c>, la fila de encabezados o el cierre
    /// <c>Total</c>.
    /// </summary>
    private static DetalleR4AseInputs ConstruirDetalleR4(Ase ase, string rutaR4, List<object?[]> filas)
    {
        var indicePrimeraFila = filas.FindIndex(f =>
            ExcelWorksheetNavigator.CeldaTexto(f.ElementAtOrDefault(2)).Equals("Vlr Servicio", StringComparison.OrdinalIgnoreCase));
        if (indicePrimeraFila <= 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: la fuente R4 ({Path.GetFileName(rutaR4)}) no trae la fila 'Vlr Servicio' (col C) que delimita el inicio del detalle.");
        }

        var indiceHeader = indicePrimeraFila - 1;
        while (indiceHeader > 0 && !FilaTieneTextoDesde(filas[indiceHeader], 3))
        {
            indiceHeader--;
        }

        var (componentes, indicesPorEncabezado) = MapearEncabezadosDesde(filas[indiceHeader], 3);
        if (componentes.Count == 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: la fuente R4 ({Path.GetFileName(rutaR4)}) no trae encabezados de componente (col D..) en la fila {indiceHeader + 1}.");
        }

        var indiceTotal = -1;
        for (var i = indicePrimeraFila; i < filas.Count; i++)
        {
            if (ExcelWorksheetNavigator.CeldaTexto(filas[i].ElementAtOrDefault(0)).Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                indiceTotal = i;
                break;
            }
        }

        if (indiceTotal < 0)
        {
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: la fuente R4 ({Path.GetFileName(rutaR4)}) no trae la fila de cierre 'Total' (col A) del detalle.");
        }

        var filasDetalle = new List<DetalleR4Fila>();
        for (var i = indicePrimeraFila; i <= indiceTotal; i++)
        {
            var fila = filas[i];
            var valores = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);
            foreach (var componente in componentes)
            {
                valores[componente] = CeldaNumeroONulo(fila.ElementAtOrDefault(indicesPorEncabezado[componente]));
            }

            var a = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(0));
            var b = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(1));
            var c = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(2));

            var hayTexto = !(string.IsNullOrWhiteSpace(a) && string.IsNullOrWhiteSpace(b) && string.IsNullOrWhiteSpace(c));
            var hayValor = componentes.Any(componente => valores[componente] is not null);
            if (!hayTexto && !hayValor)
            {
                continue;
            }

            filasDetalle.Add(new DetalleR4Fila { A = a, B = b, C = c, ValoresPorComponente = valores });
        }

        return new DetalleR4AseInputs
        {
            Ase = ase,
            Componentes = componentes,
            Filas = filasDetalle
        };
    }

    /// <summary>
    /// Plan 29 (T2): verdad si la fila tiene texto en alguna columna desde <paramref name="desde"/>
    /// (0-based). Se usa para localizar la fila de encabezados subiendo desde la primera fila de datos.
    /// </summary>
    private static bool FilaTieneTextoDesde(object?[]? fila, int desde)
    {
        if (fila is null)
        {
            return false;
        }

        for (var j = desde; j < fila.Length; j++)
        {
            if (!string.IsNullOrWhiteSpace(ExcelWorksheetNavigator.CeldaTexto(fila[j])))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Plan 29 (T2): encabezados de columna desde <paramref name="desde"/> (0-based) en orden físico,
    /// con su índice 0-based. Los vacíos se omiten y los duplicados se conservan una sola vez.
    /// </summary>
    private static (List<string> Encabezados, Dictionary<string, int> Indices) MapearEncabezadosDesde(object?[] filaHeader, int desde)
    {
        var encabezados = new List<string>();
        var indices = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var j = desde; j < filaHeader.Length; j++)
        {
            var texto = ExcelWorksheetNavigator.CeldaTexto(filaHeader[j]);
            if (string.IsNullOrWhiteSpace(texto))
            {
                continue;
            }

            var encabezado = texto.Trim();
            if (indices.ContainsKey(encabezado))
            {
                continue;
            }

            indices[encabezado] = j;
            encabezados.Add(encabezado);
        }

        return (encabezados, indices);
    }
}
