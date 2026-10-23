using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;
using Serilog;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Plan 55 (H1.1/H1.3, D4 — ENMIENDA 2026-10-23): RESOLVEDOR POR ETIQUETAS del camino por empresa
/// (R1/R2/R4). Sustituye a la tabla congelada por (empresa, ASE)
/// (<see cref="WorkbookLeafCellMapPorEmpresa"/>, retirada del camino productivo a LEGACY).
///
/// Doctrina (S0.4 del plan, ratificada por la propietaria de la plantilla):
/// <list type="bullet">
/// <item><b>La plantilla es la autoridad.</b> El destino de cada operando se resuelve por CRUCE DE
/// ETIQUETAS <c>(cuadro/sección, fuente = col C, canal = col D, concepto = col E) → celda</c>, con el
/// COMPONENTE resuelto por el ENCABEZADO de columna que declara el propio cuadro del template.</item>
/// <item>Maquinaria reusada, no reinventada: <see cref="SinonimosEmpresaR1"/> (sinonimia abierta),
/// <see cref="ContratoSeccionPeriodFreeR1.SeccionPorCierre"/> (sección por frontera que cierra) y
/// <see cref="ContratoSeccionPeriodFreeR1.Alinear"/> (alineación monotónica por tipo) — el MISMO
/// criterio del camino de bloque (<c>EscritorValoresSlotR1</c>).</item>
/// <item><b>Ausencia</b> (el cuadro declara y la fuente no trae) → cero del template: no es error.</item>
/// <item><b>Exceso</b> (la fuente trae un operando que ningún cuadro declara) → registro auditable
/// (<see cref="ConciliacionEmpresaInputs.DestinosSinCelda"/>) + <c>Log.Warning</c>; nunca descarte
/// silencioso, nunca celda inventada, nunca aborto por defecto.</item>
/// <item>Celda destino con <c>&lt;f&gt;</c> → NO se escribe (guard anti-fórmula) y se registra en
/// <see cref="ConciliacionEmpresaInputs.DestinosFormula"/>: el escritor values-only jamás
/// sobrescribe una fórmula.</item>
/// <item>Agregado por empresa ANTI-TAUTOLÓGICO (H1.3): se deriva de la FUENTE, agrupado por etiqueta,
/// sobre el MISMO conjunto de roles/secciones que alimenta el visible de bloque — nunca "la suma de lo
/// que escribimos".</item>
/// </list>
/// Cero listas por ASE, cero literales de empresa, cero ramas por quincena: la geometría sale del
/// template en tiempo de ejecución y el vocabulario de empresas del catálogo + la sinonimia.
/// </summary>
internal static class ResolvedorConciliacionEmpresa
{
    internal const string HojaR1 = "Reporte Componentes R1";
    internal const string HojaR2 = "Rem. Anticipos R2";
    internal const string HojaR4 = "Reversion Pagos R4";

    /// <summary>Encabezado del componente "Total" (varía de columna por hoja: F en R1, E en R2, D en R4).</summary>
    private const string EncabezadoTotal = "Total";

    /// <summary>Encabezados admitidos del componente Especiales (R1/R4: SERVICIO ESPECIALES; R2: Especiales).</summary>
    private static readonly string[] EncabezadosEspeciales = ["SERVICIO ESPECIALES", "Especiales"];

    private const string CuadroComponente = "Componente";
    private const string CuadroSubsCont = "Subs/Cont";

    // ── Modelo de la FUENTE ────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Operando de fuente ya tipado por las CUATRO etiquetas de la doctrina: reporte (hoja),
    /// cuadro/sección, canal y concepto, más su componente Total y Especiales.
    /// </summary>
    private sealed record Operando(
        string Reporte,
        string Cuadro,
        SeccionFronteraArriba? Seccion,
        string Etiqueta,
        string Canal,
        string Concepto,
        decimal Total,
        decimal Especiales);

    // ── API ────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Construye la conciliación por empresa del ASE: operaciones de fuente ya resueltas por etiqueta,
    /// destinos resueltos contra la plantilla (si se conoce su ruta) y agregados anti-tautológicos.
    /// </summary>
    internal static IReadOnlyList<ConciliacionEmpresaInputs> Resolver(
        Ase ase,
        IReadOnlyList<FilaEspejoR1> espejoR1,
        IReadOnlyList<object?[]> filasR2,
        IReadOnlyList<object?[]> filasR4,
        string? rutaPlantilla)
    {
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(espejoR1);
        ArgumentNullException.ThrowIfNull(filasR2);
        ArgumentNullException.ThrowIfNull(filasR4);

        var operandoR1 = OperandosR1(ase, espejoR1);
        var operandoR2 = OperandosR2(filasR2);
        var operandoR4 = OperandosR4(filasR4);
        var operando = operandoR1.Concat(operandoR2).Concat(operandoR4).ToList();

        var plantilla = string.IsNullOrWhiteSpace(rutaPlantilla)
            ? null
            : PlantillaEmpresa.Obtener(rutaPlantilla);
        var bloqueR1 = plantilla?.Bloque(HojaR1, ase.Id);
        var bloqueR2 = plantilla?.Bloque(HojaR2, ase.Id);
        var bloqueR4 = plantilla?.Bloque(HojaR4, ase.Id);

        var seccionesR1 = bloqueR1 is null
            ? null
            : AlinearSeccionesR1(espejoR1, bloqueR1);

        var acumulados = new List<Acumulado>();
        var catalogo = EmpresaFacturacion.Catalogo;
        foreach (var empresa in catalogo)
        {
            acumulados.Add(new Acumulado(empresa));
        }

        foreach (var op in operando)
        {
            var acumulado = ObtenerAcumulado(acumulados, op.Etiqueta, catalogo);
            acumulado.AseId = ase.Id;
            acumulado.Operandos.Add(op);
        }

        var resultado = new List<ConciliacionEmpresaInputs>(acumulados.Count);
        foreach (var acumulado in acumulados)
        {
            var inputs = new ConciliacionEmpresaInputs
            {
                Empresa = acumulado.Empresa,
                Ase = ase,
                CeldasR1 = ResolverCeldas(acumulado, "R1", bloqueR1, seccionesR1),
                CeldasR2 = ResolverCeldas(acumulado, "R2", bloqueR2, null),
                CeldasR4 = ResolverCeldas(acumulado, "R4", bloqueR4, null),
                VisibleR1 = VisibleR1(acumulado),
                VisibleR2 = VisibleR2(acumulado),
                VisibleR4 = VisibleR4(acumulado),
                FuentesAusentes = FuentesAusentes(acumulado, bloqueR1, bloqueR2, bloqueR4),
                DestinosFormula = acumulado.DestinosFormula.Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                DestinosSinCelda = acumulado.DestinosSinCelda.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
            };

            resultado.Add(inputs);
        }

        return resultado;
    }

    // ── Operandos de la FUENTE (mismo tipado para fuente y plantilla) ─────────────────────────

    /// <summary>
    /// R1 (<c>Recaudoporcomponente</c>): filas-dato de empresa = <c>C=etiqueta ∧ D='Total' ∧ E</c>
    /// vacío (la fila de cierre de la empresa en su sección). Su sección la fija la frontera que la
    /// CIERRA (<see cref="ContratoSeccionPeriodFreeR1.SeccionPorCierre"/>) y su cuadro la posición
    /// respecto de la fila <c>Componente/Total</c>. Valores por ENCABEZADO (Total / SERVICIO
    /// ESPECIALES), nunca por índice.
    /// </summary>
    private static IEnumerable<Operando> OperandosR1(Ase ase, IReadOnlyList<FilaEspejoR1> espejo)
    {
        var indiceComponente = IndiceComponente(espejo.Select(f => f.A).ToList(), espejo.Select(f => f.B).ToList());
        if (indiceComponente < 0)
        {
            // Invariante de ESTRUCTURA del reporte (no es una fuente opcional): la fila
            // 'Componente/Total' separa los dos cuadros. Su ausencia sigue siendo fail-fast; solo la
            // AUSENCIA DE UNA FUENTE es cero (Plan 50 S0.2 / Plan 55 S0.4).
            throw new CalculoInvalidoException(
                $"ASE {ase.Id}: no se encontró la fila 'Componente' en la fuente R1 para localizar los bloques por empresa.");
        }

        for (var i = 0; i < espejo.Count; i++)
        {
            var fila = espejo[i];
            if (!EsFilaEmpresaR1(fila))
            {
                continue;
            }

            var seccion = ContratoSeccionPeriodFreeR1.SeccionPorCierre(espejo, i);
            if (seccion is null)
            {
                continue; // firma AFaseo: no es una empresa
            }

            yield return new Operando(
                "R1",
                i < indiceComponente ? CuadroComponente : CuadroSubsCont,
                seccion,
                fila.C.Trim(),
                fila.D.Trim(),
                fila.E.Trim(),
                fila.Valor(EncabezadoTotal) ?? 0m,
                fila.Valor("SERVICIO ESPECIALES") ?? 0m);
        }
    }

    /// <summary>
    /// R2 (<c>RerpoteDetalleSaldosaFavor</c>): filas-dato de empresa = <c>B=etiqueta ∧ C='Total' ∧ D</c>
    /// vacío. Cuadro por posición respecto de <c>Componente/Total</c> y <c>Subs/Cont/Total</c>. Total en
    /// el encabezado <c>Total</c> del cuadro (col E de la fuente) y Especiales en su encabezado.
    /// </summary>
    private static IEnumerable<Operando> OperandosR2(IReadOnlyList<object?[]> filas)
    {
        var indiceComponente = BuscarFila(filas, (a, b) => Igual(a, CuadroComponente) && Igual(b, "Total"));
        var indiceSubsCont = BuscarFila(filas, (a, b) => Igual(a, CuadroSubsCont) && Igual(b, "Total"));
        if (indiceComponente < 0 || indiceSubsCont < 0)
        {
            yield break;
        }

        var (indiceTotal, indiceEspeciales) = IndicesComponentes(filas);

        for (var i = 0; i < filas.Count; i++)
        {
            if (i >= indiceSubsCont)
            {
                break;
            }

            var fila = filas[i];
            var etiqueta = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(1));
            var concepto = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(2));
            var canal = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(3));
            if (etiqueta.Length == 0 || !Igual(concepto, "Total") || canal.Length != 0)
            {
                continue;
            }

            yield return new Operando(
                "R2",
                i < indiceComponente ? CuadroComponente : CuadroSubsCont,
                null,
                etiqueta,
                canal,
                concepto,
                indiceTotal >= 0 ? ExcelWorksheetNavigator.CeldaNumero(fila.ElementAtOrDefault(indiceTotal)) : 0m,
                indiceEspeciales >= 0 ? ExcelWorksheetNavigator.CeldaNumero(fila.ElementAtOrDefault(indiceEspeciales)) : 0m);
        }
    }

    /// <summary>
    /// R4 (<c>ReversiónPorComponente</c>): filas-dato de empresa = <c>A=etiqueta ∧ B='Total'</c>; el
    /// valor es la columna <c>Total</c> (D) y el componente Especiales (P) no tiene análogo en la
    /// fuente → ausencia = cero del template.
    /// </summary>
    private static IEnumerable<Operando> OperandosR4(IReadOnlyList<object?[]> filas)
    {
        var (indiceTotal, indiceEspeciales) = IndicesComponentes(filas);

        for (var i = 0; i < filas.Count; i++)
        {
            var fila = filas[i];
            var etiqueta = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(0));
            var concepto = ExcelWorksheetNavigator.CeldaTexto(fila.ElementAtOrDefault(1));
            if (etiqueta.Length == 0 || !Igual(concepto, "Total") || Igual(etiqueta, "Total"))
            {
                continue;
            }

            yield return new Operando(
                "R4",
                string.Empty,
                null,
                etiqueta,
                concepto,
                concepto,
                indiceTotal >= 0 ? ExcelWorksheetNavigator.CeldaNumero(fila.ElementAtOrDefault(indiceTotal)) : 0m,
                indiceEspeciales >= 0 ? ExcelWorksheetNavigator.CeldaNumero(fila.ElementAtOrDefault(indiceEspeciales)) : 0m);
        }
    }

    // ── Agregado por empresa ANTI-TAUTOLÓGICO (H1.3 / CA8) ────────────────────────────────────

    /// <summary>
    /// Visible R1 por empresa derivado de la FUENTE sobre el MISMO conjunto de roles que alimenta el
    /// visible de bloque: <c>Σ Total de las secciones Oportuno − Σ Especiales de las secciones Oportuno
    /// menos la ÚLTIMA</c>. Es la descomposición exacta de
    /// <c>TotalOportunoEsperadoPorAse = ΣF(Mes/Total) − ΣEsp(Mes/Total menos la última)</c>: se agrupa
    /// por etiqueta de fuente, nunca por lo escrito en el workbook.
    /// </summary>
    private static decimal VisibleR1(Acumulado acumulado)
    {
        var oportunas = acumulado.Operandos
            .Where(o => o.Reporte == "R1" && o.Seccion?.Tipo == TipoSeccionEspejoR1.Oportuno)
            .ToList();
        if (oportunas.Count == 0)
        {
            return 0m;
        }

        var ultima = oportunas.Max(o => o.Seccion!.Value.Ordinal);
        var total = oportunas.Sum(o => o.Total);
        var especiales = oportunas.Where(o => o.Seccion!.Value.Ordinal != ultima).Sum(o => o.Especiales);
        return total - especiales;
    }

    /// <summary>
    /// Visible R2 por empresa = <c>Total(Componente) + Total(Subs/Cont) − Especiales(Componente)</c>,
    /// la misma descomposición del visible de bloque (<c>E15 + E26 − K15</c>).
    /// </summary>
    private static decimal VisibleR2(Acumulado acumulado)
    {
        var componente = acumulado.Operandos.FirstOrDefault(o => o.Reporte == "R2" && o.Cuadro == CuadroComponente);
        var subsCont = acumulado.Operandos.FirstOrDefault(o => o.Reporte == "R2" && o.Cuadro == CuadroSubsCont);
        return (componente?.Total ?? 0m) + (subsCont?.Total ?? 0m) - (componente?.Especiales ?? 0m);
    }

    /// <summary>Visible R4 por empresa = <c>Total</c> (P = 0 en la fuente): <c>D(total) = D(empresa) − P</c>.</summary>
    private static decimal VisibleR4(Acumulado acumulado) =>
        acumulado.Operandos.Where(o => o.Reporte == "R4").Sum(o => o.Total - o.Especiales);

    // ── Resolución de destinos contra la PLANTILLA ────────────────────────────────────────────

    /// <summary>
    /// Cruza cada operando del reporte con la fila que el cuadro del template declara para esa
    /// <c>(sección, etiqueta, canal, concepto)</c> y devuelve <c>celda → valor</c> usando el
    /// componente por ENCABEZADO del template. Exceso → registro; fórmula → registro (nunca se escribe).
    /// </summary>
    private static IReadOnlyDictionary<string, decimal> ResolverCeldas(
        Acumulado acumulado,
        string reporte,
        BloquePlantilla? bloque,
        IReadOnlyDictionary<(string Cuadro, SeccionFronteraArriba Seccion), SeccionFronteraArriba>? secciones)
    {
        var celdas = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        var operandos = acumulado.Operandos.Where(o => o.Reporte == reporte).ToList();
        if (operandos.Count == 0)
        {
            return celdas;
        }

        if (bloque is null)
        {
            // Sin plantilla no hay autoridad de destino: la ausencia de celdas es explícita y auditable.
            foreach (var operando in operandos)
            {
                acumulado.DestinosSinCelda.Add($"{reporte}:{operando.Etiqueta}:{operando.Cuadro}:sin-plantilla");
            }

            return celdas;
        }

        foreach (var operando in operandos)
        {
            var fila = bloque.BuscarFila(operando, secciones);
            if (fila is null)
            {
                acumulado.DestinosSinCelda.Add(DescribirExceso(reporte, operando));
                Log.ForContext("AseId", acumulado.AseId)
                    .ForContext("Reporte", reporte)
                    .ForContext("Fuente", operando.Etiqueta)
                    .ForContext("Seccion", operando.Seccion?.ToString() ?? "-")
                    .ForContext("Cuadro", operando.Cuadro)
                    .Warning(
                        "Plan 55: la fuente {Reporte} trae el operando de la empresa {Fuente} (cuadro {Cuadro}, sección {Seccion}) y ningún cuadro de la plantilla lo declara para el ASE {AseId}; se registra como exceso y NO se inventa celda destino.",
                        reporte,
                        operando.Etiqueta,
                        operando.Cuadro,
                        operando.Seccion?.ToString() ?? "-",
                        acumulado.AseId);
                continue;
            }

            Asignar(acumulado, reporte, fila, bloque.ColumnaTotal, operando.Total, celdas);
            if (bloque.ColumnaEspeciales is not null && operando.Cuadro != CuadroSubsCont)
            {
                Asignar(acumulado, reporte, fila, bloque.ColumnaEspeciales, operando.Especiales, celdas);
            }
        }

        return celdas;
    }

    private static void Asignar(
        Acumulado acumulado,
        string reporte,
        FilaPlantilla fila,
        string columna,
        decimal valor,
        IDictionary<string, decimal> celdas)
    {
        var celda = $"{columna}{fila.Numero}";
        if (fila.EsFormula(columna))
        {
            acumulado.DestinosFormula.Add($"{reporte}:{celda}");
            Log.ForContext("AseId", acumulado.AseId)
                .ForContext("Reporte", reporte)
                .ForContext("Celda", celda)
                .Warning(
                    "Plan 55: la celda destino {Reporte}!{Celda} es FÓRMULA en la plantilla vigente; el escritor values-only no la toca (guard anti-fórmula).",
                    reporte,
                    celda);
            return;
        }

        if (celdas.TryGetValue(celda, out var previo) && previo != valor)
        {
            acumulado.DestinosSinCelda.Add($"{reporte}:{celda}:conflicto");
        }

        celdas[celda] = valor;
    }

    private static string DescribirExceso(string reporte, Operando operando) =>
        $"{reporte}:{operando.Etiqueta}:{operando.Cuadro}:{operando.Seccion?.Tipo.ToString() ?? "-"}#{operando.Seccion?.Ordinal.ToString(CultureInfo.InvariantCulture) ?? "-"}";

    /// <summary>
    /// Alineación de SECCIONES fuente ↔ template del bloque R1 por tipo y orden monotónico
    /// (<see cref="ContratoSeccionPeriodFreeR1.Alinear"/>), <b>dentro de cada cuadro</b>: las
    /// secciones del cuadro <c>Componente</c> se alinean con las del cuadro <c>Componente</c> del
    /// template y las de <c>Subs/Cont</c> con las de <c>Subs/Cont</c> (la fuente puede traer más
    /// secciones Oportuno que el cuadro oficial y nunca se remapea por posición).
    /// </summary>
    private static IReadOnlyDictionary<(string Cuadro, SeccionFronteraArriba Seccion), SeccionFronteraArriba> AlinearSeccionesR1(
        IReadOnlyList<FilaEspejoR1> espejo,
        BloquePlantilla bloque)
    {
        var indiceComponente = IndiceComponente(espejo.Select(f => f.A).ToList(), espejo.Select(f => f.B).ToList());
        var resultado = new Dictionary<(string, SeccionFronteraArriba), SeccionFronteraArriba>();
        foreach (var cuadro in new[] { CuadroComponente, CuadroSubsCont })
        {
            var fuente = Distintas(Enumerable.Range(0, espejo.Count)
                .Where(i => string.Equals(i < indiceComponente ? CuadroComponente : CuadroSubsCont, cuadro, StringComparison.OrdinalIgnoreCase))
                .Select(i => ContratoSeccionPeriodFreeR1.SeccionPorCierre(espejo, i)));
            var oficial = Distintas(bloque.SeccionesPorCuadro().Where(s => string.Equals(s.Cuadro, cuadro, StringComparison.OrdinalIgnoreCase)).Select(s => s.Seccion));
            foreach (var (origen, destino) in ContratoSeccionPeriodFreeR1.Alinear(fuente, oficial))
            {
                resultado[(cuadro, origen)] = destino;
            }
        }

        return resultado;
    }

    private static List<SeccionFronteraArriba> Distintas(IEnumerable<SeccionFronteraArriba?> secciones)
    {
        var lista = new List<SeccionFronteraArriba>();
        SeccionFronteraArriba? previa = null;
        foreach (var seccion in secciones)
        {
            if (seccion is null)
            {
                continue;
            }

            if (previa is null || previa.Value != seccion.Value)
            {
                lista.Add(seccion.Value);
            }

            previa = seccion;
        }

        return lista;
    }

    // ── Huella de AUSENCIA (Plan 50 A1, intacta) ───────────────────────────────────────────────

    /// <summary>
    /// Huella <c>{Reporte}:{etiqueta}</c> por fuente ausente: la etiqueta es la que el CUADRO declara
    /// para esa empresa en ese reporte (autoridad del template) y, si no la declara, el rótulo del
    /// catálogo. Sin filas de la empresa en la fuente = ausencia legítima (cero del template).
    /// </summary>
    private static IReadOnlyList<string> FuentesAusentes(
        Acumulado acumulado,
        BloquePlantilla? bloqueR1,
        BloquePlantilla? bloqueR2,
        BloquePlantilla? bloqueR4)
    {
        var ausentes = new List<string>();
        foreach (var (reporte, bloque) in new[] { ("R1", bloqueR1), ("R2", bloqueR2), ("R4", bloqueR4) })
        {
            if (acumulado.Operandos.Any(o => o.Reporte == reporte))
            {
                continue;
            }

            // La AUSENCIA solo tiene sentido donde el CUADRO declara la fuente: si el cuadro no
            // declara ese rótulo, no hay slot que quede en cero (no aplica, no es huella).
            var declarada = bloque?.EtiquetaDeclarada(reporte, acumulado.Empresa.LabelTemplate);
            if (declarada is null)
            {
                continue;
            }

            ausentes.Add($"{reporte}:{declarada}");
        }

        return ausentes;
    }

    // ── Agrupación por etiqueta (catálogo abierto + sinonimia) ────────────────────────────────

    private static Acumulado ObtenerAcumulado(
        List<Acumulado> acumulados,
        string etiqueta,
        IReadOnlyList<EmpresaFacturacion> catalogo)
    {
        var porCatalogo = catalogo.FirstOrDefault(empresa => SinonimosEmpresaR1.SonMismaEmpresa(empresa.LabelTemplate, etiqueta));
        if (porCatalogo is not null)
        {
            return acumulados.First(a => a.Empresa.Id == porCatalogo.Id);
        }

        var existente = acumulados.FirstOrDefault(a =>
            a.Empresa.Id > catalogo.Count && SinonimosEmpresaR1.SonMismaEmpresa(a.Empresa.LabelTemplate, etiqueta));
        if (existente is not null)
        {
            return existente;
        }

        // Empresa declarada por el CUADRO pero ausente del catálogo (p. ej. EAB en R2): se registra
        // como empresa de fuente (Id sintético > catálogo) para que su dinero entre al agregado.
        var empresa = new EmpresaFacturacion
        {
            Id = catalogo.Count + 1 + acumulados.Count(a => a.Empresa.Id > catalogo.Count),
            Nombre = etiqueta.Trim(),
            LabelTemplate = etiqueta.Trim()
        };
        var nuevo = new Acumulado(empresa);
        acumulados.Add(nuevo);
        return nuevo;
    }

    // ── Helpers puros ─────────────────────────────────────────────────────────────────────────

    private static bool EsFilaEmpresaR1(FilaEspejoR1 fila) =>
        !string.IsNullOrWhiteSpace(fila.C)
        && Igual(fila.D, "Total")
        && string.IsNullOrWhiteSpace(fila.E)
        && !Igual(fila.C, "Total");

    private static int IndiceComponente(IReadOnlyList<string> a, IReadOnlyList<string> b)
    {
        for (var i = 0; i < a.Count; i++)
        {
            if (Igual(a[i], CuadroComponente) && Igual(b[i], "Total"))
            {
                return i;
            }
        }

        return -1;
    }

    private static int BuscarFila(IReadOnlyList<object?[]> filas, Func<string, string, bool> predicado)
    {
        for (var i = 0; i < filas.Count; i++)
        {
            var a = ExcelWorksheetNavigator.CeldaTexto(filas[i].ElementAtOrDefault(0));
            var b = ExcelWorksheetNavigator.CeldaTexto(filas[i].ElementAtOrDefault(1));
            if (predicado(a, b))
            {
                return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// Columnas del componente <c>Total</c> y de Especiales del bloque de datos: se resuelven en la
    /// fila de ENCABEZADOS (la inmediatamente anterior a la primera fila de concepto
    /// <c>Vlr Servicio</c>), nunca por índice fijo ni por barrido de toda la hoja.
    /// </summary>
    private static (int Total, int Especiales) IndicesComponentes(IReadOnlyList<object?[]> filas)
    {
        var filaConcepto = -1;
        for (var i = 1; i < filas.Count; i++)
        {
            if (filas[i].Take(5).Any(c => Igual(ExcelWorksheetNavigator.CeldaTexto(c), "Vlr Servicio")))
            {
                filaConcepto = i;
                break;
            }
        }

        if (filaConcepto < 1)
        {
            return (-1, -1);
        }

        var encabezados = filas[filaConcepto - 1];
        var total = IndiceEnFila(encabezados, EncabezadoTotal);
        var especiales = -1;
        foreach (var encabezado in EncabezadosEspeciales)
        {
            especiales = IndiceEnFila(encabezados, encabezado);
            if (especiales >= 0)
            {
                break;
            }
        }

        return (total, especiales);
    }

    private static int IndiceEnFila(object?[] fila, string encabezado)
    {
        for (var i = 0; i < fila.Length; i++)
        {
            if (Igual(ExcelWorksheetNavigator.CeldaTexto(fila[i]), encabezado))
            {
                return i;
            }
        }

        return -1;
    }

    internal static bool Igual(string? a, string? b) =>
        string.Equals(a?.Trim(), b?.Trim(), StringComparison.OrdinalIgnoreCase);

    /// <summary>Acumulado por empresa (catálogo o etiqueta declarada por el cuadro).</summary>
    private sealed class Acumulado(EmpresaFacturacion empresa)
    {
        public EmpresaFacturacion Empresa { get; } = empresa;

        public int AseId { get; set; }

        public List<Operando> Operandos { get; } = [];

        public List<string> DestinosFormula { get; } = [];

        public List<string> DestinosSinCelda { get; } = [];
    }

    // ═════════════════════════════════════════════════════════════════════════════════════════
    // Plantilla: geometría leída por bytes (la plantilla es la autoridad)
    // ═════════════════════════════════════════════════════════════════════════════════════════

    /// <summary>Fila del template: número, textos por LETRA de columna y letras con <c>&lt;f&gt;</c>.</summary>
    private sealed record FilaPlantilla(int Numero, IReadOnlyDictionary<string, string> Textos, IReadOnlySet<string> Formulas)
    {
        public string A => Texto("A");

        public string B => Texto("B");

        public string C => Texto("C");

        public string D => Texto("D");

        public string E => Texto("E");

        public bool EsFormula(string columna) => Formulas.Contains(columna);

        public string Texto(string columna) => Textos.TryGetValue(columna, out var texto) ? texto : string.Empty;

        public FilaEspejoR1 Etiquetas() => new() { A = A, B = B, C = C, D = D, E = E };
    }

    /// <summary>Bloque de un ASE en una hoja: filas, columnas por encabezado y fronteras de cuadro.</summary>
    private sealed class BloquePlantilla
    {
        private readonly IReadOnlyList<FilaPlantilla> _filas;
        private readonly IReadOnlyDictionary<string, string> _columnas;
        private readonly int _indiceComponente;

        public BloquePlantilla(
            string nombreHoja,
            IReadOnlyList<FilaPlantilla> filas,
            IReadOnlyDictionary<string, string> columnas,
            int indiceComponente)
        {
            NombreHoja = nombreHoja;
            _filas = filas;
            _columnas = columnas;
            _indiceComponente = indiceComponente;
            ColumnaTotal = columnas.TryGetValue(EncabezadoTotal, out var total)
                ? total
                : throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"El cuadro de '{nombreHoja}' no declara el encabezado '{EncabezadoTotal}' (componente Total) para resolver el destino por etiqueta.");
            ColumnaEspeciales = EncabezadosEspeciales.Select(e => columnas.TryGetValue(e, out var c) ? c : null).FirstOrDefault(c => c is not null);
        }

        public string NombreHoja { get; }

        /// <summary>Letra de la columna del componente <c>Total</c> (por ENCABEZADO del template).</summary>
        public string ColumnaTotal { get; }

        /// <summary>Letra del componente Especiales, si el cuadro lo declara.</summary>
        public string? ColumnaEspeciales { get; }

        /// <summary>Secciones del bloque, tipadas por la frontera que las CIERRA (misma regla de la fuente).</summary>
        public IEnumerable<SeccionFronteraArriba?> Secciones()
        {
            var etiquetas = _filas.Select(f => f.Etiquetas()).ToList();
            for (var i = 0; i < etiquetas.Count; i++)
            {
                yield return ContratoSeccionPeriodFreeR1.SeccionPorCierre(etiquetas, i);
            }
        }

        /// <summary>Secciones del bloque con el CUADRO al que pertenece cada fila.</summary>
        public IEnumerable<(string Cuadro, SeccionFronteraArriba? Seccion)> SeccionesPorCuadro()
        {
            var etiquetas = _filas.Select(f => f.Etiquetas()).ToList();
            for (var i = 0; i < etiquetas.Count; i++)
            {
                yield return (
                    i < _indiceComponente ? CuadroComponente : CuadroSubsCont,
                    ContratoSeccionPeriodFreeR1.SeccionPorCierre(etiquetas, i));
            }
        }

        public FilaPlantilla? BuscarFila(
            Operando operando,
            IReadOnlyDictionary<(string Cuadro, SeccionFronteraArriba Seccion), SeccionFronteraArriba>? seccionesR1)
        {
            var etiquetas = _filas.Select(f => f.Etiquetas()).ToList();
            for (var i = 0; i < _filas.Count; i++)
            {
                var fila = _filas[i];
                var cuadro = i < _indiceComponente ? CuadroComponente : CuadroSubsCont;

                if (!string.IsNullOrEmpty(operando.Cuadro)
                    && !string.Equals(cuadro, operando.Cuadro, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Las CUATRO etiquetas de la doctrina se leen de las columnas que el propio cuadro
                // declara para cada reporte (R1: C/D/E; R2: B/C/D; R4: A/B) — nunca por índice fijo.
                var encaja = operando.Reporte switch
                {
                    "R1" => SinonimosEmpresaR1.SonMismaEmpresa(fila.C, operando.Etiqueta)
                        && Igual(fila.D, operando.Canal)
                        && Igual(fila.E, operando.Concepto),
                    "R2" => SinonimosEmpresaR1.SonMismaEmpresa(fila.B, operando.Etiqueta)
                        && Igual(fila.C, "Total")
                        && string.IsNullOrWhiteSpace(fila.D),
                    "R4" => SinonimosEmpresaR1.SonMismaEmpresa(fila.A, operando.Etiqueta)
                        && Igual(fila.B, "Total"),
                    _ => false
                };

                if (!encaja)
                {
                    continue;
                }

                if (operando.Seccion is { } seccion)
                {
                    var oficial = ContratoSeccionPeriodFreeR1.SeccionPorCierre(etiquetas, i);
                    var esperada = seccionesR1 is not null && seccionesR1.TryGetValue((operando.Cuadro, seccion), out var destino)
                        ? destino
                        : seccion;
                    if (oficial.HasValue && oficial.Value != esperada)
                    {
                        continue;
                    }
                }

                return fila;
            }

            return null;
        }

        public string? EtiquetaDeclarada(string reporte, string etiqueta)
        {
            foreach (var fila in _filas)
            {
                var declarada = reporte switch
                {
                    "R2" => fila.B,
                    "R4" => fila.A,
                    _ => fila.C
                };

                if (SinonimosEmpresaR1.SonMismaEmpresa(declarada, etiqueta))
                {
                    return declarada;
                }
            }

            return null;
        }
    }

    /// <summary>
    /// Plantilla vigente leída UNA vez por ruta (read-only, sin Excel/COM): filas con sus etiquetas
    /// A–E, encabezados por columna y celdas con <c>&lt;f&gt;</c> (guard anti-fórmula por bytes).
    /// </summary>
    private sealed class PlantillaEmpresa
    {
        private static readonly Dictionary<string, PlantillaEmpresa> Cache = new(StringComparer.OrdinalIgnoreCase);
        private static readonly object Candado = new();

        private readonly Dictionary<string, HojaPlantilla> _hojas;

        private PlantillaEmpresa(Dictionary<string, HojaPlantilla> hojas) => _hojas = hojas;

        public static PlantillaEmpresa Obtener(string ruta)
        {
            var clave = Path.GetFullPath(ruta);
            lock (Candado)
            {
                if (Cache.TryGetValue(clave, out var existente))
                {
                    return existente;
                }

                var leida = Leer(clave);
                Cache[clave] = leida;
                return leida;
            }
        }

        public BloquePlantilla? Bloque(string hoja, int aseId) =>
            _hojas.TryGetValue(hoja, out var leida) ? leida.Bloque(aseId) : null;

        private static PlantillaEmpresa Leer(string ruta)
        {
            if (!File.Exists(ruta))
            {
                throw new ArchivoFuenteNoEncontradoException($"No se encontró la plantilla para resolver destinos por etiqueta: '{ruta}'.");
            }

            using var documento = SpreadsheetDocument.Open(ruta, false);
            var workbookPart = documento.WorkbookPart
                ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La plantilla '{ruta}' no tiene WorkbookPart válido.");

            var hojas = new Dictionary<string, HojaPlantilla>(StringComparer.OrdinalIgnoreCase);
            foreach (var nombre in new[] { HojaR1, HojaR2, HojaR4 })
            {
                hojas[nombre] = LeerHoja(workbookPart, nombre);
            }

            return new PlantillaEmpresa(hojas);
        }

        private static HojaPlantilla LeerHoja(WorkbookPart workbookPart, string nombre)
        {
            var hoja = workbookPart.Workbook?.Descendants<Sheet>()
                    .FirstOrDefault(s => string.Equals(s.Name?.Value, nombre, StringComparison.OrdinalIgnoreCase))
                ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La plantilla no trae la hoja '{nombre}' para resolver destinos por etiqueta.");
            var parte = (WorksheetPart)workbookPart.GetPartById(hoja.Id!);
            var worksheet = parte.Worksheet
                ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{nombre}' de la plantilla no tiene WorksheetPart válido.");

            var filas = new List<FilaPlantilla>();
            foreach (var fila in worksheet.Descendants<Row>())
            {
                if (fila.RowIndex is null)
                {
                    continue;
                }

                var textos = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var formulas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var celda in fila.Elements<Cell>())
                {
                    var referencia = celda.CellReference?.Value;
                    if (string.IsNullOrWhiteSpace(referencia))
                    {
                        continue;
                    }

                    var columna = new string(referencia.TakeWhile(char.IsLetter).ToArray()).ToUpperInvariant();
                    textos[columna] = TextoCelda(workbookPart, celda);
                    if (celda.CellFormula is not null)
                    {
                        formulas.Add(columna);
                    }
                }

                filas.Add(new FilaPlantilla((int)fila.RowIndex.Value, textos, formulas));
            }

            return new HojaPlantilla(nombre, filas.OrderBy(f => f.Numero).ToList());
        }

        private static string TextoCelda(WorkbookPart workbookPart, Cell celda)
        {
            if (celda.DataType is not null && celda.DataType.Value == CellValues.SharedString
                && celda.CellValue is not null
                && int.TryParse(celda.CellValue.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var indice))
            {
                var tabla = workbookPart.SharedStringTablePart?.SharedStringTable;
                var item = tabla?.Elements<SharedStringItem>().ElementAtOrDefault(indice);
                return (item?.InnerText ?? string.Empty).Trim();
            }

            if (celda.DataType is not null && celda.DataType.Value == CellValues.InlineString)
            {
                return (celda.InlineString?.InnerText ?? string.Empty).Trim();
            }

            return (celda.CellValue?.InnerText ?? string.Empty).Trim();
        }
    }

    /// <summary>Hoja del template: sus filas y la localización del bloque de un ASE por su rótulo.</summary>
    private sealed class HojaPlantilla(string nombre, IReadOnlyList<FilaPlantilla> filas)
    {
        public BloquePlantilla? Bloque(int aseId)
        {
            var esperado = AseFactory.DesdeId(aseId).NombreCompleto;
            var inicio = -1;
            for (var i = 0; i < filas.Count - 1; i++)
            {
                if (SinonimosEmpresaR1.SonMismaEmpresa(filas[i].B, esperado))
                {
                    inicio = i;
                    break;
                }
            }

            if (inicio < 0)
            {
                return null;
            }

            // Fila de ENCABEZADOS del cuadro: la inmediatamente anterior a la primera fila de datos
            // (<c>Vlr Servicio</c>). Las hojas traen filas en blanco entre el rótulo del ASE y los
            // encabezados (p. ej. ASE4/ASE5), así que se localiza por la firma del primer dato y no
            // por un offset fijo.
            var indiceEncabezado = -1;
            for (var i = inicio + 1; i < filas.Count; i++)
            {
                if (filas[i].Textos.Values.Any(v => Igual(v, "Vlr Servicio")))
                {
                    indiceEncabezado = i - 1;
                    break;
                }

                if (Igual(filas[i].A, "Total") && string.IsNullOrWhiteSpace(filas[i].B))
                {
                    break;
                }
            }

            if (indiceEncabezado < 0)
            {
                return null;
            }

            var fin = filas.Count - 1;
            for (var j = inicio; j < filas.Count; j++)
            {
                if (Igual(filas[j].A, "Total") && string.IsNullOrWhiteSpace(filas[j].B))
                {
                    fin = j;
                    break;
                }
            }

            var delBloque = filas.Skip(inicio).Take(fin - inicio + 1).ToList();
            var columnas = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var encabezado = filas[indiceEncabezado];
            foreach (var (letra, texto) in encabezado.Textos.OrderBy(kv => kv.Key, StringComparer.Ordinal))
            {
                var limpio = texto.Trim();
                if (limpio.Length == 0 || columnas.ContainsKey(limpio))
                {
                    continue;
                }

                columnas[limpio] = letra;
            }

            var indiceComponente = -1;
            for (var i = 0; i < delBloque.Count; i++)
            {
                if (Igual(delBloque[i].A, CuadroComponente) && Igual(delBloque[i].B, "Total"))
                {
                    indiceComponente = i;
                    break;
                }
            }

            if (indiceComponente < 0)
            {
                indiceComponente = delBloque.Count; // sin frontera: todo el bloque es el primer cuadro
            }

            return new BloquePlantilla(nombre, delBloque, columnas, indiceComponente);
        }
    }
}
