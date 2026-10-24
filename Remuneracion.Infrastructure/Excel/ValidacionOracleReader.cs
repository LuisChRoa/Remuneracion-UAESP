using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Core.Services;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// HU-13 (2.7): lector-oráculo de validaciones cruzadas (D1) — OpenXML READ-ONLY (nunca
/// escribe; A4 hash intacto). Lee el caché &lt;v&gt; de las hojas de validación protegidas y
/// verifica la presencia de fórmula &lt;f&gt; donde T0 lo exige (W2/D7): hoja o celda ausente →
/// fallo que NOMBRA hoja+celda; nunca 0 silencioso.
///
/// HU-14 (3.1 + deuda HU-13): S-1 booleano estricto ("false" literal → FALSE en toda rama),
/// S-2 O9/P9 leídos UNA vez (fuera del loop por ASE; el validador gatea una vez + igualdad ×5),
/// S-4 caché &lt;v&gt; exigido en celdas de gate (<see cref="LeerValorNumericoExigido"/>) y
/// W-1 sub-bloques booleanos de VALIDACION_TOTAL (C15/D25/O25/D34/F34) leídos por ASE.
///
/// El validador de dominio nunca abre .xlsx; este reader vive en Infrastructure y produce
/// <see cref="ValidacionCruzadaSnapshot"/> por ASE (matcheo estricto por Ase.Id).
/// </summary>
public sealed class ValidacionOracleReader : IValidacionOracleReader
{
    public IReadOnlyList<ValidacionCruzadaSnapshot> LeerSnapshots(string rutaWorkbook, Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(periodo);

        if (string.IsNullOrWhiteSpace(rutaWorkbook) || !File.Exists(rutaWorkbook))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.FuenteNoEncontrada, $"No se encontró el workbook para el oráculo de validaciones: '{rutaWorkbook}'.");
        }

        var snapshots = new List<ValidacionCruzadaSnapshot>(5);

        using (var workbook = SpreadsheetDocument.Open(rutaWorkbook, false))
        {
            var workbookPart = workbook.WorkbookPart
                ?? throw new CalculoInvalidoException(CodigoError.FormatoFuente, "El workbook del oráculo no tiene WorkbookPart válido.");

            // Plan 30 (D-A): el nombre de la hoja se compone del período (no de un literal).
            var hojaDetValiRetri = NombresHojaPeriodo.DetValiRetri(periodo);

            // W2: la hoja DetValiRetri y la fila Total (D21/D29) se verifican UNA vez (compartidas).
            // Plan 30 (T2, D-C/R-C-3): si la hoja DetValiRetri del período no existe, el fail-fast
            // es ERR-PLANTILLA y nombra hoja esperada + período + operación (sin fallback).
            _ = ObtenerHoja(workbookPart, hojaDetValiRetri, "oráculo DetValiRetri D21/D29", CodigoError.Plantilla, periodo.CodigoCompleto);
            _ = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, "D21", "oráculo");
            _ = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, "D29", "oráculo");

            // HU-14 (S-2, D8): VALIDACION_TOTAL O9/P9 se leen UNA vez (fuera del loop por ASE)
            // y se pueblan los 5 snapshots con el mismo valor; el validador gatea una vez
            // (snapshot[0]) y aserta igualdad ×5 (divergencia = bug del reader, fail-fast).
            var celdaTotalO = ObtenerCeldaFormula(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "O9", "oráculo");
            var celdaTotalP = ObtenerCeldaFormula(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "P9", "oráculo");
            var valorTotalO = LeerValorNumericoExigido(celdaTotalO, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "O9");
            var valorTotalP = LeerValorBooleano(celdaTotalP, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, "P9");

            // Plan 40 (T1, R-E-1/R-E-2 — D-A/D-D): el resolutor acotado de valores se construye UNA
            // vez (cachea hojas/celdas/maestros compartidos) y lo comparten los 5 snapshots.
            var resolutor = new ResolutorValorValidacion(workbookPart);

            for (var aseId = 1; aseId <= 5; aseId++)
            {
                snapshots.Add(LeerSnapshotAse(workbookPart, aseId, periodo, hojaDetValiRetri, valorTotalO, valorTotalP, resolutor));
            }
        }

        return snapshots;
    }

    private static ValidacionCruzadaSnapshot LeerSnapshotAse(
        WorkbookPart workbookPart,
        int aseId,
        Periodo periodo,
        string hojaDetValiRetri,
        decimal valorTotalO,
        bool valorTotalP,
        ResolutorValorValidacion resolutor)
    {
        var ase = CrearAse(aseId);
        var porEmpresa = new List<ValidacionEmpresaSnapshot>(5);
        var fila = WorkbookLeafCellMapValidaciones.FilaValidacionPorAse(aseId);

        foreach (var empresa in EmpresaFacturacion.Catalogo.OrderBy(e => e.Id))
        {
            // W2: celda O/P de la fila-ASE de la empresa debe existir y ser fórmula.
            var celdaO = ObtenerCeldaFormula(workbookPart, empresa.HojaValidacion, $"O{fila}", "oráculo");
            var celdaP = ObtenerCeldaFormula(workbookPart, empresa.HojaValidacion, $"P{fila}", "oráculo");

            porEmpresa.Add(new ValidacionEmpresaSnapshot
            {
                Empresa = empresa.Nombre,
                AseId = aseId,
                DiferenciaO = LeerValorNumericoExigido(celdaO, empresa.HojaValidacion, $"O{fila}"),
                VerificacionP = LeerValorBooleano(celdaP, empresa.HojaValidacion, $"P{fila}")
            });
        }

        // DetValiRetri: diferencias D(15+ase) y verificaciones D(23+ase); totales D21/D29 ya
        // verificados en presencia (W2) a nivel hoja.
        var diferencias = new List<DetValiRetriCeldaValor>(1);
        var filaDif = WorkbookLeafCellMapValidaciones.FilaDiferenciaDetValiRetri(aseId);
        var celdaDiferencia = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, $"D{filaDif}", "oráculo");
        diferencias.Add(new DetValiRetriCeldaValor
        {
            Celda = $"D{filaDif}",
            Valor = LeerValorNumericoExigido(celdaDiferencia, hojaDetValiRetri, $"D{filaDif}")
        });

        var verificaciones = new List<DetValiRetriCeldaVerificacion>(1);
        var filaVerif = WorkbookLeafCellMapValidaciones.FilaVerificacionDetValiRetri(aseId);
        var celdaVerificacion = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, $"D{filaVerif}", "oráculo");
        verificaciones.Add(new DetValiRetriCeldaVerificacion
        {
            Celda = $"D{filaVerif}",
            Verificacion = LeerValorBooleano(celdaVerificacion, hojaDetValiRetri, $"D{filaVerif}")
        });

        var celdaD21 = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, "D21", "oráculo");
        var celdaD29 = ObtenerCeldaFormula(workbookPart, hojaDetValiRetri, "D29", "oráculo");

        var detValiRetri = new DetValiRetriSnapshot
        {
            Ase = ase,
            DiferenciasAse = diferencias,
            VerificacionesAse = verificaciones,
            DiferenciaTotalD21 = LeerValorNumericoExigido(celdaD21, hojaDetValiRetri, "D21"),
            VerificacionTotalD29 = LeerValorBooleano(celdaD29, hojaDetValiRetri, "D29")
        };

        // Plan 40 (T1, R-E-1/R-E-2 — D-A/D-D; enmienda E-1): los 5 sub-bloques booleanos de
        // VALIDACION_TOTAL (C15/D25/O25/D34/F34) YA NO se leen del caché <v> (falsa verificación
        // heredada). Se aserta el TEXTO <f> esperado (tabla FormasValidacionTotal, congelada por el
        // T0 del Plan 40) y se EVALÚA la igualdad de los dos operandos resueltos desde las celdas
        // efectivamente escritas (nunca desde el caché de una celda de fórmula) — incluida la cadena
        // del espejo de remuneración. Tercera forma de texto = ERR-PLANTILLA nombrando
        // hoja+celda+esperado+observado. El gate TRUE exacto por ASE vive en el validador (que
        // re-asegura el texto y evalúa en Core).
        var subBloques = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        var evidencia = new Dictionary<string, SubBloqueValidacionTotalEvidencia>(StringComparer.OrdinalIgnoreCase);
        foreach (var celdaSub in WorkbookLeafCellMapValidaciones.SubBloquesValidacionTotal)
        {
            _ = ObtenerCeldaFormula(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidacionTotal, celdaSub, "oráculo");
            var textoEfectivo = resolutor.TextoFormulaEfectiva(WorkbookLeafCellMapValidaciones.HojaValidacionTotal, celdaSub);
            var textoObservado = NormalizarFormula(textoEfectivo);
            var textoEsperado = FormasValidacionTotal.TextoEsperado(celdaSub) ?? string.Empty;
            if (!string.Equals(textoObservado, textoEsperado, StringComparison.Ordinal))
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"{WorkbookLeafCellMapValidaciones.HojaValidacionTotal}!{celdaSub}: la fórmula <f> no coincide con la tabla del gate. Esperado: '{textoEsperado}'. Observado: '{textoEfectivo}'. (R-E-2: tercera forma = ERR-PLANTILLA.)");
            }

            var (computable, izquierdo, derecho) = ResolverOperandos(resolutor, textoObservado);
            evidencia[celdaSub] = new SubBloqueValidacionTotalEvidencia
            {
                Celda = celdaSub,
                TextoFormula = textoObservado,
                Computable = computable,
                OperandoIzquierdo = izquierdo,
                OperandoDerecho = derecho
            };

            // Plan 40 (E-1, E1-1): los 5 sub-bloques se EVALÚAN siempre desde lo escrito (prohibido
            // SeGatea=false). Un FALSE es divergencia real (hallazgo E1-2), nunca "no evaluable".
            subBloques[celdaSub] = computable && FormasValidacionTotal.Igualdad(izquierdo, derecho);
        }

        // Controles Valida -* (informativo; sin gate numérico — T0-0.5; S-4 conserva el 0).
        var controles = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase)
        {
            ["Valida -Remunera!D9"] = LeerValorNumericoSiExiste(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidaRemunera, "D9"),
            ["Valida - Anticipos!D6"] = LeerValorNumericoSiExiste(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidaAnticipos, "D6"),
            ["Valida - Control Recaudo!F10"] = LeerValorNumericoSiExiste(workbookPart, WorkbookLeafCellMapValidaciones.HojaValidaControlRecaudo, "F10")
        };

        return new ValidacionCruzadaSnapshot
        {
            Ase = ase,
            PorEmpresa = porEmpresa,
            DetValiRetri = detValiRetri,
            ValidacionTotal = valorTotalO,
            ValidacionTotalOkP = valorTotalP,
            Controles = controles,
            SubBloquesValidacionTotal = subBloques,
            SubBloquesValidacionTotalEvidencia = evidencia,
            ExtempResuelto = ResolverExtempConsolidado(resolutor, aseId),
            EvidenciaAnclasExtemp = ResolverAnclasExtempR1(resolutor, aseId)
        };
    }

    /// <summary>
    /// Plan 49 (A1-núcleo, §2.2): EXTEMP del ASE resuelto desde lo escrito en
    /// <c>CONSOLIDADO_TOTAL RECAUDO!D{46+aseId}</c>. El resolutor sigue el texto-<c>&lt;f&gt;</c>
    /// (nunca el caché <c>&lt;v&gt;</c>): en la v2, <c>D51 → 'Reporte Componentes R1'!F682</c> y
    /// <c>F682 = F660+F629-L629</c>; el valor sale de evaluar la cadena completa desde lo escrito.
    /// Celda ausente o cadena no resoluble = <c>ERR-PLANTILLA</c> nombrando hoja+celda (nunca 0
    /// silencioso). El reader NO juzga: solo expone el valor para que el validador lo compare con el
    /// dominio.
    /// </summary>
    private static decimal ResolverExtempConsolidado(ResolutorValorValidacion resolutor, int aseId)
    {
        var hoja = PublicacionExtempConsolidado.HojaConsolidado;
        var celda = PublicacionExtempConsolidado.CeldaConsolidadoExtemp(aseId);

        if (!resolutor.ExisteCelda(hoja, celda))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ASE {aseId} · {hoja} · EXTEMPORANEO: la celda '{celda}' del bloque no existe en el workbook. (Plan 49 A1: el gate absoluto necesita la celda del CONSOLIDADO; sin ella no hay 0 silencioso.)");
        }

        if (!resolutor.TryEvaluar(hoja, celda, out var valor))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ASE {aseId} · {hoja} · EXTEMPORANEO: la celda '{celda}' del bloque no es resoluble desde lo escrito. (Plan 49 A1: tercera forma = ERR-PLANTILLA, nunca 0 silencioso.)");
        }

        return valor;
    }

    /// <summary>
    /// Plan 49 (A1-interior, §2.3) + Plan 57 (§2-D2/D3): evidencia de las anclas EXTEMPORANEO del
    /// interior del bloque, <b>derivada de la plantilla en runtime</b> — sin tabla de anclas, sin
    /// direcciones congeladas y sin ramas por quincena.
    ///
    /// <para><b>Cómo se ubica el ancla (§2-D3.1).</b> La celda de publicación del bloque
    /// (<c>CONSOLIDADO!D{46+aseId}</c>) declara en su <c>&lt;f&gt;</c> UNA referencia a la hoja de
    /// componentes; ésa —y no una lista— es la celda-ancla. Se lee del workbook, así que un cambio de
    /// plantilla (otra celda, otra fila) se absorbe sin tocar código.</para>
    ///
    /// <para><b>Familia secundaria (§2-D3.2).</b> Las hojas publicadoras (<c>REMUNERACION_*</c>)
    /// declaran el EXTEMP del MISMO bloque en la <b>misma fila</b> de publicación; cada una se
    /// reconoce por la referencia que escribe (no por su nombre), y su destino entra en la evidencia.
    /// </para>
    ///
    /// <para><b>Naturaleza en runtime (§2-D2.1).</b> Si la celda-ancla trae <c>&lt;f&gt;</c> se evalúa
    /// la fórmula leída con los valores del propio workbook (su descomposición ES la expresión de
    /// secciones); si es literal se lee. Sin ramas por bloque ni por quincena.</para>
    ///
    /// <para><b>Lo que no es (§2-D4.2).</b> Esto es evidencia diagnóstica del interior, no el veredicto:
    /// el núcleo compara lo visible del workbook contra el esperado del dominio
    /// (<see cref="WorkbookLeafInputsR1.ExtemporaneoEsperadoPorAse"/>) y no se ve afectado por esta
    /// evaluación.</para>
    ///
    /// <para><b>Irresoluble (§2-D5).</b> Referencia ausente o no única, destino inexistente o fórmula
    /// que no evalúa desde lo escrito ⇒ <c>ERR-PLANTILLA</c> nombrando ASE + hoja + concepto; el reader
    /// es read-only y no se escribe nada. Un workbook que no declara la publicación del bloque (p. ej.
    /// el manual legacy, usado solo como fixture de lectura) no tiene ancla que verificar: la lista
    /// queda vacía, nunca un 0 silencioso.</para>
    /// </summary>
    private static IReadOnlyList<AnclaExtempR1Evidencia> ResolverAnclasExtempR1(
        ResolutorValorValidacion resolutor,
        int aseId)
    {
        var hojaConsolidado = PublicacionExtempConsolidado.HojaConsolidado;
        var celdaPublicacion = PublicacionExtempConsolidado.CeldaConsolidadoExtemp(aseId);
        var textoPublicacion = resolutor.TextoFormulaEfectiva(hojaConsolidado, celdaPublicacion);

        if (string.IsNullOrWhiteSpace(textoPublicacion))
        {
            // El workbook no publica el EXTEMP de este bloque por referencia (geometría ausente).
            return [];
        }

        if (!resolutor.TryReferenciaUnica(hojaConsolidado, textoPublicacion, out var hojaAncla, out var celdaAncla))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ASE {aseId} · {hojaConsolidado} · EXTEMPORANEO: la celda '{celdaPublicacion}' del bloque no declara una única celda de referencia ('{textoPublicacion}'), así que no se puede ubicar el ancla del bloque.");
        }

        if (!resolutor.ExisteCelda(hojaAncla, celdaAncla))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ASE {aseId} · {hojaAncla} · EXTEMPORANEO: la celda '{celdaPublicacion}' del bloque declara el ancla '{celdaAncla}', que no existe en la hoja. (Sin ancla no hay evidencia; nunca 0 silencioso.)");
        }

        var destinos = new List<(string Hoja, string Celda)> { (hojaAncla, celdaAncla) };
        destinos.AddRange(DestinosPublicadosDelBloque(resolutor, aseId, hojaAncla));

        var evidencias = new List<AnclaExtempR1Evidencia>(destinos.Count);
        foreach (var (hoja, celda) in destinos.Distinct())
        {
            evidencias.Add(new AnclaExtempR1Evidencia
            {
                Hoja = hoja,
                Celda = celda,
                ValorAncla = ResolverValorDelAncla(resolutor, aseId, hoja, celda, out var expresionSecciones, out var valorSecciones),
                ExpresionSecciones = expresionSecciones,
                ValorSecciones = valorSecciones
            });
        }

        return evidencias;
    }

    /// <summary>
    /// Plan 57 (§2-D3.2): destinos R1 que las hojas PUBLICADORAS declaran para el bloque en su fila de
    /// publicación. Se reconocen por la referencia que escriben —nunca por el nombre de la hoja— y se
    /// descarta la propia hoja del ancla (ahí viven las anclas, no sus consumidores).
    /// </summary>
    private static IEnumerable<(string Hoja, string Celda)> DestinosPublicadosDelBloque(
        ResolutorValorValidacion resolutor,
        int aseId,
        string hojaAncla)
    {
        var celdaPublicacion = $"D{PublicacionExtempConsolidado.FilaConsolidadoExtemp(aseId)}";

        foreach (var nombre in resolutor.NombresHojas())
        {
            if (string.Equals(nombre, hojaAncla, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var texto = resolutor.TextoFormulaEfectiva(nombre, celdaPublicacion);
            if (!resolutor.TryReferenciaUnica(nombre, texto, out var hojaDestino, out var celdaDestino)
                || !string.Equals(hojaDestino, hojaAncla, StringComparison.OrdinalIgnoreCase)
                || !resolutor.ExisteCelda(hojaDestino, celdaDestino))
            {
                continue;
            }

            yield return (hojaDestino, celdaDestino);
        }
    }

    /// <summary>
    /// Plan 57 (§2-D2.1): el valor del ancla sale de su NATURALEZA leída en el workbook. Con
    /// <c>&lt;f&gt;</c> se evalúa la fórmula leída (que es, a la vez, la descomposición declarada y la
    /// expresión de secciones); literal se lee. Irresoluble ⇒ <c>ERR-PLANTILLA</c> con ASE + hoja +
    /// concepto (nunca 0 silencioso).
    /// </summary>
    private static decimal ResolverValorDelAncla(
        ResolutorValorValidacion resolutor,
        int aseId,
        string hoja,
        string celda,
        out string expresionSecciones,
        out decimal? valorSecciones)
    {
        expresionSecciones = string.Empty;
        valorSecciones = null;

        var textoFormula = resolutor.TextoFormulaEfectiva(hoja, celda);
        if (string.IsNullOrWhiteSpace(textoFormula))
        {
            if (!resolutor.TryEvaluar(hoja, celda, out var literal))
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"ASE {aseId} · {hoja} · EXTEMPORANEO: el ancla '{celda}' del bloque no es resoluble desde lo escrito.");
            }

            return literal;
        }

        if (!resolutor.TryEvaluar(hoja, textoFormula, out var secciones))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ASE {aseId} · {hoja} · EXTEMPORANEO: la fórmula del ancla '{celda}' ('{textoFormula}') no es resoluble desde lo escrito; no se puede verificar la descomposición del bloque.");
        }

        if (!resolutor.TryEvaluar(hoja, celda, out var valor))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ASE {aseId} · {hoja} · EXTEMPORANEO: el ancla '{celda}' del bloque no es resoluble desde lo escrito.");
        }

        expresionSecciones = textoFormula;
        valorSecciones = secciones;
        return valor;
    }

    /// <summary>
    /// Plan 40 (T1): parte el texto-<c>&lt;f&gt;</c> del sub-bloque en sus dos operandos (igualdad)
    /// y los resuelve desde lo escrito. <c>Computable=false</c> si alguno no es resoluble por las
    /// formas tabuladas (no se evalúa en silencio).
    /// </summary>
    private static (bool Computable, decimal Izquierdo, decimal Derecho) ResolverOperandos(
        ResolutorValorValidacion resolutor,
        string textoFormula)
    {
        var partes = textoFormula.Split('=');
        if (partes.Length != 2)
        {
            return (false, 0m, 0m);
        }

        var okIzquierdo = resolutor.TryEvaluar(WorkbookLeafCellMapValidaciones.HojaValidacionTotal, partes[0], out var izquierdo);
        var okDerecho = resolutor.TryEvaluar(WorkbookLeafCellMapValidaciones.HojaValidacionTotal, partes[1], out var derecho);
        return (okIzquierdo && okDerecho, izquierdo, derecho);
    }

    /// <summary>
    /// Plan 40 (T1): normaliza el texto de una fórmula para compararlo con la tabla del gate
    /// (elimina espacios fuera de nombres de hoja entre comillas; preserva el texto entre comillas).
    /// </summary>
    private static string NormalizarFormula(string? texto)
    {
        if (string.IsNullOrEmpty(texto))
        {
            return string.Empty;
        }

        var sb = new StringBuilder(texto.Length);
        var enComillas = false;
        foreach (var ch in texto)
        {
            if (ch == '\'')
            {
                enComillas = !enComillas;
            }

            if (!enComillas && char.IsWhiteSpace(ch))
            {
                continue;
            }

            sb.Append(ch);
        }

        return sb.ToString();
    }

    // ── Helpers OpenXML (read-only; duplican navegación mínima del writer, sin escritura) ─────

    private static Cell ObtenerCeldaFormula(WorkbookPart workbookPart, string hoja, string celda, string operacion)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"La celda-oráculo '{hoja}!{celda}' no existe. (W2: sin asserts no hay snapshot; nunca 0 silencioso.)");

        if (cell.CellFormula is null)
        {
            throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"La celda-oráculo '{hoja}!{celda}' debería ser fórmula protegida y se detectó un valor fijo. (W2)");
        }

        return cell;
    }

    /// <summary>
    /// Lectura de controles informativos <c>Valida -*</c> (S-4): celda ausente o sin caché →
    /// 0 tolerado documentado (nunca gate). Solo un valor no numérico presente es fallo.
    /// </summary>
    private static decimal LeerValorNumericoSiExiste(WorkbookPart workbookPart, string hoja, string celda)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, "oráculo");
        var cell = worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase));
        if (cell is null || cell.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
        {
            return 0m; // control sin valor o sin caché en esta plantilla (informativo)
        }

        if (decimal.TryParse(cell.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor))
        {
            return valor;
        }

        throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"La celda-oráculo '{hoja}!{celda}' tiene un valor no numérico: '{cell.CellValue.InnerText}'.");
    }

    /// <summary>
    /// HU-14 (S-4): lectura EXIGIDA para celdas de gate. <c>&lt;v&gt;</c> ausente (fórmula sin
    /// caché = workbook nunca recalculado) → fallo que NOMBRA hoja+celda; nunca 0 silencioso
    /// en gates (doctrina W2 HU-13, extendida por S-4).
    /// </summary>
    private static decimal LeerValorNumericoExigido(Cell cell, string hoja, string celda)
    {
        if (cell.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
        {
            throw new CalculoInvalidoException(
                CodigoError.FormatoFuente,
                $"La celda-oráculo '{hoja}!{celda}' es fórmula sin caché (<v> ausente); recalcule el workbook en Excel y reintente. (S-4: nunca 0 silencioso en gates.)");
        }

        if (decimal.TryParse(cell.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor))
        {
            return valor;
        }

        throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"La celda-oráculo '{hoja}!{celda}' tiene un valor no numérico: '{cell.CellValue.InnerText}'.");
    }

    /// <summary>
    /// HU-14 (S-1): lectura booleana ESTRICTA.
    /// "1"/"true" → TRUE y "0"/"false" → FALSE en TODAS las ramas (case-insensitive; el
    /// literal "false" YA NO se interpreta como TRUE — S-1).
    /// - Con t="b": cualquier otro texto es serialización booleana inválida → fail-fast nombrado.
    /// - Sin t="b": numérico ≠ 0 → TRUE; vacío → FALSE (caché ausente; conserva S-4 para
    ///   informativos); resto → fail-fast nombrado (hoja+celda+valor).
    /// </summary>
    private static bool LeerValorBooleano(Cell cell, string hoja, string celda)
    {
        if (cell.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
        {
            return false; // caché ausente → FALSE (el gate lo reporta nombrando la celda)
        }

        var texto = cell.CellValue.InnerText.Trim();

        if (string.Equals(texto, "1", StringComparison.OrdinalIgnoreCase)
            || string.Equals(texto, "true", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (string.Equals(texto, "0", StringComparison.OrdinalIgnoreCase)
            || string.Equals(texto, "false", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (cell.DataType is not null && cell.DataType.Value == CellValues.Boolean)
        {
            throw new CalculoInvalidoException(
                CodigoError.FormatoFuente,
                $"La celda-oráculo '{hoja}!{celda}' tiene un booleano t=\"b\" no reconocido: '{texto}' (se esperaba 1/0 o true/false). (S-1)");
        }

        if (decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor))
        {
            return valor != 0m;
        }

        throw new CalculoInvalidoException(
            CodigoError.FormatoFuente,
            $"La celda-oráculo '{hoja}!{celda}' tiene un valor booleano no reconocido: '{texto}' (se esperaba 1/0 o true/false). (S-1)");
    }

    private static Worksheet ObtenerHoja(WorkbookPart workbookPart, string nombreHoja, string operacion, string codigo = CodigoError.FormatoFuente, string? codigoPeriodo = null)
    {
        var sufijoPeriodo = string.IsNullOrWhiteSpace(codigoPeriodo) ? string.Empty : $" (período {codigoPeriodo})";
        var workbook = workbookPart.Workbook ?? throw new CalculoInvalidoException(CodigoError.FormatoFuente, $"El workbook para {operacion} no tiene metadata Workbook válida.");
        var sheet = workbook.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(codigo, $"La hoja-oráculo '{nombreHoja}' no existe en el workbook para {operacion}{sufijoPeriodo}. (W2: sin asserts no hay snapshot.)");

        var worksheetPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart
            ?? throw new CalculoInvalidoException(codigo, $"No se pudo resolver la hoja '{nombreHoja}' en el workbook para {operacion}{sufijoPeriodo}.");

        return worksheetPart.Worksheet ?? throw new CalculoInvalidoException(codigo, $"La hoja '{nombreHoja}' no tiene Worksheet válido para {operacion}{sufijoPeriodo}.");
    }

    /// <summary>
    /// HU-14 (S-3): delega a <see cref="AseFactory.DesdeId"/> (fuente única desde
    /// <see cref="Constants.CarpetasAse.Prefijos"/>); el array hardcodeado desaparece.
    /// </summary>
    private static Ase CrearAse(int id) => AseFactory.DesdeId(id);
}
