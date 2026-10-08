using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Remuneracion.Core.Constants;
using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;

namespace Remuneracion.Infrastructure.Excel;

/// <summary>
/// Writer OpenXML de la plantilla real.
/// HU-04: validación estructural no destructiva (<see cref="IPlantillaWriter"/>).
/// HU-05: escritura real de celdas leaf sobre una copia (<see cref="IWorkbookLeafWriter"/>).
/// HU-07: overload multi-ASE con mapa por bloque (<see cref="WorkbookLeafCellMapPorAse"/>),
/// una sola copia, validación pre/post ampliada y borrado de parcial ante fallo.
/// HU-11 (2.5): en Q2 escribe los operandos editables de SALDOS POR NOTA / RETRIBUCION NEGATIVA
/// en la MISMA pasada atómica (D2a: T0-0.7 demostró bloques de valores editables) y amplía la
/// validación protegida a la cadena AJUSTES-SF-T (mapa <see cref="WorkbookLeafCellMapAjustesSfT"/>
/// + HU-10 parametrizado al sufijo de hoja 2026072).
/// HU-12 (2.6 ampliada): en Q2 escribe los leafs R1/R2/R4-Q2 (mapa <see cref="WorkbookLeafCellMapQ2"/>,
/// con variante ASE5 de 2 filas V0.3) + DetRetri-Q2 (V0.4: ROUND(D104:D108,0) vía
/// <see cref="DetRetriRounder"/>) en la MISMA pasada; la validación protegida se parametriza por
/// período (D5): Q1 exige las fórmulas HU-07, Q2 exige el mapa T0 (F53…, D73…, DetValiRetri/
/// VALIDACION_*/INTERVENTORIA/ANT EXT-REV protegidas) y M1 queda ejercitado contra el canónico.
/// </summary>
public class OpenXmlPlantillaWriter : IPlantillaWriter, IWorkbookLeafWriter, IEspejoR1Writer
{
    private const string HojaConsolidado = WorkbookLeafCellMap.HojaConsolidado;
    private const string HojaR1 = WorkbookLeafCellMap.HojaR1;
    private const string HojaR2 = WorkbookLeafCellMap.HojaR2;
    private const string HojaR4 = WorkbookLeafCellMap.HojaR4;
    private const string HojaBanco = WorkbookLeafCellMapReporteBanco.HojaBanco;
    private const string HojaBce = WorkbookLeafCellMapBalanceSc.HojaBce;
    private const string HojaSaldosNotas = WorkbookLeafCellMapAjustesSfT.HojaSaldosNotas;
    private const string HojaRetribucionNegativa = WorkbookLeafCellMapAjustesSfT.HojaRetribucionNegativa;
    /// <summary>
    /// Ancla-Q1 del naming (Plan 30, D-A): el mapa BCE (<see cref="WorkbookLeafCellMapBalanceSc.Protegidas"/>)
    /// está declarado con los nombres del canónico Q1; el writer sustituye esta ancla por
    /// <c>periodo.CodigoCompleto</c> para resolver el nombre REAL del período (para 2026071 la
    /// sustitución es identidad; cero literales de período en el runtime).
    /// </summary>
    private const string AnclaSufijoHojasQ1 = "2026071";

    public void EscribirConsolidado(string rutaPlantilla, ResultadoRemuneracion resultado)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(resultado);

        var workbookPath = ValidarArchivo(rutaPlantilla, nameof(EscribirConsolidado));
        using var workbook = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = workbook.WorkbookPart ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene WorkbookPart válido.");
        var worksheet = ObtenerHoja(workbook, HojaConsolidado, nameof(EscribirConsolidado));

        ValidarConsolidadoFormulario(workbookPart, worksheet);
        ValidarCadenaFormulaR1R2R4(workbookPart);
    }

    public void EscribirDetalleR1(string rutaPlantilla, Ase ase, RecaudoComponenteR1 datos)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(datos);

        var workbookPath = ValidarArchivo(rutaPlantilla, nameof(EscribirDetalleR1));
        using var workbook = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = workbook.WorkbookPart ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene WorkbookPart válido.");
        var worksheet = ObtenerHoja(workbook, HojaR1, nameof(EscribirDetalleR1));

        ValidarHojaConLabelsEsperados(workbookPart, HojaR1, ["Componente", "Total", "Mes"]);
        ValidarCeldaTieneFormula(workbookPart, worksheet, "F46", ["F25", "F41", "L25"], HojaR1, nameof(EscribirDetalleR1));
        ValidarCeldaTieneFormula(workbookPart, worksheet, "F48", ["F30", "F10", "L10"], HojaR1, nameof(EscribirDetalleR1));

        LanzarPayloadInsuficiente(nameof(EscribirDetalleR1), HojaR1,
            "Se requiere un contrato de WorkbookLeafInputs que exponga los inputs reales detrás de F25,F41,L25,F30,F10,L10; el payload actual solo expone agregados TotOpt/Extemp y detalle parcial.");
    }

    public void EscribirDetalleR2(string rutaPlantilla, Ase ase, SaldosFavorR2 datos)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(datos);

        var workbookPath = ValidarArchivo(rutaPlantilla, nameof(EscribirDetalleR2));
        using var workbook = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = workbook.WorkbookPart ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene WorkbookPart válido.");
        var worksheet = ObtenerHoja(workbook, HojaR2, nameof(EscribirDetalleR2));

        ValidarHojaConLabelsEsperados(workbookPart, HojaR2, ["Total", "Componente TDF"]);
        ValidarCeldaTieneFormula(workbookPart, worksheet, "E41", ["E15", "E26", "K15"], HojaR2, nameof(EscribirDetalleR2));

        LanzarPayloadInsuficiente(nameof(EscribirDetalleR2), HojaR2,
            "Se requiere un contrato de WorkbookLeafInputs que exponga los inputs reales detrás de E15,E26,K15; el modelo actual solo expone GrandTotal y ServEspK.");
    }

    public void EscribirDetalleR4(string rutaPlantilla, Ase ase, ReversionR4 datos)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantilla);
        ArgumentNullException.ThrowIfNull(ase);
        ArgumentNullException.ThrowIfNull(datos);

        var workbookPath = ValidarArchivo(rutaPlantilla, nameof(EscribirDetalleR4));
        using var workbook = SpreadsheetDocument.Open(workbookPath, false);
        var workbookPart = workbook.WorkbookPart ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene WorkbookPart válido.");
        var worksheet = ObtenerHoja(workbook, HojaR4, nameof(EscribirDetalleR4));

        ValidarCeldaTieneFormula(workbookPart, worksheet, "D67", ["D9", "P9"], HojaR4, nameof(EscribirDetalleR4));

        LanzarPayloadInsuficiente(nameof(EscribirDetalleR4), HojaR4,
            "Se requiere un contrato de WorkbookLeafInputs que exponga los inputs reales detrás de D9 y P9; el modelo actual solo expone TotalReversiones agregado.");
    }

    /// <inheritdoc />
    public void GenerarWorkbook(string rutaPlantillaOrigen, string rutaSalida, ResultadoRemuneracion resultado, WorkbookLeafInputs leafInputs)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantillaOrigen);
        ArgumentNullException.ThrowIfNull(rutaSalida);
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leafInputs);

        var origen = ValidarArchivo(rutaPlantillaOrigen, nameof(GenerarWorkbook));
        ValidarRutaSalida(rutaSalida);
        ValidarNoInPlace(origen, rutaSalida);

        WorkbookLeafCoherence.ValidarContraResultado(leafInputs, resultado);

        // HU-15 (W-2.3, D5): copia + escritura dentro del try de atomicidad — un fallo de I/O
        // (IOException) se envuelve en ERR-ESCRITURA con InnerException preservada y el parcial
        // se borra. Las excepciones de dominio (ERR-PLANTILLA de estructura, ERR-VALIDACION de
        // coherencia) se re-lanzan tal cual tras borrar el parcial.
        try
        {
            var directorioSalida = Path.GetDirectoryName(rutaSalida);
            if (!string.IsNullOrWhiteSpace(directorioSalida))
            {
                Directory.CreateDirectory(directorioSalida);
            }

            File.Copy(origen, rutaSalida, overwrite: true);

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, true))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene WorkbookPart válido.");
                ValidarFormulasProtegidas(workbookPart, nameof(GenerarWorkbook));
                EscribirCeldasLeaf(workbookPart, leafInputs);

                // Plan 29 (T3, Unidad R): desglose-detalle R2/R4 en la MISMA pasada atómica.
                EscribirCeldasDetalleR2R4(workbookPart, leafInputs);

                // Plan 29 (T5, Unidad P — R-P-3/T0f): total de control de recaudo (F10).
                EscribirControlRecaudoTotal(workbookPart, [leafInputs]);

                // Plan 28 (Unidad F): sella las fechas del período desde el R10 (solo si el
                // orquestador las aportó). El path single-ASE sin R10 no trae fechas → no-op.
                EscribirFechasPeriodo(workbookPart, resultado);

                var workbookXml = workbookPart.Workbook
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene metadata Workbook válida.");

                // Plan 28 (Unidad S): punto único de guardado — saneamiento calcChain + fullCalcOnLoad.
                SaneadorCadenaCalculo.Sanear(workbookPart);
                workbookXml.Save();
            }

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, false))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook generado no tiene WorkbookPart válido.");
                ValidarFormulasProtegidas(workbookPart, nameof(GenerarWorkbook));
            }
        }
        catch (Exception ex)
        {
            // HU-15 (W-2.3, D5): atomicidad intacta (borrar parcial) + causa preservada.
            try
            {
                if (File.Exists(rutaSalida))
                {
                    File.Delete(rutaSalida);
                }
            }
            catch
            {
                // Best-effort: si el archivo quedó bloqueado (antivirus/handle), no enmascarar
                // la causa real con un error de borrado.
            }

            if (ex is ArchivoFuenteNoEncontradoException or CalculoInvalidoException)
            {
                throw;
            }

            throw new CalculoInvalidoException(
                CodigoError.Escritura,
                $"No se pudo generar el workbook de salida: {ex.Message}",
                ex);
        }
    }

    /// <inheritdoc />
    public void GenerarWorkbook(string rutaPlantillaOrigen, string rutaSalida, ResultadoRemuneracion resultado, IReadOnlyList<WorkbookLeafInputs> leafInputs)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantillaOrigen);
        ArgumentNullException.ThrowIfNull(rutaSalida);
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leafInputs);

        if (leafInputs.Count == 0)
        {
            throw new CalculoInvalidoException("La lista de leafs del período está vacía; no hay nada que escribir.");
        }

        var origen = ValidarArchivo(rutaPlantillaOrigen, nameof(GenerarWorkbook));
        ValidarRutaSalida(rutaSalida);
        ValidarNoInPlace(origen, rutaSalida);

        WorkbookLeafCoherence.ValidarContraResultadoMultiAse(leafInputs, resultado);

        // HU-08 (2.2): gate Σ empresas = visible de bloque por ASE antes de escribir (D4/G5).
        foreach (var leaf in leafInputs)
        {
            WorkbookLeafCoherence.ValidarSigmaEmpresas(leaf.Conciliacion, leaf);
        }

        // HU-11 (2.5) + Plan 30 (D-A): el nombre de las hojas DetRetri/DetValiRetri depende del
        // período y se COMPONE del dominio (periodo.CodigoCompleto), no de literales.
        var periodoReferencia = leafInputs[0].Periodo;
        var esQuincena2 = leafInputs.Any(l => l.Periodo.NumeroQuincena == 2);

        // Plan 21 (T4): si el orquestador adjuntó el bloque espejo R1 de los 5 ASE, la hoja
        // Reporte Componentes R1 se gobierna por el espejo (dimensionado + escritura por
        // encabezado); las celdas R1 del mapa absoluto HU-07/HU-12 quedan superseded dentro del
        // bloque y NO se escriben por el mapa. El resto de las hojas queda intacto.
        var bloquesEspejo = leafInputs
            .Where(l => l.EspejoR1 is not null)
            .Select(l => l.EspejoR1!)
            .OrderBy(b => b.Ase.Id)
            .ToList();
        var espejoR1 = bloquesEspejo.Count == 5;

        // HU-15 (W-2.3, D5): copia + escritura dentro del try de atomicidad — un fallo de I/O
        // (IOException) se envuelve en ERR-ESCRITURA con InnerException preservada y el parcial
        // se borra. Las excepciones de dominio (ERR-PLANTILLA de estructura, ERR-VALIDACION de
        // coherencia) se re-lanzan tal cual tras borrar el parcial.
        try
        {
            var directorioSalida = Path.GetDirectoryName(rutaSalida);
            if (!string.IsNullOrWhiteSpace(directorioSalida))
            {
                Directory.CreateDirectory(directorioSalida);
            }

            File.Copy(origen, rutaSalida, overwrite: true);

            // Plan 21 (T4): el gate de evidencia del espejo se computa ANTES de mutar el workbook
            // (estado real de la plantilla respecto de la forma de la fuente) y se reutiliza en la
            // validación post-escritura (CRITICAL #1: recomputarlo tras AjustarEnWorkbook siempre
            // daría false con Δ≠0 → validación por direcciones absolutas sobre un R1 desplazado).
            var espejoDesplazado = false;

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, true))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene WorkbookPart válido.");

                // Plan 21 (T4): el espejo se aplica ANTES de la escritura de hojas y solo cuando
                // hace falta (algún bloque difiere de la forma de la fuente). Con Δ=0 (julio) el
                // espejo reescribe el bloque ya dimensionado y las direcciones absolutas quedan
                // intactas; la validación protegida corre normal. Con Δ≠0 (agosto) las direcciones
                // R1 se desplazan y su validación por dirección fija queda superseded por el gate
                // estructural del espejo (T3), evitando falsos ERR-PLANTILLA.
                if (espejoR1)
                {
                    espejoDesplazado = OpenXmlEspejoR1Mutador.RequiereAjuste(workbookPart, bloquesEspejo);
                    OpenXmlEspejoR1Mutador.AjustarEnWorkbook(workbookPart, bloquesEspejo);
                }

                if (!espejoDesplazado)
                {
                    ValidarFormulasProtegidasMultiAse(workbookPart, nameof(GenerarWorkbook), periodoReferencia);
                }

                // Con Δ≠0 (espejo desplazado) NO se valida aquí por direcciones absolutas de julio
                // (R1 se movió). La red es el gate de forma recompuesta, que corre POST-escritura
                // (bloque de solo-lectura más abajo) porque el archivo sigue abierto en escritura.
                foreach (var leaf in leafInputs.OrderBy(l => l.Ase.Id))
                {
                    EscribirCeldasLeafPorAse(workbookPart, leaf, esQuincena2, omitirR1: espejoR1);
                    EscribirCeldasEmpresa(workbookPart, leaf, omitirR1: espejoR1);
                    EscribirCeldasBanco(workbookPart, leaf);
                    EscribirCeldasBalanceSc(workbookPart, leaf);
                    EscribirCeldasAjustesSfT(workbookPart, leaf);

                    // Plan 29 (T3, Unidad R): desglose-detalle R2/R4 en la MISMA pasada atómica
                    // (tras los agregados E43/D73 y el espejo R1; el guard anti-fórmula es la red).
                    EscribirCeldasDetalleR2R4(workbookPart, leaf);
                }

                // HU-12 (2.6 ampliada, V0.4): DetRetri-Q2 (enteros por ASE + total) en la MISMA
                // pasada. SOLO en Q2: en Q1 el DetRetri calculado es oráculo de validación contra
                // el R10 (HU-20/G3) y la hoja DetRetri2026071 queda con sus fórmulas protegidas
                // intactas (NUNCA se sobrescribe).
                if (esQuincena2)
                {
                    EscribirCeldasDetRetriQ2(workbookPart, leafInputs);
                }

                // Plan 29 (T5, Unidad P — R-P-3/T0f): total de control de recaudo (F10) como
                // literal en la MISMA pasada atómica, tras las validaciones/control.
                EscribirControlRecaudoTotal(workbookPart, leafInputs);

                // Plan 28 (Unidad F): sello de fechas del período desde el R10 del período.
                EscribirFechasPeriodo(workbookPart, resultado);

                var workbookXml = workbookPart.Workbook
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene metadata Workbook válida.");

                // Plan 28 (Unidad S): punto único de guardado — saneamiento calcChain + fullCalcOnLoad.
                SaneadorCadenaCalculo.Sanear(workbookPart);
                workbookXml.Save();
            }

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, false))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook generado no tiene WorkbookPart válido.");

                // CRITICAL #1 (auditoría PR3): reutilizar el resultado del PRIMER gate
                // (espejoDesplazado) en lugar de recomputar RequiereAjuste aquí. Este bloque corre
                // DESPUÉS de AjustarEnWorkbook, que ya mutó el workbook: recomputar devolvería
                // false (los bloques ya coinciden con la fuente) y dispararía la validación por
                // direcciones absolutas de julio sobre un R1 desplazado → ERR-PLANTILLA falso
                // garantizado con Δ≠0. El gate de evidencia es el estado ANTES de mutar.
                if (!espejoDesplazado)
                {
                    ValidarFormulasProtegidasMultiAse(workbookPart, nameof(GenerarWorkbook), periodoReferencia);
                }
                else
                {
                    // Plan 31 (T2, R-B-3/D-B): con el espejo desplazado la validación por
                    // direcciones absolutas de julio no aplica (R1 se movió); su sucesor valida la
                    // FORMA RECOMPUESTA de los visibles (texto-fórmula por firma contra las filas
                    // reales del bloque ya recompuesto por el mutador). Corre post-escritura, sobre
                    // el archivo cerrado, con lectura BCL (sin Excel/COM).
                    new ValidadorTotalesR1Workbook().ExigirFormaRecompuesta(rutaSalida, leafInputs);
                }
            }
        }
        catch (Exception ex)
        {
            // HU-15 (W-2.3, D5): atomicidad intacta (borrar parcial) + causa preservada.
            try
            {
                if (File.Exists(rutaSalida))
                {
                    File.Delete(rutaSalida);
                }
            }
            catch
            {
                // Best-effort: si el archivo quedó bloqueado (antivirus/handle), no enmascarar
                // la causa real con un error de borrado.
            }

            if (ex is ArchivoFuenteNoEncontradoException or CalculoInvalidoException)
            {
                throw;
            }

            throw new CalculoInvalidoException(
                CodigoError.Escritura,
                $"No se pudo generar el workbook de salida: {ex.Message}",
                ex);
        }
    }

    /// <summary>
    /// Plan 21 (T3, R-E-2/R-E-3/R-E-4): aplica el espejo estructural R1 sobre una COPIA de la
    /// plantilla: dimensiona cada bloque ASE a la forma de la fuente del período actual
    /// (inserta/borra filas preservando estilos y fórmulas), reancla las referencias A1 afectadas
    /// (fórmulas, rangos compartidos, celdas combinadas y nombres definidos) y escribe los valores
    /// por ENCABEZADO de columna. No recalcula aritmética de negocio.
    /// </summary>
    /// <inheritdoc />
    public void EscribirEspejoR1(string rutaPlantillaOrigen, string rutaSalida, IReadOnlyList<BloqueEspejoAseInputs> bloques)
    {
        _ = EscribirEspejoR1ConResultado(rutaPlantillaOrigen, rutaSalida, bloques);
    }

    /// <summary>
    /// Plan 21 (W-2, auditoría PR3): variante interna que devuelve la instrumentación del
    /// reanclaje (Δ y referencias reancladas por bloque) para que los tests aserten evidencia
    /// estructural. La API pública (<see cref="EscribirEspejoR1"/>) mantiene su firma <c>void</c>.
    /// </summary>
    internal EspejoR1MutacionResultado EscribirEspejoR1ConResultado(string rutaPlantillaOrigen, string rutaSalida, IReadOnlyList<BloqueEspejoAseInputs> bloques)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantillaOrigen);
        ArgumentNullException.ThrowIfNull(rutaSalida);
        ArgumentNullException.ThrowIfNull(bloques);

        try
        {
            return OpenXmlEspejoR1Mutador.Ajustar(rutaPlantillaOrigen, rutaSalida, bloques);
        }
        catch (Exception ex)
        {
            try
            {
                if (File.Exists(rutaSalida))
                {
                    File.Delete(rutaSalida);
                }
            }
            catch
            {
                // Best-effort: no enmascarar la causa real con un error de borrado del parcial.
            }

            if (ex is ArchivoFuenteNoEncontradoException or CalculoInvalidoException)
            {
                throw;
            }

            throw new CalculoInvalidoException(
                CodigoError.Escritura,
                $"No se pudo aplicar el espejo estructural R1: {ex.Message}",
                ex);
        }
    }

    private static void ValidarConsolidadoFormulario(WorkbookPart workbookPart, Worksheet worksheet)
    {
        var reglas = new Dictionary<string, string[]>
        {
            ["D9"] = ["F46", "Reporte Componentes R1"],
            ["D28"] = ["E41", "Rem. Anticipos R2"],
            ["D47"] = ["F48", "Reporte Componentes R1"],
            ["D66"] = ["D67", "Reversion Pagos R4"],
            ["D85"] = ["D47", "AJUSTES"],
            ["D104"] = ["D9", "D28", "D47", "D66", "D85"],
            ["D109"] = ["D104", "D108", "SUM"]
        };

        foreach (var (celda, fragmentos) in reglas)
        {
            var formula = ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, HojaConsolidado, nameof(EscribirConsolidado));
            if (string.IsNullOrWhiteSpace(formula))
            {
                // HU-15 (W-2.3, D5): estructura de plantilla = ERR-PLANTILLA.
                throw new CalculoInvalidoException(CodigoError.Plantilla, $"La celda '{celda}' de '{HojaConsolidado}' no contiene una fórmula válida.");
            }
        }
    }

    private static void ValidarCadenaFormulaR1R2R4(WorkbookPart workbookPart)
    {
        var hojaR1 = ObtenerHoja(workbookPart, HojaR1, nameof(EscribirConsolidado));
        var hojaR2 = ObtenerHoja(workbookPart, HojaR2, nameof(EscribirConsolidado));
        var hojaR4 = ObtenerHoja(workbookPart, HojaR4, nameof(EscribirConsolidado));

        ValidarCeldaTieneFormula(workbookPart, hojaR1, "F46", ["F25", "F41", "L25"], HojaR1, nameof(EscribirConsolidado));
        ValidarCeldaTieneFormula(workbookPart, hojaR1, "F48", ["F30", "F10", "L10"], HojaR1, nameof(EscribirConsolidado));
        ValidarCeldaTieneFormula(workbookPart, hojaR2, "E41", ["E15", "E26", "K15"], HojaR2, nameof(EscribirConsolidado));
        ValidarCeldaTieneFormula(workbookPart, hojaR4, "D67", ["D9", "P9"], HojaR4, nameof(EscribirConsolidado));
    }

    private static void ValidarHojaConLabelsEsperados(WorkbookPart workbookPart, string hoja, params string[] labelsEsperados)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(EscribirConsolidado));
        var textos = worksheet.Descendants<Cell>()
            .Select(c => LeerTextoCelda(c, workbookPart))
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .Select(t => t.Trim())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var label in labelsEsperados)
        {
            var coincidencia = textos.Any(t => t.Contains(label, StringComparison.OrdinalIgnoreCase));
            if (!coincidencia)
            {
                // HU-15 (W-2.3, D5): plantilla sin el label esperado = ERR-PLANTILLA.
                throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{hoja}' no incluye el label esperado '{label}' y no es compatible con la semántica actual del workbook.");
            }
        }
    }

    private static string ValidarCeldaTieneFormula(WorkbookPart workbookPart, Worksheet worksheet, string celda, string[] fragmentosEsperados, string hoja, string operacion)
    {
        // HU-15 (W-2.3, D5): estructura de plantilla (celda ausente / valor fijo / fórmula no
        // resoluble / referencias faltantes) = ERR-PLANTILLA, nunca ERR-VALIDACION.
        var cell = ObtenerCelda(worksheet, celda)
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La celda '{celda}' no existe en la hoja '{hoja}' para {operacion}. El workbook no es compatible con la estructura esperada.");

        if (cell.CellFormula is null)
        {
            throw new CalculoInvalidoException(CodigoError.Plantilla, $"La celda '{celda}' de '{hoja}' debería seguir siendo fórmula; se detectó un valor fijo. No se puede continuar con la validación del workbook derivado.");
        }

        var formula = NormalizarFormula(LeerTextoCelda(cell, workbookPart));

        // Shared-formula awareness (HU-07 task 2.3): un follower con SharedIndex y texto vacío
        // es una réplica del maestro; se valida por presencia de fórmula, no por fragmentos.
        var esSharedFollower = string.IsNullOrWhiteSpace(formula) && cell.CellFormula.SharedIndex is not null;
        if (string.IsNullOrWhiteSpace(formula) && !esSharedFollower)
        {
            throw new CalculoInvalidoException(CodigoError.Plantilla, $"La celda '{celda}' de '{hoja}' no tiene una fórmula resoluble por OpenXML.");
        }

        if (fragmentosEsperados.Length > 0 && !esSharedFollower)
        {
            var formulaNormalizada = NormalizarFormula(formula);
            var faltan = fragmentosEsperados
                .Select(NormalizarFormula)
                .Where(fragmento => !string.IsNullOrWhiteSpace(fragmento))
                .Where(fragmento => !formulaNormalizada.Contains(fragmento, StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (faltan.Length > 0)
            {
                var texto = cell.CellFormula.Text ?? formula;
                throw new CalculoInvalidoException(CodigoError.Plantilla, $"La celda '{celda}' de '{hoja}' no mantiene todas las referencias esperadas: faltan [{string.Join(", ", faltan)}]. Fórmula actual: '{texto}'.");
            }
        }

        return formula;
    }

    private static string ValidarArchivo(string rutaPlantilla, string operacion)
    {
        // HU-15 (W-2.3, D5): plantilla ausente = problema de PLANTILLA (ERR-PLANTILLA → salida 2),
        // no de fuente (antes default ERR-FUENTE-NO-ENCONTRADA, mismo código de salida pero
        // semántica incorrecta).
        if (string.IsNullOrWhiteSpace(rutaPlantilla))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.Plantilla, $"La ruta de plantilla para {operacion} es requerida.");
        }

        if (!File.Exists(rutaPlantilla))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.Plantilla, $"No se encontró la plantilla para {operacion}: '{rutaPlantilla}'.");
        }

        return rutaPlantilla;
    }

    private static void ValidarRutaSalida(string rutaSalida)
    {
        if (string.IsNullOrWhiteSpace(rutaSalida))
        {
            throw new ArchivoFuenteNoEncontradoException(CodigoError.Plantilla, "La ruta de salida del workbook es requerida.");
        }
    }

    private static void ValidarNoInPlace(string origen, string rutaSalida)
    {
        if (string.Equals(Path.GetFullPath(origen), Path.GetFullPath(rutaSalida), StringComparison.OrdinalIgnoreCase))
        {
            // HU-15 (W-2.3, D5): salida == plantilla es problema de PLANTILLA → ERR-PLANTILLA
            // (antes default ERR-VALIDACION → salida 1, errónea; ahora salida 2).
            throw new CalculoInvalidoException(CodigoError.Plantilla, "La escritura real no puede mutar la plantilla original in-place. Use una ruta de salida distinta.");
        }
    }

    /// <summary>
    /// Resuelve una hoja del workbook. Plan 30 (T2, D-C/R-C-3): si la hoja esperada es una hoja
    /// cuyo nombre depende del período (<c>DetRetri{AAAAMMQ}</c>/<c>DetValiRetri{AAAAMMQ}</c>), el
    /// fail-fast de hoja ausente nombra la hoja esperada + su período + la operación. El período se
    /// EXTRAE del nombre esperado (que el dominio ya compuso) — NUNCA se enumera el workbook para
    /// descubrir hojas (D-A): la resolución sigue siendo por nombre exacto, sin fallback a otro
    /// período.
    /// </summary>
    private static Worksheet ObtenerHoja(WorkbookPart workbookPart, string nombreHoja, string operacion)
    {
        // HU-15 (W-2.3, D5): estructura de plantilla (metadata/hoja ausente) = ERR-PLANTILLA.
        var detallePeriodo = DetalleHojaDePeriodo(nombreHoja);
        var workbook = workbookPart.Workbook ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"El workbook para {operacion} no tiene metadata Workbook válida.");
        var sheet = workbook.Descendants<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{nombreHoja}' no existe en el workbook para {operacion}{detallePeriodo}.");

        var worksheetPart = workbookPart.GetPartById(sheet.Id!) as WorksheetPart
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"No se pudo resolver la hoja '{nombreHoja}' en el workbook para {operacion}{detallePeriodo}.");

        return worksheetPart.Worksheet ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{nombreHoja}' no tiene Worksheet válido para {operacion}{detallePeriodo}.");
    }

    /// <summary>
    /// Plan 30 (T2, D-C): detalle para el mensaje de fail-fast cuando la hoja esperada es del
    /// período (<c>DetRetri*</c>/<c>DetValiRetri*</c>): <c>" (hoja Det del período {AAAAMMQ})"</c>.
    /// Vacío para las demás hojas. NO resuelve hojas: solo anota el período que ya viaja en el
    /// nombre compuesto por el dominio (la resolución sigue siendo por nombre exacto).
    /// </summary>
    private static string DetalleHojaDePeriodo(string nombreHoja)
    {
        foreach (var prefijo in new[] { "DetValiRetri", "DetRetri" })
        {
            if (!nombreHoja.StartsWith(prefijo, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var sufijo = nombreHoja[prefijo.Length..];
            if (sufijo.Length == 7 && sufijo.All(char.IsAsciiDigit))
            {
                return $" (hoja Det del período {sufijo})";
            }
        }

        return string.Empty;
    }

    private static Worksheet ObtenerHoja(SpreadsheetDocument workbook, string nombreHoja, string operacion)
    {
        var workbookPart = workbook.WorkbookPart ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"El workbook para {operacion} no tiene WorkbookPart.");
        return ObtenerHoja(workbookPart, nombreHoja, operacion);
    }

    private static void LanzarPayloadInsuficiente(string operacion, string hoja, string detalle)
    {
        throw new CalculoInvalidoException($"La escritura funcional de '{operacion}' sobre '{hoja}' no está soportada con el payload actual. {detalle} Requiere un nuevo contrato WorkbookLeafInputs para la hoja real.");
    }

    private static Cell? ObtenerCelda(Worksheet worksheet, string cellReference)
    {
        var sheetData = worksheet.Elements<SheetData>().FirstOrDefault();
        if (sheetData is null)
        {
            return null;
        }

        return sheetData.Descendants<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));
    }

    private static string LeerTextoCelda(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.CellFormula is not null)
        {
            return cell.CellFormula.Text ?? string.Empty;
        }

        if (cell.CellValue is null)
        {
            return string.Empty;
        }

        if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString)
        {
            var sharedStrings = workbookPart.SharedStringTablePart;
            if (sharedStrings is not null && sharedStrings.SharedStringTable is not null
                && int.TryParse(cell.CellValue.Text, out var index) && index >= 0
                && index < sharedStrings.SharedStringTable.Count())
            {
                var item = sharedStrings.SharedStringTable.ElementAt(index);
                var text = item.InnerText;
                return text ?? string.Empty;
            }
        }

        if (cell.DataType is not null && cell.DataType.Value == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText ?? string.Empty;
        }

        return cell.CellValue.InnerText;
    }

    private static string NormalizarFormula(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return string.Empty;
        }

        var value = formula.Trim();
        value = value.Replace("'", string.Empty, StringComparison.Ordinal);
        value = value.Replace(" ", string.Empty, StringComparison.Ordinal);
        value = value.Replace("_xlfn.", string.Empty, StringComparison.OrdinalIgnoreCase);
        return value;
    }

    private static void ValidarFormulasProtegidas(WorkbookPart workbookPart, string operacion)
    {
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMap.ProtectedFormulas)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }
    }

    /// <summary>
    /// Valida el mapa ampliado: visibles de cada bloque ASE (R1/R2/R4) + filas CONSOLIDADO
    /// D9:D13/D28:D32/D47:D51/D66:D70/D85:D89/D104:D108/D109 con shared-formula awareness,
    /// + mapa 2.2 (HU-08): <c>REMUNERACION_*</c>, <c>VALIDACION_*</c>, <c>GERENTES_*</c>,
    /// <c>Recaudo *</c> fila 29+ (D6) + HU-09 (banco) + HU-10 (BCE, parametrizado al período).
    /// HU-11 (2.5): en Q2 además valida la cadena AJUSTES-SF-T (<see cref="WorkbookLeafCellMapAjustesSfT.Protegidas"/>).
    /// HU-12 (2.6 ampliada, D5): validación PROTEGIDA parametrizada por período — Q1 exige las
    /// fórmulas HU-07 (F46/F176/…/D67, invariante Q1); Q2 exige el mapa T0 (<see cref="WorkbookLeafCellMapQ2"/>:
    /// F53…, E43…, D73…, DetValiRetri/VALIDACION_*/INTERVENTORIA/ANT EXT-REV siempre protegidas).
    /// </summary>
    private static void ValidarFormulasProtegidasMultiAse(WorkbookPart workbookPart, string operacion, Periodo periodo)
    {
        if (periodo.NumeroQuincena == 2)
        {
            ValidarFormulasProtegidasMultiAseQ2(workbookPart, operacion, periodo);
            return;
        }

        foreach (var aseId in WorkbookLeafCellMapPorAse.EditableLeafCellsPorAse.Keys.OrderBy(k => k))
        {
            var bloque = WorkbookLeafCellMapPorAse.ProtectedFormulasPorAse[aseId];
            foreach (var (hoja, celda, fragmentos) in bloque)
            {
                if (string.Equals(hoja, HojaConsolidado, StringComparison.OrdinalIgnoreCase))
                {
                    continue; // las filas CONSOLIDADO se validan una sola vez abajo
                }

                var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
                ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
            }
        }

        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapPorAseProtectedConsolidado)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapPorEmpresa.ProtectedFormulasPorEmpresa)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-09 (2.3, D6): mapa de fórmulas protegidas del reporte por banco (consolidado 1–7,
        // Total de bloque, verificación I/J, validación 59–80 y TOTAL RECAUDO). Nunca se escriben.
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapReporteBanco.Protegidas)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-10 (2.4, D6): mapa de fórmulas protegidas 2.4 (BCE F/I/H + filas 9/10/11/12/13 +
        // bloque 18–24 + CONSOLIDADO J/K/M + refs DetRetri/DetValiRetri). Jamás se escriben.
        // Plan 30 (D-A): el mapa HU-10 (anclado a Q1) se resuelve al período vía
        // ProtegidasBceParaPeriodo(periodo) (para 2026071 la sustitución es identidad).
        foreach (var (hoja, celda, fragmentos) in ProtegidasBceParaPeriodo(periodo))
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-11 (2.5, Requirement 7): cadena AJUSTES-SF-T protegida (visibles SALDOS/RETRIBUCION,
        // AJUSTES D9:D13/D28:D32/D47:D51, CONSOLIDADO D85:D89, INTERVENTORIA, ANT EXT-REV).
        if (periodo.NumeroQuincena == 2)
        {
            foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapAjustesSfT.Protegidas)
            {
                var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
                ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
            }
        }

        // HU-13 (2.7, Requirement 4/D4): mapa protegido extendido de TODAS las hojas de
        // validación (VALIDACION_*, VALIDACION_TOTAL, DetRetri/DetValiRetri col D, Valida -*,
        // GERENTES_*). El writer falla si alguna deja de ser fórmula donde T0 lo exige.
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapValidaciones.ProtegidasValidacionesParaPeriodo(periodo))
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-16 (D2b): assert estructural de INTERVENTORIA (bloque presente con el carácter T0:
        // totales en fórmula + filas ASE 26..30 en VALORES con L = Id ASE). Aplica en ambos
        // períodos (Q1 gana la protección; Q2 re-asegura el mapa existente sin duplicarlo).
        ValidarInterventoriaEstructura(workbookPart, operacion);
    }

    /// <summary>
    /// HU-12 (2.6 ampliada, D5): validación protegida Q2 contra el mapa T0 congelado
    /// (<see cref="WorkbookLeafCellMapQ2"/>). Q1 NO se valida aquí (rama por período): las celdas
    /// F46/F48/F519/F521 del mapa HU-07 son VALORES en Q2 (V0.2) y quedan EXCLUIDAS de
    /// protegidas-fórmula. Incluye: visibles R1/R2/R4-Q2 por ASE, CONSOLIDADO Q2, banco HU-09,
    /// BCE parametrizado 2026072 (M1), cadena AJUSTES-SF-T, DetRetri/DetValiRetri Q2 y las
    /// protegidas adicionales (REMUNERACION_* genéricas + VALIDACION_*/GERENTES_*/INTERVENTORIA/
    /// ANT EXT-REV — Requirement 5). El writer falla si alguna deja de ser fórmula.
    /// </summary>
    private static void ValidarFormulasProtegidasMultiAseQ2(WorkbookPart workbookPart, string operacion, Periodo periodo)
    {
        foreach (var aseId in WorkbookLeafCellMapQ2.R1Q2EditablesPorAse.Keys.OrderBy(k => k))
        {
            foreach (var (celda, fragmentos) in WorkbookLeafCellMapQ2.ObtenerR1Q2Protegidos(aseId))
            {
                var worksheet = ObtenerHoja(workbookPart, HojaR1, operacion);
                ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, HojaR1, operacion);
            }

            foreach (var (celda, fragmentos) in WorkbookLeafCellMapQ2.ObtenerR2Q2Protegidos(aseId))
            {
                var worksheet = ObtenerHoja(workbookPart, HojaR2, operacion);
                ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, HojaR2, operacion);
            }

            foreach (var (celda, fragmentos) in WorkbookLeafCellMapQ2.ObtenerR4Q2Protegidos(aseId))
            {
                var worksheet = ObtenerHoja(workbookPart, HojaR4, operacion);
                ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, HojaR4, operacion);
            }
        }

        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.ConsolidadoProtected)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-09 (2.3): mapa banco parametrizado al período — el Q2 excluye el TOTAL RECAUDO fila 81
        // (C81/D81) que no existe en el template Q2 (T0-0.5; la hoja termina en la fila 79).
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.BancoProtegidasQ2)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-10 (2.4) + Plan 30 (D-A): BCE resuelto al período (M1: matchea hojas/celdas reales,
        // incl. DetRetri{periodo}/DetValiRetri{periodo}).
        foreach (var (hoja, celda, fragmentos) in ProtegidasBceParaPeriodo(periodo))
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-11 (2.5): cadena AJUSTES-SF-T (Q2).
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapAjustesSfT.Protegidas)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-12 (2.6 ampliada, Requirement 5): DetRetri/DetValiRetri Q2 (T0-0.4) + protegidas
        // adicionales (REMUNERACION_* genéricas, VALIDACION_*, GERENTES_*, INTERVENTORIA, ANT EXT-REV).
        // Plan 30 (T2, D-C): la resolución de la hoja Det del período anota hoja + período + operación.
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.DetRetriProtected(periodo))
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.ProtegidasAdicionalesQ2)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-13 (2.7, Requirement 4/D4): mapa protegido extendido 2.7 en Q2 (VALIDACION_* O/P
        // filas 3..7, VALIDACION_TOTAL O/P, Valida -*, GERENTES_* SUM). DetRetri/DetValiRetri Q2
        // ya cubiertos por WorkbookLeafCellMapQ2.DetRetriProtected (no se duplica).
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapValidaciones.ProtegidasValidacionesParaPeriodo(periodo))
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-16 (D2b): assert estructural de INTERVENTORIA (ambos períodos; Requirement 5).
        ValidarInterventoriaEstructura(workbookPart, operacion);
    }

    /// <summary>
    /// HU-16 (D2b, §2.5 regla 3): assert estructural de <c>INTERVENTORIA</c> con el carácter T0
    /// congelado: los totales K31/M31/N31 (SUM) y el gran total K32 (SUM(M31:N31)) siguen siendo
    /// FÓRMULA, y las filas ASE 26..30 (K/L/M/N) son VALORES (no fórmula) con L26..L30 = Id ASE.
    /// Detecta stale futuro (el riesgo que motiva esta HU): si la plantilla cambia el carácter
    /// del bloque, fail-fast que nombra la hoja y la celda — nunca se escribe ni se inventa.
    /// </summary>
    private static void ValidarInterventoriaEstructura(WorkbookPart workbookPart, string operacion)
    {
        var hoja = WorkbookLeafCellMapInterventoria.HojaInterventoria;
        var worksheet = ObtenerHoja(workbookPart, hoja, operacion);

        foreach (var (_, celda, fragmentos) in WorkbookLeafCellMapInterventoria.FormulasProtegidas)
        {
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        foreach (var celda in WorkbookLeafCellMapInterventoria.CeldasValoresBloque())
        {
            var cell = ObtenerCelda(worksheet, celda)
                ?? throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La celda '{hoja}!{celda}' no existe en la plantilla para {operacion}. INTERVENTORIA perdió el bloque por ASE (T0-0.2) — fail-fast, nunca valor inventado.");

            if (cell.CellFormula is not null)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La celda '{hoja}!{celda}' debería ser VALOR estático (insumo externo declarado) y es fórmula. INTERVENTORIA cambió su carácter T0 — fail-fast, nunca se escribe.");
            }
        }

        // L26..L30 = Id ASE (1..5) — el bloque mantiene el orden congelado por T0-0.2.
        for (var i = 0; i < 5; i++)
        {
            var aseId = i + 1;
            var celda = WorkbookLeafCellMapInterventoria.CeldaBloque("L", WorkbookLeafCellMapInterventoria.FilaPrimerAse + i);
            var valor = LeerCeldaNumerica(workbookPart, worksheet, hoja, celda);
            if (Math.Abs(valor - aseId) > 0.5m)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"La celda '{hoja}!{celda}' debería contener el Id del ASE {aseId} (bloque T0-0.2) y vale {valor}. INTERVENTORIA cambió su estructura — fail-fast.");
            }
        }
    }

    private static decimal LeerCeldaNumerica(WorkbookPart workbookPart, Worksheet worksheet, string hoja, string celda)
    {
        var cell = ObtenerCelda(worksheet, celda)
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La celda '{hoja}!{celda}' no existe en el workbook.");
        if (cell.CellValue is null || string.IsNullOrWhiteSpace(cell.CellValue.InnerText))
        {
            return 0m;
        }

        if (!decimal.TryParse(cell.CellValue.InnerText, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var valor))
        {
            // HU-17 (S-3 HU-16): texto en celda de gate ≠ 0 — fail-fast que nombra la celda
            // (doctrina W2: nunca 0 silencioso en gates). Si INTERVENTORIA trae "ASE1" en L26
            // en vez del Id numérico, el fallo lo dice tal cual, no "vale 0".
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La celda '{hoja}!{celda}' contiene un valor NO numérico ('{cell.CellValue.InnerText}'); se esperaba el Id del ASE (bloque T0-0.2). INTERVENTORIA cambió su estructura — fail-fast.");
        }

        return valor;
    }

    /// <summary>
    /// HU-11 (2.5) + Plan 30 (D-A): devuelve el mapa protegido BCE con el nombre de hoja
    /// DetRetri/DetValiRetri del período, sustituyendo el ancla-Q1 (<see cref="AnclaSufijoHojasQ1"/>)
    /// por <c>periodo.CodigoCompleto</c>. Para 2026071 la sustitución es identidad; para 2026072/…
    /// resuelve el nombre real. El mapa HU-10 NO se toca (regla del plan): solo se parametriza su
    /// interpretación por período.
    /// </summary>
    private static IEnumerable<(string Hoja, string Celda, string[] Fragmentos)> ProtegidasBceParaPeriodo(Periodo periodo)
    {
        var codigoCompleto = periodo.CodigoCompleto;
        return WorkbookLeafCellMapBalanceSc.Protegidas
            .Select(p => (
                ReemplazarSufijo(p.Hoja, codigoCompleto),
                p.Celda,
                p.Fragmentos.Select(f => ReemplazarSufijo(f, codigoCompleto)).ToArray()));
    }

    private static string ReemplazarSufijo(string texto, string codigoCompleto) =>
        texto.Replace(AnclaSufijoHojasQ1, codigoCompleto, StringComparison.Ordinal);

    /// <summary>
    /// Filas CONSOLIDADO del mapa ampliado multi-ASE (plan §2.3): D9:D13, D28:D32, D47:D51,
    /// D66:D70, D85:D89, D104:D108 (maestro + réplicas shared), D109. Fórmulas jamás se escriben.
    /// </summary>
    private static readonly (string Hoja, string Celda, string[] Fragmentos)[] WorkbookLeafCellMapPorAseProtectedConsolidado =
    [
        (HojaConsolidado, "D9", ["F46", "Reporte Componentes R1"]),
        (HojaConsolidado, "D10", ["F176", "Reporte Componentes R1"]),
        (HojaConsolidado, "D11", ["F316", "Reporte Componentes R1"]),
        (HojaConsolidado, "D12", ["F437", "Reporte Componentes R1"]),
        (HojaConsolidado, "D13", ["F519", "Reporte Componentes R1"]),
        (HojaConsolidado, "D28", ["E41", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D29", ["E135", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D30", ["E247", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D31", ["E343", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D32", ["E413", "Rem. Anticipos R2"]),
        (HojaConsolidado, "D47", ["F48", "Reporte Componentes R1"]),
        (HojaConsolidado, "D48", ["F178", "Reporte Componentes R1"]),
        (HojaConsolidado, "D49", ["F318", "Reporte Componentes R1"]),
        (HojaConsolidado, "D50", ["F439", "Reporte Componentes R1"]),
        (HojaConsolidado, "D51", ["F521", "Reporte Componentes R1"]),
        (HojaConsolidado, "D66", ["D67", "Reversion Pagos R4"]),
        (HojaConsolidado, "D67", ["D161", "Reversion Pagos R4"]),
        (HojaConsolidado, "D68", ["D198", "Reversion Pagos R4"]),
        (HojaConsolidado, "D69", ["D312", "Reversion Pagos R4"]),
        (HojaConsolidado, "D70", ["D347", "Reversion Pagos R4"]),
        (HojaConsolidado, "D85", ["D47", "AJUSTES"]),
        (HojaConsolidado, "D86", ["D48", "AJUSTES"]),
        (HojaConsolidado, "D87", ["D49", "AJUSTES"]),
        (HojaConsolidado, "D88", ["D50", "AJUSTES"]),
        (HojaConsolidado, "D89", ["D51", "AJUSTES"]),
        (HojaConsolidado, "D104", ["D9", "D28", "D47", "D66", "D85"]),
        (HojaConsolidado, "D105", ["D10", "D29", "D48", "D67", "D86"]),
        (HojaConsolidado, "D106", []), // shared follower del maestro D104
        (HojaConsolidado, "D107", ["D12", "D31", "D50", "D69", "D88"]),
        (HojaConsolidado, "D108", ["D13", "D32", "D51", "D70", "D89"]),
        (HojaConsolidado, "D109", ["D104", "D108", "SUM"])
    ];

    private static void EscribirCeldasLeaf(WorkbookPart workbookPart, WorkbookLeafInputs leafInputs)
    {
        var valores = new Dictionary<(string Hoja, string Celda), decimal>(new LeafCellComparer())
        {
            [(WorkbookLeafCellMap.HojaR1, "F25")] = leafInputs.R1.F25,
            [(WorkbookLeafCellMap.HojaR1, "F41")] = leafInputs.R1.F41,
            [(WorkbookLeafCellMap.HojaR1, "L25")] = leafInputs.R1.L25,
            [(WorkbookLeafCellMap.HojaR1, "F30")] = leafInputs.R1.F30,
            [(WorkbookLeafCellMap.HojaR1, "F10")] = leafInputs.R1.F10,
            [(WorkbookLeafCellMap.HojaR1, "L10")] = leafInputs.R1.L10,
            [(WorkbookLeafCellMap.HojaR2, "E15")] = leafInputs.R2.E15,
            [(WorkbookLeafCellMap.HojaR2, "E26")] = leafInputs.R2.E26,
            [(WorkbookLeafCellMap.HojaR2, "K15")] = leafInputs.R2.K15,
            [(WorkbookLeafCellMap.HojaR4, "D9")] = leafInputs.R4.D9,
            [(WorkbookLeafCellMap.HojaR4, "P9")] = leafInputs.R4.P9
        };

        foreach (var (hoja, celda, nombre) in WorkbookLeafCellMap.EditableLeafCells)
        {
            if (WorkbookLeafCellMap.ProtectedFormulas.Any(p =>
                    string.Equals(p.Hoja, hoja, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(p.Celda, celda, StringComparison.OrdinalIgnoreCase)))
            {
                throw new CalculoInvalidoException(
                    $"El cell-map intentó escribir '{nombre}' sobre la fórmula protegida {hoja}!{celda}.");
            }

            if (!valores.TryGetValue((hoja, celda), out var valor))
            {
                throw new CalculoInvalidoException($"No hay valor leaf mapeado para {nombre} ({hoja}!{celda}).");
            }

            EscribirValorNumerico(workbookPart, hoja, celda, valor, nombre);
        }
    }

    /// <summary>
    /// Escribe SOLO las celdas del bloque del ASE (mapa por ASE congelado por T0).
    /// Los valores provienen de <see cref="WorkbookLeafInputs"/> (CeldasPorAse por hoja).
    /// HU-12 (2.6 ampliada, D1/G2): en Q2 escribe el mapa hermano <see cref="WorkbookLeafCellMapQ2"/>
    /// (R1/R2/R4-Q2 por ASE, incluida la variante ASE5 de 2 filas V0.3); Q1 queda intacto.
    /// </summary>
    private static void EscribirCeldasLeafPorAse(WorkbookPart workbookPart, WorkbookLeafInputs leaf, bool esQuincena2, bool omitirR1)
    {
        if (esQuincena2)
        {
            EscribirCeldasLeafPorAseQ2(workbookPart, leaf, omitirR1);
            return;
        }

        var editables = WorkbookLeafCellMapPorAse.ObtenerEditables(leaf.Ase.Id);

        foreach (var (hoja, celda, nombre) in editables)
        {
            if (omitirR1 && string.Equals(hoja, HojaR1, StringComparison.OrdinalIgnoreCase))
            {
                continue; // Plan 21 (T4): la hoja R1 la gobierna el espejo, no el mapa absoluto.
            }

            if (WorkbookLeafCellMapPorAse.ProtectedFormulasPorAse[leaf.Ase.Id].Any(p =>
                    string.Equals(p.Hoja, hoja, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(p.Celda, celda, StringComparison.OrdinalIgnoreCase)))
            {
                throw new CalculoInvalidoException(
                    $"El cell-map del ASE {leaf.Ase.Id} intentó escribir '{nombre}' sobre la fórmula protegida {hoja}!{celda}.");
            }

            var valor = ObtenerValorLeaf(leaf, hoja, celda, nombre);
            EscribirValorNumerico(workbookPart, hoja, celda, valor, nombre);
        }
    }

    private static decimal ObtenerValorLeaf(WorkbookLeafInputs leaf, string hoja, string celda, string nombre)
    {
        IReadOnlyDictionary<string, decimal> celdas = hoja switch
        {
            _ when string.Equals(hoja, HojaR1, StringComparison.OrdinalIgnoreCase) => leaf.R1.CeldasPorAse,
            _ when string.Equals(hoja, HojaR2, StringComparison.OrdinalIgnoreCase) => leaf.R2.CeldasPorAse,
            _ when string.Equals(hoja, HojaR4, StringComparison.OrdinalIgnoreCase) => leaf.R4.CeldasPorAse,
            _ => throw new CalculoInvalidoException($"Hoja '{hoja}' no tiene valores leaf mapeados para '{nombre}'.")
        };

        if (!celdas.TryGetValue(celda, out var valor))
        {
            throw new CalculoInvalidoException($"No hay valor leaf mapeado para {nombre} ({hoja}!{celda}) del ASE {leaf.Ase.Id}.");
        }

        return valor;
    }

    /// <summary>
    /// HU-12 (2.6 ampliada, T0-0.2/0.3): escribe SOLO las celdas Q2 del bloque del ASE
    /// (R1/R2/R4 del mapa <see cref="WorkbookLeafCellMapQ2"/>), incluida la variante ASE5 de 2
    /// filas (V0.3). Fail-fast si falta un valor mapeado (nombra ASE + hoja + celda; nunca 0
    /// silencioso). El guard de <see cref="EscribirValorNumerico"/> impide tocar fórmulas.
    /// </summary>
    private static void EscribirCeldasLeafPorAseQ2(WorkbookPart workbookPart, WorkbookLeafInputs leaf, bool omitirR1)
    {
        if (!omitirR1)
        {
            foreach (var (celda, _) in WorkbookLeafCellMapQ2.ObtenerR1Q2Editables(leaf.Ase.Id))
            {
                if (!leaf.R1.CeldasPorAse.TryGetValue(celda, out var valorR1))
                {
                    throw new CalculoInvalidoException($"No hay valor leaf R1-Q2 mapeado para {celda} del ASE {leaf.Ase.Id}.");
                }

                EscribirValorNumerico(workbookPart, HojaR1, celda, valorR1, $"R1-Q2.ASE{leaf.Ase.Id}.{celda}");
            }
        }

        var r2 = WorkbookLeafCellMapQ2.ObtenerR2Q2Editables(leaf.Ase.Id);
        foreach (var celda in new[] { r2.Componente, r2.SubsCont, r2.Especiales })
        {
            if (!leaf.R2.CeldasPorAse.TryGetValue(celda, out var valorR2))
            {
                throw new CalculoInvalidoException($"No hay valor leaf R2-Q2 mapeado para {celda} del ASE {leaf.Ase.Id}.");
            }

            EscribirValorNumerico(workbookPart, HojaR2, celda, valorR2, $"R2-Q2.ASE{leaf.Ase.Id}.{celda}");
        }

        var r4 = WorkbookLeafCellMapQ2.ObtenerR4Q2Editables(leaf.Ase.Id);
        foreach (var celda in new[] { r4.Total, r4.P })
        {
            if (!leaf.R4.CeldasPorAse.TryGetValue(celda, out var valorR4))
            {
                throw new CalculoInvalidoException($"No hay valor leaf R4-Q2 mapeado para {celda} del ASE {leaf.Ase.Id}.");
            }

            EscribirValorNumerico(workbookPart, HojaR4, celda, valorR4, $"R4-Q2.ASE{leaf.Ase.Id}.{celda}");
        }
    }

    /// <summary>
    /// HU-12 (2.6 ampliada, V0.4): escribe DetRetri-Q2 (hoja <c>DetRetri2026072</c>) en la MISMA
    /// pasada atómica: D9:D13 = ROUND(D104:D108) por ASE (vía <see cref="DetRetriRounder"/>, única
    /// regla) y D14 = ROUND(Σ D104:D108). La composición está CONGELADA (probada 5/5 contra el
    /// golden, V0.4); nunca se inventa. El llamador la ejecuta SOLO en Q2 (HU-20/G3): en Q1 el
    /// DetRetri calculado existe como oráculo de validación contra el R10, pero la hoja
    /// <c>DetRetri2026071</c> queda con sus fórmulas protegidas intactas.
    /// </summary>
    private static void EscribirCeldasDetRetriQ2(WorkbookPart workbookPart, IReadOnlyList<WorkbookLeafInputs> leafInputs)
    {
        var conDetalle = leafInputs
            .Where(l => l.DetRetriQ2 is not null)
            .OrderBy(l => l.Ase.Id)
            .ToList();
        if (conDetalle.Count == 0)
        {
            return; // Q1 puro: sin escrituras 2.6.
        }

        // Plan 30 (D-A): la hoja DetRetri se resuelve del dominio del período (no de un literal).
        var periodo = conDetalle[0].Periodo;
        var hojaDetRetri = WorkbookLeafCellMapQ2.HojaDetRetri(periodo);

        foreach (var leaf in conDetalle)
        {
            var celda = WorkbookLeafCellMapQ2.ObtenerDetRetriDestino(leaf.Ase.Id);
            EscribirValorNumericoEnCeldaExistente(workbookPart, hojaDetRetri, celda, leaf.DetRetriQ2!.Detalle, $"escritura DetRetri-Q2 ASE{leaf.Ase.Id} {celda}");

            // Plan 29 (T4, Unidad D): desglose trazable por columna (gate T0d) en la MISMA pasada.
            EscribirDesgloseTrazableDetRetri(workbookPart, leaf);
        }

        var totalD104 = conDetalle.Sum(l => l.DetRetriQ2!.TotalD104);
        EscribirValorNumericoEnCeldaExistente(workbookPart, hojaDetRetri, WorkbookLeafCellMapQ2.DetRetriTotal, DetRetriRounder.Round(totalD104), "escritura DetRetri-Q2 total D14");
    }

    /// <summary>
    /// Plan 29 (T4, Unidad D — desglose trazable): escribe como LITERAL redondeado las columnas de
    /// <c>DetRetri2026072</c> / <c>DetValiRetri2026072</c> cuyo valor workbook-interno es de UNA
    /// arista (mapa <see cref="WorkbookLeafCellMapDetRetri"/>): J ← BCE F3 (Subsidio+Contribución),
    /// L ← INTERVENTORIA F15, DetValiRetri I ← −L y O ← 0.
    ///
    /// Las columnas cuyo origen es una fórmula de la cadena R1/R2/R4/AJUSTES NO se escriben y quedan
    /// listadas con motivo en <see cref="WorkbookLeafCellMapDetRetri.ColumnasExcluidas"/> (el plan
    /// prohíbe replicar la cadena en C#). El manual las deja literales (V7); la app no las inventa.
    /// El guard anti-fórmula de <see cref="EscribirValorNumericoEnCeldaExistente"/> sigue intacto
    /// (destino fórmula ⇒ ERR-PLANTILLA, nunca sobrescritura).
    /// </summary>
    private static void EscribirDesgloseTrazableDetRetri(WorkbookPart workbookPart, WorkbookLeafInputs leaf)
    {
        var aseId = leaf.Ase.Id;
        var periodo = leaf.Periodo;
        var fila = WorkbookLeafCellMapDetRetri.ObtenerFila(aseId);
        var hojaDetRetri = WorkbookLeafCellMapDetRetri.HojaDetRetri(periodo);
        var hojaDetValiRetri = WorkbookLeafCellMapDetRetri.HojaDetValiRetri(periodo);

        // Valor-trazable de la interventoría (INTERVENTORIA!F15, 2ª quincena, literal declarado).
        var interventoria = InterventoriaDeclarada.SegundaQuincenaPorAse[aseId];

        // J = BCE SC POR FACT.!F3 = D3+E3 (Subsidio + Contribución), ya escritos por la Unidad B.
        var balance = leaf.BalanceSc ?? throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"El leaf del ASE {aseId} no trae BalanceSc para derivar DetRetri!J{fila} (BCE SC POR FACT.!F3); fail-fast, nunca se inventa el desglose.");
        var bce = balance.Ases.SingleOrDefault(a => a.Ase.Id == aseId) ?? throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"El BalanceSc del ASE {aseId} no trae su fila; no se puede derivar DetRetri!J{fila}.");
        var j = DetRetriRounder.Round(bce.TotalBsc);

        foreach (var (columna, fuente) in WorkbookLeafCellMapDetRetri.ColumnasDetRetri)
        {
            var celda = $"{columna}{fila}";
            var valor = fuente switch
            {
                WorkbookLeafCellMapDetRetri.FuenteDetRetri.BalanceBceF3 => j,
                WorkbookLeafCellMapDetRetri.FuenteDetRetri.InterventoriaF15 => interventoria,
                _ => throw new ArgumentOutOfRangeException(nameof(fuente), fuente, "Fuente DetRetri no soportada.")
            };
            EscribirValorNumericoEnCeldaExistente(workbookPart, hojaDetRetri, celda, valor, $"escritura DetRetri-Q2 desglose ASE{aseId} {celda}");
        }

        foreach (var (columna, fuente) in WorkbookLeafCellMapDetRetri.ColumnasDetValiRetri)
        {
            var celda = $"{columna}{fila}";
            var valor = fuente switch
            {
                WorkbookLeafCellMapDetRetri.FuenteDetValiRetri.InterventoriaF15Negada => -interventoria,
                WorkbookLeafCellMapDetRetri.FuenteDetValiRetri.Cero => 0m,
                _ => throw new ArgumentOutOfRangeException(nameof(fuente), fuente, "Fuente DetValiRetri no soportada.")
            };
            EscribirValorNumericoEnCeldaExistente(workbookPart, hojaDetValiRetri, celda, valor, $"escritura DetValiRetri-Q2 desglose ASE{aseId} {celda}");
        }
    }

    /// <summary>
    /// HU-13 (2.7, W1-guarda D8): escribe un valor SOLO si la celda destino YA EXISTE en la
    /// plantilla (nunca crea filas ni celdas: si una plantilla futura trae más/menos filas ASE en
    /// DetRetri D9:D14, el fail-fast NOMBRA la hoja; jamás truncado silencioso ni insert/delete).
    /// Usado por la escritura DetRetri-Q2 (V0.4: D9:D14 son VALORES editables del template).
    ///
    /// Plan 30 (T2, D-C/R-C-3): la resolución de hoja (<see cref="ObtenerHoja"/>) falla nombrando
    /// hoja esperada + período + operación si la hoja Det del período no existe.
    /// </summary>
    private static void EscribirValorNumericoEnCeldaExistente(WorkbookPart workbookPart, string hoja, string celda, decimal valor, string operacion)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, operacion);

        if (ObtenerCelda(worksheet, celda) is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"Guarda de capacidad: la celda '{hoja}!{celda}' no existe en la plantilla para {operacion}. La plantilla futura cambió la capacidad D9:D14 del ASE — fail-fast, nunca truncado silencioso.");
        }

        EscribirValorNumerico(workbookPart, hoja, celda, valor, operacion);
    }

    /// <summary>
    /// HU-08 (2.2): escribe en la MISMA pasada los operandos por empresa (R1/R2/R4) y las hojas
    /// <c>Recaudo *</c> en valores. Solo celdas del mapa congelado; el guard de fórmulas
    /// (<see cref="EscribirValorNumerico"/>) impide tocar cualquier celda con <c>&lt;f&gt;</c>.
    /// </summary>
    private static void EscribirCeldasEmpresa(WorkbookPart workbookPart, WorkbookLeafInputs leaf, bool omitirR1)
    {
        foreach (var conc in leaf.Conciliacion)
        {
            EscribirCeldasConciliacion(workbookPart, leaf.Ase.Id, conc, omitirR1);
        }

        foreach (var recaudo in leaf.Recaudos)
        {
            foreach (var (celda, valor) in recaudo.Celdas)
            {
                EscribirValorNumerico(workbookPart, recaudo.HojaRecaudo, celda, valor, $"Recaudo.{recaudo.Empresa.Nombre}.{celda}");
            }
        }
    }

    private static void EscribirCeldasConciliacion(WorkbookPart workbookPart, int aseId, ConciliacionEmpresaInputs conc, bool omitirR1)
    {
        if (!omitirR1 && WorkbookLeafCellMapPorEmpresa.EditablesR1PorEmpresa.TryGetValue((conc.Empresa.Id, aseId), out var r1))
        {
            foreach (var (celda, _) in r1)
            {
                if (!conc.CeldasR1.TryGetValue(celda, out var valor))
                {
                    throw new CalculoInvalidoException($"No hay valor R1 por empresa para {conc.Empresa.Nombre} ({HojaR1}!{celda}) del ASE {aseId}.");
                }

                EscribirValorNumerico(workbookPart, HojaR1, celda, valor, $"R1.{conc.Empresa.Nombre}.{celda}");
            }
        }

        if (WorkbookLeafCellMapPorEmpresa.EditablesR2PorEmpresa.TryGetValue((conc.Empresa.Id, aseId), out var r2))
        {
            foreach (var (celda, _) in r2)
            {
                if (!conc.CeldasR2.TryGetValue(celda, out var valor))
                {
                    throw new CalculoInvalidoException($"No hay valor R2 por empresa para {conc.Empresa.Nombre} ({HojaR2}!{celda}) del ASE {aseId}.");
                }

                EscribirValorNumerico(workbookPart, HojaR2, celda, valor, $"R2.{conc.Empresa.Nombre}.{celda}");
            }
        }

        if (WorkbookLeafCellMapPorEmpresa.EditablesR4PorEmpresa.TryGetValue((conc.Empresa.Id, aseId), out var r4))
        {
            foreach (var (celda, _) in r4)
            {
                if (!conc.CeldasR4.TryGetValue(celda, out var valor))
                {
                    throw new CalculoInvalidoException($"No hay valor R4 por empresa para {conc.Empresa.Nombre} ({HojaR4}!{celda}) del ASE {aseId}.");
                }

                EscribirValorNumerico(workbookPart, HojaR4, celda, valor, $"R4.{conc.Empresa.Nombre}.{celda}");
            }
        }
    }

    /// <summary>
    /// HU-09 (2.3): escribe SOLO las celdas del mapa banco del ASE (bloques 9–58 en valores
    /// + C59 = quincena) en la MISMA pasada atómica HU-07/HU-08 (G4/D3). <c>ReporteBanco == null</c>
    /// = comportamiento HU-08 puro. El guard de <see cref="EscribirValorNumerico"/> impide tocar
    /// cualquier celda con &lt;f&gt; (filas 1–7, Total de bloque, 59–80 y TOTAL quedan intactos).
    /// </summary>
    private static void EscribirCeldasBanco(WorkbookPart workbookPart, WorkbookLeafInputs leaf)
    {
        if (leaf.ReporteBanco is null)
        {
            return; // HU-08 puro (lista 2.3 vacía = comportamiento existente intacto).
        }

        var bloque = leaf.ReporteBanco.Ases.SingleOrDefault(b => b.Ase.Id == leaf.Ase.Id)
            ?? throw new CalculoInvalidoException(
                $"El leaf del ASE {leaf.Ase.Id} no trae su bloque de reporte banco para escribir.");

        var editables = WorkbookLeafCellMapReporteBanco.ObtenerEditables(leaf.Ase.Id);

        foreach (var empresa in bloque.Empresas)
        {
            if (!editables.TryGetValue(empresa.Empresa, out var celdas))
            {
                throw new CalculoInvalidoException(
                    $"El bloque banco del ASE {leaf.Ase.Id} trae la empresa '{empresa.Empresa}' pero el mapa T0-0.7 no la declara para ese ASE.");
            }

            var valores = new[] { empresa.AplicadosFacturacion, empresa.SaldosFavorGenerados, empresa.FinanciacionesNuevas, empresa.RecibosServEspeciales };
            for (var i = 0; i < celdas.Length && i < valores.Length; i++)
            {
                EscribirValorNumerico(workbookPart, HojaBanco, celdas[i], valores[i], $"Banco.{empresa.Empresa}.{celdas[i]}");
            }
        }

        // C59 = quincena (dominio, nunca fuente; Requirement 4).
        EscribirValorNumerico(workbookPart, HojaBanco, "C59", leaf.ReporteBanco.Quincena, "Banco.C59");
    }

    /// <summary>
    /// HU-10 (2.4): escribe SOLO las celdas D/E de la hoja <c>BCE SC POR FACT.</c> del ASE
    /// (Subsidio → D, Contribución → E: header de la plantilla, veredicto corregido T0-V4 del
    /// Plan 29 — el veredicto original del Plan 10 "Contribución → D" quedó refutado) en la MISMA
    /// pasada atómica HU-07/HU-08/HU-09 (G5/D3). El nombre del log sigue al concepto, no a la
    /// columna: <c>BCE.Contribucion.ASE{n}</c> siempre escribe el valor de contribución (→ celda E)
    /// y <c>BCE.Subsidio.ASE{n}</c> el de subsidio (→ celda D). <c>BalanceSc == null</c> =
    /// comportamiento HU-09 puro. La columna H es fórmula (D2(b), T0-0.6) → nunca se escribe;
    /// F/I/filas 9/11/18–24 y los downstream J/K/M/DetRetri/DetValiRetri quedan protegidos por
    /// <see cref="WorkbookLeafCellMapBalanceSc.Protegidas"/> y el guard de
    /// <see cref="EscribirValorNumerico"/> impide tocar cualquier celda con &lt;f&gt;.
    /// </summary>
    private static void EscribirCeldasBalanceSc(WorkbookPart workbookPart, WorkbookLeafInputs leaf)
    {
        if (leaf.BalanceSc is null)
        {
            return; // HU-09 puro (lista 2.4 vacía = comportamiento existente intacto).
        }

        var bloque = leaf.BalanceSc.Ases.SingleOrDefault(b => b.Ase.Id == leaf.Ase.Id)
            ?? throw new CalculoInvalidoException(
                $"El leaf del ASE {leaf.Ase.Id} no trae su fila de balance SC para escribir.");

        var editables = WorkbookLeafCellMapBalanceSc.ObtenerEditables(leaf.Ase.Id);
        EscribirValorNumerico(workbookPart, HojaBce, editables.Contribucion, bloque.Contribucion, $"BCE.Contribucion.ASE{leaf.Ase.Id}");
        EscribirValorNumerico(workbookPart, HojaBce, editables.Subsidio, bloque.Subsidio, $"BCE.Subsidio.ASE{leaf.Ase.Id}");
    }

    /// <summary>
    /// HU-11 (2.5): escribe los operandos editables de SALDOS POR NOTA y RETRIBUCION NEGATIVA
    /// del ASE (bloques de valores T0-0.7) en la MISMA pasada atómica HU-07..HU-10 (D2a/G8).
    /// <c>AjustesSfT == null</c> = comportamiento HU-10 puro (Q1, G3). Los visibles Cn-In de
    /// cada bloque, la hoja AJUSTES-SF-T, INTERVENTORIA y ANT EXT-REV son fórmulas protegidas
    /// (mapa <see cref="WorkbookLeafCellMapAjustesSfT.Protegidas"/> + guard de
    /// <see cref="EscribirValorNumerico"/>) → jamás se escriben. La columna I (Especiales) del
    /// template se escribe con 0 (fuente sin esa columna, T0-0.5).
    /// </summary>
    private static void EscribirCeldasAjustesSfT(WorkbookPart workbookPart, WorkbookLeafInputs leaf)
    {
        if (leaf.AjustesSfT is null)
        {
            return; // HU-10 puro (Q1): sin escrituras 2.5.
        }

        foreach (var (celda, valor) in leaf.AjustesSfT.SaldosNotas.Celdas)
        {
            EscribirValorNumerico(workbookPart, HojaSaldosNotas, celda, valor, $"SaldosNotas.ASE{leaf.Ase.Id}.{celda}");
        }

        foreach (var (celda, valor) in leaf.AjustesSfT.RetribucionNegativa.Celdas)
        {
            EscribirValorNumerico(workbookPart, HojaRetribucionNegativa, celda, valor, $"RetribucionNegativa.ASE{leaf.Ase.Id}.{celda}");
        }
    }

    /// <summary>
    /// Plan 29 (T3, Unidad R — escritura): escribe el desglose-detalle por componente de las hojas
    /// <c>Rem. Anticipos R2</c> y <c>Reversion Pagos R4</c> en la MISMA pasada atómica de los
    /// agregados (V2: <c>E43</c>/<c>D73</c> intactos) y el espejo R1. Solo en Q2: el reader puebla
    /// <c>DetalleR2</c>/<c>DetalleR4</c> únicamente en Q2, por lo que en Q1 el método es no-op.
    ///
    /// Resuelve la fila destino por LABEL (firma A|B|C|D para R2, A|B|C para R4 leída de la propia
    /// plantilla) con una alineación LCS que tolera la malla derivada de agosto: las filas-fuente
    /// sin contraparte en la plantilla se omiten y las filas-destino sin fuente quedan en su 0.
    /// Las columnas se resuelven por ENCABEZADO (<see cref="WorkbookLeafCellMapDetalleR2R4.ColumnasR2"/>
    /// / <c>ColumnasR4</c>), nunca por índice. Cero insert/delete de filas (D-C); el guard de
    /// <see cref="EscribirValorNumerico"/> impide tocar fórmulas (destino-detalle fórmula ⇒
    /// <c>ERR-PLANTILLA</c>, nunca sobrescritura — R-R-3).
    /// </summary>
    private static void EscribirCeldasDetalleR2R4(WorkbookPart workbookPart, WorkbookLeafInputs leaf)
    {
        if (leaf.DetalleR2 is not null)
        {
            EscribirDetalleR2(workbookPart, leaf.DetalleR2);
        }

        if (leaf.DetalleR4 is not null)
        {
            EscribirDetalleR4(workbookPart, leaf.DetalleR4);
        }
    }

    private static void EscribirDetalleR2(WorkbookPart workbookPart, DetalleR2AseInputs detalle)
    {
        var aseId = detalle.Ase.Id;
        if (!WorkbookLeafCellMapDetalleR2R4.BloquesR2.TryGetValue(aseId, out var bloque))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"No hay bloque de detalle R2 declarado para el ASE {aseId}. El mapa T0b no cubre ese ASE — fail-fast, nunca se inventa la fila.");
        }

        var hoja = WorkbookLeafCellMapDetalleR2R4.HojaR2;
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));
        var pares = EmparejarBloque(
            workbookPart,
            worksheet,
            bloque,
            detalle.Filas.Select(f => f.Firma).ToList(),
            columnas: ["A", "B", "C", "D"]);

        foreach (var (filaDestino, indiceFuente) in pares)
        {
            var filaFuente = detalle.Filas[indiceFuente];
            foreach (var (encabezado, columna) in WorkbookLeafCellMapDetalleR2R4.ColumnasR2)
            {
                var valor = encabezado is null ? 0m : filaFuente.Valor(encabezado) ?? 0m;
                var celda = columna + filaDestino.ToString(CultureInfo.InvariantCulture);
                EscribirValorNumerico(workbookPart, hoja, celda, valor, $"DetalleR2.ASE{aseId}.{celda}");
            }
        }
    }

    private static void EscribirDetalleR4(WorkbookPart workbookPart, DetalleR4AseInputs detalle)
    {
        var aseId = detalle.Ase.Id;
        if (!WorkbookLeafCellMapDetalleR2R4.BloquesR4.TryGetValue(aseId, out var bloque))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"No hay bloque de detalle R4 declarado para el ASE {aseId}. El mapa T0b no cubre ese ASE — fail-fast, nunca se inventa la fila.");
        }

        var hoja = WorkbookLeafCellMapDetalleR2R4.HojaR4;
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));
        var pares = EmparejarBloque(
            workbookPart,
            worksheet,
            bloque,
            detalle.Filas.Select(f => f.Firma).ToList(),
            columnas: ["A", "B", "C"]);

        foreach (var (filaDestino, indiceFuente) in pares)
        {
            var filaFuente = detalle.Filas[indiceFuente];
            foreach (var (encabezado, columna) in WorkbookLeafCellMapDetalleR2R4.ColumnasR4)
            {
                var valor = encabezado is null ? 0m : filaFuente.Valor(encabezado) ?? 0m;
                var celda = columna + filaDestino.ToString(CultureInfo.InvariantCulture);
                EscribirValorNumerico(workbookPart, hoja, celda, valor, $"DetalleR4.ASE{aseId}.{celda}");
            }
        }
    }

    /// <summary>
    /// Empareja las filas de un bloque-destino (labels leídos de la plantilla) con las firmas de las
    /// filas-fuente mediante LCS (ver <see cref="AlinearFirmas"/>). Devuelve pares
    /// (fila-destino absoluta, índice-fuente). Solo las filas emparejadas se escriben: una fila
    /// destino sin contraparte queda en su 0 de plantilla; una fila fuente intercalada sin destino
    /// se omite. Cero insert/delete (D-C).
    /// </summary>
    private static IReadOnlyList<(int FilaDestino, int IndiceFuente)> EmparejarBloque(
        WorkbookPart workbookPart,
        Worksheet worksheet,
        (int FilaInicio, int FilaFin) bloque,
        IReadOnlyList<string> firmasFuente,
        string[] columnas)
    {
        var firmasDestino = new List<string>(bloque.FilaFin - bloque.FilaInicio + 1);
        for (var fila = bloque.FilaInicio; fila <= bloque.FilaFin; fila++)
        {
            firmasDestino.Add(NormalizarFirma(LeerFirmaDestino(workbookPart, worksheet, fila, columnas)));
        }

        var normalizadasFuente = firmasFuente.Select(NormalizarFirma).ToList();
        return AlinearFirmas(firmasDestino, normalizadasFuente)
            .Select(p => (bloque.FilaInicio + p.IndiceDestino, p.IndiceFuente))
            .ToList();
    }

    /// <summary>
    /// Empareja dos secuencias de firmas por la subsecuencia común más larga (LCS). Devuelve los
    /// índices (0-based) emparejados; las posiciones no emparejadas son inserciones/borrados.
    /// </summary>
    private static IReadOnlyList<(int IndiceDestino, int IndiceFuente)> AlinearFirmas(
        IReadOnlyList<string> firmasDestino,
        IReadOnlyList<string> firmasFuente)
    {
        var n = firmasDestino.Count;
        var m = firmasFuente.Count;
        var dp = new int[n + 1, m + 1];
        for (var i = n - 1; i >= 0; i--)
        {
            for (var j = m - 1; j >= 0; j--)
            {
                dp[i, j] = string.Equals(firmasDestino[i], firmasFuente[j], StringComparison.OrdinalIgnoreCase)
                    ? dp[i + 1, j + 1] + 1
                    : Math.Max(dp[i + 1, j], dp[i, j + 1]);
            }
        }

        var pares = new List<(int, int)>();
        var x = 0;
        var y = 0;
        while (x < n && y < m)
        {
            if (string.Equals(firmasDestino[x], firmasFuente[y], StringComparison.OrdinalIgnoreCase))
            {
                pares.Add((x, y));
                x++;
                y++;
            }
            else if (dp[x + 1, y] >= dp[x, y + 1])
            {
                x++;
            }
            else
            {
                y++;
            }
        }

        return pares;
    }

    /// <summary>
    /// Lee la firma de una fila-destino concatenando el texto de las columnas indicadas (labels de
    /// la plantilla), normalizada para comparar con la firma de la fuente.
    /// </summary>
    private static string LeerFirmaDestino(WorkbookPart workbookPart, Worksheet worksheet, int fila, string[] columnas)
    {
        var partes = new string[columnas.Length];
        for (var i = 0; i < columnas.Length; i++)
        {
            var celda = columnas[i] + fila.ToString(CultureInfo.InvariantCulture);
            var cell = ObtenerCelda(worksheet, celda);
            partes[i] = cell is null ? string.Empty : LeerTextoCelda(cell, workbookPart).Trim();
        }

        return string.Join("|", partes);
    }

    private static string NormalizarFirma(string firma)
    {
        var partes = firma.Split('|');
        for (var i = 0; i < partes.Length; i++)
        {
            partes[i] = partes[i].Trim();
        }

        return string.Join("|", partes);
    }

    /// <summary>
    /// Plan 29 (T5, Unidad P — R-P-3 / T0f): puebla como LITERAL el total de control de recaudo en
    /// <c>'Valida - Control Recaudo'!F10</c>
    /// (<see cref="WorkbookLeafCellMapValidaciones.CeldaControlRecaudoTotal"/>).
    ///
    /// El valor es la Σ de los totales de las hojas <c>Recaudo *</c>
    /// (<see cref="RecaudoEmpresaInputs.Total"/> = fila 27 de la columna de valor de la quincena:
    /// D27 en Q1, F27 en Q2). Es el MISMO origen que alimenta
    /// <c>VALIDACION_TOTAL!C9 = SUM(C3:C8)</c>: cada <c>C3:C8</c> es <c>Σ 'Recaudo *'!F21:F26</c>
    /// por el TEXTO de su fórmula y, por construcción del template, <c>F27 = Σ F21:F26</c>. No se
    /// recalcula por otro camino: se reusan los inputs <c>Recaudo *</c> ya leídos por el reader.
    ///
    /// Sin este literal, post-recálculo <c>C9 (recaudo del app) ≠ 0 = F10</c> ⇒
    /// <c>C15 = (C14 = C9)</c> queda FALSE (T0f). Guard anti-fórmula intacto: si F10 fuera fórmula →
    /// <c>ERR-PLANTILLA</c> (nunca sobrescritura). Sin hojas <c>Recaudo *</c> aportadas (path legacy
    /// HU-07 puro) es no-op: no se inventa un 0.
    /// </summary>
    private static void EscribirControlRecaudoTotal(WorkbookPart workbookPart, IReadOnlyList<WorkbookLeafInputs> leafInputs)
    {
        // Los 5 leafs comparten la MISMA lista de Recaudos (ProcesadorPeriodo la lee una vez): se
        // deduplica por empresa para no sumar la lista 5 veces si el orquestador la replicara.
        var recaudos = leafInputs
            .SelectMany(l => l.Recaudos)
            .GroupBy(r => r.Empresa.Id)
            .Select(g => g.First())
            .ToList();
        if (recaudos.Count == 0)
        {
            return; // Sin Recaudo * no hay total que poblar (HU-07 puro / single-ASE legacy).
        }

        var totalControl = recaudos.Sum(r => r.Total);
        EscribirValorNumerico(
            workbookPart,
            WorkbookLeafCellMapValidaciones.HojaValidaControlRecaudo,
            WorkbookLeafCellMapValidaciones.CeldaControlRecaudoTotal,
            totalControl,
            $"ValidaControlRecaudo.{WorkbookLeafCellMapValidaciones.CeldaControlRecaudoTotal}");
    }

    /// <summary>
    /// Plan 28 (Unidad F / D-C/D-D/D-E): sella las fechas del período en
    /// <c>CONSOLIDADO_TOTAL RECAUDO</c>: <c>FechaDesde → G7</c> y <c>FechaHasta → K7</c>
    /// (columna REAL del template; T0 fijó K7 — J7 es el rótulo "Feha Hasta:").
    ///
    /// SOLO valores (seriales OADate) sobre celdas estáticas, con el guard anti-fórmula de
    /// <see cref="EscribirValorNumerico"/> como red (celda destino fórmula → ERR-PLANTILLA
    /// nombrando hoja+celda, sin escritura). Si el resultado no trae fechas (path single-ASE sin
    /// R10) es no-op; si trae una sola, fail-fast (nunca sello a medias). El origen es el R10 ya
    /// leído por <c>ProcesadorPeriodo</c>; este método no lee I/O.
    ///
    /// Plan 31 (T3, R-S-2/D-F): extiende el sello al PROCESO — <c>FechaProceso → DetRetri*/DetValiRetri*
    /// D6</c> y <c>HoraProceso → DetRetri*/DetValiRetri* D7</c> (nombres por
    /// <see cref="NombresHojaPeriodo"/>). Origen congelado por mini-T0 (R-S-1): el encabezado del
    /// R10 (D6 "Fecha de Proceso" / D7 "Hora"), que coincide al carácter con el manual del
    /// administrativo (julio 04/08/2026 10:15 AM; agosto 02/09/2026 07:42 AM). Las 4 celdas traen
    /// <c>numFmtId=49</c> ("@" Texto) y valor de texto: se sellan como TEXTO (no serial OADate, que
    /// mostraría el número crudo) preservando el estilo, con guard anti-fórmula. <c>CONSOLIDADO!D6</c>
    /// NO se toca (ya viene sellado en la base; R-S-3 no-duplicación). Sin sello de proceso legible
    /// el orquestador ya falló con <c>ERR-FORMATO-FUENTE</c>; aquí solo se escribe.
    /// </summary>
    internal static void EscribirFechasPeriodo(WorkbookPart workbookPart, ResultadoRemuneracion resultado)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);
        ArgumentNullException.ThrowIfNull(resultado);

        var hayRango = resultado.FechaDesde.HasValue || resultado.FechaHasta.HasValue;
        var hayProceso = resultado.FechaProceso.HasValue || resultado.HoraProceso.HasValue;
        if (!hayRango && !hayProceso)
        {
            return; // Path single-ASE sin R10: comportamiento previo intacto (no se inventan fechas).
        }

        if (resultado.FechaDesde.HasValue != resultado.FechaHasta.HasValue)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El sello de fechas del período {resultado.Periodo.CodigoCompleto} requiere 'Fecha Desde' y 'Fecha Hasta'; se recibió una sola. Fail-fast, nunca se sella a medias.");
        }

        if (resultado.FechaProceso.HasValue != resultado.HoraProceso.HasValue)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El sello de proceso del período {resultado.Periodo.CodigoCompleto} requiere 'Fecha de Proceso' y 'Hora'; se recibió una sola. Fail-fast, nunca se sella a medias.");
        }

        if (resultado.FechaDesde.HasValue)
        {
            EscribirFechaSerie(workbookPart, "G7", resultado.FechaDesde.Value, "FechaDesde");
            EscribirFechaSerie(workbookPart, "K7", resultado.FechaHasta!.Value, "FechaHasta");
        }

        if (resultado.FechaProceso.HasValue)
        {
            SellarProcesoEnDetRetri(workbookPart, resultado.Periodo, resultado.FechaProceso.Value, resultado.HoraProceso!.Value);
        }
    }

    /// <summary>
    /// Plan 31 (T3, R-S-2/D-F): sella <c>Fecha de Proceso</c> (D6) y <c>Hora</c> (D7) en las hojas
    /// <c>DetRetri{AAAAMMQ}</c> y <c>DetValiRetri{AAAAMMQ}</c> (nombres compuestos por
    /// <see cref="NombresHojaPeriodo"/>, nunca detectados enumerando hojas).
    ///
    /// Representación = TEXTO idéntica a la del manual/base/R10 (mini-T0 R-S-1): las 4 celdas traen
    /// formato Texto (<c>numFmtId=49</c>, valor compartido). Escribir un serial OADate mostraría el
    /// número crudo — no la fecha — así que se sella el texto <c>dd/MM/yyyy</c> / <c>hh:mm AM/PM</c>
    /// preservando el estilo. La "Hora" es hora-del-día y su origen es el R10 (D7), no la hora de la
    /// corrida actual: el manual del administrativo coincide con el R10 al carácter.
    /// </summary>
    private static void SellarProcesoEnDetRetri(WorkbookPart workbookPart, Periodo periodo, DateTime fechaProceso, TimeSpan horaProceso)
    {
        var textoFecha = fechaProceso.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        var textoHora = FormatearHoraProceso(horaProceso);

        foreach (var hoja in new[] { NombresHojaPeriodo.DetRetri(periodo), NombresHojaPeriodo.DetValiRetri(periodo) })
        {
            EscribirValorTexto(workbookPart, hoja, "D6", textoFecha, $"FechaProceso.{hoja}.D6");
            EscribirValorTexto(workbookPart, hoja, "D7", textoHora, $"HoraProceso.{hoja}.D7");
        }
    }

    /// <summary>
    /// Plan 31 (T3): formatea la hora-del-día del sello de proceso como el manual/base/R10
    /// (<c>"07:42 AM"</c>: 12h, 2 dígitos, designador AM/PM en mayúsculas).
    /// </summary>
    private static string FormatearHoraProceso(TimeSpan hora)
    {
        var baseFecha = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).Add(hora);
        return baseFecha.ToString("hh:mm tt", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Plan 31 (T3): escribe un TEXTO sobre una celda EXISTENTE preservando su estilo (a diferencia
    /// de <see cref="EscribirValorNumerico"/>, que fija <c>CellValues.Number</c>). Guard
    /// anti-fórmula intacto: celda destino fórmula → <c>ERR-PLANTILLA</c>; celda ausente → fail-fast.
    /// </summary>
    private static void EscribirValorTexto(WorkbookPart workbookPart, string hoja, string celda, string valor, string nombre)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));
        var cell = ObtenerCelda(worksheet, celda)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La celda '{hoja}!{celda}' no existe en la plantilla para sellar '{nombre}'. Fail-fast, nunca se inventa la celda.");

        if (cell.CellFormula is not null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La celda sello '{nombre}' ({hoja}!{celda}) es fórmula en la plantilla. Plan 31 no puede sobrescribir fórmulas.");
        }

        // Representación de texto (inline string) para preservar el estilo Texto de la celda; se
        // limpia cualquier <v> previo (una celda no puede tener <v> y <is> a la vez).
        cell.CellValue = null;
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(new Text(valor));
    }

    /// <summary>
    /// Plan 28 (Unidad F): escribe una fecha como serial OADate (double→decimal) sobre una celda
    /// EXISTENTE de <c>CONSOLIDADO_TOTAL RECAUDO</c>. No crea celdas ni toca estilos: preserva el
    /// formato de fecha de la celda (T0: G7/K7 traen estilo s=9, numFmt 14). Fail-fast si la celda
    /// no existe; el guard de <see cref="EscribirValorNumerico"/> cubre el caso celda-fórmula.
    /// </summary>
    private static void EscribirFechaSerie(WorkbookPart workbookPart, string celda, DateTime fecha, string etiqueta)
    {
        var worksheet = ObtenerHoja(workbookPart, HojaConsolidado, nameof(GenerarWorkbook));
        if (ObtenerCelda(worksheet, celda) is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La celda '{HojaConsolidado}!{celda}' no existe en la plantilla para sellar '{etiqueta}' del período. Fail-fast, nunca se inventa la celda.");
        }

        EscribirValorNumerico(workbookPart, HojaConsolidado, celda, (decimal)fecha.ToOADate(), $"{etiqueta}.{HojaConsolidado}.{celda}");
    }

    private static void EscribirValorNumerico(WorkbookPart workbookPart, string hoja, string celda, decimal valor, string nombre)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));
        var cell = ObtenerOCrearCelda(worksheet, celda);

        if (cell.CellFormula is not null)
        {
            // HU-15 (W-2.3, D5): una celda editable del mapa que es fórmula en la plantilla =
            // plantilla incompatible (nunca sobrescribir fórmulas) → ERR-PLANTILLA.
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La celda leaf {nombre} ({hoja}!{celda}) es fórmula en la plantilla. HU-05 no puede sobrescribir fórmulas.");
        }

        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture));
    }

    private static Cell ObtenerOCrearCelda(Worksheet worksheet, string cellReference)
    {
        var existente = ObtenerCelda(worksheet, cellReference);
        if (existente is not null)
        {
            return existente;
        }

        var sheetData = worksheet.Elements<SheetData>().FirstOrDefault()
            ?? throw new CalculoInvalidoException($"La hoja no tiene SheetData para crear la celda '{cellReference}'.");

        var (_, filaNumero) = ParsearReferencia(cellReference);
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == filaNumero);
        if (row is null)
        {
            row = new Row { RowIndex = filaNumero };
            var siguientes = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value > filaNumero);
            if (siguientes is null)
            {
                sheetData.Append(row);
            }
            else
            {
                sheetData.InsertBefore(row, siguientes);
            }
        }

        var nueva = new Cell { CellReference = cellReference };
        var siguienteCelda = row.Elements<Cell>().FirstOrDefault(c =>
            string.Compare(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase) > 0);
        if (siguienteCelda is null)
        {
            row.Append(nueva);
        }
        else
        {
            row.InsertBefore(nueva, siguienteCelda);
        }

        return nueva;
    }

    private static (string Columna, uint Fila) ParsearReferencia(string cellReference)
    {
        var match = Regex.Match(cellReference, @"^(?<col>[A-Za-z]+)(?<row>\d+)$");
        if (!match.Success)
        {
            throw new CalculoInvalidoException($"Referencia de celda inválida: '{cellReference}'.");
        }

        return (match.Groups["col"].Value.ToUpperInvariant(), uint.Parse(match.Groups["row"].Value, CultureInfo.InvariantCulture));
    }

    private sealed class LeafCellComparer : IEqualityComparer<(string Hoja, string Celda)>
    {
        public bool Equals((string Hoja, string Celda) x, (string Hoja, string Celda) y) =>
            string.Equals(x.Hoja, y.Hoja, StringComparison.OrdinalIgnoreCase)
            && string.Equals(x.Celda, y.Celda, StringComparison.OrdinalIgnoreCase);

        public int GetHashCode((string Hoja, string Celda) obj) =>
            HashCode.Combine(obj.Hoja.ToUpperInvariant(), obj.Celda.ToUpperInvariant());
    }
}
