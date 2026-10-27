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

    /// <summary>
    /// LEGACY-only, FUERA DE PARIDAD (Plan 54, H1.2, D2): API histórica de detalle R1 que no expone los
    /// operandos del template (F25/F41/L25/F30/F10/L10) y por diseño falla honestamente con
    /// <c>ERR-PAYLOAD</c>. No forma parte del camino de producción (la escritura R1 real la gobiernan
    /// <see cref="EscritorValoresSlotR1"/> + <c>EscribirCeldasLeafPorAse</c>) ni de la paridad manual.
    /// Se conserva por compatibilidad de contrato (<see cref="IPlantillaWriter"/>) y para las
    /// herramientas de diagnóstico. Prohibido añadir lectores nuevos de F30/F10/L10.
    /// </summary>
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
        VerificadorDestinoProtegido.ValidarBloquesDeRecaudo(workbookPart);

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
        VerificadorDestinoProtegido.ValidarBloquesDeRecaudo(workbookPart);

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

        VerificadorDestinoProtegido.ValidarBloquesDeRecaudo(workbookPart);

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

            // Plan 39 (T2, Phase 0): base del período escrita en la copia de trabajo (el origen se
            // abre SOLO lectura), antes de la validación/escritura. La quincena llega por DOMINIO.
            PreparadorBasePipeline.Ejecutar(origen, leafInputs.Periodo, rutaSalida);

            // Plan 58 (H1.2): foto de la base que recibe el escritor (gate de cierre por bytes).
            var basePreparada = File.ReadAllBytes(rutaSalida);

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, true))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene WorkbookPart válido.");
                ValidarFormulasProtegidas(workbookPart, nameof(GenerarWorkbook), leafInputs.Periodo);
                EscribirCeldasLeaf(workbookPart, leafInputs);

                // Plan 29 (T3, Unidad R): desglose-detalle R2/R4 en la MISMA pasada atómica.
                EscribirCeldasDetalleR2R4(workbookPart, leafInputs);

                // Plan 29 (T5, Unidad P — R-P-3/T0f) + Plan 48 (H1.2): control de recaudo.
                EscribirControlRecaudoTotal(workbookPart, [leafInputs]);
                // Plan 48 (H1b): sella los cachés EVALUADOS de E7/E10/F10 (la v2 los trae como
                // fórmulas); jamás se inventa el valor: si no es resoluble, no se sella.
                SellarCachesControlRecaudo(workbookPart);

                // Plan 28 (Unidad F): sella las fechas del período desde el R10 (solo si el
                // orquestador las aportó). El path single-ASE sin R10 no trae fechas → no-op.
                EscribirFechasPeriodo(workbookPart, resultado);

                var workbookXml = workbookPart.Workbook
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene metadata Workbook válida.");

                // Plan 40 (T2): sellado de cachés computables por read-back (cadena C3:C8→C9→C14→C15)
                // ANTES del saneamiento calcChain, para que el caché de la salida no mienta.
                SellarCachesFormula(workbookPart, leafInputs.Periodo);

                // Plan 28 (Unidad S): punto único de guardado — saneamiento calcChain + fullCalcOnLoad.
                SaneadorCadenaCalculo.Sanear(workbookPart);
                workbookXml.Save();
            }

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, false))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook generado no tiene WorkbookPart válido.");
                ValidarFormulasProtegidas(workbookPart, nameof(GenerarWorkbook), leafInputs.Periodo);
            }

            // Plan 58 (H1.2): gate de integridad del paquete POST-escritura (por bytes, sin Excel).
            ValidadorIntegridadPaquete.ValidarEscritura(basePreparada, File.ReadAllBytes(rutaSalida));
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

        // Plan 21 (T4) / Plan 36 (T2, R-B-3/D-A/D-B): si el orquestador adjuntó el bloque espejo R1
        // de los 5 ASE, la hoja Reporte Componentes R1 se gobierna por el ESCRITOR VALUES-ONLY POR
        // SLOT (EscritorValoresSlotR1): pega SOLO <v> en los slots del contrato T0 de la plantilla
        // estandarizada única (firma + frontera + sinonimia). Cero insert/delete, cero <f>, cero
        // reanclaje: la geometría y las fórmulas del template quedan byte-idénticas (gate R-G-1).
        // Las celdas R1 del mapa absoluto HU-07/HU-12 quedan superseded dentro del bloque y NO se
        // escriben por el mapa. El resto de las hojas queda intacto.
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

            // Plan 39 (T2, Phase 0): prepara la base del período (renombra las 3 hojas de período,
            // reescribe los 11 tokens <f> de período, sella N3 y selecciona L9:L13 por quincena
            // —R-INTERV-Q—) y la escribe en la COPIA de trabajo. El OFICIAL (origen) se abre SOLO
            // lectura y nunca se modifica; corre ANTES de ValidadorBasePeriodo.Exigir y de cualquier
            // escritura de negocio. El procesamiento sigue values-only (cero <f> aguas abajo).
            PreparadorBasePipeline.Ejecutar(origen, periodoReferencia, rutaSalida);

            // Plan 58 (H1.2): foto de la base que recibe el escritor. El gate de cierre exige que la
            // salida conserve EXACTAMENTE el conjunto de fórmulas (celda → texto + t/ref/si).
            var basePreparada = File.ReadAllBytes(rutaSalida);

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, true))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene WorkbookPart válido.");

                // Plan 36 (T2, Fase 1): el R1 se pega values-only por slot (mismo orden del espejo
                // retirado: antes de la escritura de hojas y solo cuando están los 5 bloques). El
                // writer recibe los bloques ya leídos y resuelve cada fila contra la geometría del
                // template; fuente sin slot libre → ERR-SIN-SLOT (nunca insert, nunca remap).
                if (espejoR1)
                {
                    // Plan 36 (T3, R-B-5/D-G): la base del período se valida ANTES de escribir
                    // (hojas del período + N3 + sellos G7/K7/D6 + D6/D7 + 0 tokens de período stale
                    // en <f>). Base equivocada → fail-fast sin salida parcial; el runtime NUNCA
                    // reescribe tokens de período (eso lo hace el procedimiento offline, §2.3-A).
                    ValidadorBasePeriodo.Exigir(workbookPart, periodoReferencia);

                    foreach (var bloque in bloquesEspejo)
                    {
                        EscritorValoresSlotR1.Escribir(workbookPart, bloque);
                    }
                }

                // Plan 36 (T2, D-A/D-E): con el escritor values-only el bloque R1 NUNCA se
                // desplaza, así que la validación PROTEGIDA por direcciones absolutas corre
                // SIEMPRE: desaparece la rama "espejo desplazado" que la salteaba con Δ≠0 (y con
                // ella su sucesor por FORMA recompuesta, que ya no se recompone nada).
                ValidarFormulasProtegidasMultiAse(workbookPart, nameof(GenerarWorkbook), periodoReferencia);

                foreach (var leaf in leafInputs.OrderBy(l => l.Ase.Id))
                {
                    // Plan 54 (H1.3, D3): mapa productivo ÚNICO para ambas quincenas (geometría v2);
                    // la rama Q1 del escritor leaf por ASE quedó colapsada.
                    EscribirCeldasLeafPorAse(workbookPart, leaf, omitirR1: espejoR1);
                    EscribirCeldasEmpresa(workbookPart, leaf, omitirR1: espejoR1);
                    EscribirCeldasBanco(workbookPart, leaf);
                    EscribirCeldasBalanceSc(workbookPart, leaf);
                    EscribirCeldasAjustesSfT(workbookPart, leaf);

                    // Plan 29 (T3, Unidad R): desglose-detalle R2/R4 en la MISMA pasada atómica
                    // (tras los agregados E43/D73 y el espejo R1; el guard anti-fórmula es la red).
                    EscribirCeldasDetalleR2R4(workbookPart, leaf);
                }

                // HU-12 (2.6 ampliada, V0.4) / Plan 59 (H4): el asiento DetRetri (columna de valor de
                // la tabla por ASE, encabezado 'Recaudo Total') se escribe en AMBAS quincenas. La v3
                // declara ese destino como VALOR en las dos —medido por bytes:
                // DetRetri2026092!D9:D14 son literales, sin <f>— y ValidarDetRetriEditable ya exige
                // que siga siendo editable. La rama por quincena dejaba en Q1 el literal 0 de la
                // plantilla dentro del entregable; en Q1 el valor ya está probado contra el R10
                // (HU-20/G3), así que no hay oráculo que preservar en esa celda.
                EscribirCeldasDetRetri(workbookPart, leafInputs);

                // Plan 59 (§11 — FRENTE 1): el DESGLOSE trazable (DetRetri J/L/Cierre · DetValiRetri
                // I/O/Cierre) se escribe en AMBAS quincenas. La plantilla declara esas celdas como
                // VALOR en las dos (medido por bytes: literales sin <f>), la verificación por ROL del
                // detalle ya las exige en ambas y el oráculo administrativo las trae llenas en Q1. La
                // ÚNICA pieza que depende de la quincena es el costo de interventoría, y su selección
                // vive en el DOMINIO (InterventoriaDeclarada.PorQuincena: primera 'N' / segunda 'M'),
                // nunca aquí. Antes esta rama dejaba J/K/L e I/O en el literal 0 de la plantilla
                // durante Q1.
                EscribirDesgloseTrazableDetRetri(workbookPart, leafInputs);

                // Plan 29 (T5, Unidad P — R-P-3/T0f) + Plan 48 (H1.2): control de recaudo.
                EscribirControlRecaudoTotal(workbookPart, leafInputs);
                // Plan 48 (H1b): sella los cachés EVALUADOS de E7/E10/F10 (fórmulas v2).
                SellarCachesControlRecaudo(workbookPart);

                // Plan 28 (Unidad F): sello de fechas del período desde el R10 del período.
                EscribirFechasPeriodo(workbookPart, resultado);

                var workbookXml = workbookPart.Workbook
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene metadata Workbook válida.");

                // Plan 40 (T2): sellado de cachés computables por read-back (cadena C3:C8→C9→C14→C15).
                SellarCachesFormula(workbookPart, periodoReferencia);

                // Plan 28 (Unidad S): punto único de guardado — saneamiento calcChain + fullCalcOnLoad.
                SaneadorCadenaCalculo.Sanear(workbookPart);
                workbookXml.Save();
            }

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, false))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook generado no tiene WorkbookPart válido.");

                // Plan 36 (T2): sin desplazamiento de R1 (values-only sobre la geometría del
                // template) la validación protegida post-escritura es la MISMA que la pre-escritura,
                // en ambos caminos y sin ramas por período ni por "espejo desplazado".
                ValidarFormulasProtegidasMultiAse(workbookPart, nameof(GenerarWorkbook), periodoReferencia);

                // Plan 59 (H4): gate de VALORES del asiento DetRetri sobre el archivo recién
                // guardado — lo escrito contra el esperado del dominio. Es el contrapeso del gate
                // de forma: la forma puede cerrar (0 − 0) con el asiento vacío, los valores no.
                VerificarAsientoDetRetri(workbookPart, leafInputs);

                // Plan 59 (§11-bis): mismo contrapeso para el desglose-detalle de R2/R4. El bloque se
                // escribe ahora en AMBAS quincenas; este gate relee lo guardado y exige la cadena de
                // agregación del bloque (detalle → totales → cierre de sección → gran total) y que el
                // gran total coincida con lo que la fuente declara por componente. Un bloque vacío en
                // Q1 dejaba CONSOLIDADO!D28:D32 en cero y la corrida certificaba OK.
                VerificarDetalleR2R4(workbookPart, leafInputs);
            }

            // Plan 58 (H1.2): gate de integridad del paquete POST-escritura (por bytes, sin Excel).
            // Grupos shared, fórmulas preservadas, duplicadas, orden, sharedStrings y hojas. Un
            // defecto aborta con ERR-INTEGRIDAD y la atomicidad borra el parcial: nunca se entrega un
            // workbook que Excel tendría que reparar.
            ValidadorIntegridadPaquete.ValidarEscritura(basePreparada, File.ReadAllBytes(rutaSalida));
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
    /// Plan 21 (T3, R-E-2/R-E-3/R-E-4) / Plan 36 (T2, R-B-3/D-A/D-B): aplica el bloque espejo R1
    /// sobre una COPIA de la plantilla estandarizada única.
    ///
    /// Tras el retiro del motor geométrico (Plan 36/T2) la operación es <b>values-only por slot</b>:
    /// pega SOLO <c>&lt;v&gt;</c> en los slots del contrato T0 de cada ASE (firma + frontera +
    /// sinonimia), sin insertar/borrar filas y sin escribir <c>&lt;f&gt;</c>. La geometría y las
    /// fórmulas del template quedan byte-idénticas; lo no entregado queda en el cero del template y
    /// una fuente sin slot libre hace fail-fast <c>ERR-SIN-SLOT</c>. No recalcula aritmética de
    /// negocio.
    /// </summary>
    /// <inheritdoc />
    public void EscribirEspejoR1(string rutaPlantillaOrigen, string rutaSalida, IReadOnlyList<BloqueEspejoAseInputs> bloques, Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(rutaPlantillaOrigen);
        ArgumentNullException.ThrowIfNull(rutaSalida);
        ArgumentNullException.ThrowIfNull(bloques);
        ArgumentNullException.ThrowIfNull(periodo);

        if (bloques.Count == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                "El espejo R1 no recibió bloques de ASE; no hay nada que pegar en los slots del contrato.");
        }

        var origen = ValidarArchivo(rutaPlantillaOrigen, nameof(EscribirEspejoR1));
        ValidarRutaSalida(rutaSalida);
        ValidarNoInPlace(origen, rutaSalida);

        try
        {
            var directorioSalida = Path.GetDirectoryName(rutaSalida);
            if (!string.IsNullOrWhiteSpace(directorioSalida))
            {
                Directory.CreateDirectory(directorioSalida);
            }

            File.Copy(origen, rutaSalida, overwrite: true);

            // Plan 39 (T2, Phase 0): base del período escrita en la copia de trabajo (el origen se
            // abre SOLO lectura), antes de la escritura values-only por slot. La quincena = dominio.
            PreparadorBasePipeline.Ejecutar(origen, periodo, rutaSalida);

            // Plan 58 (H1.2): foto de la base que recibe el escritor (gate de cierre por bytes).
            var basePreparada = File.ReadAllBytes(rutaSalida);

            using (var workbook = SpreadsheetDocument.Open(rutaSalida, true))
            {
                var workbookPart = workbook.WorkbookPart
                    ?? throw new CalculoInvalidoException(CodigoError.Plantilla, "El workbook abierto no tiene WorkbookPart válido.");

                foreach (var bloque in bloques.OrderBy(b => b.Ase.Id))
                {
                    EscritorValoresSlotR1.Escribir(workbookPart, bloque);
                }

                // Plan 40 (T2): sellado de cachés computables por read-back (cadena C3:C8→C9→C14→C15)
                // en el camino standalone del espejo (read-back de lo que la base traiga).
                SellarCachesFormula(workbookPart, periodo);

                // Plan 28 (Unidad S): el espejo standalone ya no mueve filas, pero se conserva el
                // punto de guardado con el saneamiento calcChain + fullCalcOnLoad (misma doctrina
                // del writer multi-ASE). El sellado (Plan 40) toca SOLO <v> de la cadena C; las <f>
                // quedan intactas.
                SaneadorCadenaCalculo.Sanear(workbookPart);
                workbookPart.Workbook?.Save();
            }

            // Plan 58 (H1.2): gate de integridad del paquete POST-escritura (por bytes, sin Excel).
            ValidadorIntegridadPaquete.ValidarEscritura(basePreparada, File.ReadAllBytes(rutaSalida));
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
        _ = worksheet;
        // Plan 59 (H2.2): las 30 celdas del CONSOLIDADO y su cierre se verifican por ROL + FORMA
        // (tabla por ASE + destino resuelto por identidad). Ya no hay diccionario de celdas ni lista de
        // fragmentos: los fragmentos también eran geometría (§2-D2.3-bis).
        VerificadorDestinoProtegido.ValidarConsolidado(workbookPart);
    }

    private static void ValidarCadenaFormulaR1R2R4(WorkbookPart workbookPart)
    {
        // Plan 59 (H2.2): los totales de bloque de R1/R2/R4 se verifican por identidad (bloque + rótulo
        // del cuadro de totales) y por forma derivada de sus cierres. Cero literales de celda.
        VerificadorDestinoProtegido.ValidarBloquesDeRecaudo(workbookPart);
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

    /// <summary>
    /// Plan 59 (H5): protecciones del camino SINGLE-ASE. Antes era una foto v1
    /// (<see cref="WorkbookLeafCellMap.ProtectedFormulas"/>: F75/F77/E47/D73 y las filas del
    /// CONSOLIDADO) que con la v3 ya no describe la plantilla — F75 es valor, no fórmula —. Ahora usa el
    /// MISMO verificador derivado que el camino multi-ASE: las protecciones se exigen por ROL + FORMA
    /// sobre la geometría que la plantilla declara, y la hoja de detalle del período se resuelve
    /// leyéndola del workbook (<see cref="HojasPeriodo"/>). Cero literales de fila/celda/período.
    /// </summary>
    private static void ValidarFormulasProtegidas(WorkbookPart workbookPart, string operacion, Periodo periodo) =>
        ValidarFormulasProtegidasMultiAse(workbookPart, operacion, periodo);

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
        // Plan 59 (§11 — FRENTE 2): la protección del DETALLE del período NO depende de la quincena.
        // El escritor escribe el asiento en las dos, la plantilla declara esas celdas como VALOR en
        // las dos (medido por bytes) y la gramática de roles/formas es la misma hoja con otro nombre.
        // Antes estas dos verificaciones vivían SOLO dentro de la rama Q2: en Q1 el asiento se
        // escribía sin que nadie comprobara que la celda siguiera siendo editable ni que la forma de
        // sus familias fuera la que la plantilla declara. Una coherencia desactivada en silencio.
        ObtenerHoja(workbookPart, NombresHojaPeriodo.DetRetri(periodo), operacion);
        VerificadorDestinoProtegido.ValidarDetRetriEditable(workbookPart, periodo);
        DetallePeriodoDerivado.Validar(workbookPart, periodo);

        if (periodo.NumeroQuincena == 2)
        {
            ValidarFormulasProtegidasMultiAseQ2(workbookPart, operacion, periodo);
            return;
        }

        // Plan 54 (H1.3, D3) + Plan 59 (H2.2): las fórmulas protegidas de R1/R2/R4 y de las filas del
        // CONSOLIDADO se verifican por ROL + FORMA derivada de la geometría que la plantilla declara.
        // El mapa literal por ASE (geometría v1) queda fuera del camino productivo: con la v3 sus
        // celdas y sus operandos ya no son los del template.
        VerificadorDestinoProtegido.ValidarBloquesDeRecaudo(workbookPart);
        // RESIDUO DECLARADO (STOP por concepto): las hojas REMUNERACION_* por empresa siguen validadas
        // por su mapa histórico, porque el T0 de este plan no midió su gramática de tablas.
        VerificadorDestinoProtegido.ValidarConsolidado(workbookPart);

        // HU-09 (2.3, D6): mapa de fórmulas protegidas del reporte por banco (consolidado 1–7,
        // Total de bloque, verificación I/J, validación 59–80 y TOTAL RECAUDO). Nunca se escriben.
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapReporteBanco.Protegidas)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-10 (2.4, D6): protecciones de la hoja de balance de subsidio/contribución (totales por
        // ASE, diferencias, sumas, cuadre y bloque de validación) y sus dos vínculos externos: el
        // balance del ASE en la hoja de detalle REAL del período y la tabla del CONSOLIDADO.
        // Plan 59 (H5): la foto congelada `WorkbookLeafCellMapBalanceSc.Protegidas` (celdas de la v1 y
        // nombre de hoja de un período concreto) salió del camino productivo: el verificador deriva
        // bandas, columnas y filas de lo que la hoja declara y resuelve la hoja del período leyendo el
        // workbook (cero literales de período).
        VerificadorBalanceScDerivado.Validar(workbookPart, periodo);

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
        // Plan 59 (H2.2): los totales de bloque de R1/R2/R4 se verifican por ROL (bloque + rótulo del
        // cuadro de totales) y FORMA derivada de sus cierres; el CONSOLIDADO, por tabla + ASE; el
        // DetRetri, por etiqueta (slot de valor). Se retiraron los mapas Q2 de celdas editables y
        // protegidas y su lista de fragmentos.
        // RESIDUO DECLARADO (STOP por concepto): las hojas REMUNERACION_* por empresa siguen validadas
        // por su mapa histórico (ProtegidasAdicionalesQ2), porque el T0 de este plan no midió su
        // gramática de tablas. No se inventa su derivación.
        VerificadorDestinoProtegido.ValidarBloquesDeRecaudo(workbookPart);
        VerificadorDestinoProtegido.ValidarConsolidado(workbookPart);

        // Plan 59 (§11 — FRENTE 2): la verificación del detalle del período (hoja editable + roles y
        // formas de las tres familias) se ejecuta ANTES de esta rama, para AMBAS quincenas. Aquí ya no
        // se repite: una sola llamada, un solo criterio.

        // HU-09 (2.3): mapa banco parametrizado al período — el Q2 excluye el TOTAL RECAUDO fila 81
        // (C81/D81) que no existe en el template Q2 (T0-0.5; la hoja termina en la fila 79).
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapQ2.BancoProtegidasQ2)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-10 (2.4) + Plan 59 (H5): mismas protecciones derivadas del balance por identidad/forma
        // (hoja y celdas del período resueltas del workbook; cero literales de período).
        VerificadorBalanceScDerivado.Validar(workbookPart, periodo);

        // HU-11 (2.5): cadena AJUSTES-SF-T (Q2).
        foreach (var (hoja, celda, fragmentos) in WorkbookLeafCellMapAjustesSfT.Protegidas)
        {
            var worksheet = ObtenerHoja(workbookPart, hoja, operacion);
            ValidarCeldaTieneFormula(workbookPart, worksheet, celda, fragmentos, hoja, operacion);
        }

        // HU-12 (2.6 ampliada, Requirement 5): DetRetri/DetValiRetri — Plan 59 (H2-bis): las 23 celdas
        // protegidas se verifican por ROL + FORMA sobre la gramática que la propia hoja declara
        // (tres familias de filas por ASE + rótulos de columna), no por la lista congelada de celdas y
        // operandos de la v2 (`WorkbookLeafCellMapQ2.DetRetriProtected`, hoy LEGACY).
        // Plan 59 (§11 — FRENTE 2): la verificación corre en el tronco común (antes de la rama de
        // quincena), así que aquí NO se repite: una sola llamada para ambas quincenas.

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
    /// LEGACY (geometría v1), fuera de paridad (Plan 54, H1.2/H1.3, D2/D3): escribe las celdas
    /// absolutas del mapa v1 (<see cref="WorkbookLeafCellMapPorAse"/>) — F25/F41/L25/F30/F10/L10 para
    /// R1 —. El camino de producción NO pasa por aquí: usa <see cref="EscribirCeldasLeafPorAse"/> con
    /// la geometría v2. Se conserva porque lo consumen fixtures de caracterización
    /// (<c>GoldenMultiAseTests</c>, <c>ParidadCliTests</c>) y la API histórica de leaf.
    /// </summary>
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
    /// Plan 59 (H2.1): escribe las celdas de valor del bloque resolviéndolas por ROL contra la
    /// geometría que la plantilla declara (<see cref="ResolvedorDestinoValor"/>). El transporte
    /// (<c>ValoresPorRol</c>) no trae referencias: las referencias eran el mapa congelado.
    /// Fail-fast si falta un valor (nombra ASE + hoja + rol; nunca 0 silencioso) y el guard de
    /// <see cref="EscribirValorNumerico"/> impide tocar fórmulas.
    /// </summary>
    private static void EscribirCeldasLeafPorAse(WorkbookPart workbookPart, WorkbookLeafInputs leaf, bool omitirR1)
    {
        if (!omitirR1)
        {
            EscribirRolesDeBloque(
                workbookPart, leaf, DelimitadorBloquesHoja.R1, leaf.R1.ValoresPorRol, "R1");
        }

        EscribirRolesDeBloque(
            workbookPart, leaf, DelimitadorBloquesHoja.R2, leaf.R2.ValoresPorRol, "R2");
        EscribirRolesDeBloque(
            workbookPart, leaf, DelimitadorBloquesHoja.R4, leaf.R4.ValoresPorRol, "R4");
    }

    /// <summary>
    /// Pega los valores del bloque en las celdas que cada ROL designa, resolviendo con el resolvedor
    /// de destinos (identidad conceptual, cero coordenadas).
    /// </summary>
    private static void EscribirRolesDeBloque(
        WorkbookPart workbookPart,
        WorkbookLeafInputs leaf,
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        IReadOnlyDictionary<string, decimal> valoresPorRol,
        string sufijoOperacion)
    {
        var bloques = DelimitadorBloquesHoja.Delimitar(workbookPart, hoja);
        var worksheet = ObtenerHoja(workbookPart, hoja.Nombre, $"escritura {sufijoOperacion} por rol");

        foreach (var (clave, valor) in valoresPorRol)
        {
            var rol = ResolvedorDestinoValor.RolValor.DesdeClave(clave);
            var destino = ResolvedorDestinoValor.Resolver(
                workbookPart, worksheet, hoja, bloques, rol.IdentidadEn(hoja.Nombre, leaf.Ase.Id));

            EscribirValorNumerico(
                workbookPart, hoja.Nombre, destino.Celda, valor, $"{sufijoOperacion}.ASE{leaf.Ase.Id}.{clave}");
        }
    }

    /// <summary>
    /// HU-12 (2.6 ampliada, V0.4) / Plan 59 (H4): escribe el asiento DetRetri del período (hoja
    /// <c>DetRetri{AAAAMMQ}</c>) en la MISMA pasada atómica: por ASE, la columna de valor de la tabla
    /// por ASE (encabezado <c>'Recaudo Total'</c>) recibe <see cref="DetRetriQ2Inputs.Detalle"/> =
    /// ROUND(D104:D108) vía <see cref="DetRetriRounder"/> (única regla) y la fila de cierre recibe la
    /// SUMA de las filas por ASE (lo que Excel computa al sumar la columna). La composición está
    /// congelada (probada 5/5 contra el golden, V0.4); nunca se inventa, y el destino se resuelve por
    /// ETIQUETA (tabla por ASE + id leído de la fila), nunca por coordenada.
    ///
    /// Corre en AMBAS quincenas: la plantilla declara ese destino como valor en las dos (ver la nota
    /// del llamador) y el DetRetri de Q1 ya viene validado contra el R10 (HU-20/G3).
    /// </summary>
    private static void EscribirCeldasDetRetri(WorkbookPart workbookPart, IReadOnlyList<WorkbookLeafInputs> leafInputs)
    {
        var conDetalle = leafInputs
            .Where(l => l.DetRetriQ2 is not null)
            .OrderBy(l => l.Ase.Id)
            .ToList();
        if (conDetalle.Count == 0)
        {
            return; // Sin DetRetri en el dominio (modo 1-ASE legado): no hay asiento que escribir.
        }

        // Plan 30 (D-A): la hoja DetRetri se resuelve del dominio del período (no de un literal).
        var periodo = conDetalle[0].Periodo;
        var hojaDetRetri = NombresHojaPeriodo.DetRetri(periodo);
        // Plan 59 (H2.3): el destino se deriva por ETIQUETA (encabezado único de la tabla por ASE +
        // id numérico de la fila). Ya no hay mapa congelado D9..D13.
        var worksheet = ObtenerHoja(workbookPart, hojaDetRetri, "escritura DetRetri por rol");

        foreach (var leaf in conDetalle)
        {
            var destino = ResolvedorDestinoValor.ResolverTabla(
                workbookPart, worksheet,
                new ResolvedorDestinoValor.IdentidadTabla(
                    hojaDetRetri, Titulo: null, leaf.Ase.Id,
                    ResolvedorDestinoValor.EncabezadoRecaudoTotal));

            EscribirValorNumericoEnCeldaExistente(
                workbookPart, hojaDetRetri, destino.Celda, leaf.DetRetriQ2!.Detalle,
                $"escritura DetRetri ASE{leaf.Ase.Id} · tabla por ASE · encabezado '{ResolvedorDestinoValor.EncabezadoRecaudoTotal}'");
        }

        // Plan 59 (H4): el cierre es la SUMA de las filas por ASE (lo que el propio asiento declara),
        // no un re-redondeo del total. Medido contra el oráculo oficial: el R10 del período 2026071
        // (DetRetri2026071) declara D14 = 58.210.094.822 = Σ de sus cinco filas, mientras
        // ROUND(Σ D104:D108) = 58.210.094.821 — un peso de diferencia que ninguna tolerancia de
        // composición cubre porque el gate del R10 compara Σ Detalle (HU-20/G3).
        var cierreEsperado = DetRetriRounder.Round(conDetalle.Sum(l => l.DetRetriQ2!.Detalle));
        var cierre = ResolvedorDestinoValor.ResolverTabla(
            workbookPart, worksheet,
            new ResolvedorDestinoValor.IdentidadTabla(
                hojaDetRetri, Titulo: null, AseId: null,
                ResolvedorDestinoValor.EncabezadoRecaudoTotal));
        EscribirValorNumericoEnCeldaExistente(
            workbookPart, hojaDetRetri, cierre.Celda, cierreEsperado,
            "escritura DetRetri · fila de cierre de la tabla por ASE");
    }

    /// <summary>
    /// Plan 59 (H4, gate anti-falso-verde): contrasta el asiento <b>YA ESCRITO</b> en la tabla
    /// principal de <c>DetRetri{AAAAMMQ}</c> contra el valor que el DOMINIO declara
    /// (<see cref="DetRetriQ2Inputs.Detalle"/> por ASE; cierre = Σ de esas filas, que es lo que el
    /// asiento suma y lo que el R10 oficial declara como total).
    ///
    /// Antes de este gate la única verificación de valores leía la fila de diferencias de
    /// <c>DetValiRetri</c> desde el caché de la salida: con el asiento en el literal 0 de la
    /// plantilla, «0 − 0» cerraba y la corrida certificaba un entregable en cero. Ahora la
    /// comparación es contra el esperado del dominio y sobre los bytes recién guardados:
    /// una celda del asiento que no traiga el valor esperado (incluido el 0 escrito donde el
    /// dominio declara un monto) aborta la corrida y el parcial se borra.
    ///
    /// Todo se resuelve por identidad (hoja · tabla por ASE · ASE · rótulo de columna); la celda
    /// hallada solo se nombra como contexto. Sin filas ni columnas quemadas y sin ramas por período.
    ///
    /// <c>internal</c> como costura de prueba: el fixture de inyección del Plan 59 (H4) fuerza ceros
    /// en la salida y demuestra que este gate enrojece (antes, ese escenario se entregaba como OK).
    /// </summary>
    internal static void VerificarAsientoDetRetri(WorkbookPart workbookPart, IReadOnlyList<WorkbookLeafInputs> leafInputs)
    {
        var conDetalle = leafInputs
            .Where(l => l.DetRetriQ2 is not null)
            .OrderBy(l => l.Ase.Id)
            .ToList();
        if (conDetalle.Count == 0)
        {
            return;
        }

        var hoja = NombresHojaPeriodo.DetRetri(conDetalle[0].Periodo);
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(VerificarAsientoDetRetri));

        decimal sumaEscrita = 0m;
        foreach (var leaf in conDetalle)
        {
            var identidadAse = $"ASE {leaf.Ase.Id}";
            var destino = ResolvedorDestinoValor.ResolverTabla(
                workbookPart, worksheet,
                new ResolvedorDestinoValor.IdentidadTabla(
                    hoja, Titulo: null, leaf.Ase.Id,
                    ResolvedorDestinoValor.EncabezadoRecaudoTotal));

            var escrito = LeerValorDelAsiento(destino, hoja, identidadAse);
            var esperado = leaf.DetRetriQ2!.Detalle;
            if (escrito != esperado)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Escritura,
                    $"ERR-ESCRITURA: el asiento DetRetri no quedó escrito para {identidadAse}: " +
                    $"la hoja '{hoja}' · tabla por ASE · columna '{ResolvedorDestinoValor.EncabezadoRecaudoTotal}' " +
                    $"trae {escrito} y el dominio espera {esperado} (ROUND(D104:D108) del ASE). " +
                    $"El workbook NO se entrega: un asiento en cero o desalineado no certifica la remuneración. " +
                    $"Celda hallada {destino.Celda}.");
            }

            sumaEscrita += escrito;
        }

        var cierre = ResolvedorDestinoValor.ResolverTabla(
            workbookPart, worksheet,
            new ResolvedorDestinoValor.IdentidadTabla(
                hoja, Titulo: null, AseId: null,
                ResolvedorDestinoValor.EncabezadoRecaudoTotal));

        var cierreEscrito = LeerValorDelAsiento(cierre, hoja, "fila de cierre");
        var cierreEsperado = DetRetriRounder.Round(conDetalle.Sum(l => l.DetRetriQ2!.Detalle));
        if (cierreEscrito != cierreEsperado)
        {
            throw new CalculoInvalidoException(
                CodigoError.Escritura,
                $"ERR-ESCRITURA: la fila de cierre del asiento DetRetri de '{hoja}' trae {cierreEscrito} y el dominio espera " +
                $"{cierreEsperado} (Σ de las filas por ASE del asiento). El workbook NO se entrega. " +
                $"Celda hallada {cierre.Celda}.");
        }

        if (cierreEscrito != sumaEscrita)
        {
            throw new CalculoInvalidoException(
                CodigoError.Escritura,
                $"ERR-ESCRITURA: el asiento DetRetri de '{hoja}' no cuadra consigo mismo: el cierre declara {cierreEscrito} " +
                $"y la suma de las filas por ASE escrita es {sumaEscrita}. El workbook NO se entrega.");
        }

        VerificarDesgloseTrazableDetRetri(workbookPart, conDetalle);
    }

    /// <summary>
    /// Tolerancia del gate del desglose-detalle: la doctrina del proyecto (±0,5) cubre el redondeo
    /// con que las fuentes publican sus propios cierres. No se amplía nunca para «hacer pasar» una
    /// identidad: si el bloque quedó vacío la diferencia es el monto completo, no un céntimo.
    /// </summary>
    private const decimal ToleranciaDetalle = 0.5m;

    /// <summary>Una fila del detalle tal como el DOMINIO la declara: firma A–D leída de la fuente + valores por componente.</summary>
    private sealed record FilaDetalleDelDominio(FilaEspejoR1 Etiquetas, IReadOnlyDictionary<string, decimal?> ValoresPorComponente)
    {
        /// <summary>Firma A–D de la fila, solo como contexto de los mensajes (no decide el nivel).</summary>
        internal string Firma => string.Join(
            "|",
            new[] { Etiquetas.A, Etiquetas.B, Etiquetas.C, Etiquetas.D }.Select(t => t.Trim()));
    }

    /// <summary>
    /// Plan 59 (§11-bis): gate de VALORES del DESGLOSE-DETALLE de <c>Rem. Anticipos R2</c> y
    /// <c>Reversion Pagos R4</c>, sobre los bytes YA GUARDADOS.
    ///
    /// <para><b>El defecto que cierra.</b> Hasta aquí el detalle por componente se construía solo en
    /// la 2.ª quincena: el entregable de la 1.ª salía con los tres niveles del bloque —detalle
    /// (<c>Vlr Servicio</c>/<c>Vlr Intereses</c> por canal y concepto), totales por canal y cierre de
    /// sección— en los literales <c>0</c> de la plantilla. Su cierre de cuadro
    /// (<c>=E25+E42-K25</c>) daba 0, la línea de anticipos de <c>CONSOLIDADO!D28:D32</c> se caía y el
    /// archivo se contradecía a sí mismo: el <c>DetRetri</c> se escribía desde el dominio y el bloque
    /// del que depende quedaba vacío. Ningún gate de forma lo veía, porque la forma siempre estuvo
    /// bien (literales, sin <c>&lt;f&gt;</c>).</para>
    ///
    /// <para><b>Qué exige, por ASE, por NIVEL del bloque y por componente (encabezado leído de la
    /// hoja):</b> que cada nivel escrito lleve exactamente lo que la FUENTE declara para ese nivel —
    /// detalle, totales por canal, totales por empresa, cierre de sección y gran total. Los niveles se
    /// clasifican con la MISMA gramática de cierres de la hoja, aplicada a la fuente y a la hoja.
    /// Un nivel en cero donde la fuente declara monto aborta con el monto completo, no con un céntimo.</para>
    ///
    /// <para><b>Por qué nivel a nivel y no por «la cadena debe cerrar»:</b> medido en agosto-2 · ASE4, la
    /// propia fuente publica su cierre de sección sin que sea la suma de sus cierres por empresa (hay
    /// filas de detalle fuera de todo grupo en el cuadro <c>Subs/Cont</c>). El entregable copia la
    /// fuente: exigir una identidad que la fuente no cumple sería inventar una regla de negocio y
    /// enrojecería un entregable correcto.</para>
    ///
    /// <para><b>Cero coordenadas y cero ramas.</b> El bloque se delimita por su cabecera de ASE y su
    /// gran total declarados (<see cref="DelimitadorBloquesHoja"/>), los niveles salen de la gramática
    /// de cierres de la hoja y las columnas de la fila de encabezados de la propia hoja. La fila
    /// hallada solo se nombra como contexto, igual que en <see cref="VerificarAsientoDetRetri"/>.</para>
    ///
    /// <para><b>Qué NO juzga:</b> los niveles/componentes que la fuente no declara (capacidad de
    /// plantilla), los componentes que el dominio declara en cero y la plantilla no declara (cero
    /// documentado: no hay nada que perder) y las columnas del bloque que traen <c>&lt;f&gt;</c> (la app
    /// nunca escribe fórmulas; esas no son celdas de valor).</para>
    ///
    /// <c>internal</c> como costura de prueba: el fixture de inyección vacía el bloque en una copia y
    /// demuestra que este gate enrojece.
    /// </summary>
    internal static void VerificarDetalleR2R4(WorkbookPart workbookPart, IReadOnlyList<WorkbookLeafInputs> leafInputs)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);
        ArgumentNullException.ThrowIfNull(leafInputs);

        foreach (var leaf in leafInputs.OrderBy(l => l.Ase.Id))
        {
            if (leaf.DetalleR2 is { } detalleR2)
            {
                VerificarBloqueDetalle(
                    workbookPart,
                    DelimitadorBloquesHoja.R2,
                    leaf.Ase.Id,
                    detalleR2.Filas
                        .Select(f => new FilaDetalleDelDominio(
                            new FilaEspejoR1 { A = f.A, B = f.B, C = f.C, D = f.D },
                            f.ValoresPorComponente))
                        .ToList());
            }

            if (leaf.DetalleR4 is { } detalleR4)
            {
                VerificarBloqueDetalle(
                    workbookPart,
                    DelimitadorBloquesHoja.R4,
                    leaf.Ase.Id,
                    detalleR4.Filas
                        .Select(f => new FilaDetalleDelDominio(
                            new FilaEspejoR1 { A = f.A, B = f.B, C = f.C },
                            f.ValoresPorComponente))
                        .ToList());
            }
        }
    }

    private static void VerificarBloqueDetalle(
        WorkbookPart workbookPart,
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        int aseId,
        IReadOnlyList<FilaDetalleDelDominio> dominio)
    {
        var identidad = $"hoja '{hoja.Nombre}' · ASE {aseId}";
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, hoja, aseId);
        var worksheet = ObtenerHoja(workbookPart, hoja.Nombre, nameof(VerificarDetalleR2R4));
        var tabla = worksheet.Elements<SheetData>().FirstOrDefault()
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La {identidad} no tiene SheetData: no se puede verificar el desglose-detalle que la app acaba de escribir.");

        var porNumero = tabla.Elements<Row>()
            .Where(f => f.RowIndex is not null)
            .GroupBy(f => (int)f.RowIndex!.Value)
            .ToDictionary(g => g.Key, g => g.First());

        var columnas = ColumnasDeValor(workbookPart, porNumero, bloque);
        var esperadoDelDominio = SumaDelDominio(hoja, dominio);

        foreach (var (etiqueta, columna) in columnas)
        {
            var niveles = new (DelimitadorBloquesHoja.TipoFilaBloqueHoja Nivel, string Nombre)[]
            {
                (DelimitadorBloquesHoja.TipoFilaBloqueHoja.Dato, "las filas de detalle"),
                (DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreCanal, "los totales por canal"),
                (DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreEmpresa, "los totales por empresa"),
                (DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSeccion, "los cierres de sección"),
            };

            foreach (var (nivel, nombre) in niveles)
            {
                var escrito = SumarNivel(porNumero, bloque, nivel, columna);
                if (escrito.TieneFormula)
                {
                    // La app nunca escribe fórmulas: una columna con <f> no es zona de valor y no se juzga.
                    continue;
                }

                if (!esperadoDelDominio.TryGetValue((nivel, etiqueta), out var esperado))
                {
                    continue; // el nivel no existe en la fuente para ese componente (capacidad de plantilla).
                }

                if (Math.Abs(escrito.Suma - esperado) <= ToleranciaDetalle)
                {
                    continue;
                }

                throw new CalculoInvalidoException(
                    CodigoError.Escritura,
                    $"ERR-ESCRITURA: el desglose-detalle no quedó escrito en la {identidad} · componente '{etiqueta}' · {nombre}: " +
                    $"la hoja declara {escrito.Suma} y la fuente del ASE declara {esperado} (Δ {escrito.Suma - esperado}). " +
                    "El bloque se escribe en sus tres niveles —detalle por canal/concepto, totales por canal/empresa y cierre de " +
                    "sección— y un nivel en cero (o desalineado) deja la línea de anticipos en cero en CONSOLIDADO!D28:D32. " +
                    $"El workbook NO se entrega. Filas halladas {bloque.Cabecera}-{bloque.GranTotal}.");
            }

            var granTotal = ValorDeFila(porNumero, bloque.GranTotal, columna);
            if (granTotal.TieneFormula
                || !esperadoDelDominio.TryGetValue((DelimitadorBloquesHoja.TipoFilaBloqueHoja.GranTotal, etiqueta), out var cierreEsperado))
            {
                continue;
            }

            if (Math.Abs(granTotal.Suma - cierreEsperado) > ToleranciaDetalle)
            {
                throw new CalculoInvalidoException(
                    CodigoError.Escritura,
                    $"ERR-ESCRITURA: el gran total del bloque no quedó escrito en la {identidad} · componente '{etiqueta}': " +
                    $"la hoja declara {granTotal.Suma} y la fuente del ASE declara {cierreEsperado} (Δ {granTotal.Suma - cierreEsperado}). " +
                    $"El workbook NO se entrega. Fila hallada {bloque.GranTotal}.");
            }
        }

        // Un componente con monto en la fuente y sin columna en la plantilla es un destino ausente:
        // se nombra el componente (nunca se remapea a otra columna ni se descarta en silencio).
        var columnasPorEtiqueta = columnas.ToDictionary(c => c.Etiqueta, c => c.Columna, StringComparer.OrdinalIgnoreCase);
        foreach (var ((nivel, etiqueta), esperado) in esperadoDelDominio)
        {
            if (nivel == DelimitadorBloquesHoja.TipoFilaBloqueHoja.GranTotal
                || Math.Abs(esperado) <= ToleranciaDetalle
                || columnasPorEtiqueta.ContainsKey(etiqueta))
            {
                continue;
            }

            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"ERR-PLANTILLA: la fuente del ASE {aseId} declara {esperado} en el componente '{etiqueta}' y la {identidad} " +
                $"no declara esa columna en la fila de encabezados de su bloque ({bloque.EncabezadoColumnas}). " +
                "La plantilla no tiene destino para ese monto: fail-fast accionable, nunca remapeo por posición.");
        }
    }

    /// <summary>
    /// Encabezados de las columnas de valor del bloque, leídos de la PROPIA fila de encabezados: la
    /// etiqueta es el texto que la celda declara y la columna sale de su referencia. Cero letras de
    /// columna escritas en el código (la firma A–E ya está consumida por el delimitador).
    /// </summary>
    private static IReadOnlyList<(string Etiqueta, string Columna)> ColumnasDeValor(
        WorkbookPart workbookPart,
        IReadOnlyDictionary<int, Row> porNumero,
        DelimitadorBloquesHoja.BloqueHoja bloque)
    {
        if (!porNumero.TryGetValue(bloque.EncabezadoColumnas, out var encabezado))
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La hoja '{bloque.Hoja}' no trae la fila de encabezados ({bloque.EncabezadoColumnas}) del bloque del ASE {bloque.AseId}: no se pueden resolver las columnas del desglose-detalle.");
        }

        var columnas = new List<(string, string)>();
        foreach (var celda in encabezado.Elements<Cell>())
        {
            var etiqueta = LectorTextoHoja.Texto(workbookPart, celda).Trim();
            if (etiqueta.Length == 0)
            {
                continue;
            }

            var columna = ColumnaDeReferencia(celda.CellReference?.Value);
            if (columna.Length > 0)
            {
                columnas.Add((etiqueta, columna));
            }
        }

        if (columnas.Count == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La fila de encabezados ({bloque.EncabezadoColumnas}) del bloque del ASE {bloque.AseId} de la hoja '{bloque.Hoja}' no declara rótulos de columna: no se puede verificar el desglose-detalle por componente.");
        }

        return columnas;
    }

    /// <summary>
    /// Σ de lo que el DOMINIO declara, por NIVEL del bloque y encabezado de componente. Los niveles se
    /// obtienen clasificando cada fila de la fuente con la MISMA gramática de cierres de la hoja
    /// (<see cref="DelimitadorBloquesHoja.NivelDe"/>): la fuente y la plantilla comparten el vocabulario,
    /// así que el detalle, los totales por canal/empresa, el cierre de sección y el gran total son
    /// comparables uno a uno.
    ///
    /// <para><b>Por qué nivel a nivel y no por identidad de suma:</b> medido en agosto-2 (ASE4), la
    /// propia fuente publica un cierre de sección que NO es la suma de sus cierres por empresa
    /// (<c>A='Componente'</c> = 126.782.453,43 frente a Σ empresas = 126.782.453,43 · pero el cuadro
    /// <c>Subs/Cont</c> trae filas de detalle fuera de todo grupo). El entregable copia la fuente; exigir
    /// una identidad que la fuente no cumple sería inventar una regla de negocio. Lo que sí se exige es
    /// que CADA nivel escrito lleve exactamente lo que la fuente declara para ese nivel.</para>
    /// </summary>
    private static Dictionary<(DelimitadorBloquesHoja.TipoFilaBloqueHoja Nivel, string Etiqueta), decimal> SumaDelDominio(
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        IReadOnlyList<FilaDetalleDelDominio> dominio)
    {
        var suma = new Dictionary<(DelimitadorBloquesHoja.TipoFilaBloqueHoja, string), decimal>(NivelEtiquetaComparer.Instancia);
        foreach (var fila in dominio)
        {
            var nivel = NivelDelDominio(hoja, fila);
            foreach (var (etiqueta, valor) in fila.ValoresPorComponente)
            {
                var clave = (nivel, etiqueta);
                suma[clave] = suma.GetValueOrDefault(clave) + (valor ?? 0m);
            }
        }

        return suma;
    }

    private static DelimitadorBloquesHoja.TipoFilaBloqueHoja NivelDelDominio(
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        FilaDetalleDelDominio fila) =>
        DelimitadorBloquesHoja.Clasificar(hoja, fila.Etiquetas, $"fila fuente '{fila.Firma}'") switch
        {
            DelimitadorBloquesHoja.NivelCierreHoja.GranTotal => DelimitadorBloquesHoja.TipoFilaBloqueHoja.GranTotal,
            DelimitadorBloquesHoja.NivelCierreHoja.Seccion => DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSeccion,
            DelimitadorBloquesHoja.NivelCierreHoja.Subseccion => DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreSubseccion,
            DelimitadorBloquesHoja.NivelCierreHoja.EmpresaBloque => DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreEmpresa,
            DelimitadorBloquesHoja.NivelCierreHoja.Canal => DelimitadorBloquesHoja.TipoFilaBloqueHoja.CierreCanal,
            _ => DelimitadorBloquesHoja.TipoFilaBloqueHoja.Dato,
        };

    /// <summary>
    /// Comparador de la clave (nivel, etiqueta): el nivel es el enum de la gramática y la etiqueta se
    /// compara como la lee la hoja (sin distinguir mayúsculas ni espacios sobrantes).
    /// </summary>
    private sealed class NivelEtiquetaComparer : IEqualityComparer<(DelimitadorBloquesHoja.TipoFilaBloqueHoja Nivel, string Etiqueta)>
    {
        internal static readonly NivelEtiquetaComparer Instancia = new();

        public bool Equals(
            (DelimitadorBloquesHoja.TipoFilaBloqueHoja Nivel, string Etiqueta) x,
            (DelimitadorBloquesHoja.TipoFilaBloqueHoja Nivel, string Etiqueta) y) =>
            x.Nivel == y.Nivel && string.Equals(x.Etiqueta.Trim(), y.Etiqueta.Trim(), StringComparison.OrdinalIgnoreCase);

        public int GetHashCode((DelimitadorBloquesHoja.TipoFilaBloqueHoja Nivel, string Etiqueta) obj) =>
            HashCode.Combine(obj.Nivel, StringComparer.OrdinalIgnoreCase.GetHashCode(obj.Etiqueta.Trim()));
    }

    private static (decimal Suma, bool TieneFormula) SumarNivel(
        IReadOnlyDictionary<int, Row> porNumero,
        DelimitadorBloquesHoja.BloqueHoja bloque,
        DelimitadorBloquesHoja.TipoFilaBloqueHoja nivel,
        string columna)
    {
        var suma = 0m;
        var formula = false;
        foreach (var fila in bloque.Filas.Where(f => f.Tipo == nivel))
        {
            var valor = ValorDeFila(porNumero, fila.Fila, columna);
            suma += valor.Suma;
            formula |= valor.TieneFormula;
        }

        return (suma, formula);
    }

    /// <summary>
    /// Valor numérico de la celda (columna, fila) tal como quedó guardado. Celda ausente o vacía = 0
    /// (el literal 0 de la plantilla); una celda con <c>&lt;f&gt;</c> no es valor escrito por la app y
    /// se informa para que el llamador no la juzgue desde el caché.
    /// </summary>
    private static (decimal Suma, bool TieneFormula) ValorDeFila(
        IReadOnlyDictionary<int, Row> porNumero,
        int numeroFila,
        string columna)
    {
        if (!porNumero.TryGetValue(numeroFila, out var fila))
        {
            return (0m, false);
        }

        var referencia = columna + numeroFila.ToString(CultureInfo.InvariantCulture);
        var celda = fila.Elements<Cell>().FirstOrDefault(c =>
            string.Equals(c.CellReference?.Value, referencia, StringComparison.OrdinalIgnoreCase));
        if (celda is null)
        {
            return (0m, false);
        }

        if (celda.CellFormula is not null)
        {
            return (0m, true);
        }

        var texto = celda.CellValue?.InnerText;
        if (string.IsNullOrWhiteSpace(texto))
        {
            return (0m, false);
        }

        return decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor)
            ? (valor, false)
            : (0m, false);
    }

    /// <summary>Letras de columna de una referencia A1 (sin dígitos ni <c>$</c>).</summary>
    private static string ColumnaDeReferencia(string? referencia)
    {
        if (string.IsNullOrEmpty(referencia))
        {
            return string.Empty;
        }

        var fin = 0;
        while (fin < referencia.Length && (char.IsLetter(referencia[fin]) || referencia[fin] == '$'))
        {
            fin++;
        }

        return referencia[..fin].Replace("$", string.Empty, StringComparison.Ordinal);
    }

    /// <summary>
    /// Plan 59 (§11 — FRENTE 1): gate de VALORES del DESGLOSE trazable sobre el archivo ya guardado.
    ///
    /// Hasta este gate, el desglose solo tenía verificación de FORMA (que la celda siga siendo el
    /// valor editable que la plantilla declara) y se escribía únicamente en la 2.ª quincena: el
    /// entregable de 2026101 salió con 'Balance de subsidio o contribuciones', 'Recaudo base de
    /// remuneración', 'Costo de la Interventoría' y 'Costo de la Interventoría' de DetValiRetri en el
    /// literal 0 de la plantilla, y la corrida certificó OK. Este gate relee lo escrito y lo contrasta
    /// con lo que el DOMINIO declara para cada ASE: un 0 donde el dominio declara un monto aborta la
    /// corrida y el parcial se borra (nunca se entrega).
    ///
    /// Qué cubre y qué no, sin ambigüedad: cubre las columnas que el dominio declara
    /// (<see cref="DesgloseDeclaradoPorElDominio"/>). NO cubre 'RECAUDO BASE DE REMUNERACIÓN': su
    /// valor es la cola de la cadena de fórmulas del CONSOLIDADO y sus operandos no están en el
    /// dominio (ver el motivo medido en <see cref="EscribirDesgloseTrazableDetRetri"/>); gatearla
    /// obligaría a inventar el número o a copiarlo del R10.
    /// </summary>
    private static void VerificarDesgloseTrazableDetRetri(
        WorkbookPart workbookPart,
        IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        var hojas = new Dictionary<string, Worksheet>(StringComparer.OrdinalIgnoreCase);
        foreach (var leaf in leafs)
        {
            foreach (var (hoja, rotulo, esperado) in DesgloseDeclaradoPorElDominio(leaf))
            {
                VerificarCeldaDeDesglose(workbookPart, hojas, hoja, leaf.Ase.Id, rotulo, esperado);
            }
        }

        // La fila de cierre de cada columna también es un valor del entregable: se contrasta con la
        // Σ de las filas por ASE que el dominio declara (el oráculo administrativo coincide con esa
        // suma, no con un re-redondeo).
        foreach (var (hoja, rotulo) in ColumnasDelDesglose(leafs[0].Periodo))
        {
            var esperado = leafs.Aggregate(0m, (suma, leaf) => suma + ValorDeclarado(leaf, hoja, rotulo));
            VerificarCeldaDeDesglose(workbookPart, hojas, hoja, aseId: null, rotulo, esperado);
        }
    }

    private static void VerificarCeldaDeDesglose(
        WorkbookPart workbookPart,
        IDictionary<string, Worksheet> hojas,
        string hoja,
        int? aseId,
        string rotulo,
        decimal esperado)
    {
        if (!hojas.TryGetValue(hoja, out var worksheet))
        {
            worksheet = ObtenerHoja(workbookPart, hoja, nameof(VerificarAsientoDetRetri));
            hojas[hoja] = worksheet;
        }

        var destino = ResolvedorDestinoValor.ResolverTablaPorConcepto(
            workbookPart, worksheet,
            new ResolvedorDestinoValor.IdentidadTabla(hoja, Titulo: null, aseId, rotulo));

        var identidad = aseId is { } ase ? $"ASE {ase}" : "la fila de cierre de la tabla por ASE";
        var escrito = LeerValorDelAsiento(destino, hoja, identidad);
        if (escrito != esperado)
        {
            throw new CalculoInvalidoException(
                CodigoError.Escritura,
                $"ERR-ESCRITURA: el desglose trazable no quedó escrito para {identidad}: " +
                $"la hoja '{hoja}' · tabla por ASE · columna '{rotulo}' trae {escrito} y el dominio espera {esperado}. " +
                "El workbook NO se entrega: un desglose en cero o desalineado no certifica la remuneración. " +
                $"Celda hallada {destino.Celda}.");
        }
    }

    /// <summary>
    /// Lee el valor que el asiento dejó en el destino resuelto. Una celda sin valor numérico (vacía,
    /// texto o fórmula sin caché) es un asiento no escrito: se informa como 0 y el gate lo contrasta
    /// contra el esperado del dominio — nunca se asume que «no hay valor» equivale a «cero correcto».
    /// </summary>
    private static decimal LeerValorDelAsiento(ResolvedorDestinoValor.Destino destino, string hoja, string identidad)
    {
        var texto = destino.CeldaXml.CellValue?.InnerText;
        if (string.IsNullOrWhiteSpace(texto))
        {
            return 0m;
        }

        if (!decimal.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out var valor))
        {
            throw new CalculoInvalidoException(
                CodigoError.Escritura,
                $"ERR-ESCRITURA: el asiento DetRetri de '{hoja}' para {identidad} no quedó como número ('{texto}'). " +
                $"El workbook NO se entrega. Celda hallada {destino.Celda}.");
        }

        return valor;
    }

    /// <summary>
    /// Encabezado de la columna donde las hojas de detalle publican el BALANCE de subsidio o
    /// contribuciones (concepto; su valor es <c>BCE SC POR FACT.</c> = Subsidio + Contribución).
    /// </summary>
    private const string RotuloBalanceSubsidioContribuciones = "Balance de subsidio o contribuciones";

    /// <summary>
    /// Encabezado de la columna donde las hojas de detalle publican el COSTO DE LA INTERVENTORÍA
    /// (la plantilla lo decora con <c>(*)</c> en una hoja y sin él en otra: la identidad es el
    /// concepto, no la decoración).
    /// </summary>
    private const string RotuloCostoInterventoria = "Costo de la Interventoría";

    /// <summary>Encabezado de la columna de la BANDERA de validación de retribución de <c>DetValiRetri</c>.</summary>
    private const string RotuloValidacionRetribucion = "Validación retribución";

    /// <summary>
    /// Plan 29 (T4, Unidad D — desglose trazable) / Plan 59 (§11 — FRENTE 1): escribe como LITERAL
    /// las columnas de las hojas de detalle del período cuyo valor el DOMINIO declara:
    /// <list type="bullet">
    /// <item><c>DetRetri</c> · 'Balance de subsidio o contribuciones' ← <c>BCE SC POR FACT.</c>
    /// (Subsidio + Contribución, ya escritos por el ASE) redondeado con <see cref="DetRetriRounder"/>;</item>
    /// <item><c>DetRetri</c> · 'Costo de la Interventoría' ← <see cref="InterventoriaDeclarada.PorQuincena"/>
    /// (primera o segunda quincena según el PERÍODO, decidido en el dominio);</item>
    /// <item><c>DetValiRetri</c> · 'Costo de la Interventoría' ← el mismo valor negado;</item>
    /// <item><c>DetValiRetri</c> · 'Validación retribución' ← 0 (bandera sin desglose).</item>
    /// </list>
    ///
    /// <para><b>Se escribe en AMBAS quincenas.</b> Antes vivía dentro de <c>if (esQuincena2)</c>: la
    /// plantilla declara esas columnas como VALOR en las dos y el oráculo administrativo las trae
    /// llenas también en la primera, así que la rama dejaba el entregable de Q1 en el literal 0.</para>
    ///
    /// <para><b>Lo que NO se escribe, con su motivo medido:</b> la columna
    /// 'RECAUDO BASE DE REMUNERACIÓN' (<c>DetRetri</c>) y sus hermanas de la cadena
    /// (RECAUDO POR DISPOSICIÓN FINAL, LIXIVIADOS, VIAT, APROVECHAMIENTO, COMERCIALIZACIÓN 11,4%,
    /// REMUNERACIÓN FINAL, CONCEPTOS ESPECIALES, REMUNERACIÓN FINAL MÁS CONCEPTOS ESPECIALES).
    /// Medido en el oráculo administrativo: cada una es el resultado de la fórmula que el
    /// CONSOLIDADO declara por sección —<c>K104 = D104−E104−F104−G104−H104−I104−J104</c> sobre las
    /// cinco tablas de recaudo— y sus operandos (E..I) NO están en el dominio del período.
    /// Reproducir esa cadena en C# duplicaría la fuente de verdad (Plan 29 §2.2, «motor C#»
    /// DESCARTADA) y copiarla del R10 sería tautología. Queda declarada como pendiente de negocio
    /// en el gate del Plan 59 §11 (la columna se deja en el 0 de la plantilla, nunca se inventa).</para>
    ///
    /// El guard anti-fórmula de <see cref="EscribirValorNumericoEnCeldaExistente"/> sigue intacto
    /// (destino fórmula ⇒ ERR-PLANTILLA, nunca sobrescritura).
    /// </summary>
    private static void EscribirDesgloseTrazableDetRetri(
        WorkbookPart workbookPart,
        IReadOnlyList<WorkbookLeafInputs> leafInputs)
    {
        var conDetalle = leafInputs
            .Where(l => l.DetRetriQ2 is not null)
            .OrderBy(l => l.Ase.Id)
            .ToList();
        if (conDetalle.Count == 0)
        {
            return;
        }

        foreach (var leaf in conDetalle)
        {
            foreach (var (hoja, rotulo, valor) in DesgloseDeclaradoPorElDominio(leaf))
            {
                EscribirCeldaDeDesglose(workbookPart, hoja, leaf.Ase.Id, rotulo, valor);
            }
        }

        // Cierre de cada columna = Σ de las filas por ASE, que es lo que la tabla totaliza y lo que
        // el oráculo administrativo declara (p. ej. 'Costo de la Interventoría' = Σ de los cinco
        // costos, sin re-redondeo). La fila de cierre es también un VALOR de la plantilla.
        var periodo = conDetalle[0].Periodo;
        foreach (var (hoja, rotulo) in ColumnasDelDesglose(periodo))
        {
            var cierre = conDetalle.Aggregate(0m, (suma, leaf) => suma + ValorDeclarado(leaf, hoja, rotulo));
            EscribirCeldaDeDesglose(workbookPart, hoja, aseId: null, rotulo, cierre);
        }
    }

    /// <summary>
    /// Escribe un valor del desglose en la celda que el RÓTULO de la columna y la identidad de la
    /// fila (ASE o cierre) resuelven. <paramref name="aseId"/> null = fila de cierre de la tabla.
    /// </summary>
    private static void EscribirCeldaDeDesglose(
        WorkbookPart workbookPart,
        string hoja,
        int? aseId,
        string rotulo,
        decimal valor)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, "escritura del desglose trazable de las hojas de detalle");
        var destino = ResolvedorDestinoValor.ResolverTablaPorConcepto(
            workbookPart, worksheet,
            new ResolvedorDestinoValor.IdentidadTabla(hoja, Titulo: null, aseId, rotulo));

        var identidad = aseId is { } ase ? $"ASE{ase}" : "fila de cierre";
        EscribirValorNumericoEnCeldaExistente(
            workbookPart, hoja, destino.Celda, valor,
            $"escritura {identidad} · '{hoja}' · columna '{rotulo}'");
    }

    /// <summary>
    /// Valores del desglose trazable que el DOMINIO declara para un ASE, como tríos
    /// (hoja, rótulo de columna, valor). Es la ÚNICA fuente que consumen la escritura y la
    /// verificación: si el dominio no declara un monto, no aparece aquí y no se inventa.
    /// </summary>
    private static IReadOnlyList<(string Hoja, string Rotulo, decimal Valor)> DesgloseDeclaradoPorElDominio(
        WorkbookLeafInputs leaf)
    {
        var aseId = leaf.Ase.Id;
        var periodo = leaf.Periodo;

        // La ÚNICA pieza del desglose que depende de la quincena: el costo declarado en la tabla
        // INTERVENTORIA (primera 'N' / segunda 'M'). La selección se resuelve en el dominio.
        var interventoria = InterventoriaDeclarada.PorQuincena(periodo.NumeroQuincena)[aseId];

        return
        [
            (NombresHojaPeriodo.DetRetri(periodo), RotuloBalanceSubsidioContribuciones, BalanceSubsidioContribuciones(leaf)),
            (NombresHojaPeriodo.DetRetri(periodo), RotuloCostoInterventoria, interventoria),
            (NombresHojaPeriodo.DetValiRetri(periodo), RotuloCostoInterventoria, -interventoria),
            (NombresHojaPeriodo.DetValiRetri(periodo), RotuloValidacionRetribucion, 0m),
        ];
    }

    /// <summary>
    /// Columnas del desglose trazable que el dominio declara para el período (hoja + rótulo), en el
    /// orden en que se escriben. Es la lista que la fila de cierre totaliza; cero coordenadas.
    /// </summary>
    private static IReadOnlyList<(string Hoja, string Rotulo)> ColumnasDelDesglose(Periodo periodo) =>
    [
        (NombresHojaPeriodo.DetRetri(periodo), RotuloBalanceSubsidioContribuciones),
        (NombresHojaPeriodo.DetRetri(periodo), RotuloCostoInterventoria),
        (NombresHojaPeriodo.DetValiRetri(periodo), RotuloCostoInterventoria),
        (NombresHojaPeriodo.DetValiRetri(periodo), RotuloValidacionRetribucion),
    ];

    /// <summary>
    /// Valor que el dominio declara para una columna del desglose de un ASE. Una columna pedida que
    /// no esté declarada es defecto de programación (no hay valor que inventar).
    /// </summary>
    private static decimal ValorDeclarado(WorkbookLeafInputs leaf, string hoja, string rotulo) =>
        DesgloseDeclaradoPorElDominio(leaf)
            .Where(x => string.Equals(x.Hoja, hoja, StringComparison.OrdinalIgnoreCase)
                && string.Equals(x.Rotulo, rotulo, StringComparison.Ordinal))
            .Select(x => x.Valor)
            .Single();

    /// <summary>
    /// Balance de subsidio o contribuciones del ASE: <c>BCE SC POR FACT.</c> = Subsidio +
    /// Contribución, redondeado con la regla única (<see cref="DetRetriRounder"/>). Sin el balance
    /// en el dominio es fail-fast: nunca se escribe un 0 que parezca un dato.
    /// </summary>
    private static decimal BalanceSubsidioContribuciones(WorkbookLeafInputs leaf)
    {
        var aseId = leaf.Ase.Id;
        var balance = leaf.BalanceSc ?? throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"El leaf del ASE {aseId} no trae BalanceSc para derivar el desglose trazable de DetRetri " +
            $"(columna '{RotuloBalanceSubsidioContribuciones}' = BCE SC POR FACT. Subsidio + Contribución); " +
            "fail-fast, nunca se inventa el desglose.");
        var bce = balance.Ases.SingleOrDefault(a => a.Ase.Id == aseId) ?? throw new CalculoInvalidoException(
            CodigoError.Plantilla,
            $"El BalanceSc del ASE {aseId} no trae su fila; no se puede derivar el desglose trazable de DetRetri.");
        return DetRetriRounder.Round(bce.TotalBsc);
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
        // Plan 55 (H1.2, D7/CA7): el mapa congelado por (empresa, ASE)
        // (`WorkbookLeafCellMapPorEmpresa`) quedó RETIRADO del camino productivo (LEGACY). El READER
        // resuelve cada celda destino por CRUCE DE ETIQUETAS contra la plantilla (autoridad del
        // cuadro) y la declara en `CeldasR1/R2/R4`; el writer solo pega VALORES en esas celdas y el
        // guard de `EscribirValorNumerico` impide tocar cualquier fórmula.
        if (!omitirR1)
        {
            EscribirCeldasResueltas(workbookPart, HojaR1, "R1", conc, aseId, conc.CeldasR1);
        }

        EscribirCeldasResueltas(workbookPart, HojaR2, "R2", conc, aseId, conc.CeldasR2);
        EscribirCeldasResueltas(workbookPart, HojaR4, "R4", conc, aseId, conc.CeldasR4);
    }

    private static void EscribirCeldasResueltas(
        WorkbookPart workbookPart,
        string hoja,
        string reporte,
        ConciliacionEmpresaInputs conc,
        int aseId,
        IReadOnlyDictionary<string, decimal> celdas)
    {
        foreach (var (celda, valor) in celdas)
        {
            EscribirValorNumerico(workbookPart, hoja, celda, valor, $"{reporte}.{conc.Empresa.Nombre}.{celda}.ASE{aseId}");
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
    /// Plan 29 (T3, Unidad R — escritura) / Plan 47 (H1, S2.1): escribe el desglose-detalle por
    /// componente de las hojas <c>Rem. Anticipos R2</c> y <c>Reversion Pagos R4</c> en la MISMA pasada
    /// atómica de los agregados (V2: <c>E43</c>/<c>D73</c> intactos) y el espejo R1.
    ///
    /// <para><b>Se escribe en AMBAS quincenas.</b> Antes el método era no-op en Q1 porque el reader
    /// solo poblaba <c>DetalleR2</c>/<c>DetalleR4</c> en Q2; el entregable de la primera quincena
    /// salía con el bloque entero en los literales 0 de la plantilla (detalle, totales por canal y
    /// cierre de sección), su cierre <c>E57</c> daba 0 y la línea de anticipos de
    /// <c>CONSOLIDADO!D28:D32</c> se caía. El detalle es del PERÍODO, no de la quincena.</para>
    ///
    /// La fila destino se resuelve por la clave <c>(firma + canal efectivo)</c> — <b>nunca por orden
    /// físico</b> (Plan 41 E1-1): el canal efectivo es el de la fila <c>Total</c> que CIERRA el grupo
    /// (R46-1 análogo, <see cref="CanalEfectivoDetalleR2R4"/>). Una fila-fuente sin slot es cero
    /// documentado sólo si está en <see cref="CerosDocumentadosDetalleR2R4"/>; fuera de la lista ⇒
    /// <c>ERR-SIN-SLOT</c>. Las filas-destino sin fuente quedan en su 0 de plantilla.
    /// Las columnas se resuelven por ENCABEZADO (<see cref="WorkbookLeafCellMapDetalleR2R4.ColumnasR2"/>
    /// / <c>ColumnasR4</c>, que incluye <c>Otros → R</c>), nunca por índice. Cero insert/delete de
    /// filas (D-C); el guard de <see cref="EscribirValorNumerico"/> impide tocar fórmulas
    /// (destino-detalle fórmula ⇒ <c>ERR-PLANTILLA</c>, nunca sobrescritura — R-R-3).
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
        // Plan 59 (H2-bis, barrido del camino de escritura): el bloque destino se DELIMITA leyendo la
        // propia hoja (cabecera de bloque + gran total declarado). Antes se tomaba de un mapa congelado
        // por ASE (BloquesR2) que la v3 ya no respeta: la fuente traía firmas que el bloque 3..33 no
        // contenía y la corrida abortaba con ERR-SIN-SLOT sobre un destino que la plantilla SÍ declara.
        var bloque = BloqueDeDetalle(workbookPart, DelimitadorBloquesHoja.R2, aseId);

        var hoja = WorkbookLeafCellMapDetalleR2R4.HojaR2;
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));
        var pares = EmparejarBloque(
            HojaDetalleR2R4.R2,
            hoja,
            aseId,
            bloque,
            LeerRotulos(workbookPart, worksheet, bloque),
            detalle.Filas.Select(f => new CanalEfectivoDetalleR2R4.Rotulo(f.A, f.B, f.C, f.D)).ToList(),
            columnasFirma: [0, 1, 2, 3],
            columnasEtiquetaPropia: WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaPropiaR2,
            columnasGrupo: WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaGrupoR2,
            columnasSegundoCierre: WorkbookLeafCellMapDetalleR2R4.ColumnasSegundoCierreR2);

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
        // Plan 59 (H2-bis): bloque del detalle delimitado por la cabecera de bloque y el gran total que
        // la hoja declara (mismo criterio que R2).
        var bloque = BloqueDeDetalle(workbookPart, DelimitadorBloquesHoja.R4, aseId);

        var hoja = WorkbookLeafCellMapDetalleR2R4.HojaR4;
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));
        var pares = EmparejarBloque(
            HojaDetalleR2R4.R4,
            hoja,
            aseId,
            bloque,
            LeerRotulos(workbookPart, worksheet, bloque),
            detalle.Filas.Select(f => new CanalEfectivoDetalleR2R4.Rotulo(f.A, f.B, f.C, string.Empty)).ToList(),
            columnasFirma: [0, 1, 2],
            columnasEtiquetaPropia: WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaPropiaR4,
            columnasGrupo: WorkbookLeafCellMapDetalleR2R4.ColumnasEtiquetaGrupoR4,
            columnasSegundoCierre: WorkbookLeafCellMapDetalleR2R4.ColumnasSegundoCierreR4);

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
    /// Plan 59 (H2-bis): rango de filas del detalle de un ASE, derivado de la ESTRUCTURA que la hoja
    /// declara — primera fila de datos (la siguiente a la fila de encabezados de columna del bloque) y
    /// última fila de datos (la anterior a su gran total). Cero ventanas congeladas por ASE.
    /// </summary>
    private static (int FilaInicio, int FilaFin) BloqueDeDetalle(
        WorkbookPart workbookPart,
        DelimitadorBloquesHoja.HojaDelimitable hoja,
        int aseId)
    {
        var bloque = DelimitadorBloquesHoja.DelimitarBloque(workbookPart, hoja, aseId);
        return (bloque.EncabezadoColumnas + 1, bloque.GranTotal);
    }

    /// <summary>
    /// Plan 47 (H1, S2.1 — §2.2 + autorización H1.1/H1.4): empareja las filas-fuente del detalle con
    /// los slots del bloque-destino por la clave <c>(firma + CANAL EFECTIVO)</c>, eliminando el
    /// casamiento por orden del LCS anterior.
    ///
    /// <list type="bullet">
    /// <item>El canal efectivo se calcula con <see cref="CanalEfectivoDetalleR2R4"/> (Core, puro,
    /// R46-1 análogo) sobre la MISMA secuencia de rótulos: la fila <c>Total</c> que CIERRA el grupo
    /// tanto en la fuente como en el oficial, COMPUESTO con la empresa del segundo cierre cuando el mapa
    /// la declara (Plan 47 W-3: <c>E/ENEL</c> vs <c>E/ENERBIT</c> en R2 ASE5, desambiguado por geometría
    /// y no por orden). Las filas con etiqueta propia conservan su firma (canal <c>null</c> en ambos
    /// lados).</item>
    /// <item>Una fila-fuente sin slot es <b>cero documentado</b> si está en
    /// <see cref="CerosDocumentadosDetalleR2R4"/> (consumo acotado por <c>Cantidad</c>); fuera de esa
    /// lista es <c>ERR-SIN-SLOT</c> accionable. Prohibido el fallback posicional (Plan 41 E1-1).</item>
    /// <item>Las filas-destino sin fuente quedan en su 0 de plantilla (nunca se inventa valor). Cero
    /// insert/delete (D-C).</item>
    /// </list>
    /// </summary>
    private static IReadOnlyList<(int FilaDestino, int IndiceFuente)> EmparejarBloque(
        HojaDetalleR2R4 hojaDetalle,
        string hoja,
        int aseId,
        (int FilaInicio, int FilaFin) bloque,
        IReadOnlyList<CanalEfectivoDetalleR2R4.Rotulo> destino,
        IReadOnlyList<CanalEfectivoDetalleR2R4.Rotulo> fuente,
        int[] columnasFirma,
        IReadOnlyList<int> columnasEtiquetaPropia,
        IReadOnlyList<int> columnasGrupo,
        IReadOnlyList<int> columnasSegundoCierre)
    {
        var firmasDestino = destino.Select(rotulo => Firma(rotulo, columnasFirma)).ToList();
        var canalesDestino = Enumerable.Range(0, destino.Count)
            .Select(i => CanalEfectivoDetalleR2R4.CanalEfectivo(destino, i, columnasEtiquetaPropia, columnasGrupo, columnasSegundoCierre))
            .ToList();

        var consumido = new bool[destino.Count];
        var documentados = new Dictionary<(string Firma, string? Canal), int>();
        var pares = new List<(int FilaDestino, int IndiceFuente)>(fuente.Count);

        for (var i = 0; i < fuente.Count; i++)
        {
            var firma = Firma(fuente[i], columnasFirma);
            var canal = CanalEfectivoDetalleR2R4.CanalEfectivo(fuente, i, columnasEtiquetaPropia, columnasGrupo, columnasSegundoCierre);

            var encontrado = -1;
            for (var j = 0; j < destino.Count; j++)
            {
                if (consumido[j])
                {
                    continue;
                }

                if (!string.Equals(firmasDestino[j], firma, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.Equals(canalesDestino[j], canal, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                encontrado = j;
                break;
            }

            if (encontrado >= 0)
            {
                consumido[encontrado] = true;
                pares.Add((bloque.FilaInicio + encontrado, i));
                continue;
            }

            // Sin slot: sólo es válido si la clave está en la lista documentada (E-44.1).
            var clave = (firma, canal);
            var yaDocumentadas = documentados.GetValueOrDefault(clave);
            if (yaDocumentadas < CerosDocumentadosDetalleR2R4.CantidadDocumentada(hojaDetalle, aseId, firma, canal))
            {
                documentados[clave] = yaDocumentadas + 1;
                continue;
            }

            throw new CalculoInvalidoException(
                CodigoError.SinSlot,
                MensajeSinSlot(hoja, aseId, bloque, i, firma, canal));
        }

        return pares;
    }

    /// <summary>Lee los rótulos A–D de cada fila del bloque-destino (la firma y el canal efectivo se
    /// derivan de ellos; el canal NO se lee de una columna fija por índice de negocio).</summary>
    private static IReadOnlyList<CanalEfectivoDetalleR2R4.Rotulo> LeerRotulos(
        WorkbookPart workbookPart,
        Worksheet worksheet,
        (int FilaInicio, int FilaFin) bloque)
    {
        var rotulos = new List<CanalEfectivoDetalleR2R4.Rotulo>(bloque.FilaFin - bloque.FilaInicio + 1);
        for (var fila = bloque.FilaInicio; fila <= bloque.FilaFin; fila++)
        {
            rotulos.Add(new CanalEfectivoDetalleR2R4.Rotulo(
                TextoDestino(workbookPart, worksheet, "A", fila),
                TextoDestino(workbookPart, worksheet, "B", fila),
                TextoDestino(workbookPart, worksheet, "C", fila),
                TextoDestino(workbookPart, worksheet, "D", fila)));
        }

        return rotulos;
    }

    private static string TextoDestino(WorkbookPart workbookPart, Worksheet worksheet, string columna, int fila)
    {
        var cell = ObtenerCelda(worksheet, columna + fila.ToString(CultureInfo.InvariantCulture));
        return cell is null ? string.Empty : LeerTextoCelda(cell, workbookPart).Trim();
    }

    /// <summary>Firma normalizada de un rótulo con las columnas declaradas (A..D en R2, A..C en R4).</summary>
    private static string Firma(CanalEfectivoDetalleR2R4.Rotulo rotulo, IReadOnlyList<int> columnas)
    {
        var partes = new string[columnas.Count];
        for (var i = 0; i < columnas.Count; i++)
        {
            partes[i] = rotulo.Texto(columnas[i]).Trim();
        }

        return string.Join("|", partes);
    }

    /// <summary>Mensaje accionable de <c>ERR-SIN-SLOT</c> (Plan 47 H1.4 / E-44.1): nombra la hoja, el
    /// ASE, la fila-fuente (posición + firma + canal efectivo) y el bloque destino revisado.</summary>
    private static string MensajeSinSlot(
        string hoja,
        int aseId,
        (int FilaInicio, int FilaFin) bloque,
        int indiceFuente,
        string firma,
        string? canal)
    {
        var canalTexto = canal is null
            ? "sin canal efectivo (la fila trae su propia etiqueta de grupo)"
            : $"canal efectivo '{canal}'";
        return $"ERR-SIN-SLOT: ASE {aseId}: la fuente del detalle de '{hoja}' trae la fila fuente " +
            $"#{indiceFuente + 1} con firma '{firma}' y {canalTexto} que NO tiene slot en el bloque " +
            $"{bloque.FilaInicio}..{bloque.FilaFin} (ninguna fila-destino con la misma firma y el mismo " +
            "canal efectivo) ni figura en la lista de ceros documentados. Nunca remap/insert ni " +
            "colocación por posición: revise la plantilla o la fuente.";
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

        // Plan 48 (H1.2): la v2 deja D7:D10 = 0 y hace E7/E10/F10 fórmulas. Se pueblan D7:D10 POR
        // ETIQUETA de la columna C (tabla citada) desde el TOTAL (fila 27) de cada Recaudo *.
        var hoja = WorkbookLeafCellMapValidaciones.HojaValidaControlRecaudo;
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));
        var porEmpresa = recaudos.ToDictionary(r => r.Empresa.Id, r => r.Total);
        var porEtiqueta = WorkbookLeafCellMapValidaciones.ControlRecaudoPorEtiqueta
            .ToDictionary(e => NormalizarEtiqueta(e.Etiqueta), e => e.EmpresaId, StringComparer.Ordinal);

        var totalControl = 0m;
        foreach (var cell in worksheet.Descendants<Cell>())
        {
            var referencia = cell.CellReference?.Value;
            if (referencia is null || referencia.Length < 2 || referencia[0] != 'C' || !char.IsAsciiDigit(referencia[1]))
            {
                continue;
            }

            var etiqueta = NormalizarEtiqueta(LeerTextoCelda(cell, workbookPart));
            if (!porEtiqueta.TryGetValue(etiqueta, out var empresaId)
                || !porEmpresa.TryGetValue(empresaId, out var total))
            {
                continue;
            }

            var destino = "D" + referencia[1..];
            // Guard anti-fórmula de EscribirValorNumerico: si el destino D fuera fórmula lanza
            // ERR-PLANTILLA (nunca sobrescritura). No se hace skip silencioso.
            EscribirValorNumerico(workbookPart, hoja, destino, total, $"ValidaControlRecaudo.{destino}");
            totalControl += total;
        }

        // F10: en la v1 es literal de control; en la v2 es la fórmula E7+E10 (la App NO la escribe:
        // la calcula Excel y su caché se resuelve por read-back para validar C14). Solo se puebla
        // cuando la celda NO es fórmula.
        var cellF10 = ObtenerCelda(worksheet, WorkbookLeafCellMapValidaciones.CeldaControlRecaudoTotal);
        if (cellF10?.CellFormula is null)
        {
            EscribirValorNumerico(
                workbookPart,
                hoja,
                WorkbookLeafCellMapValidaciones.CeldaControlRecaudoTotal,
                totalControl,
                $"ValidaControlRecaudo.{WorkbookLeafCellMapValidaciones.CeldaControlRecaudoTotal}");
        }
    }

    /// <summary>Normaliza un rótulo para el match por etiqueta (sin espacios, mayúsculas).</summary>
    private static string NormalizarEtiqueta(string texto) =>
        new string(texto.Where(c => !char.IsWhiteSpace(c)).ToArray()).ToUpperInvariant();

    /// <summary>
    /// Plan 48 (H1b — patrón Plan 40 «evalúa y sella»): tras escribir <c>D7:D10</c>, sella los
    /// <c>&lt;v&gt;</c> de <c>E7 = SUM(D7:D9)</c>, <c>E10 = SUM(D10)</c> y <c>F10 = E7+E10</c> con el
    /// valor EVALUADO por read-back (nunca se inventa: si no es resoluble, no se sella). Preserva la
    /// <c>&lt;f&gt;</c> intacta. En la v1 estas celdas son literales (las escribe la App) y el método
    /// es no-op; en la v2 son fórmulas y el gate <c>VALIDACION_TOTAL!C14</c> lee su valor.
    /// </summary>
    private static void SellarCachesControlRecaudo(WorkbookPart workbookPart)
    {
        var hoja = WorkbookLeafCellMapValidaciones.HojaValidaControlRecaudo;
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));
        var resolutor = new ResolutorValorValidacion(workbookPart);
        foreach (var celda in WorkbookLeafCellMapValidaciones.CeldasSelladoControlRecaudo)
        {
            var cell = ObtenerCelda(worksheet, celda);
            if (cell?.CellFormula is null)
            {
                continue; // literal escrito por la App (v1) o celda sin fórmula: no se toca.
            }

            if (!resolutor.TryEvaluar(hoja, celda, out var valor))
            {
                continue; // no evaluable ⇒ no se sella (E1-1: nunca un valor inventado).
            }

            cell.CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture));
            cell.DataType = CellValues.Number;
        }
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

        // Plan 58 (H1.1): el guard anti-fórmula es UNO (GuardEscrituraFormula); la condición se
        // evalúa sobre el contenido de la celda, nunca sobre su nombre ni su hoja.
        GuardEscrituraFormula.ExigirSinFormula(cell, hoja, celda, $"sello '{nombre}'");

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

    /// <summary>
    /// Plan 40 (T2, R-S-1/R-S-2/R-S-3 — D-E): <b>sellado de cachés computables</b>. Tras la
    /// escritura de negocio y ANTES del saneamiento calcChain, escribe los <c>&lt;v&gt;</c> de la
    /// cadena <c>C3:C8 → C9 → C14 → C15</c> y de los 5 sub-bloques <c>C15/D25/O25/D34/F34</c>,
    /// calculados por READ-BACK de los valores efectivamente escritos en el workbook de salida
    /// (jamás desde la intención del dominio), preservando el <c>&lt;f&gt;</c> intacto.
    ///
    /// Allow-list del writer (R-S-3, ampliada por la enmienda E-1): SOLO las celdas de
    /// <c>VALIDACION_TOTAL</c> cuya cadena es resoluble por read-back de lo escrito —
    /// <c>C3:C8</c> (= Σ 'Recaudo *'!F21:F26 en Q2 / D21:D26 en Q1), <c>C9</c> (= SUM(C3:C8)),
    /// <c>C14</c> (= 'Valida - Control Recaudo'!F10) y los 5 sub-bloques de
    /// <see cref="FormasValidacionTotal.SubBloques"/> (<c>C15</c>, <c>D25</c>, <c>O25</c>,
    /// <c>D34</c>, <c>F34</c>). Fuera de la allow-list NADA se sella. Las celdas del espejo de
    /// REMUNERACIÓN que alimentan <c>O25/D34/F34</c> se resuelven por read-back (referencias + SUM
    /// + fórmulas compartidas); el sellado escribe el valor EVALUADO, incluido FALSE (jamás se
    /// inventa un TRUE — E1-2/D-E).
    ///
    /// El guard anti-fórmula de <see cref="EscribirValorNumerico"/> (celda-fórmula → ERR-PLANTILLA)
    /// NO se relaja: este es el camino separado <c>SellarCacheFormula</c>, que aserta el TEXTO
    /// esperado de la <c>&lt;f&gt;</c> y escribe SOLO <c>CellValue</c>/<c>DataType</c>.
    /// </summary>
    private static void SellarCachesFormula(WorkbookPart workbookPart, Periodo periodo)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        var hoja = WorkbookLeafCellMapValidaciones.HojaValidacionTotal;
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));
        var resolutor = new ResolutorValorValidacion(workbookPart);

        // C3:C8 = Σ 'Recaudo *'!{col}{18+k}: los 5 recaudos escritos por la app (fragmentos
        // orden-agnósticos). La columna la decide la quincena (E1-4): D en Q1, F en Q2 — el mismo
        // eje que PreparadorBasePipeline usa para reescribir la <f> de C3:C8.
        var columna = PreparadorBasePipeline.ColumnaRecaudoDe(periodo);
        for (var k = 3; k <= 8; k++)
        {
            var fila = 18 + k;
            var fragmentos = EmpresaFacturacion.Catalogo
                .OrderBy(e => e.Id)
                .Select(e => $"'{e.HojaRecaudo}'!{columna}{fila}")
                .ToArray();
            SellarCeldaNumerica(worksheet, hoja, $"C{k}", fragmentos, resolutor);
        }

        SellarCeldaNumerica(worksheet, hoja, "C9", ["SUM(C3:C8)"], resolutor);
        SellarCeldaNumerica(worksheet, hoja, "C14", ["'Valida - Control Recaudo'!F10"], resolutor);

        // Plan 40 (enmienda E-1): los 5 sub-bloques booleanos se sellan con el valor EVALUADO por
        // read-back (C15/D25 = recaudo; O25/D34/F34 = espejo de remuneración). Un sub-bloque que no
        // cierra se sella FALSE (0) — el sellado NUNCA inventa un TRUE: el caché del archivo
        // entregado debe coincidir con lo que Excel recalcularía (D-B/D-E), no mentir.
        foreach (var (celda, _) in FormasValidacionTotal.SubBloques)
        {
            SellarCeldaBooleana(worksheet, hoja, celda, resolutor);
        }
    }

    /// <summary>Plan 40 (T2, R-S-2): aserta el texto-<c>&lt;f&gt;</c> (fragmentos) y sella el <c>&lt;v&gt;</c> numérico por read-back; no computable → no se sella.</summary>
    private static void SellarCeldaNumerica(Worksheet worksheet, string hoja, string celda, string[] fragmentosEsperados, ResolutorValorValidacion resolutor)
    {
        var cell = ExigirCeldaFormula(worksheet, hoja, celda, fragmentosEsperados);
        if (!resolutor.TryEvaluar(hoja, celda, out var valor))
        {
            return; // dependencia no escrita por la app: sin read-back no se sella (R-S-1).
        }

        cell.CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture));
        cell.DataType = CellValues.Number;
    }

    /// <summary>Plan 40 (T2, R-S-2): sella el booleano de un sub-bloque de igualdad (C15 = C14=C9) evaluado por read-back.</summary>
    private static void SellarCeldaBooleana(Worksheet worksheet, string hoja, string celda, ResolutorValorValidacion resolutor)
    {
        var textoEsperado = FormasValidacionTotal.TextoEsperado(celda) ?? string.Empty;
        var cell = ExigirCeldaFormula(worksheet, hoja, celda, [textoEsperado]);
        var texto = cell.CellFormula?.Text ?? string.Empty;
        var partes = texto.Split('=');
        if (partes.Length != 2
            || !resolutor.TryEvaluar(hoja, partes[0], out var izquierdo)
            || !resolutor.TryEvaluar(hoja, partes[1], out var derecho))
        {
            return; // no computable → no se sella.
        }

        cell.CellValue = new CellValue(FormasValidacionTotal.Igualdad(izquierdo, derecho) ? "1" : "0");
        cell.DataType = CellValues.Boolean;
    }

    private static Cell ExigirCeldaFormula(Worksheet worksheet, string hoja, string celda, string[] fragmentosEsperados)
    {
        var cell = ObtenerCelda(worksheet, celda);
        if (cell?.CellFormula is null)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El sellado de caché esperaba {hoja}!{celda} como celda-fórmula y no lo es. (R-S-2: el sellado nunca crea ni convierte celdas.)");
        }

        var texto = cell.CellFormula.Text ?? string.Empty;
        foreach (var fragmento in fragmentosEsperados)
        {
            if (!texto.Contains(fragmento, StringComparison.Ordinal))
            {
                throw new CalculoInvalidoException(
                    CodigoError.Plantilla,
                    $"El sellado de caché {hoja}!{celda}: la fórmula <f> no coincide con la esperada ('{fragmento}'). Observado: '{texto}'. (R-S-2: nunca se sella sobre una <f> distinta.)");
            }
        }

        return cell;
    }

    private static void EscribirValorNumerico(WorkbookPart workbookPart, string hoja, string celda, decimal valor, string nombre)
    {
        var worksheet = ObtenerHoja(workbookPart, hoja, nameof(GenerarWorkbook));
        var cell = ObtenerOCrearCelda(worksheet, celda);

        // Plan 58 (H1.1): guard anti-fórmula ÚNICO. Una celda editable del mapa que es fórmula en la
        // plantilla = plantilla incompatible (nunca sobrescribir fórmulas) → ERR-PLANTILLA.
        GuardEscrituraFormula.ExigirSinFormula(cell, hoja, celda, nombre);

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
