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
/// Plan 36 (T1, R-B-1/R-B-2/R-B-4 — D-A/D-B/D-F) + Plan 41 (T1/Enmienda E-1) + Plan 44 (F1/Enmienda
/// E-44.1): escritor values-only MULTICOLUMNA del interior de <c>Reporte Componentes R1</c> sobre la
/// plantilla estandarizada única.
///
/// Doctrina slot: la app NUNCA inserta/borra filas y NUNCA escribe <c>&lt;f&gt;</c>; solo pega
/// VALORES (<c>&lt;v&gt;</c>) en el slot de datos que corresponde por firma + SECCIÓN + sinonimia;
/// lo no entregado queda en el cero del template; la fuente que trae fila sin slot hace fail-fast
/// <c>ERR-SIN-SLOT</c> (nunca remap silencioso, nunca insert, nunca colocación por posición).
///
/// Resolución (Plan 56 — identidad de cuadro A/B, §2-D2):
/// <list type="bullet">
/// <item>IDENTIDAD del cuadro = <c>(sección A, subsección B, ocurrencia de B dentro de A)</c>, leída
/// de las etiquetas de CIERRE de la propia hoja (<see cref="ContratoSeccionPeriodFreeR1.ParticionarJerarquia"/>):
/// <c>B=&lt;sub&gt;/C='Total'</c> cierra la subsección y <c>A=&lt;sec&gt;/B='Total'</c> cierra la
/// sección; <c>A='Total'</c> con B vacía es el gran total del bloque. NO hay lista de secciones ni
/// identificadores de bloque/empresa/período en el camino de resolución (D-H);</item>
/// <item>LLAVE de colocación dentro del cuadro: fuente <c>C</c> (por el catálogo abierto de sinonimia,
/// <see cref="SinonimosEmpresaR1.SonMismaEmpresa"/>) → canal <c>D</c> efectivo
/// (<see cref="ContratoSeccionPeriodFreeR1.FirmaCuadroR1"/>) → concepto <c>E</c>, con ordinal dentro
/// de la firma cuando se repite en el cuadro. Los cierres de subsección/sección/gran total son destino
/// declarado por su etiqueta;</item>
/// <item>reglas de §2-D3 (<see cref="ContratoSeccionPeriodFreeR1.PlanificarJerarquiaR1"/>): ausencia en
/// la fuente ⇒ cero del template; combinación ausente en TODO el bloque ⇒ <c>ERR-SIN-SLOT</c>
/// accionable ANTES de escribir nada (sección + subsección + fuente + canal + concepto + hoja + rango);
/// fila sin destino en SU cuadro pero declarada en el bloque ⇒ cero DOCUMENTADO (registro durable) si
/// el cuadro cierra al centavo por sus filas-hoja, error si no cierra. Prohibido remap/reacomodo,
/// insert/delete, <c>&lt;f&gt;</c> y colocación por posición;</item>
/// <item>una fuente SIN secciones declaradas (secuencia sintética sin cierres A/B) conserva el camino
/// histórico por firma (<see cref="TryPlanificar"/>): ningún bloque real de R1 cae ahí porque las
/// invariantes T0e <c>Componente/Total</c> y <c>Subs/Cont/Total</c> son cierres de sección.</item>
/// </list>
///
/// Contrato de columnas (Plan 44 F1, §2.3 del plan + T0 §1/§7): la plantilla resuelve las columnas
/// destino por ENCABEZADO del template (F..Z; la PRIMERA ocurrencia de un encabezado repetido manda:
/// el `Intereses Cte Financiac. Aseo` de la fuente va a **T**, y **Y** queda intacta). El
/// `Rec.Intereses` de la fuente (col U) va a **Z** de la plantilla. Las columnas `Exonera-*` (U/V/W/X)
/// del template NO tienen contraparte en la fuente y quedan INTACTAS por construcción.
///
/// Contrato de valores (W1 del review de T1): el slot recibe el DECIMAL EXACTO de la fuente (sin
/// redondeo); el redondeo a entero es exclusivo del <c>DetRetri</c> (Plan 29 Unidad D). Ver la
/// justificación completa en <see cref="EscribirValor"/>.
///
/// Guard pre-write (REGLA, no excepción): la celda destino debe existir y NO traer <c>&lt;f&gt;</c>;
/// si trae fórmula → <c>ERR-PLANTILLA</c> con hoja + celda + slot. Cero <c>InsertRow</c>/<c>DeleteRow</c>,
/// cero <c>new CellFormula</c>, cero Reanclar en este path (el grep de cierre T1 lo demuestra).
///
/// El writer NO recibe ningún período (D-H) y no nombra ninguna empresa (catálogo abierto).
/// <see cref="CompositorInteriorR1"/> queda como referencia de firma, no se usa aquí (D-C).
/// </summary>
internal static class EscritorValoresSlotR1
{
    public const string HojaR1 = DelimitadorBloquesHoja.NombreHojaR1;

    /// <summary>
    /// Resultado del pegado values-only: cuántas filas fuente se colocaron, cuántas celdas
    /// <c>&lt;v&gt;</c> se escribieron en las columnas F (Total) y L (especiales) — las del contrato
    /// histórico —, cuántas filas quedaron en el cero del template por valor nulo y cuántas cayeron en
    /// <b>cero documentado</b> (Plan 45 §2.2: fila sin slot de la lista T0; no se escribe). Los slots
    /// no consumidos quedan intactos por construcción. <see cref="EncabezadosSinDestino"/> recoge los
    /// encabezados de la fuente SIN columna destino en el template (H1.4; nunca descarte silencioso).
    /// </summary>
    internal sealed record EscrituraSlotR1Resultado(
        int FilasColocadas,
        int CeldasF,
        int CeldasL,
        int OmitidasPorNulo,
        int CerosDocumentados,
        IReadOnlyList<string> EncabezadosSinDestino)
    {
        /// <summary>
        /// Registro durable de los ceros documentados del Plan 56 (§2-D3): cuadro (A/B/ocurrencia),
        /// firma fuente, canal, concepto y valor de cada fila que quedó en el cero del template porque
        /// su cuadro cierra al centavo sin ella. Nunca descarte silencioso.
        /// </summary>
        public IReadOnlyList<ContratoSeccionPeriodFreeR1.CeroDocumentadoR1> CerosDetalle { get; init; } = [];
    }

    /// <summary>
    /// Pega los valores del bloque espejo (leído por <c>LeerEspejoR1</c>, intacto) en los slots del
    /// ASE sobre un <see cref="WorkbookPart"/> ya abierto en escritura. Solo <c>&lt;v&gt;</c>.
    /// </summary>
    internal static EscrituraSlotR1Resultado Escribir(WorkbookPart workbookPart, BloqueEspejoAseInputs bloque)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);
        ArgumentNullException.ThrowIfNull(bloque);

        var aseId = bloque.Ase.Id;
        var hoja = ObtenerHoja(workbookPart, HojaR1);
        // Plan 59 (H1.2): el bloque se DELIMITA de la propia hoja (cabecera por rótulo del ASE +
        // gran total declarado) y la identidad se verifica por ese mismo rótulo. Ya no hay ventana
        // de filas congelada: la plantilla es la que declara la geometría y una plantilla nueva con
        // la misma gramática se procesa sin cambiar código.
        var bloqueHoja = VerificarCabeceraBloque(hoja, workbookPart, bloque);
        var mapaColumnas = LeerMapaColumnasDestino(hoja, workbookPart, aseId, bloqueHoja);
        var fuente = bloque.Filas;

        // Plan 56 (§2-D2/D3): la resolución es la IDENTIDAD DE CUADRO (sección A, subsección B,
        // ocurrencia de B dentro de A) leída de las etiquetas de cierre, con llave C/D(efectivo)/E
        // dentro del cuadro. Una fuente sin secciones declaradas (fixture sintético sin cierres A/B)
        // conserva el camino histórico por firma: ningún bloque real cae ahí (las invariantes T0e
        // Componente/Total y Subs/Cont/Total son cierres de sección).
        // [P56-JERARQUIA-AB-INICIO] camino de resolución: geometría pura, SIN ids de bloque/ASE,
        // nombres de empresa, ramas por período ni direcciones de celda congeladas (candado §2-D5).
        var slots = LeerSlots(aseId, bloqueHoja);
        var planificado = ContratoSeccionPeriodFreeR1.PlanificarJerarquiaR1(fuente, slots.Select(FilaDesdeSlot).ToList());
        // [P56-JERARQUIA-AB-FIN]
        List<(FilaEspejoR1 Fila, SlotR1 Slot)> plan;
        int ceros;
        IReadOnlyList<ContratoSeccionPeriodFreeR1.CeroDocumentadoR1> cerosDetalle;
        if (planificado.AplicaJerarquia)
        {
            if (planificado.Fallo is { } fallo)
            {
                throw new CalculoInvalidoException(
                    CodigoError.SinSlot,
                    MensajeSinSlotJerarquia(aseId, fuente, fallo, bloqueHoja));
            }

            plan = planificado.Colocaciones.Select(c => (fuente[c.Fuente], slots[c.Slot])).ToList();
            ceros = planificado.Ceros.Count;
            cerosDetalle = planificado.Ceros;
        }
        else
        {
            if (!TryPlanificar(fuente, slots, out plan, out ceros, out var fallo))
            {
                throw new CalculoInvalidoException(
                    CodigoError.SinSlot,
                    MensajeSinSlot(aseId, fallo.Fila, fallo.Seccion, fallo.Ocurrencia, fallo.CeldaCandidata));
            }

            cerosDetalle = [];
        }

        // Registro DURABLE del cero documentado (§2-D3 rama 1): no se escribe y NO se aborta, pero
        // queda huella con cuadro, firma, canal, concepto y valor (nunca descarte silencioso).
        foreach (var cero in cerosDetalle)
        {
            Log.Warning(
                "{Hoja}: cero documentado (ASE {Ase}) — cuadro ({Seccion}, {Subseccion} #{Ocurrencia}) fuente '{Fuente}' canal '{Canal}' concepto '{Concepto}' valor {Valor}: sin destino en su cuadro y el subtotal del cuadro cierra al centavo sin esa fila (Plan 56 §2-D3).",
                HojaR1,
                aseId,
                cero.Cuadro.Seccion,
                cero.Cuadro.Subseccion,
                cero.Cuadro.Ocurrencia,
                cero.Fuente,
                cero.Canal,
                cero.Concepto,
                cero.Total);
        }

        var colocadas = 0;
        var celdasF = 0;
        var celdasL = 0;
        var omitidas = 0;
        var encabezadosSinDestino = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var (fila, slot) in plan)
        {
            EscribirComponentes(hoja, mapaColumnas, slot, fila, aseId, ref celdasF, ref celdasL, encabezadosSinDestino);
            colocadas++;
            if (fila.Valor("Total") is null)
            {
                omitidas++;
            }
        }

        hoja.Save();
        return new EscrituraSlotR1Resultado(colocadas, celdasF, celdasL, omitidas, ceros, encabezadosSinDestino.ToList())
        {
            CerosDetalle = cerosDetalle,
        };
    }

    /// <summary>Punto de fail-fast planificado (alimenta el mensaje <c>ERR-SIN-SLOT</c>).</summary>
    private readonly record struct FalloSinSlot(FilaEspejoR1 Fila, SeccionFronteraArriba? Seccion, int Ocurrencia, string? CeldaCandidata);

    /// <summary>
    /// Planifica (SIN escribir) un bloque que NO declara jerarquía A/B (secuencia sintética sin
    /// cierres de sección/subsección): comportamiento HISTÓRICO del Plan 41 — veto de sección SOLO en
    /// filas-dato de empresa (frontera ARRIBA) y cursor por firma, con cero documentado de la lista
    /// <see cref="CerosDocumentadosR1"/>. Ningún bloque real llega aquí (la fuente R1 siempre declara
    /// <c>Componente/Total</c> y <c>Subs/Cont/Total</c>, que son cierres de sección). Devuelve las
    /// colocaciones (fila → slot) en orden, el nº de ceros documentados y, si no es colocable, el
    /// punto de fail-fast. Un intento fallido no escribe nada (la escritura ocurre tras elegir el plan).
    /// </summary>
    private static bool TryPlanificar(
        IReadOnlyList<FilaEspejoR1> fuente,
        IReadOnlyList<SlotR1> slots,
        out List<(FilaEspejoR1 Fila, SlotR1 Slot)> plan,
        out int cerosDocumentados,
        out FalloSinSlot fallo)
    {
        var grupos = ConstruirGruposSeccion(slots);
        var secciones = ClasificarSeccionesBaseline(fuente);
        plan = [];
        cerosDocumentados = 0;
        fallo = default;

        for (var i = 0; i < fuente.Count; i++)
        {
            var fila = fuente[i];
            var seccion = secciones[i];
            var grupo = BuscarGrupo(grupos, fila, seccion, canal: null);
            var slot = grupo?.SiguienteLibre();

            if (slot is null)
            {
                if (CerosDocumentadosR1.EsCeroDocumentado(fila, seccion))
                {
                    cerosDocumentados++;
                    continue;
                }

                fallo = new FalloSinSlot(fila, seccion, ContarOcurrenciaPrevia(fuente, secciones, i) + 1, grupo?.UltimoConsumidoF);
                return false;
            }

            plan.Add((fila, slot));
        }

        return true;
    }

    /// <summary>
    /// Slot de dato del contrato: fila del template con su firma A–E, su SECCIÓN-CONTRATO (frontera
    /// que la CIERRA para las filas de datos; la sección que la propia frontera cierra para las
    /// filas-frontera; <c>null</c> para AFaseo) y su CANAL EFECTIVO (Plan 46 §2.1: <c>null</c> salvo
    /// en las filas de canal sin <c>D</c> propio). Las celdas F/L son referencias A1.
    /// </summary>
    internal sealed record SlotR1(int Fila, string CeldaF, string CeldaL, string A, string B, string C, string D, string E, SeccionFronteraArriba? Seccion, string? Canal = null);

    /// <summary>
    /// Grupo de slots de la MISMA firma y SECCIÓN-CONTRATO (Plan 41/T1 + Plan 45 §2.2) con su propio
    /// cursor ordenado. Cada (firma, sección-contrato) avanza su cursor sin interferir con los demás:
    /// un slot-resumen intercalado antes en el template sigue disponible para una fila que la fuente
    /// emite después. Las filas-frontera (<c>Mes/Total</c>, <c>Aplicacion…/Total</c>) llevan la sección
    /// que ABREN; solo AFaseo queda <c>null</c> (grupo por firma, comportamiento histórico). Plan 41
    /// (Enmienda E-1, E1-1): NO existe ningún cursor de respaldo — este es el ÚNICO mecanismo.
    /// </summary>
    private sealed class GrupoSeccion
    {
        private readonly List<SlotR1> _slots = [];
        private int _cursor;

        public GrupoSeccion(SlotR1 representante) => Representante = representante;

        /// <summary>Primer slot del grupo (identidad del grupo para el matcher firma+sección).</summary>
        public SlotR1 Representante { get; }

        /// <summary>Celda F del último slot consumido por el grupo ("(ninguno)" si aún no consumió).</summary>
        public string UltimoConsumidoF => _cursor == 0 ? "(ninguno)" : _slots[_cursor - 1].CeldaF;

        public void Agregar(SlotR1 slot) => _slots.Add(slot);

        /// <summary>Próximo slot libre del grupo en orden, o <c>null</c> si el grupo se agotó.</summary>
        public SlotR1? SiguienteLibre() => _cursor < _slots.Count ? _slots[_cursor++] : null;

        /// <summary>Verdad si el grupo aún tiene un slot libre (sin consumirlo).</summary>
        public bool TieneLibre => _cursor < _slots.Count;
    }

    /// <summary>
    /// Verdad si la fila fuente y el slot pertenecen a la misma firma + sección-contrato (localizador
    /// puro): AFaseo dual-firma primero (cualquiera de las dos variantes), luego A/B/D/E exactos
    /// (OrdinalIgnoreCase) + C por sinonimia; y —cuando AMBOS lados traen sección— la SECCIÓN debe
    /// coincidir. Solo AFaseo viaja sin sección (se localiza por firma). Cero ramas por nombre de
    /// empresa (catálogo abierto por construcción).
    ///
    /// Plan 46 (§2.1, sobrecarga con canal): cuando la fila trae CANAL EFECTIVO y el slot también,
    /// deben coincidir — así las filas sin <c>D</c> propio (VS/VI/Subsidio) caen en el slot de su
    /// canal real y no por cursor de orden. La sobrecarga de 3 argumentos (sin canal) preserva el
    /// comportamiento histórico y la usan los tests de sinonimia del Plan 36.
    /// </summary>
    internal static bool EncajaEnSlot(FilaEspejoR1 fila, SeccionFronteraArriba? seccionFuente, SlotR1 slot) =>
        EncajaEnSlot(fila, seccionFuente, canalFuente: null, slot);

    internal static bool EncajaEnSlot(FilaEspejoR1 fila, SeccionFronteraArriba? seccionFuente, string? canalFuente, SlotR1 slot)
    {
        ArgumentNullException.ThrowIfNull(fila);
        ArgumentNullException.ThrowIfNull(slot);

        if (EsFirmaAFaseo(fila.A, fila.B, fila.D, fila.E))
        {
            return EsFirmaAFaseo(slot.A, slot.B, slot.D, slot.E);
        }

        if (EsFirmaAFaseo(slot.A, slot.B, slot.D, slot.E))
        {
            return false;
        }

        if (!IgualEtiqueta(fila.A, slot.A)
            || !IgualEtiqueta(fila.B, slot.B)
            || !IgualEtiqueta(fila.D, slot.D)
            || !IgualEtiqueta(fila.E, slot.E))
        {
            return false;
        }

        if (!ColumnasCCompatibles(fila.C, slot.C))
        {
            return false;
        }

        // Plan 41-T1 (D-A) + Plan 45 (§2.2) como parte de la llave: Oportuno vs Aplicación y las
        // ocurrencias NO se mezclan. Si AMBOS lados traen sección (filas de datos/componentes) deben
        // coincidir; si algún lado no participa (null: fila-frontera/AFaseo — se localizan por firma),
        // la sección no veta.
        if (seccionFuente.HasValue
            && slot.Seccion.HasValue
            && seccionFuente.Value != slot.Seccion.Value)
        {
            return false;
        }

        // Plan 46 (§2.1, DEFECTO 1): el canal efectivo es parte de la llave. Si ambos lados lo traen
        // (filas/Slots de canal) y difieren, no encajan.
        if (canalFuente is not null
            && slot.Canal is not null
            && !string.Equals(canalFuente, slot.Canal, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return true;
    }

    /// <summary>
    /// Firma AFaseo dual (R-T-2): variante fuente <c>A='AFaseo' B='Total'</c> o variante predicado
    /// <c>D='AFASEO' E='Total'</c>. Ambas resuelven al mismo slot F0 del ASE.
    /// </summary>
    internal static bool EsFirmaAFaseo(string a, string b, string d, string e) =>
        (IgualEtiqueta(a, "AFaseo") && IgualEtiqueta(b, "Total"))
        || (IgualEtiqueta(d, "AFASEO") && IgualEtiqueta(e, "Total"));

    private static bool IgualEtiqueta(string? x, string? y) =>
        string.Equals(x?.Trim(), y?.Trim(), StringComparison.OrdinalIgnoreCase);

    private static bool ColumnasCCompatibles(string? fuente, string? slot)
    {
        if (string.IsNullOrWhiteSpace(fuente) && string.IsNullOrWhiteSpace(slot))
        {
            return true;
        }

        if (string.IsNullOrWhiteSpace(fuente) || string.IsNullOrWhiteSpace(slot))
        {
            return false;
        }

        return SinonimosEmpresaR1.SonMismaEmpresa(fuente, slot);
    }

    /// <summary>
    /// Agrupa los slots del bloque por FIRMA + SECCIÓN (Plan 41/T1, D-A): cada grupo conserva su
    /// propia lista ordenada de slots (en el orden físico del template) y su propio cursor. Un slot
    /// se suma al primer grupo cuyo representante coincide (misma firma A/B/D/E + C por sinonimia +
    /// AFaseo dual + misma sección en filas de datos); si no encaja en ninguno, abre un grupo nuevo.
    /// </summary>
    private static List<GrupoSeccion> ConstruirGruposSeccion(IReadOnlyList<SlotR1> slots)
    {
        var grupos = new List<GrupoSeccion>();
        foreach (var slot in slots)
        {
            var firmaSlot = FilaDesdeSlot(slot);
            var grupo = grupos.Find(g => EncajaEnSlot(firmaSlot, slot.Seccion, slot.Canal, g.Representante));
            if (grupo is null)
            {
                grupo = new GrupoSeccion(slot);
                grupos.Add(grupo);
            }

            grupo.Agregar(slot);
        }

        return grupos;
    }

    private static FilaEspejoR1 FilaDesdeSlot(SlotR1 slot) => new()
    {
        A = slot.A,
        B = slot.B,
        C = slot.C,
        D = slot.D,
        E = slot.E,
    };

    /// <summary>
    /// Primer grupo (en orden físico del template) que encaja con la fila por firma + (<paramref name="seccion"/>
    /// o sin sección si es <c>null</c>) + (canal efectivo o sin canal) Y aún tiene un slot libre. Con
    /// <paramref name="seccion"/> no nulo es la resolución por SECCIÓN-CONTRATO + CANAL (Plan 45 §2.2 +
    /// Plan 46 §2.1); con <c>null</c> es el POOL HISTÓRICO (firma sola, Plan 41) que preserva la
    /// neutralidad Q1 (nunca un fail-fast antes que el previo).
    /// </summary>
    private static GrupoSeccion? BuscarGrupo(IReadOnlyList<GrupoSeccion> grupos, FilaEspejoR1 fila, SeccionFronteraArriba? seccion, string? canal)
    {
        foreach (var grupo in grupos)
        {
            if (grupo.TieneLibre && EncajaEnSlot(fila, seccion, canal, grupo.Representante))
            {
                return grupo;
            }
        }

        return null;
    }

    /// <summary>
    /// Mensaje de <c>ERR-SIN-SLOT</c> (Plan 41 E1-3 + Plan 44 E-44.1): nombra la FILA (firma
    /// A|B|C|D|E + ocurrencia), su RÓTULO DE EMPRESA (columna C), el CANAL (columna D) y el CONCEPTO
    /// (columna E) de la fila, y la SECCIÓN ESPERADA (la resuelta por su frontera inmediatamente
    /// arriba y mapeada por el contrato) — información accionable, nunca un fallo mudo.
    /// </summary>
    private static string MensajeSinSlot(int aseId, FilaEspejoR1 fila, SeccionFronteraArriba? seccion, int ocurrencia, string? celdaCandidata)
    {
        var razon = celdaCandidata is null
            ? "sin ningún slot con firma + sección compatibles en el contrato T0"
            : $"sin slot libre en el contrato T0 (celda candidata: último slot consumido {celdaCandidata})";
        return $"ERR-SIN-SLOT: ASE {aseId}: la fuente R1 (Recaudoporcomponente) trae la fila " +
            $"A|B|C|D|E='{fila.Firma}' con rótulo de empresa '{fila.C}', canal '{fila.D}' y concepto '{fila.E}' " +
            $"en la sección esperada '{TipoSeccion(seccion)}' (ocurrencia #{ocurrencia} de esa firma) {razon} del ASE {aseId}. " +
            $"Nunca remap/insert ni colocación por posición: revise el contrato T0 (slot por sección) o la fuente.";
    }

    /// <summary>
    /// Mensaje de <c>ERR-SIN-SLOT</c> del camino JERÁRQUICO (Plan 56 §2-D3 / G-TRAZA): nombra el
    /// cuadro (sección A + subsección B + ocurrencia), la fuente (C), el canal (D), el concepto (E),
    /// la hoja y el rango físico del bloque, en vocabulario de la propietaria. Describe el hecho sin
    /// culpar a la fuente y deja explícito que no se escribe nada del bloque.
    /// </summary>
    private static string MensajeSinSlotJerarquia(
        int aseId,
        IReadOnlyList<FilaEspejoR1> fuente,
        ContratoSeccionPeriodFreeR1.FalloJerarquiaR1 fallo,
        DelimitadorBloquesHoja.BloqueHoja bloque)
    {
        var fila = fuente[fallo.IndiceFilaFuente];
        var cuadro = fallo.Seccion is null
            ? "el gran total del bloque"
            : fallo.Subseccion is null
                ? $"la sección '{fallo.Seccion}'"
                : $"el cuadro (sección '{fallo.Seccion}', subsección '{fallo.Subseccion}', ocurrencia #{fallo.Ocurrencia})";
        var fuenteRotulo = fallo.Fuente.Length == 0 ? "(sin rótulo)" : fallo.Fuente;
        var canalRotulo = fallo.Canal.Length == 0 ? "(sin canal propio)" : fallo.Canal;
        var conceptoRotulo = fallo.Concepto.Length == 0 ? "(sin concepto)" : fallo.Concepto;
        return $"ERR-SIN-SLOT: ASE {aseId}: la fila A|B|C|D|E='{fila.Firma}' de la fuente R1 (Recaudoporcomponente) " +
            $"no tiene destino en {cuadro}: fuente (C) '{fuenteRotulo}', canal (D) '{canalRotulo}', " +
            $"concepto (E) '{conceptoRotulo}'; hoja '{HojaR1}', rango del bloque {bloque.Cabecera}-{bloque.GranTotal} " +
            $"(delimitado por la cabecera '{bloque.Rotulo}' y su gran total declarado). " +
            $"Detalle: {fallo.Motivo}. No se escribe nada del bloque: la plantilla no declara ese destino, " +
            $"y no se reacomoda a otro cuadro ni se coloca por posición.";
    }

    private static int ContarOcurrenciaPrevia(
        IReadOnlyList<FilaEspejoR1> fuente,
        IReadOnlyList<SeccionFronteraArriba?> secciones,
        int indice)
    {
        var conteo = 0;
        for (var k = 0; k < indice; k++)
        {
            if (EncajaEnSlot(fuente[indice], secciones[indice], SlotDesdeFila(fuente[k], secciones[k], canal: null)))
            {
                conteo++;
            }
        }

        return conteo;
    }

    private static SlotR1 SlotDesdeFila(FilaEspejoR1 fila, SeccionFronteraArriba? seccion, string? canal) =>
        new(0, string.Empty, string.Empty, fila.A, fila.B, fila.C, fila.D, fila.E, seccion, canal);

    private static string TipoSeccion(SeccionFronteraArriba? seccion) => seccion switch
    {
        null => "sin-sección",
        { Tipo: TipoSeccionEspejoR1.Primera } => "1.ª sección",
        { Tipo: TipoSeccionEspejoR1.Oportuno } => $"Oportuno #{seccion.Value.Ordinal}",
        { Tipo: TipoSeccionEspejoR1.Aplicacion } => $"Aplicación #{seccion.Value.Ordinal}",
        _ => $"desconocida #{seccion.Value.Ordinal}",
    };

    /// <summary>
    /// Sección del comportamiento HISTÓRICO (Plan 41): la frontera inmediatamente ARRIBA SOLO en las
    /// filas-dato de empresa (<see cref="R1FirmaInterior.EsDatoEmpresa"/>); <c>null</c> para el resto
    /// (cursor por firma). Es la reproducción byte-fiel del comportamiento previo a la regla sin-C —
    /// solo aplica a secuencias SIN jerarquía declarada (fixtures sintéticos del contrato por firma).
    /// </summary>
    private static IReadOnlyList<SeccionFronteraArriba?> ClasificarSeccionesBaseline(IReadOnlyList<FilaEspejoR1> filas)
    {
        var resultado = new SeccionFronteraArriba?[filas.Count];
        for (var i = 0; i < filas.Count; i++)
        {
            resultado[i] = R1FirmaInterior.EsDatoEmpresa(filas[i])
                ? SeccionEspejoR1.PorFronteraArriba(filas, i)
                : null;
        }

        return resultado;
    }

    private static List<SlotR1> LeerSlots(
        int aseId,
        DelimitadorBloquesHoja.BloqueHoja bloque)
    {
        // Plan 59 (H1.2): los slots son las filas DELIMITADAS del bloque (cabecera..gran total)
        // menos las dos filas que no son datos: la cabecera de bloque (rótulo del ASE) y la fila de
        // encabezados de columna. La fila que no declara firma de cierre es dato por definición: no
        // se calcula ninguna "primera fila de datos" por offset.
        var filas = bloque.FilasSlot;
        var firmas = filas.Select(f => f.Etiquetas).ToList();
        var secciones = ClasificarSeccionesBaseline(firmas);
        var slots = new List<SlotR1>(firmas.Count);
        for (var i = 0; i < firmas.Count; i++)
        {
            var numero = filas[i].Fila;
            slots.Add(new SlotR1(
                numero, $"F{numero}", $"L{numero}",
                firmas[i].A, firmas[i].B, firmas[i].C, firmas[i].D, firmas[i].E,
                secciones[i]));
        }

        if (slots.Count == 0)
        {
            throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El bloque R1 del ASE {aseId} ('{bloque.Rotulo}', {bloque.Cabecera}-{bloque.GranTotal}) no declara filas de slot: la delimitación por cabecera + gran total no deja ninguna fila de datos ni de cierre.");
        }

        return slots;
    }

    /// <summary>
    /// Plan 44 (F1, §2.3): resuelve el mapeo ENCABEZADO→LETRA de las columnas del bloque R1 desde la
    /// fila de encabezados de columna que el DELIMITADOR localizó por estructura (declara rótulos y
    /// ningún valor; Plan 59 H1.2: ya no es «inicio − 1»). Un encabezado repetido conserva la
    /// PRIMERA ocurrencia (T0 §7: `Intereses Cte Financiac. Aseo` → **T**, no Y). Las columnas
    /// `Exonera-*` quedan en el mapa pero la fuente no las trae → nunca se escriben.
    /// </summary>
    private static IReadOnlyDictionary<string, string> LeerMapaColumnasDestino(
        Worksheet hoja,
        WorkbookPart workbookPart,
        int aseId,
        DelimitadorBloquesHoja.BloqueHoja bloque)
    {
        var filaHeader = hoja.Elements<SheetData>().FirstOrDefault()?
            .Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == bloque.EncabezadoColumnas)
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"El bloque R1 del ASE {aseId} ('{bloque.Rotulo}') no trae la fila de encabezados de columna {bloque.EncabezadoColumnas} que el delimitador localizó (contrato de columnas por encabezado, Plan 44 F1).");

        var mapa = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var letra in DelimitadorBloquesHoja.ColumnasComponentes)
        {
            var texto = LectorTextoHoja.Texto(workbookPart, filaHeader, $"{letra}{bloque.EncabezadoColumnas}").Trim();
            if (texto.Length == 0 || mapa.ContainsKey(texto))
            {
                continue;
            }

            mapa[texto] = letra;
        }

        return mapa;
    }

    /// <summary>
    /// Plan 44 (F1): pega TODOS los componentes de la fila fuente en el slot, por ENCABEZADO
    /// (F..Z). Solo escribe los encabezados que existen en la fuente Y tienen columna destino en el
    /// template; los valores nulos se omiten (queda el cero del template). Decimal exacto.
    /// </summary>
    private static void EscribirComponentes(
        Worksheet hoja,
        IReadOnlyDictionary<string, string> mapaColumnas,
        SlotR1 slot,
        FilaEspejoR1 fila,
        int aseId,
        ref int celdasF,
        ref int celdasL,
        HashSet<string> encabezadosSinDestino)
    {
        foreach (var (encabezado, valor) in fila.ValoresPorColumna)
        {
            if (valor is null)
            {
                continue;
            }

            if (!mapaColumnas.TryGetValue(encabezado, out var columna))
            {
                // H1.4 (review F1 SPT-1): encabezado de la fuente SIN columna destino en el template.
                // Nunca descarte silencioso: se registra accionable (ASE + fila + encabezado) y se
                // recoge en el resultado. Hoy inocuo (T0 §1.2: F..U fuente mapea 1:1 a F..Z).
                encabezadosSinDestino.Add(encabezado);
                Log.Warning(
                    "{Hoja}: encabezado de fuente sin columna destino '{Encabezado}' (ASE {Ase}, fila slot {Fila}); valor descartado — revise el contrato de columnas (H1.4).",
                    HojaR1,
                    encabezado,
                    aseId,
                    slot.Fila);
                continue;
            }

            EscribirValor($"{columna}{slot.Fila}", hoja, valor.Value, aseId, slot.Fila);
            if (string.Equals(columna, "F", StringComparison.OrdinalIgnoreCase))
            {
                celdasF++;
            }
            else if (string.Equals(columna, "L", StringComparison.OrdinalIgnoreCase))
            {
                celdasL++;
            }
        }
    }

    /// <summary>
    /// Pega el valor del slot como <c>&lt;v&gt;</c> numérico. Contrato de redondeo (W1 del review de
    /// T1, decidido aquí): el slot R1 conserva el DECIMAL EXACTO de la fuente — NO se redondea.
    ///
    /// Por qué: el §2.1 del plan dice "número redondeado según contrato", y el contrato que manda
    /// (T0 congelado desde disco + el comportamiento histórico del writer por encabezado) es
    /// "mismo valor de la fuente": los slots F/L guardan importes con decimales (p. ej.
    /// <c>F649 = 13359789539.23</c> en el template) y el gate los compara contra el dato de la
    /// fuente. El redondeo a ENTERO es exclusivo del <c>DetRetri</c> (columna D, Plan 29 Unidad D,
    /// <c>DetRetriRounder</c>) y vive allí, no aquí. Redondear en el slot introduciría una
    /// divergencia silenciosa contra la fuente dentro de la tolerancia ±0.5 y rompería la
    /// igualdad exacta que el contrato T0/T1 aserta. Guard anti-fórmula: celda destino con
    /// <c>&lt;f&gt;</c> → <c>ERR-PLANTILLA</c> (nunca se sobrescribe una fórmula).
    /// </summary>
    private static void EscribirValor(string celda, Worksheet hoja, decimal valor, int aseId, int filaSlot)
    {
        var cell = hoja.Descendants<Cell>().FirstOrDefault(c =>
                string.Equals(c.CellReference?.Value, celda, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(
                CodigoError.Plantilla,
                $"La celda slot '{HojaR1}!{celda}' (fila {filaSlot} del contrato T0, ASE {aseId}) no existe en la plantilla. Fail-fast, nunca se inventa la celda.");

        // Plan 58 (H1.1): guard anti-fórmula único del escritor values-only.
        GuardEscrituraFormula.ExigirSinFormula(
            cell, HojaR1, celda, $"slot de la fila {filaSlot} del contrato T0, ASE {aseId}");

        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(valor.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Plan 59 (H1.2 / S2.1): verificación de la CABECERA DE BLOQUE por IDENTIDAD DE RÓTULO.
    ///
    /// Sustituye la verificación por fila congelada (cabecera deducida por offset de la ventana): el
    /// bloque se delimita leyendo la hoja (<see cref="DelimitadorBloquesHoja"/>) y su identidad se
    /// contrasta contra el catálogo de dominio (<see cref="AseFactory"/>) con el catálogo abierto de
    /// sinonimia (<see cref="SinonimosEmpresaR1.SonMismaEmpresa"/>), la ÚNICA normalización de
    /// rótulos de empresa del proyecto. Un workbook cuyo bloque no declara el ASE esperado (o declara
    /// su estructura de cierres incompleta) hace fail-fast antes de pegar ningún valor, nombrando
    /// hoja + rótulo esperado + bloques hallados. Cero remapeo por posición.
    /// </summary>
    private static DelimitadorBloquesHoja.BloqueHoja VerificarCabeceraBloque(
        Worksheet hoja,
        WorkbookPart workbookPart,
        BloqueEspejoAseInputs bloque) =>
        DelimitadorBloquesHoja.DelimitarBloque(workbookPart, DelimitadorBloquesHoja.R1, bloque.Ase.Id, hoja);

    private static Worksheet ObtenerHoja(WorkbookPart workbookPart, string nombreHoja)
    {
        var hoja = workbookPart.Workbook?.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, nombreHoja, StringComparison.OrdinalIgnoreCase))
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"El workbook no trae la hoja '{nombreHoja}' para el escritor de slots R1.");
        var parte = (WorksheetPart)workbookPart.GetPartById(hoja.Id!);
        return parte.Worksheet
            ?? throw new CalculoInvalidoException(CodigoError.Plantilla, $"La hoja '{nombreHoja}' no tiene WorksheetPart válido.");
    }
}
