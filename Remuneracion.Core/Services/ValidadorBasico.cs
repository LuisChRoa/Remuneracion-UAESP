using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;

namespace Remuneracion.Core.Services;

/// <summary>
/// Validador de dominio de la remuneración.
/// HU-07: overload multi-ASE con matcheo ESTRICTO por <see cref="Ase.Id"/> (se elimina el
/// fallback a <see cref="Enumerable.FirstOrDefault"/>); el gate R1 compara la primera fila
/// Mes/Total (F25-equivalente por bloque) contra <see cref="ConsolidadoAse.Extemp"/> — nunca
/// <see cref="ConsolidadoAse.TotOpt"/> contra visibles R1.
/// </summary>
public sealed class ValidadorBasico : IValidador
{
    private const decimal Tolerancia = 0.5m;

    /// <summary>
    /// Plan 46 (review W-1): catálogo de divergencias declaradas que consulta la clasificación.
    /// Default = catálogo REAL del producto (<see cref="DivergenciasDeclaradasCatalogo.Entradas"/>;
    /// el matcher puro ya ignora las entradas retiradas). La costura permite inyectar un catálogo
    /// sintético de test SIN duplicar la lógica de clasificación: el fixture ejercita ESTA misma
    /// implementación (antes el fixture espejaba la clasificación línea a línea).
    /// </summary>
    private readonly IReadOnlyList<DivergenciasDeclaradasCatalogo.Entrada> _catalogo;

    /// <summary>
    /// Crea el validador. <paramref name="catalogo"/> <c>null</c> = catálogo real del producto
    /// (comportamiento de producción intacto). Un catálogo explícito solo lo usa el fixture de test.
    /// </summary>
    public ValidadorBasico(IReadOnlyList<DivergenciasDeclaradasCatalogo.Entrada>? catalogo = null)
    {
        _catalogo = catalogo ?? DivergenciasDeclaradasCatalogo.Entradas;
    }

    /// <summary>
    /// HU-14 (3.1): agrega un error con el prefijo <c>[ERR-VALIDACION]</c> del catálogo
    /// (todo error del validador es una validación de dominio que no cierra).
    /// </summary>
    private static void AgregarError(List<string> errores, string mensaje) =>
        errores.Add($"[{CodigoError.Validacion}] {mensaje}");

    public List<string> Validar(ResultadoRemuneracion resultado)
    {
        ArgumentNullException.ThrowIfNull(resultado);

        var errores = new List<string>();

        if (resultado.Consolidados.Count != 1)
        {
            AgregarError(errores, "Debe existir exactamente 1 consolidado para validar la hoja de salida en Fase 1.");
        }

        if (resultado.Consolidados.Count > 0 && resultado.Consolidados.Any(c => Math.Abs(c.AjustesSfT) > Tolerancia))
        {
            AgregarError(errores, $"AjustesSfT debe ser 0 en Fase 1, pero se encontró {resultado.Consolidados.First().AjustesSfT}.");
        }

        var consolidado = resultado.Consolidados.FirstOrDefault();
        if (consolidado is not null)
        {
            var diffGranTotal = Math.Abs(resultado.GranTotal - consolidado.TotalAse);
            if (diffGranTotal > Tolerancia)
            {
                AgregarError(errores, $"GranTotal no coincide con el único TotalAse: {resultado.GranTotal} vs {consolidado.TotalAse}. Diferencia={diffGranTotal}.");
            }
        }

        if (resultado.Consolidados.Count == 1)
        {
            var totalAse = resultado.Consolidados[0].TotalAse;
            var granTotal = resultado.GranTotal;
            if (Math.Abs(granTotal - totalAse) > Tolerancia)
            {
                AgregarError(errores, $"GranTotal ≠ único TotalAse ({granTotal} ≠ {totalAse}).");
            }
        }

        return errores;
    }

    public List<string> Validar(ResultadoRemuneracion resultado, WorkbookLeafInputs leaf)
    {
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leaf);

        var errores = Validar(resultado);

        // HU-07: matcheo estricto por Ase.Id — si falta el ASE, es error; nunca fallback.
        var consolidado = resultado.Consolidados.SingleOrDefault(c => c.Ase.Id == leaf.Ase.Id);
        if (consolidado is null)
        {
            AgregarError(errores, $"No se encontró el consolidado del ASE {leaf.Ase.Id} para comparar contra los inputs leaf.");
            return errores;
        }

        if (Math.Abs(consolidado.AjustesSfT) > Tolerancia)
        {
            AgregarError(errores, $"AjustesSfT debe ser 0: {consolidado.AjustesSfT}.");
        }

        ValidarGatesPorAse(errores, leaf, consolidado);

        return errores;
    }

    public List<string> Validar(ResultadoRemuneracion resultado, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leafs);

        var errores = new List<string>();

        // §2.5 regla 1: Consolidados.Count == leafs.Count; en modo período se exige 5.
        if (resultado.Consolidados.Count != leafs.Count)
        {
            AgregarError(errores, $"El modo período exige la misma cantidad de consolidados y leafs: {resultado.Consolidados.Count} consolidados vs {leafs.Count} leafs.");
        }

        if (leafs.Count != 5)
        {
            AgregarError(errores,
                leafs.Count == 1
                    ? "El modo período exige exactamente 5 ASE; se recibió 1 (¿usó el modo single-ASE?)."
                    : $"El modo período exige exactamente 5 ASE; se recibieron {leafs.Count}.");
        }

        // §2.5 regla 2: sin duplicados de Ase.Id.
        var duplicados = leafs
            .GroupBy(l => l.Ase.Id)
            .FirstOrDefault(g => g.Count() > 1);
        if (duplicados is not null)
        {
            AgregarError(errores, $"Se detectaron leafs con ASE duplicado: {duplicados.Key}.");
        }

        // §2.5 regla 3: AjustesSfT por quincena — Q1 exige 0 (intacto); Q2 exige el TotalAjustes
        // de la composición T0-0.3 ±0.5 con matcheo estricto por Ase.Id (D5). Si en Q2 un leaf no
        // trae AjustesSfT, es error que nombra el ASE (nunca 0 silencioso).
        var esQuincena2 = resultado.Periodo.NumeroQuincena == 2;
        foreach (var consolidado in resultado.Consolidados)
        {
            if (!esQuincena2)
            {
                if (Math.Abs(consolidado.AjustesSfT) > Tolerancia)
                {
                    AgregarError(errores, $"AjustesSfT debe ser 0 en Q1; ASE {consolidado.Ase.Id} tiene {consolidado.AjustesSfT}.");
                }

                continue;
            }

            var ajustes = leafs.SingleOrDefault(l => l.Ase.Id == consolidado.Ase.Id)?.AjustesSfT;
            if (ajustes is null)
            {
                AgregarError(errores, $"ASE {consolidado.Ase.Id}: en Q2 el leaf debe traer AjustesSfT (SALDOS POR NOTA + RETRIBUCION NEGATIVA) para validar el gate; no puede ser null.");
                continue;
            }

            var diferencia = Math.Abs(consolidado.AjustesSfT - ajustes.TotalAjustes);
            if (diferencia > Tolerancia)
            {
                AgregarError(errores, $"ASE {consolidado.Ase.Id}: AjustesSfT del consolidado ({consolidado.AjustesSfT}) no coincide con TotalAjustes de la fuente ({ajustes.TotalAjustes}). Diferencia={diferencia} > ±{Tolerancia}.");
            }
        }

        // §2.5 regla 4: gate leaf-vs-consolidado POR ASE con matcheo estricto (Single por Id).
        foreach (var leaf in leafs)
        {
            var consolidado = resultado.Consolidados.SingleOrDefault(c => c.Ase.Id == leaf.Ase.Id);
            if (consolidado is null)
            {
                AgregarError(errores, $"No se encontró el consolidado del ASE {leaf.Ase.Id} para el gate leaf-vs-consolidado.");
                continue;
            }

            ValidarGatesPorAse(errores, leaf, consolidado);

            // §2.5 regla 2.2 (HU-08): gate Σ empresas = visible de bloque por ASE y hoja
            // (R1/R2/R4), tolerancia ±0.5, ceros legítimos. Lista vacía = comportamiento HU-07.
            ValidarSigmaEmpresasPorAse(errores, leaf);
        }

        // Plan 50 (A3, §2.5 — control mínimo alternativo): evidencia de etiquetas de fuente no
        // reconocidas + ausencias declaradas en el MISMO reporte = posible fuente presente no
        // reconocida → fail-fast (nunca un 0 silencioso).
        ValidarFuentesRecaudoNoReconocidas(errores, leafs);

        // Plan 51 (A3-banco, §2.3): el reporte por banco tiene su propio juicio de etiquetas
        // (grano ASE, canal durable `BANCO:`). Ausencia coherente = normal (0 de plantilla +
        // huella); etiqueta desconocida CON ausencia en el mismo bloque = posible fuente presente
        // no reconocida → fail-fast. Sin ausencia, la etiqueta desconocida es solo evidencia.
        ValidarEtiquetasBancoNoReconocidas(errores, leafs);

        // HU-12 (2.6 ampliada, §2.5 regla 2) / Plan 59 (§11 — FRENTE 2): DetRetri por ASE — en AMBAS
        // quincenas. El gate se ACTIVA cuando el período declara el asiento (algún leaf trae
        // DetRetriQ2, que es lo que hace siempre el procesador del período completo en las dos
        // quincenas) y entonces exige que TODOS los ASE lo traigan — un leaf sin DetRetri en un
        // período que sí lo declara es un hueco, no un 0 legítimo. Antes la regla vivía dentro de
        // `if (esQuincena2)`: en Q1 la comprobación simplemente no ocurría.
        var declaraDetRetri = leafs.Any(l => l.DetRetriQ2 is not null);
        foreach (var leaf in leafs)
        {
            var detalle = leaf.DetRetriQ2;
            if (detalle is null)
            {
                if (declaraDetRetri)
                {
                    AgregarError(errores, $"ASE {leaf.Ase.Id}: el leaf debe traer DetRetriQ2 (ROUND(D104:D108,0)) para validar el gate; no puede ser null.");
                }

                continue;
            }

            var redondeado = DetRetriRounder.Round(detalle.TotalD104);
            var diferencia = Math.Abs(detalle.Detalle - redondeado);
            if (diferencia > Tolerancia)
            {
                AgregarError(errores, $"ASE {leaf.Ase.Id}: Detalle DetRetri ({detalle.Detalle}) no coincide con ROUND(D104:D108,0) ({redondeado}). Diferencia={diferencia} > ±{Tolerancia}.");
            }
        }

        // §2.5 regla 5: GranTotal == Σ TotalAse (±0.5) + TotalAse aritmético por ASE.
        var sumaTotalAse = resultado.Consolidados.Sum(c => c.TotalAse);
        if (Math.Abs(resultado.GranTotal - sumaTotalAse) > Tolerancia)
        {
            AgregarError(errores, $"GranTotal ({resultado.GranTotal}) no coincide con Σ TotalAse ({sumaTotalAse}).");
        }

        foreach (var consolidado in resultado.Consolidados)
        {
            var totalAritmetico = consolidado.TotOpt
                + consolidado.R2TotalOportuno
                + consolidado.Extemp
                + consolidado.ReversionR4
                + consolidado.AjustesSfT;
            if (Math.Abs(consolidado.TotalAse - totalAritmetico) > Tolerancia)
            {
                AgregarError(errores, $"TotalAse del ASE {consolidado.Ase.Id} ({consolidado.TotalAse}) no coincide con su aritmética ({totalAritmetico}).");
            }
        }

        // HU-09 (2.3, §2.5 reglas 2-4): gates D5 del reporte por banco.
        ValidarReporteBanco(errores, leafs);

        // HU-10 (2.4, §2.5 reglas 2-4): gates D5 del balance de subsidios y contribuciones.
        ValidarBalanceSc(errores, leafs);

        return errores;
    }

    public List<string> Validar(
        ResultadoRemuneracion resultado,
        IReadOnlyList<WorkbookLeafInputs> leafs,
        IReadOnlyList<ValidacionCruzadaSnapshot> snapshots)
    {
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leafs);
        ArgumentNullException.ThrowIfNull(snapshots);

        // Todo HU-01..HU-12 intacto (gates existentes con matcheo estricto por Ase.Id).
        var errores = Validar(resultado, leafs);

        // §2.5 regla 2.7-0: la lista de snapshots debe tener la misma cantidad de ASE que los
        // consolidados/leafs. Snapshot ausente/vacío para un ASE = HU-12 puro para ese ASE.
        if (resultado.Consolidados.Count != snapshots.Count)
        {
            AgregarError(errores, $"El modo validaciones cruzadas exige la misma cantidad de consolidados y snapshots: {resultado.Consolidados.Count} consolidados vs {snapshots.Count} snapshots.");
        }

        if (snapshots.Count != 5)
        {
            AgregarError(errores, $"El modo validaciones cruzadas exige exactamente 5 snapshots (uno por ASE); se recibieron {snapshots.Count}.");
        }

        var snapshotsOrdenados = snapshots.OrderBy(s => s.Ase.Id).ToList();
        foreach (var snapshot in snapshotsOrdenados)
        {
            ValidarGatesValidacionesCruzadasPorAse(errores, snapshot, resultado, leafs);
        }

        // HU-14 (S-2, D8): VALIDACION_TOTAL con gate ÚNICO (no ×5) + igualdad entre snapshots.
        // El reader lee O9/P9 una sola vez; el gate evalúa el primer snapshot y la igualdad
        // detecta una divergencia del lector (fail-fast con 1 error, no 5).
        if (snapshotsOrdenados.Count > 0)
        {
            var baseTotal = snapshotsOrdenados[0];
            if (Math.Abs(baseTotal.ValidacionTotal) > Tolerancia)
            {
                AgregarError(errores, $"ASE {baseTotal.Ase.Id}: VALIDACION_TOTAL (O9) no cierra: {baseTotal.ValidacionTotal} (debe ser 0 ±{Tolerancia}).");
            }

            if (!baseTotal.ValidacionTotalOkP)
            {
                AgregarError(errores, $"ASE {baseTotal.Ase.Id}: VALIDACION_TOTAL P9 (INT(O9)=0) es falso en el workbook.");
            }

            foreach (var snapshot in snapshotsOrdenados.Skip(1))
            {
                if (snapshot.ValidacionTotal != baseTotal.ValidacionTotal
                    || snapshot.ValidacionTotalOkP != baseTotal.ValidacionTotalOkP)
                {
                    AgregarError(errores, $"VALIDACION_TOTAL: el snapshot del ASE {snapshot.Ase.Id} (O9={snapshot.ValidacionTotal}, P9={snapshot.ValidacionTotalOkP}) diverge del ASE {baseTotal.Ase.Id}; el lector-oráculo debe leer O9/P9 una sola vez (S-2).");
                    break;
                }
            }
        }

        return errores;
    }

    /// <summary>
    /// Plan 42 (T1, R-CLS-1/R-CLS-2/R-CLS-3, D-B/D-F): clasifica cada sub-bloque FALSE de
    /// <c>VALIDACION_TOTAL</c> en declarado (entrada vigente de
    /// <see cref="DivergenciasDeclaradasCatalogo"/>) vs NO declarado, y separa los errores de la
    /// validación cruzada que no son sub-bloques declarados. Reutiliza <see cref="Validar(ResultadoRemuneracion, IReadOnlyList{WorkbookLeafInputs}, IReadOnlyList{ValidacionCruzadaSnapshot})"/>
    /// (criterio estricto E1-1..E1-5 INTACTO para lo no declarado) y solo RETIRA del listado los
    /// errores de los sub-bloques que el catálogo reconoce como declarados. Una cadena no resoluble
    /// (E1-1) NUNCA se declara: queda como no declarada.
    /// </summary>
    public ClasificacionDivergenciasDeclaradas ClasificarDivergenciasDeclaradas(
        ResultadoRemuneracion resultado,
        IReadOnlyList<WorkbookLeafInputs> leafs,
        IReadOnlyList<ValidacionCruzadaSnapshot> snapshots)
    {
        ArgumentNullException.ThrowIfNull(resultado);
        ArgumentNullException.ThrowIfNull(leafs);
        ArgumentNullException.ThrowIfNull(snapshots);

        // Criterio estricto vigente (E1-1..E1-5): todos los gates, mismo nombre y severidad.
        var errores = Validar(resultado, leafs, snapshots);

        var veredictos = new List<VeredictoDivergenciaValidacion>();
        foreach (var snapshot in snapshots.OrderBy(s => s.Ase.Id))
        {
            foreach (var (celda, valor) in snapshot.SubBloquesValidacionTotal.OrderBy(kv => kv.Key))
            {
                var (falso, clasificable, observado, esperado) = EvaluarSubBloque(snapshot, celda, valor);
                if (!falso)
                {
                    continue;
                }

                // Solo una divergencia EVALUABLE desde lo escrito puede corresponder a una entrada
                // del catálogo; "no evaluable" (E1-1) y los snapshots sin evidencia = no declaradas.
                var entrada = clasificable
                    ? DivergenciasDeclaradasCatalogo.Buscar(_catalogo, celda, observado, esperado)
                    : null;

                veredictos.Add(new VeredictoDivergenciaValidacion
                {
                    AseId = snapshot.Ase.Id,
                    Celda = celda,
                    Declarada = entrada is not null,
                    EntradaId = entrada?.Id,
                    MontoObservado = observado,
                    MontoEsperado = esperado
                });
            }
        }

        var declaradas = veredictos
            .Where(v => v.Declarada)
            .Select(v => (v.AseId, v.Celda))
            .ToHashSet();
        var erroresNoDeclarados = declaradas.Count == 0
            ? errores
            : errores.Where(e => !EsErrorDeSubBloqueDeclarado(e, declaradas)).ToList();

        return new ClasificacionDivergenciasDeclaradas
        {
            Veredictos = veredictos,
            ErroresNoDeclarados = erroresNoDeclarados
        };
    }

    /// <summary>
    /// Espeja la evaluación de <see cref="ValidarGatesValidacionesCruzadasPorAse"/> para UN
    /// sub-bloque: con evidencia, el veredicto sale de <c>Computable</c> + <c>Igualdad</c>; sin
    /// evidencia (snapshots construidos a mano), del booleano entregado. <c>Clasificable</c> solo
    /// cuando la cadena es EVALUABLE desde lo escrito (una cadena no resoluble nunca se declara).
    /// </summary>
    private static (bool Falso, bool Clasificable, decimal Observado, decimal Esperado) EvaluarSubBloque(
        ValidacionCruzadaSnapshot snapshot,
        string celda,
        bool valor)
    {
        if (snapshot.SubBloquesValidacionTotalEvidencia.TryGetValue(celda, out var evidencia))
        {
            if (!evidencia.Computable)
            {
                return (true, false, evidencia.OperandoIzquierdo, evidencia.OperandoDerecho);
            }

            var cierra = FormasValidacionTotal.Igualdad(evidencia.OperandoIzquierdo, evidencia.OperandoDerecho);
            return (!cierra, !cierra, evidencia.OperandoIzquierdo, evidencia.OperandoDerecho);
        }

        return (!valor, false, 0m, 0m);
    }

    /// <summary>
    /// Verdad si el error corresponde a un sub-bloque declarado (mismo marcador
    /// <c>ASE n · VALIDACION_TOTAL · celda</c> que emite el gate W-1).
    /// </summary>
    private static bool EsErrorDeSubBloqueDeclarado(string error, HashSet<(int AseId, string Celda)> declaradas)
    {
        foreach (var (aseId, celda) in declaradas)
        {
            if (error.Contains($"ASE {aseId} · VALIDACION_TOTAL · {celda}", StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// HU-20 (G3-D1): valida el DetRetri CALCULADO por ASE (<see cref="DetRetriQ2Inputs.Detalle"/> =
    /// ROUND(D104:D108,0)) contra el oráculo R10 del período, tolerancia ±0.5. El R10 es oráculo de
    /// VALIDACIÓN (nunca se escribe al template). Matcheo estricto por <see cref="Ase.Id"/>. El
    /// proceso puebla <see cref="WorkbookLeafInputs.DetRetriQ2"/> en AMBAS quincenas, por lo que el
    /// gate corre en Q1 y Q2; los leafs sin detalle (construcciones manuales de test) se omiten.
    /// El procesador agrega período + archivo al fail-fast.
    /// </summary>
    public List<string> ValidarDetRetriContraR10(IReadOnlyList<WorkbookLeafInputs> leafs, DetRetriInputs r10)
    {
        ArgumentNullException.ThrowIfNull(leafs);
        ArgumentNullException.ThrowIfNull(r10);

        var errores = new List<string>();
        var conDetalle = leafs
            .Where(l => l.DetRetriQ2 is not null)
            .OrderBy(l => l.Ase.Id)
            .ToList();
        if (conDetalle.Count == 0)
        {
            return errores; // Leafs sin DetRetri calculado (construcciones manuales de test).
        }

        foreach (var leaf in conDetalle)
        {
            var calculado = leaf.DetRetriQ2!.Detalle;
            if (!r10.DetRetriPorAse.TryGetValue(leaf.Ase.Id, out var oraculo))
            {
                AgregarError(errores, $"El R10 ({r10.CodigoRemuneracion}) no trae el DetRetri del ASE {leaf.Ase.Id} para validar el calculado ({calculado}).");
                continue;
            }

            var diferencia = Math.Abs(calculado - oraculo);
            if (diferencia > Tolerancia)
            {
                AgregarError(errores, $"ASE {leaf.Ase.Id}: DetRetri calculado ({calculado}) no coincide con el R10 ({oraculo}). Diferencia={diferencia} > ±{Tolerancia}.");
            }
        }

        var sumaCalculada = conDetalle.Sum(l => l.DetRetriQ2!.Detalle);
        var diferenciaTotal = Math.Abs(sumaCalculada - r10.Total);
        if (diferenciaTotal > Tolerancia)
        {
            AgregarError(errores, $"DetRetri total calculado (Σ {sumaCalculada}) no coincide con el R10 total ({r10.Total}). Diferencia={diferenciaTotal} > ±{Tolerancia}.");
        }

        return errores;
    }

    /// <summary>
    /// HU-13 (2.7, §2.5): gates aditivos por ASE contra el snapshot-oráculo (D1/D3). El validador
    /// NO abre .xlsx: solo compara números. Matcheo estricto por <see cref="Ase.Id"/> (Single,
    /// nunca fallback). Semántica congelada por T0 en ambos canónicos (Plan 13 §4 Fase 0).
    ///
    /// HU-14 (W-2, veredicto documentado — sin mover): este gate de validaciones cruzadas vive
    /// AQUÍ (ValidadorBasico) y NO en <c>WorkbookLeafCoherence</c> por diseño: todos los gates
    /// cruzados HU-08..HU-12 (<c>ValidarSigmaEmpresasPorAse</c>, <c>ValidarReporteBanco</c>,
    /// <c>ValidarBalanceSc</c>) viven en el validador con AGREGACIÓN de errores por lista;
    /// moverlos a <c>WorkbookLeafCoherence</c> (throw al primer fallo) cambiaría la semántica
    /// de agregación. Ver Plan 13 §2.8/§4-3.1 enmendado por Plan 14 §9.7.
    ///
    /// HU-14 (W-1): además de O/P por empresa y DetValiRetri, se gatean los sub-bloques
    /// booleanos de <c>VALIDACION_TOTAL</c> (C15/D25/O25/D34/F34) == true exacto por ASE.
    /// HU-14 (S-2): el gate numérico/booleano de VALIDACION_TOTAL (O9/P9) NO se evalúa aquí
    /// por ASE (sería ×5); se evalúa UNA vez en <c>Validar(resultado, leafs, snapshots)</c>.
    /// </summary>
    private static void ValidarGatesValidacionesCruzadasPorAse(
        List<string> errores,
        ValidacionCruzadaSnapshot snapshot,
        ResultadoRemuneracion resultado,
        IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(leafs);

        var consolidado = resultado.Consolidados.SingleOrDefault(c => c.Ase.Id == snapshot.Ase.Id);
        if (consolidado is null)
        {
            AgregarError(errores, $"No se encontró el consolidado del ASE {snapshot.Ase.Id} para el gate de validaciones cruzadas.");
            return;
        }

        // (i) Por empresa (VALIDACION_*): fila 2+aseId de cada hoja. O = H−N == 0 ±0.5 y
        //     P (INT(O)=0) == true exacto (falla aunque O esté en tolerancia). Error ASE+empresa.
        foreach (var empresa in snapshot.PorEmpresa)
        {
            if (empresa.AseId != snapshot.Ase.Id)
            {
                AgregarError(errores, $"ASE {snapshot.Ase.Id}: el snapshot por empresa trae la fila del ASE {empresa.AseId} (empresa {empresa.Empresa}); matcheo estricto violado.");
                continue;
            }

            var diferenciaO = Math.Abs(empresa.DiferenciaO);
            if (diferenciaO > Tolerancia)
            {
                AgregarError(errores, $"ASE {snapshot.Ase.Id} · {empresa.Empresa}: la validación Recaudo vs REMUNERACION (O) no cierra: O={empresa.DiferenciaO} (debe ser 0 ±{Tolerancia}).");
            }

            if (!empresa.VerificacionP)
            {
                AgregarError(errores, $"ASE {snapshot.Ase.Id} · {empresa.Empresa}: la verificación P (INT(O)=0) es falsa en el workbook aunque O esté en tolerancia.");
            }
        }

        // (ii) DetValiRetri: D16..D20 (diferencias por ASE) == 0 ±0.5; D24..D28 == true exacto;
        //     D29 (total) == true exacto. D21 (fila Total) y D9:D14/J9:J14 ≠ ROUND quedan
        //     EXCLUIDOS (D6: divergencia documentada en ambos canónicos).
        if (snapshot.DetValiRetri is not null)
        {
            foreach (var celda in snapshot.DetValiRetri.DiferenciasAse)
            {
                var diferencia = Math.Abs(celda.Valor);
                if (diferencia > Tolerancia)
                {
                    AgregarError(errores, $"ASE {snapshot.Ase.Id}: DetValiRetri {celda.Celda} no cierra: {celda.Valor} (debe ser 0 ±{Tolerancia}).");
                }
            }

            foreach (var celda in snapshot.DetValiRetri.VerificacionesAse)
            {
                if (!celda.Verificacion)
                {
                    AgregarError(errores, $"ASE {snapshot.Ase.Id}: DetValiRetri {celda.Celda} es falso (composición no verificada).");
                }
            }

            if (!snapshot.DetValiRetri.VerificacionTotalD29)
            {
                AgregarError(errores, $"ASE {snapshot.Ase.Id}: DetValiRetri D29 (verificación TOTAL) es falso.");
            }
        }

        // (iii) Plan 40 (T1, R-E-1/R-E-2/R-E-3 — D-A/D-D; enmienda E-1): los 5 sub-bloques booleanos
        //     de VALIDACION_TOTAL (C15/D25/O25/D34/F34) se EVALÚAN desde lo escrito (nunca desde el
        //     caché <v>, que era la falsa verificación heredada). Para cada sub-bloque con evidencia
        //     se aserta el TEXTO <f> esperado (tabla FormasValidacionTotal, congelada por el T0)
        //     ANTES de evaluar (tercera forma = ERR-PLANTILLA) y se gatea Igualdad(izq, der) ±0.5.
        //     E1-1: se gatean los 5 SIN excepción; una cadena no resoluble NO es verde (se reporta).
        //     Snapshots sin evidencia (construidos a mano en tests) conservan el gate del booleano
        //     entregado (compatibilidad). El gate TOTAL (O9/P9) es ÚNICO (S-2) y vive en
        //     Validar(...snapshots).
        foreach (var (celda, valor) in snapshot.SubBloquesValidacionTotal.OrderBy(kv => kv.Key))
        {
            if (snapshot.SubBloquesValidacionTotalEvidencia.TryGetValue(celda, out var evidencia))
            {
                var esperado = FormasValidacionTotal.TextoEsperado(celda);
                if (esperado is null || !string.Equals(evidencia.TextoFormula, esperado, StringComparison.Ordinal))
                {
                    throw new CalculoInvalidoException(
                        CodigoError.Plantilla,
                        $"{FormasValidacionTotal.Hoja}!{celda}: la fórmula <f> no coincide con la tabla del gate. Esperado: '{esperado}'. Observado: '{evidencia.TextoFormula}'. (R-E-2: tercera forma = ERR-PLANTILLA.)");
                }

                if (!evidencia.Computable)
                {
                    // E1-1: "no evaluable" no es verde. La cadena del sub-bloque no se pudo resolver
                    // desde lo escrito → hallazgo que el gate nombra (nunca un pase silencioso).
                    AgregarError(errores, $"ASE {snapshot.Ase.Id} · VALIDACION_TOTAL · {celda}: la cadena del sub-bloque no es resoluble desde lo escrito (no evaluable). (W-1/E1-1: nunca verde.)");
                    continue;
                }

                if (!FormasValidacionTotal.Igualdad(evidencia.OperandoIzquierdo, evidencia.OperandoDerecho))
                {
                    AgregarError(errores, $"ASE {snapshot.Ase.Id} · VALIDACION_TOTAL · {celda}: el sub-bloque evaluado desde lo escrito es falso ({evidencia.OperandoIzquierdo} != {evidencia.OperandoDerecho}). (W-1: debe ser TRUE exacto.)");
                }

                continue;
            }

            if (!valor)
            {
                AgregarError(errores, $"ASE {snapshot.Ase.Id} · VALIDACION_TOTAL · {celda}: sub-bloque booleano es falso en el workbook (debe ser TRUE exacto; W-1).");
            }
        }

        // (iv) Plan 49 (A1-núcleo + A1-interior): gate ABSOLUTO anti-cancelacion. El gate 2.7 (iii)
        //     compara dos derivaciones del MISMO workbook (D33 vs D109): una pérdida común entra en
        //     ambos operandos y cancela (T0 §6.3). Aquí un operando es el DOMINIO por firma
        //     (WorkbookLeafInputsR1.ExtemporaneoEsperadoPorAse) y el otro es lo escrito resuelto, así
        //     que la igualdad no puede cancelarse por construcción.
        ValidarExtempAbsolutoPorAse(errores, snapshot, resultado, leafs);
    }

    /// <summary>
    /// Plan 49 (A1, §2.2/§2.3 — D-A/D-D): gate absoluto del EXTEMPORANEO por ASE.
    ///
    /// <para>
    /// <b>Núcleo (obligatorio, ambos lados del CONSOLIDADO):</b> el EXTEMP resuelto desde lo escrito
    /// en <c>CONSOLIDADO_TOTAL RECAUDO!D{46+aseId}</c> (D47..D51) debe coincidir ±0.5 con el agregado
    /// EXTEMP del dominio por firma
    /// (<see cref="WorkbookLeafInputsR1.ExtemporaneoEsperadoPorAse"/>). Como el operando derecho NO
    /// es otra celda del workbook, una pérdida común a ambas derivaciones (la del EXTEMP ASE5 en la
    /// plantilla v2) ya no es invisible. Firma del falso verde de julio-2/agosto-2: <c>D34</c> sigue
    /// TRUE (cancela) y este gate falla con el Δ exacto (<c>441.849</c> / <c>1.249.810</c>).
    /// </para>
    ///
    /// <para>
    /// <b>Desviación documentada (H1, 2026-10-19):</b> el plan §2.2 nombra como operando derecho
    /// <see cref="ConsolidadoAse.Extemp"/>, pero ese campo lo alimenta el lector legacy
    /// <c>ExcelDataReaderRecaudoReader.LeerR1</c> con la PRIMERA fila <c>Mes</c> de la fuente (no el
    /// agregado EXTEMP): en julio-2 da <c>13.084.233.937,44</c> (ASE5) y en agosto-2 <c>-4.485,01</c>,
    /// mientras el manual y el trabajo por firma valen <c>441.849</c> / <c>1.249.810</c>. Comparar
    /// contra él dispararía 5/5 ASE con Δ falsos (rojo fuera de la lista cerrada). El operando que
    /// produce EXACTAMENTE los Δ del plan es el agregado del dominio por firma
    /// (<see cref="WorkbookLeafInputsR1.ExtemporaneoEsperadoPorAse"/>), que coincide al centavo con
    /// <c>CONSOLIDADO!D47:D51</c> del manual en julio-2 y agosto-2. La costura `ConsolidadoAse.Extemp`
    /// queda intacta (fuera del alcance: no se corrige el EXTEMP).
    /// </para>
    ///
    /// <para>
    /// <b>Interior (acotado, §2.3; Plan 57 §2-D2):</b> la evidencia de las anclas la DERIVA el reader de
    /// la propia plantilla (posición por la referencia publicada, naturaleza leída). Aquí se compara la
    /// celda ancla con la descomposición que ella misma declara, de modo que el error nombra la hoja y
    /// la celda ancla donde se perdió el importe. Es diagnóstico: nunca altera el veredicto del núcleo.
    /// Un ancla literal no declara descomposición (<see cref="AnclaExtempR1Evidencia.ValorSecciones"/>
    /// = <c>null</c>) y no aporta comparación. Sin evidencia del reader
    /// (<see cref="ValidacionCruzadaSnapshot.ExtempResuelto"/> = <c>null</c> y lista de anclas vacía),
    /// no se agrega nada: snapshots construidos a mano conservan la compatibilidad.
    /// </para>
    ///
    /// Sin rama por quincena (§2.6): el gate corre en Q1 y Q2; compara el dominio del período.
    /// </summary>
    private static void ValidarExtempAbsolutoPorAse(
        List<string> errores,
        ValidacionCruzadaSnapshot snapshot,
        ResultadoRemuneracion resultado,
        IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        var consolidado = resultado.Consolidados.SingleOrDefault(c => c.Ase.Id == snapshot.Ase.Id);
        var leaf = leafs.SingleOrDefault(l => l.Ase.Id == snapshot.Ase.Id);
        if (consolidado is null || leaf is null)
        {
            return; // el gate de validaciones cruzadas ya reportó el consolidado/leaf ausente.
        }

        if (snapshot.ExtempResuelto is { } resuelto)
        {
            var esperado = leaf.R1.ExtemporaneoEsperadoPorAse;
            var diferencia = Math.Abs(resuelto - esperado);
            if (diferencia > Tolerancia)
            {
                var celda = PublicacionExtempConsolidado.CeldaConsolidadoExtemp(snapshot.Ase.Id);
                AgregarError(errores,
                    $"ASE {snapshot.Ase.Id} · {PublicacionExtempConsolidado.HojaConsolidado} · {celda}: EXTEMP resuelto desde lo escrito ({resuelto}) != Extemp del dominio ({esperado}). Diferencia={diferencia} > ±{Tolerancia}. (Plan 49 A1: gate absoluto anti-cancelacion.)");
            }
        }

        foreach (var ancla in snapshot.EvidenciaAnclasExtemp)
        {
            if (ancla.ValorSecciones is not { } valorSecciones)
            {
                continue; // ancla literal: la plantilla no declara descomposición que contrastar.
            }

            var diferencia = Math.Abs(ancla.ValorAncla - valorSecciones);
            if (diferencia > Tolerancia)
            {
                var expresionVisible = ancla.ExpresionSecciones.Replace('-', '−');
                AgregarError(errores,
                    $"ASE {snapshot.Ase.Id} · {ancla.Hoja} · {ancla.Celda}: ancla EXTEMP resuelta ({ancla.ValorAncla}) != suma de secciones {expresionVisible} ({valorSecciones}). Diferencia={diferencia} > ±{Tolerancia}. (Plan 49 A1-interior.)");
            }
        }
    }

    /// <summary>
    /// HU-10 (2.4, §2.5): gates D5 del balance SC.
    /// (i) Por ASE: Total BSC (Contribucion+Subsidio) == Total General fuente (col G) ±0.5;
    ///     fail-fast nombra el ASE.
    /// (ii) Σ fila 11 en dominio: Σ Total BSC a través de los 5 ASE == Σ TotalFuente ±0.5
    ///     (las celdas D11/E11/F11 del template son fórmulas protegidas; la aritmética de
    ///     dominio verifica que la suma cierra).
    /// (iii) H≈F solo cuando el desenlace T0 es D2(a) (<see cref="BalanceScAseInputs.Sistema"/>
    ///     no nulo): redondeo a entero (AwayFromZero). En Q1 el veredicto es D2(b) (H es
    ///     fórmula → Sistema null) y H≈F se verifica contra el caché golden en Capa A (G7).
    /// Todos los leafs con <c>BalanceSc == null</c> = comportamiento HU-09 puro (sin gates 2.4).
    /// </summary>
    private static void ValidarBalanceSc(List<string> errores, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        var conBalance = leafs.Where(l => l.BalanceSc is not null).ToList();
        if (conBalance.Count == 0)
        {
            return; // HU-09 puro: sin gates 2.4
        }

        foreach (var leaf in conBalance)
        {
            var balance = leaf.BalanceSc!;
            var bloque = balance.Ases.SingleOrDefault(b => b.Ase.Id == leaf.Ase.Id);
            if (bloque is null)
            {
                AgregarError(errores, $"El leaf del ASE {leaf.Ase.Id} no trae su fila de balance SC para validar.");
                continue;
            }

            // Gate (i): BCE = fuente ±0.5 (TotalBsc vs TotalFuente; col G del Total General).
            var diferencia = Math.Abs(bloque.TotalBsc - bloque.TotalFuente);
            if (diferencia > Tolerancia)
            {
                AgregarError(errores, $"ASE {leaf.Ase.Id}: Total BSC ({bloque.TotalBsc}) != Total General fuente ({bloque.TotalFuente}). Diferencia={diferencia} > ±{Tolerancia}.");
            }

            // Gate (iii): H≈F con redondeo a entero SOLO en desenlace D2(a); en D2(b) (Q1)
            // Sistema es null y H≈F se verifica en Capa A contra el caché golden.
            if (bloque.Sistema is not null)
            {
                var redondeado = Math.Round(bloque.TotalBsc, MidpointRounding.AwayFromZero);
                if (Math.Abs(bloque.Sistema.Value - redondeado) > Tolerancia)
                {
                    AgregarError(errores, $"ASE {leaf.Ase.Id}: H (SISTEMA) {bloque.Sistema.Value} no coincide con ROUND(F,0)={redondeado}.");
                }
            }
        }

        // Gate (ii): sumas fila 11 en dominio — Σ Total BSC == Σ TotalFuente ±0.5.
        var sumaBsc = conBalance.Sum(l => l.BalanceSc!.Ases.Sum(b => b.TotalBsc));
        var sumaFuente = conBalance.Sum(l => l.BalanceSc!.Ases.Sum(b => b.TotalFuente));
        if (Math.Abs(sumaBsc - sumaFuente) > Tolerancia)
        {
            AgregarError(errores, $"Σ Total BSC ({sumaBsc}) != Σ Total General fuente ({sumaFuente}). Diferencia={Math.Abs(sumaBsc - sumaFuente)} > ±{Tolerancia}.");
        }
    }

    /// <summary>
    /// HU-09 (2.3, §2.5): gates D5 del reporte por banco.
    /// (i) Por ASE y empresa: bloque (Σ conceptos) == resumen fuente (<see cref="ReporteBancoEmpresaInputs.TotalFuente"/>)
    ///     ±0.5; fail-fast nombra ASE y empresa-columna.
    /// (ii) Consolidado 1–7 == Σ bloques por empresa ±0.5: si el dominio trae
    ///     <see cref="ReporteBancoInputs.Consolidado"/> se compara y falla si difiere; en el
    ///     desenlace T0 D2(b) (filas 1–7 son fórmulas) el campo es null y la verificación Σ
    ///     contra el caché golden vive en la Capa A (G3: "solo se verifica Σ").
    /// (iii) C59 (<see cref="ReporteBancoInputs.Quincena"/>) == <see cref="Periodo.NumeroQuincena"/>.
    /// Todos los leafs con <c>ReporteBanco == null</c> = comportamiento HU-08 puro (sin gates 2.3).
    /// </summary>
    private static void ValidarReporteBanco(List<string> errores, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        var conBanco = leafs.Where(l => l.ReporteBanco is not null).ToList();
        if (conBanco.Count == 0)
        {
            return; // HU-08 puro: sin gates 2.3
        }

        foreach (var leaf in conBanco)
        {
            var banco = leaf.ReporteBanco!;

            // Gate (iii): C59 == quincena (dominio, nunca fuente).
            if (banco.Quincena != leaf.Periodo.NumeroQuincena)
            {
                AgregarError(errores, $"ASE {leaf.Ase.Id}: C59 del reporte por banco ({banco.Quincena}) no coincide con la quincena del período ({leaf.Periodo.NumeroQuincena}).");
            }

            // Gate (i): bloque = resumen fuente por empresa ±0.5 (fail-fast nombra ASE + empresa).
            foreach (var bloque in banco.Ases)
            {
                foreach (var empresa in bloque.Empresas)
                {
                    var diferencia = Math.Abs(empresa.Total - empresa.TotalFuente);
                    if (diferencia > Tolerancia)
                    {
                        AgregarError(errores, $"ASE {leaf.Ase.Id} · {empresa.Empresa}: bloque banco ({empresa.Total}) != resumen fuente ({empresa.TotalFuente}). Diferencia={diferencia} > ±{Tolerancia}.");
                    }
                }
            }
        }

        // Gate (ii): consolidado 1–7 == Σ bloques por empresa ±0.5. En D2(b) el Consolidado
        // del dominio es null (filas 1–7 son fórmulas): la Σ se verifica contra el caché
        // golden en Capa A y se expone en el log/UI, nunca como error aquí.
        var consolidado = conBanco.Select(l => l.ReporteBanco!.Consolidado).FirstOrDefault(c => c is not null);
        if (consolidado is null)
        {
            return;
        }

        var sumaPorEmpresa = conBanco
            .SelectMany(l => l.ReporteBanco!.Ases)
            .SelectMany(b => b.Empresas)
            .GroupBy(e => e.Empresa, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.Sum(e => e.Total), StringComparer.OrdinalIgnoreCase);

        foreach (var (empresa, esperado) in consolidado)
        {
            if (!sumaPorEmpresa.TryGetValue(empresa, out var suma))
            {
                AgregarError(errores, $"Consolidado banco declara la empresa '{empresa}' pero no hay bloques con esa empresa en los leafs.");
                continue;
            }

            if (Math.Abs(suma - esperado) > Tolerancia)
            {
                AgregarError(errores, $"Consolidado 1–7 de '{empresa}' ({esperado}) no coincide con Σ bloques por empresa ({suma}).");
            }
        }
    }

    private static void ValidarGatesPorAse(
        List<string> errores,
        WorkbookLeafInputs leaf,
        ConsolidadoAse consolidado)
    {
        // Gate R1: primera fila Mes/Total col F (F25-equivalente por bloque) vs Extemp.
        // Prohibido comparar TotOpt HU-02 contra visibles R1.
        if (Math.Abs(leaf.R1.F25 - consolidado.Extemp) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: la hoja R1 F25-equivalente ({leaf.R1.F25}) no coincide con Extemp del consolidado ({consolidado.Extemp}).");
        }

        // Gate R2: visible del bloque vs R2TotalOportuno.
        if (Math.Abs(leaf.R2.TotalOportunoEsperado - consolidado.R2TotalOportuno) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: la hoja R2 TotalOportunoEsperado ({leaf.R2.TotalOportunoEsperado}) no coincide con R2TotalOportuno del consolidado ({consolidado.R2TotalOportuno}).");
        }

        // Gate R4: visible del bloque vs ReversionR4.
        if (Math.Abs(leaf.R4.TotalReversionEsperada - consolidado.ReversionR4) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: la hoja R4 TotalReversionEsperada ({leaf.R4.TotalReversionEsperada}) no coincide con ReversionR4 del consolidado ({consolidado.ReversionR4}).");
        }
    }

    /// <summary>
    /// HU-08 (§2.5): Σ visibles-empresa == visible de bloque por hoja (R1/R2/R4) ±0.5.
    /// Ceros legítimos: EAAB-CL todo 0 en Q1 es válido (V8). El error nombra ASE + empresa.
    /// Lista vacía = HU-07 puro.
    /// </summary>
    private static void ValidarSigmaEmpresasPorAse(List<string> errores, WorkbookLeafInputs leaf)
    {
        if (leaf.Conciliacion.Count == 0)
        {
            return;
        }

        // Plan 50 (A2, §2.4 — RESIDUO Q2 DECLARADO `R-SIGMA-Q2`):
        // este gate (Σ visibles-empresa == visible de bloque) se definió sobre los mapas Q1 de HU-08 y
        // NO es una descomposición válida del visible de bloque en Q2. Evidencia medida en bytes
        // (`Docs/T0-Plan50-FuentesRecaudoOpcionales.md` §4):
        //   - R4 solo declara las empresas con celda destino (p. ej. ASE1: ENEL no tiene celda), por lo
        //     que la Σ omite dinero real del reporte (julio-2 ASE1 Δ = 15.533.961,70);
        //   - R2 incluye etiquetas reales sin destino en el mapa (EAB, ENERBIT) → Δ = 73.523,10 /
        //     6.242.402,48 / 6.055.022,95;
        //   - el visible R1-Q2 se deriva por firma (espejo: roles Mes/Aplic), no de estos operandos.
        // Plan 59 (§15, 2026-10-27): el recorte de la conciliación por empresa en Q2 se RETIRÓ (la v4
        // ya no tiene fórmula en R2!E285), así que la lista llega POBLADA en Q2 y esta rama es la que
        // impide evaluar el gate allí. El gate NO debe reactivarse en Q2 sin el contrato nuevo (T0
        // propio, residuo `R-SIGMA-Q2`).
        if (leaf.Periodo.NumeroQuincena != 1)
        {
            return;
        }

        var sumaR1 = leaf.Conciliacion.Sum(c => c.VisibleR1);
        if (Math.Abs(sumaR1 - leaf.R1.TotalOportunoEsperadoPorAse) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: Σ visibles de empresas en R1 ({sumaR1}) no coincide con el visible de bloque ({leaf.R1.TotalOportunoEsperadoPorAse}). Detalle: {DetalleEmpresas(leaf, c => c.VisibleR1)}");
        }

        var sumaR2 = leaf.Conciliacion.Sum(c => c.VisibleR2);
        if (Math.Abs(sumaR2 - leaf.R2.TotalOportunoEsperado) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: Σ visibles de empresas en R2 ({sumaR2}) no coincide con el visible de bloque ({leaf.R2.TotalOportunoEsperado}). Detalle: {DetalleEmpresas(leaf, c => c.VisibleR2)}");
        }

        var sumaR4 = leaf.Conciliacion.Sum(c => c.VisibleR4);
        if (Math.Abs(sumaR4 - leaf.R4.TotalReversionEsperada) > Tolerancia)
        {
            AgregarError(errores, $"ASE {leaf.Ase.Id}: Σ visibles de empresas en R4 ({sumaR4}) no coincide con el visible de bloque ({leaf.R4.TotalReversionEsperada}). Detalle: {DetalleEmpresas(leaf, c => c.VisibleR4)}");
        }
    }

    private static string DetalleEmpresas(WorkbookLeafInputs leaf, Func<ConciliacionEmpresaInputs, decimal> selector) =>
        string.Join(", ", leaf.Conciliacion.Select(c => $"{c.Empresa.Nombre}={selector(c)}"));

    /// <summary>
    /// Plan 50 (A3, §2.5 — control mínimo alternativo): el T0 de H0 demostró con bytes que la vía
    /// aritmética (Σ fuentes presentes vs total del reporte) NO es viable en R1 (el total del cuadro
    /// Componente no es la suma de las filas-etiqueta) ni en R2 (el total incluye la etiqueta real
    /// <c>EAB</c>, sin destino en el mapa). El control mínimo es el de ETIQUETAS:
    ///
    /// - <see cref="ConciliacionEmpresaInputs.EtiquetasNoReconocidas"/> = etiquetas observadas en las
    ///   filas de cierre del reporte que no resuelven a ninguna empresa del mapa (evidencia del reader,
    ///   read-only, sin workbooks extra).
    /// - Si además hay una <see cref="ConciliacionEmpresaInputs.FuentesAusentes"/> en el MISMO reporte
    ///   del MISMO ASE, hay dinero en el reporte que ninguna fuente reconocida explica: es la firma de
    ///   "fuente presente no reconocida" (cambio de etiqueta) → FAIL en duro, nunca 0 silencioso.
    ///
    /// Sin ausencias en el mismo reporte la etiqueta desconocida es sólo evidencia (p. ej. <c>EAB</c>,
    /// fuente real sin destino en el mapa: deuda declarada del Plan 49, no se convierte en rojo).
    /// </summary>
    private static void ValidarFuentesRecaudoNoReconocidas(List<string> errores, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        foreach (var leaf in leafs.OrderBy(l => l.Ase.Id))
        {
            if (leaf.Conciliacion.Count == 0)
            {
                continue;
            }

            var noReconocidas = leaf.Conciliacion
                .SelectMany(conc => conc.EtiquetasNoReconocidas)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (noReconocidas.Count == 0)
            {
                continue;
            }

            var ausentes = leaf.Conciliacion
                .SelectMany(conc => conc.FuentesAusentes)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (ausentes.Count == 0)
            {
                continue;
            }

            foreach (var reporte in noReconocidas.Select(n => ReporteDe(n)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var ausentesDelReporte = ausentes.Where(a => ReporteDe(a).Equals(reporte, StringComparison.OrdinalIgnoreCase)).ToList();
                if (ausentesDelReporte.Count == 0)
                {
                    continue; // etiqueta desconocida sin ausencia en SU reporte = sólo evidencia.
                }

                var etiquetas = noReconocidas.Where(n => ReporteDe(n).Equals(reporte, StringComparison.OrdinalIgnoreCase)).Select(EtiquetaDe);
                var fuentes = ausentesDelReporte.Select(EtiquetaDe);
                AgregarError(
                    errores,
                    $"ASE {leaf.Ase.Id} · {reporte} · control de fuentes: etiqueta(s) no reconocida(s) [{string.Join(", ", etiquetas)}] con fuente(s) ausente(s) [{string.Join(", ", fuentes)}]. (Plan 50 A3: posible fuente presente no reconocida.)");
            }
        }
    }

    /// <summary>
    /// Plan 51 (A3-banco, §2.3): el reporte por banco tiene su propio control de etiquetas, con
    /// grano ASE y el canal durable <c>BANCO:</c> (no se toca el juicio R1/R2/R4 del Plan 50).
    ///
    /// - <see cref="ReporteBancoAseInputs.FuentesAusentesBanco"/> = fuente del catálogo que la
    ///   plantilla declara para el ASE y no vino en el Resumen (ausencia legítima: se omite, sus
    ///   celdas quedan en el 0 de la plantilla).
    /// - <see cref="ReporteBancoAseInputs.EtiquetasBancoNoReconocidas"/> = celda de header que no
    ///   resuelve al vocabulario de fuentes y no es <c>Total</c> (evidencia del reader).
    ///
    /// Si ambas coexisten en el MISMO ASE hay dinero en el reporte que ninguna fuente reconocida
    /// explica: es la firma de "fuente presente no reconocida" (cambio de etiqueta) → FAIL en
    /// duro, nunca 0 silencioso. Sin ausencia en ese ASE la etiqueta desconocida es sólo evidencia.
    /// Sin umbral numérico: el juicio es de etiquetas (el gate Σ por empresa presente sigue
    /// intacto en <see cref="ValidarReporteBanco"/> y la fuente ausente no entra a él).
    /// </summary>
    private static void ValidarEtiquetasBancoNoReconocidas(List<string> errores, IReadOnlyList<WorkbookLeafInputs> leafs)
    {
        foreach (var leaf in leafs.OrderBy(l => l.Ase.Id))
        {
            var ases = leaf.ReporteBanco?.Ases;
            if (ases is null || ases.Count == 0)
            {
                continue;
            }

            var noReconocidas = ases
                .SelectMany(ase => ase.EtiquetasBancoNoReconocidas)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (noReconocidas.Count == 0)
            {
                continue; // sin hallazgos: nada que juzgar.
            }

            var ausentes = ases
                .SelectMany(ase => ase.FuentesAusentesBanco)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (ausentes.Count == 0)
            {
                continue; // etiqueta desconocida sin ausencia en el mismo ASE = sólo evidencia.
            }

            AgregarError(
                errores,
                $"ASE {leaf.Ase.Id} · BANCO · control de fuentes: etiqueta(s) no reconocida(s) [{string.Join(", ", noReconocidas.Select(EtiquetaDe))}] con fuente(s) ausente(s) [{string.Join(", ", ausentes.Select(EtiquetaDe))}]. (Plan 51 A3: posible fuente presente no reconocida.)");
        }
    }

    private static string ReporteDe(string entrada)
    {
        var separador = entrada.IndexOf(':');
        return separador < 0 ? entrada : entrada[..separador];
    }

    private static string EtiquetaDe(string entrada)
    {
        var separador = entrada.IndexOf(':');
        return separador < 0 ? entrada : entrada[(separador + 1)..];
    }
}
