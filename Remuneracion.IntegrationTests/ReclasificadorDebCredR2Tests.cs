using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Models;
using Remuneracion.Core.Rules;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// Plan 60 (H1.2 — §6 U1..U13): unitarios SINTÉTICOS de la regla pura
/// <see cref="ReclasificadorDebCredR2"/>. Sin I/O, sin plantilla, sin período: el modelo se arma en
/// memoria con la firma A|B|C|D y los dos puertos <c>Deb/Cred</c> / <c>Deb/Cred-otros</c>.
///
/// <para>La forma de los modelos reproduce la gramática de cierres de <c>Rem. Anticipos R2</c>
/// (canal <c>C/D=Total</c>, empresa <c>B/C=Total</c>, sección <c>A/B=Total</c>, gran total
/// <c>A=Total</c>) con la aritmética medida en octubre-1: <c>OCCIDENTE = H+O+T</c>,
/// <c>Componente = Σ cierres de canal</c> y <c>Total = Σ de los dos cuadros</c>. Los rótulos son
/// DATOS del fixture; la regla no los conoce.</para>
/// </summary>
public sealed class ReclasificadorDebCredR2Tests
{
    private const decimal Tol = 0.5m;

    // Índices del fixture ModeloOctubre (posición en la secuencia del detalle).
    private const int PrimeraDeCanalO = 20;        // Vlr Servicio
    private const int InteresesDeCanalO = 21;      // Vlr Intereses
    private const int CierreDeCanalO = 22;         // ||O|Total
    private const int Occidente = 25;              // |OCCIDENTE|Total|
    private const int Componente = 26;             // Componente|Total||
    private const int CierreDeCanalOSubsCont = 45; // segunda instancia de O (cuadro Subs/Cont)
    private const int GranTotal = 50;              // Total|||

    // ── U1 ──────────────────────────────────────────────────────────────────────────────────────

    /// <summary>U1: un cierre de canal sobre el umbral mueve el grupo COMPLETO (conceptos + cierre).</summary>
    [Fact]
    public void DisparoSobreUmbral_MueveGrupoCompleto()
    {
        var modelo = ModeloOctubre();

        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);
        var filas = resultado.Clasificado.Filas;

        // Las TRES filas del grupo del canal O quedan con Q'=0 y R'=Q original.
        Assert.Equal(0m, Q(filas, PrimeraDeCanalO));
        Assert.Equal(54.00m, R(filas, PrimeraDeCanalO));
        Assert.Equal(0m, Q(filas, InteresesDeCanalO));
        Assert.Equal(6_759_294.81m, R(filas, InteresesDeCanalO));
        Assert.Equal(0m, Q(filas, CierreDeCanalO));
        Assert.Equal(6_759_348.81m, R(filas, CierreDeCanalO));

        // Los agregados literales se re-propagan: Q' = Q − Σ y R' = R + Σ.
        Assert.Equal(-6.15m, Q(filas, Occidente));
        Assert.Equal(6_759_348.81m, R(filas, Occidente));
        Assert.Equal(22.96m, Q(filas, Componente));
        Assert.Equal(6_759_348.81m, R(filas, Componente));

        // El gran total reconcilia los DOS cuadros: el 100,00 del cuadro Subs/Cont no se movió,
        // así que permanece en Q (22,96 + 100,00). En la hoja real ese cuadro está en cero y el
        // gran total da 22,96 (lo fija el gate e2e I1 contra el oráculo).
        Assert.Equal(122.96m, Q(filas, GranTotal));
        Assert.Equal(6_759_348.81m, R(filas, GranTotal));

        // El cuadro homónimo (Subs/Cont) NO se toca: su cierre O vale 100 y no dispara.
        Assert.Equal(100m, Q(filas, CierreDeCanalOSubsCont));
        Assert.Equal(0m, R(filas, CierreDeCanalOSubsCont));

        Assert.Equal(3, Conceptos(resultado).Count);
    }

    /// <summary>U1-bis (invariante §3.2): la instancia y las filas de ENTRADA no se mutan.</summary>
    [Fact]
    public void Aplicar_NoMutaElModeloDeEntrada()
    {
        var modelo = ModeloOctubre();
        var antes = modelo.Filas.Select(f => $"{f.A}|{f.B}|{f.C}|{f.D}|{Q([f], 0)}").ToList();

        _ = ReclasificadorDebCredR2.Aplicar(modelo);

        Assert.Equal(antes, modelo.Filas.Select(f => $"{f.A}|{f.B}|{f.C}|{f.D}|{Q([f], 0)}").ToList());
        Assert.Equal(54.00m, Q(modelo.Filas, PrimeraDeCanalO));
        Assert.Equal(6_759_348.81m, Q(modelo.Filas, CierreDeCanalO));
        Assert.Null(modelo.Filas[CierreDeCanalO].Valor(ReclasificadorDebCredR2.ComponenteDebCredOtros));
    }

    // ── U2 ──────────────────────────────────────────────────────────────────────────────────────

    /// <summary>U2 (D1): exactamente 500.000 NO dispara — la regla dice «mayor a», no «mayor o igual».</summary>
    [Fact]
    public void Exactamente500000_NoDispara()
    {
        var modelo = ModeloSimple(cierre: 500_000m);

        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);

        Assert.Empty(resultado.Movimientos);
        Assert.Equal(500_000m, Q(resultado.Clasificado.Filas, 1));
        Assert.Equal(0m, R(resultado.Clasificado.Filas, 1));
        Assert.Equal(0m, R(resultado.Clasificado.Filas, 0));
    }

    // ── U3 ──────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// U3 (D1 ratificada por el Ingeniero): el disparo usa VALOR ABSOLUTO. Un cierre de −6.000.000
    /// dispara igual que uno positivo; la lectura literal (<c>Q &gt; umbral</c>) NO dispararía. El
    /// test fija la decisión en código: si alguien cambia la comparación, enrojece aquí.
    /// </summary>
    [Fact]
    public void NegativoGrande_DisparaSegunD1()
    {
        var modelo = ModeloSimple(cierre: -6_000_000m);

        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);

        var filas = resultado.Clasificado.Filas;
        Assert.Equal(0m, Q(filas, 1));
        Assert.Equal(-6_000_000m, R(filas, 1));
        Assert.Equal(-6_000_000m, R(filas, 4)); // el gran total también se re-propaga
        Assert.NotEmpty(resultado.Movimientos);
    }

    // ── U4 ──────────────────────────────────────────────────────────────────────────────────────

    /// <summary>U4: dos cierres sobre el umbral en el mismo bloque ⇒ ambos grupos se mueven.</summary>
    [Fact]
    public void DosCanalesSobreUmbral_MueveAmbos()
    {
        var modelo = ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 600_000m),
            Fila("", "", "E", "Total", q: 600_000m),
            Fila("", "", "", "Vlr Servicio", q: 700_000m),
            Fila("", "", "O", "Total", q: 700_000m),
            Fila("", "EMPRESA", "Total", string.Empty, q: 1_300_000m),
            Fila("Seccion", "Total", "", string.Empty, q: 1_300_000m),
            Fila("Total", "", "", string.Empty, q: 1_300_000m),
        ]);

        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);
        var filas = resultado.Clasificado.Filas;

        Assert.Equal(0m, Q(filas, 0));
        Assert.Equal(600_000m, R(filas, 0));
        Assert.Equal(0m, Q(filas, 2));
        Assert.Equal(700_000m, R(filas, 2));

        // Agregado = Q − Σ(ambos) y R = + Σ(ambos).
        Assert.Equal(0m, Q(filas, 4));
        Assert.Equal(1_300_000m, R(filas, 4));
        Assert.Equal(0m, Q(filas, 6));
        Assert.Equal(1_300_000m, R(filas, 6));

        Assert.Equal(4, Conceptos(resultado).Count);
    }

    // ── U5 ──────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// U5 (R-ETIQUETA-COLAPSO): la etiqueta <c>O</c> aparece DOS veces (cuadro Componente y cuadro
    /// Subs/Cont). La identidad es la INSTANCIA del cierre: solo se mueve el grupo cuya instancia
    /// dispara; el homónimo queda intacto. Un <c>Map&lt;etiqueta,…&gt;</c> colapsaría ambos.
    /// </summary>
    [Fact]
    public void EtiquetaRepetidaEnDosCuadros_NoColapsa()
    {
        var modelo = ModeloOctubre();

        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);

        var conceptos = Conceptos(resultado);
        Assert.Equal(3, conceptos.Count);
        Assert.All(conceptos, m => Assert.Equal("O", m.Coordenada.Canal));
        Assert.All(conceptos, m => Assert.Equal("Componente", m.Coordenada.Seccion));

        // El grupo homónimo del segundo cuadro (sección Subs/Cont, Q=100) no produjo movimiento.
        Assert.DoesNotContain(resultado.Movimientos, m => m.Coordenada.Seccion == "Subs/Cont");
        Assert.Equal(100m, Q(resultado.Clasificado.Filas, CierreDeCanalOSubsCont));
        Assert.Equal(0m, R(resultado.Clasificado.Filas, CierreDeCanalOSubsCont));
    }

    // ── U6 ──────────────────────────────────────────────────────────────────────────────────────

    /// <summary>U6 (§4.3): para TODA fila tocada se conserva el par <c>Q' + R' = Q + R</c>.</summary>
    [Fact]
    public void Conservacion_QMasR_Invariante()
    {
        var modelo = ModeloOctubre();
        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);

        var tocadas = resultado.Movimientos.Select(m => m.IndiceFila).Distinct().ToList();
        Assert.NotEmpty(tocadas);

        foreach (var indice in tocadas)
        {
            var antes = Q(modelo.Filas, indice) + R(modelo.Filas, indice);
            var despues = Q(resultado.Clasificado.Filas, indice) + R(resultado.Clasificado.Filas, indice);
            Assert.True(Math.Abs(despues - antes) <= Tol, $"fila #{indice + 1}: Q'+R'={despues} vs Q+R={antes}.");
        }

        // La suma de TODO el bloque en Q+R tampoco cambia (CONSOLIDADO/VALIDACION siguen cerrando).
        var sumaAntes = Enumerable.Range(0, modelo.Filas.Count).Sum(i => Q(modelo.Filas, i) + R(modelo.Filas, i));
        var sumaDespues = Enumerable.Range(0, resultado.Clasificado.Filas.Count)
            .Sum(i => Q(resultado.Clasificado.Filas, i) + R(resultado.Clasificado.Filas, i));
        Assert.True(Math.Abs(sumaDespues - sumaAntes) <= Tol, $"Σ(Q+R) del bloque: antes={sumaAntes} después={sumaDespues}.");
    }

    // ── U7 ──────────────────────────────────────────────────────────────────────────────────────

    /// <summary>U7 (§4.4): <c>Aplicar(Aplicar(x)) == Aplicar(x)</c> fila por fila y sin movimientos nuevos.</summary>
    [Fact]
    public void Idempotencia_SegundaAplicacionNoOp()
    {
        var primera = ReclasificadorDebCredR2.Aplicar(ModeloOctubre());
        var segunda = ReclasificadorDebCredR2.Aplicar(primera.Clasificado);

        Assert.Empty(segunda.Movimientos);
        Assert.Equal(primera.Clasificado.Filas.Count, segunda.Clasificado.Filas.Count);
        for (var i = 0; i < primera.Clasificado.Filas.Count; i++)
        {
            Assert.Equal(Q(primera.Clasificado.Filas, i), Q(segunda.Clasificado.Filas, i));
            Assert.Equal(R(primera.Clasificado.Filas, i), R(segunda.Clasificado.Filas, i));
        }
    }

    // ── U8 ──────────────────────────────────────────────────────────────────────────────────────

    /// <summary>U8 (D5): sin cierres sobre el umbral el modelo no cambia y <c>R</c> queda en 0 explícito.</summary>
    [Fact]
    public void SinCanalesSobreUmbral_NoOp_R0Explicito()
    {
        var modelo = ModeloSimple(cierre: 2_232.90m);

        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);

        Assert.Empty(resultado.Movimientos);
        for (var i = 0; i < modelo.Filas.Count; i++)
        {
            Assert.Equal(Q(modelo.Filas, i), Q(resultado.Clasificado.Filas, i));
            Assert.Equal(0m, R(resultado.Clasificado.Filas, i));
        }

        // R es 0 EXPLÍCITO en el clasificado (nunca null: el writer escribe 0 y el par se conserva).
        Assert.NotNull(resultado.Clasificado.Filas[1].Valor(ReclasificadorDebCredR2.ComponenteDebCredOtros));
    }

    /// <summary>U8-bis: un modelo vacío es no-op (no hay nada que reclasificar, no hay error).</summary>
    [Fact]
    public void ModeloVacio_NoOp()
    {
        var modelo = new DetalleR2AseInputs
        {
            Ase = new Ase { Id = 2 },
            Componentes = [ReclasificadorDebCredR2.ComponenteDebCred],
            Fuente = "R2.xlsx",
        };

        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);

        Assert.Empty(resultado.Movimientos);
        Assert.Empty(resultado.Clasificado.Filas);
    }

    // ── U9 ──────────────────────────────────────────────────────────────────────────────────────
    /// <summary>U9 (D2): una fila de DETALLE cuyo importe no reconcilia con ningún tramo no se toca.</summary>
    [Fact]
    public void AgregadoSinTramo_FilaDeDetalle_NoSeToca()
    {
        var modelo = ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 10m),
            Fila("", "", "E", "Total", q: 10m),
            Fila("", "", "", "Vlr Intereses", q: 4_321.99m),
            Fila("", "EMPRESA", "Total", string.Empty, q: 10m),
            Fila("Seccion", "Total", "", string.Empty, q: 10m),
            Fila("Total", "", "", string.Empty, q: 10m),
        ]);

        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);

        Assert.Empty(resultado.Movimientos);
        Assert.Equal(4_321.99m, Q(resultado.Clasificado.Filas, 2));
        Assert.Equal(0m, R(resultado.Clasificado.Filas, 2));
    }

    // ── U10 ─────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// U10 (D2, R-SIN-TRAMO): una fila de CIERRE (empresa/sección/gran total) cuyo importe no es la
    /// suma de ningún tramo ⇒ fail-fast <c>ERR-RECLASIFICACION</c> con ASE + fila + firma + importe.
    ///
    /// <para>El fixture trae un canal SOBRE el umbral (dispara) además del cierre que no reconcilia:
    /// la validación de tramos es post-movimiento (ver U14), así que el fail-fast solo tiene sentido
    /// en una corrida donde algo se movió. El aserto de D2 no cambia.</para>
    /// </summary>
    [Fact]
    public void CierreSinTramo_FailFast()
    {
        var modelo = ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 600_000m),
            Fila("", "", "E", "Total", q: 600_000m),   // dispara
            Fila("", "", "", "Vlr Servicio", q: 10m),
            Fila("", "", "A", "Total", q: 10m),
            Fila("", "ENEL", "Total", string.Empty, q: 999.99m), // no es suma de ningún tramo
            Fila("Seccion", "Total", "", string.Empty, q: 10m),
            Fila("Total", "", "", string.Empty, q: 10m),
        ]);

        var ex = Assert.Throws<CalculoInvalidoException>(() => ReclasificadorDebCredR2.Aplicar(modelo));

        Assert.Equal(CodigoError.Reclasificacion, ex.Codigo);
        Assert.Contains("ERR-RECLASIFICACION", ex.Message, StringComparison.Ordinal);
        Assert.Contains("ASE 2", ex.Message, StringComparison.Ordinal);
        Assert.Contains("fila #5", ex.Message, StringComparison.Ordinal);
        Assert.Contains("|ENEL|Total|", ex.Message, StringComparison.Ordinal);
        Assert.Contains("importe", ex.Message, StringComparison.Ordinal);
        Assert.Contains("NINGÚN tramo", ex.Message, StringComparison.Ordinal);
    }

    // ── U14 ─────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// U14 (corrección de borde, D5/D4): **sin disparo la regla es NO-OP**. Un cierre estructural cuyo
    /// <c>Q</c> no es la suma de ningún tramo contiguo de cierres de canal NO es un error cuando
    /// ningún canal superó el umbral: no hay movimiento que reconciliar, así que la regla devuelve el
    /// modelo idéntico, cero movimientos y NO lanza. La validación de tramos es post-movimiento;
    /// convertirla en un fallo de un período sano sería un falso rojo.
    /// </summary>
    [Fact]
    public void SinDisparo_AgregadoQueNoReconcilia_NoFalla()
    {
        // (a) Todos los cierres de canal por debajo del umbral — (b) el cierre de empresa 999,99 no
        // es la suma de ningún tramo (10,00 / 0,00) y aun así el período debe salir intacto.
        var modelo = ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 10m),
            Fila("", "", "E", "Total", q: 10m),
            Fila("", "ENEL", "Total", string.Empty, q: 999.99m),
            Fila("", "", "", "Vlr Sustitutivo", q: 0m),
            Fila("", "", "A", "Total", q: 0m),
            Fila("", "EAB", "Total", string.Empty, q: 0m),
            Fila("Seccion", "Total", "", string.Empty, q: 1_009.99m),
            Fila("Total", "", "", string.Empty, q: 1_009.99m),
        ]);

        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);

        Assert.Empty(resultado.Movimientos);
        Assert.Equal(modelo.Filas.Count, resultado.Clasificado.Filas.Count);
        for (var i = 0; i < modelo.Filas.Count; i++)
        {
            Assert.Equal(Firma(modelo.Filas[i]), Firma(resultado.Clasificado.Filas[i]));
            Assert.Equal(Q(modelo.Filas, i), Q(resultado.Clasificado.Filas, i));
            Assert.Equal(R(modelo.Filas, i), R(resultado.Clasificado.Filas, i));
        }

        // El modelo de ENTRADA sigue intacto (el no-op también clona).
        Assert.Equal(999.99m, Q(modelo.Filas, 2));
    }

    // ── U15 ─────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// U15 (corrección de borde, hallazgo del review): una fila de **DETALLE** cuyo importe coincide
    /// por casualidad con la suma de un tramo de cierres anteriores NO se reconcilia. La propagación
    /// aplica solo a los CIERRES estructurales (empresa, sección, gran total): el <c>Q/R</c> de un
    /// concepto es su importe, no un agregado.
    ///
    /// <para>Fixture: el canal <c>E</c> dispara (600.000); el canal <c>A</c> cierra en 10,00 con una
    /// fila de detalle de 600.000,00 —exactamente el tramo <c>[E]</c>— y por tanto sin la guarda esa
    /// fila se habría movido a <c>Deb/Cred-otros</c> y la aritmética del grupo A habría quedado rota.</para>
    /// </summary>
    [Fact]
    public void DetalleQueCoincideConTramo_NoSeToca()
    {
        var modelo = ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 600_000m),                 // 0 · concepto del canal E
            Fila("", "", "E", "Total", q: 600_000m),                       // 1 · cierre: dispara
            Fila("", "", "", "Vlr Servicio", q: 600_000m),                 // 2 · DETALLE del canal A (coincide con [E])
            Fila("", "", "A", "Total", q: 10m),                            // 3 · cierre del canal A (bajo el umbral)
            Fila("", "EMPRESA", "Total", string.Empty, q: 600_010m),       // 4 · empresa = E + A
            Fila("Seccion", "Total", "", string.Empty, q: 600_010m),       // 5 · sección
            Fila("Total", "", "", string.Empty, q: 600_010m),              // 6 · gran total
        ]);

        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);
        var filas = resultado.Clasificado.Filas;

        // La fila de detalle queda INTACTA (ni Q ni R se mueven).
        Assert.Equal(600_000m, Q(filas, 2));
        Assert.Equal(0m, R(filas, 2));
        Assert.DoesNotContain(resultado.Movimientos, m => m.IndiceFila == 2);

        // El grupo del canal E sí se mueve…
        Assert.Equal(0m, Q(filas, 0));
        Assert.Equal(600_000m, R(filas, 0));
        Assert.Equal(0m, Q(filas, 1));
        Assert.Equal(600_000m, R(filas, 1));
        Assert.Equal(10m, Q(filas, 3));   // el cierre del canal A no dispara: se queda en Q

        // …y los cierres REALES se reconcilian con el tramo [E, A].
        Assert.Equal(10m, Q(filas, 4));
        Assert.Equal(600_000m, R(filas, 4));
        Assert.Equal(10m, Q(filas, 5));
        Assert.Equal(600_000m, R(filas, 5));
        Assert.Equal(10m, Q(filas, 6));
        Assert.Equal(600_000m, R(filas, 6));

        // Traza: 2 conceptos del grupo E + 3 cierres estructurales. Nada más.
        Assert.Equal(2, Conceptos(resultado).Count);
        Assert.Equal(3, resultado.Movimientos.Count(m =>
            m.Motivo.StartsWith(ReclasificadorDebCredR2.MotivoAgregado, StringComparison.Ordinal)));
        Assert.Equal(5, resultado.Movimientos.Count);
    }

    // ── U11 ─────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// U11 (D6/§4.5): sin coordenada COMPLETA no se mueve «lo que se pueda». Cuatro faltantes:
    /// fuente no declarada, sección no derivable, subsección (empresa) no derivable y ASE sin id.
    /// </summary>
    [Fact]
    public void CoordenadaIncompleta_FailFast()
    {
        // (a) Fuente vacía.
        var sinFuente = ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 600_000m),
            Fila("", "", "O", "Total", q: 600_000m),
            Fila("", "EMPRESA", "Total", string.Empty, q: 600_000m),
            Fila("Seccion", "Total", "", string.Empty, q: 600_000m),
            Fila("Total", "", "", string.Empty, q: 600_000m),
        ], fuente: string.Empty);
        var exFuente = Assert.Throws<CalculoInvalidoException>(() => ReclasificadorDebCredR2.Aplicar(sinFuente));
        Assert.Equal(CodigoError.Reclasificacion, exFuente.Codigo);
        Assert.Contains("no está declarada", exFuente.Message, StringComparison.Ordinal);

        // (b) Sin cierre de sección: nada encierra el grupo.
        var sinSeccion = ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 600_000m),
            Fila("", "", "O", "Total", q: 600_000m),
            Fila("", "EMPRESA", "Total", string.Empty, q: 600_000m),
            Fila("Total", "", "", string.Empty, q: 600_000m),
        ]);
        var exSeccion = Assert.Throws<CalculoInvalidoException>(() => ReclasificadorDebCredR2.Aplicar(sinSeccion));
        Assert.Equal(CodigoError.Reclasificacion, exSeccion.Codigo);
        Assert.Contains("sección derivable", exSeccion.Message, StringComparison.Ordinal);

        // (c) Sin cierre de empresa: no hay subsección.
        var sinEmpresa = ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 600_000m),
            Fila("", "", "O", "Total", q: 600_000m),
            Fila("Seccion", "Total", "", string.Empty, q: 600_000m),
            Fila("Total", "", "", string.Empty, q: 600_000m),
        ]);
        var exEmpresa = Assert.Throws<CalculoInvalidoException>(() => ReclasificadorDebCredR2.Aplicar(sinEmpresa));
        Assert.Equal(CodigoError.Reclasificacion, exEmpresa.Codigo);
        Assert.Contains("subsección (empresa) derivable", exEmpresa.Message, StringComparison.Ordinal);

        // (d) ASE no identificado.
        var sinAse = ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 600_000m),
            Fila("", "", "O", "Total", q: 600_000m),
            Fila("", "EMPRESA", "Total", string.Empty, q: 600_000m),
            Fila("Seccion", "Total", "", string.Empty, q: 600_000m),
            Fila("Total", "", "", string.Empty, q: 600_000m),
        ], aseId: 0);
        var exAse = Assert.Throws<CalculoInvalidoException>(() => ReclasificadorDebCredR2.Aplicar(sinAse));
        Assert.Equal(CodigoError.Reclasificacion, exAse.Codigo);
        Assert.Contains("ASE del bloque no está identificado", exAse.Message, StringComparison.Ordinal);
    }

    // ── U12 ─────────────────────────────────────────────────────────────────────────────────────

    /// <summary>U12 (D6/§4.5): un tramo que mezcla canales o que trae una etiqueta de sección/empresa ⇒ fail-fast.</summary>
    [Fact]
    public void GrupoNoContiguoOFilaDeOtroCanal_FailFast()
    {
        // (a) Una fila del tramo declara OTRO canal.
        var otroCanal = ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 600_000m),
            Fila("", "", "X", "Vlr Intereses", q: 0m),
            Fila("", "", "O", "Total", q: 600_000m),
            Fila("", "EMPRESA", "Total", string.Empty, q: 600_000m),
            Fila("Seccion", "Total", "", string.Empty, q: 600_000m),
            Fila("Total", "", "", string.Empty, q: 600_000m),
        ]);
        var exCanal = Assert.Throws<CalculoInvalidoException>(() => ReclasificadorDebCredR2.Aplicar(otroCanal));
        Assert.Equal(CodigoError.Reclasificacion, exCanal.Codigo);
        Assert.Contains("canal 'X'", exCanal.Message, StringComparison.Ordinal);
        Assert.Contains("se mueve nada", exCanal.Message, StringComparison.Ordinal);

        // (b) Una fila del tramo trae etiqueta de empresa (la expulsa del grupo).
        var conEmpresa = ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 600_000m),
            Fila("", "EMPRESA", "", "Vlr Intereses", q: 0m),
            Fila("", "", "O", "Total", q: 600_000m),
            Fila("", "EMPRESA2", "Total", string.Empty, q: 600_000m),
            Fila("Seccion", "Total", "", string.Empty, q: 600_000m),
            Fila("Total", "", "", string.Empty, q: 600_000m),
        ]);
        var exEmpresa = Assert.Throws<CalculoInvalidoException>(() => ReclasificadorDebCredR2.Aplicar(conEmpresa));
        Assert.Equal(CodigoError.Reclasificacion, exEmpresa.Codigo);
        Assert.Contains("etiqueta de sección/empresa", exEmpresa.Message, StringComparison.Ordinal);
    }

    // ── U13 ─────────────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// U13 (D6/§4.5, CA-R8): la traza lleva la COORDENADA COMPLETA por cada concepto movido y por
    /// cada agregado reconciliado: hoja · fuente · ASE · sección · subsección/empresa · canal · concepto.
    /// </summary>
    [Fact]
    public void TrazaPorConcepto_LlevaCoordenadaCompleta()
    {
        const string Fuente = "C:\\insumos\\2-Lime\\R2.xlsx";
        var modelo = ModeloOctubre(fuente: Fuente);

        var resultado = ReclasificadorDebCredR2.Aplicar(modelo);

        var conceptos = Conceptos(resultado);
        Assert.Equal(3, conceptos.Count);
        Assert.Equal(
            new[] { "Vlr Servicio", "Vlr Intereses", "Total" },
            conceptos.Select(m => m.Coordenada.Concepto).ToArray());
        Assert.Equal(
            new[] { 54.00m, 6_759_294.81m, 6_759_348.81m },
            conceptos.Select(m => m.MovidoQ).ToArray());

        foreach (var movimiento in conceptos)
        {
            var c = movimiento.Coordenada;
            Assert.Equal(ReclasificadorDebCredR2.HojaRemAnticiposR2, c.Hoja);
            Assert.Equal(Fuente, c.Fuente);
            Assert.Equal(2, c.AseId);
            Assert.Equal("Componente", c.Seccion);
            Assert.Equal("OCCIDENTE", c.Subseccion);
            Assert.Equal("O", c.Canal);
            Assert.Equal(Firma(modelo.Filas[movimiento.IndiceFila]), movimiento.Firma);
        }

        // Los agregados reconciliados también viajan con su coordenada y su tramo.
        var agregados = resultado.Movimientos
            .Where(m => m.Motivo.StartsWith(ReclasificadorDebCredR2.MotivoAgregado, StringComparison.Ordinal))
            .ToList();
        Assert.Equal(3, agregados.Count);

        var occidente = Assert.Single(agregados, m => m.Coordenada.Subseccion == "OCCIDENTE");
        Assert.Equal("Componente", occidente.Coordenada.Seccion);
        Assert.Equal("H,O,T", occidente.Coordenada.Canal);
        Assert.Equal(6_759_348.81m, occidente.MovidoQ);
        Assert.Contains("|OCCIDENTE|Total|", occidente.Firma, StringComparison.Ordinal);

        Assert.Contains(agregados, m => m.Coordenada.Concepto == "Componente" && m.Coordenada.Seccion == "Componente");
        Assert.Contains(agregados, m => m.Coordenada.Seccion.Length == 0 && m.Coordenada.Concepto == "Total");
    }

    // ── modelo sintético ────────────────────────────────────────────────────────────────────────

    private static IReadOnlyList<MovimientoReclasificacion> Conceptos(ResultadoReclasificacion resultado) =>
        resultado.Movimientos
            .Where(m => m.Motivo.StartsWith(ReclasificadorDebCredR2.MotivoCierreCanal, StringComparison.Ordinal))
            .ToList();

    private static DetalleR2Fila Fila(string a, string b, string c, string d, decimal q, decimal? r = null) =>
        new()
        {
            A = a,
            B = b,
            C = c,
            D = d,
            ValoresPorComponente = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase)
            {
                ["Total"] = 0m,
                [ReclasificadorDebCredR2.ComponenteDebCred] = q,
                [ReclasificadorDebCredR2.ComponenteDebCredOtros] = r,
            },
        };

    private static DetalleR2AseInputs ModeloDe(DetalleR2Fila[] filas, int aseId = 2, string fuente = "R2.xlsx") =>
        new()
        {
            Ase = new Ase { Id = aseId, NombreCorto = "ASE", NombreCompleto = "Ase", NumeroCarpeta = aseId },
            Componentes = ["Total", ReclasificadorDebCredR2.ComponenteDebCred],
            Fuente = fuente,
            Filas = filas,
        };

    /// <summary>
    /// Fixture con la aritmética medida en octubre-1: los 9 cierres de canal del cuadro
    /// <c>Componente</c> (A 5,17 · E 31,40 · 1 12,95 · 2 2,44 · 3 4,90 · 5 −27,75 · H −21,95 ·
    /// O 6.759.348,81 · T 15,80), las 4 empresas (<c>EAB</c> = A, <c>ENEL</c> = E,
    /// <c>NUEVO ESQUEMA</c> = 1+2+3+5, <c>OCCIDENTE</c> = H+O+T) y el cuadro <c>Subs/Cont</c>, que
    /// repite los mismos canales con importes bajo el umbral (O = 100,00).
    /// </summary>
    private static DetalleR2AseInputs ModeloOctubre(string fuente = "R2.xlsx") =>
        ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: 5.17m),                      // 0
            Fila("", "", "A", "Total", q: 5.17m),                            // 1
            Fila("", "EAB", "Total", string.Empty, q: 5.17m),                // 2
            Fila("", "", "", "Vlr Servicio", q: 30.25m),                     // 3
            Fila("", "", "", "Vlr Intereses", q: 1.15m),                     // 4
            Fila("", "", "E", "Total", q: 31.40m),                           // 5
            Fila("", "ENEL", "Total", string.Empty, q: 31.40m),              // 6
            Fila("", "", "", "Vlr Servicio", q: 12.99m),                     // 7
            Fila("", "", "", "Vlr Intereses", q: -0.04m),                    // 8
            Fila("", "", "1", "Total", q: 12.95m),                           // 9
            Fila("", "", "", "Vlr Servicio", q: 2.44m),                      // 10
            Fila("", "", "2", "Total", q: 2.44m),                            // 11
            Fila("", "", "", "Vlr Servicio", q: 4.90m),                      // 12
            Fila("", "", "3", "Total", q: 4.90m),                            // 13
            Fila("", "", "", "Vlr Servicio", q: -27.76m),                    // 14
            Fila("", "", "", "Vlr Intereses", q: 0.01m),                     // 15
            Fila("", "", "5", "Total", q: -27.75m),                          // 16
            Fila("", "NUEVO ESQUEMA", "Total", string.Empty, q: -7.46m),     // 17
            Fila("", "", "", "Vlr Servicio", q: -21.95m),                    // 18
            Fila("", "", "H", "Total", q: -21.95m),                          // 19
            Fila("", "", "", "Vlr Servicio", q: 54.00m),                     // 20
            Fila("", "", "", "Vlr Intereses", q: 6_759_294.81m),             // 21
            Fila("", "", "O", "Total", q: 6_759_348.81m),                    // 22
            Fila("", "", "", "Vlr Servicio", q: 15.80m),                     // 23
            Fila("", "", "T", "Total", q: 15.80m),                           // 24
            Fila("", "OCCIDENTE", "Total", string.Empty, q: 6_759_342.66m),  // 25
            Fila("Componente", "Total", "", string.Empty, q: 6_759_371.77m), // 26
            Fila("", "", "", "Subsidio(-)/Contribucion(+)", q: 0m),          // 27
            Fila("", "", "A", "Total", q: 0m),                               // 28
            Fila("", "EAB", "Total", string.Empty, q: 0m),                   // 29
            Fila("", "", "", "Subsidio(-)/Contribucion(+)", q: 0m),          // 30
            Fila("", "", "E", "Total", q: 0m),                               // 31
            Fila("", "ENEL", "Total", string.Empty, q: 0m),                  // 32
            Fila("", "", "", "Subsidio(-)/Contribucion(+)", q: 0m),          // 33
            Fila("", "", "1", "Total", q: 0m),                               // 34
            Fila("", "", "", "Subsidio(-)/Contribucion(+)", q: 0m),          // 35
            Fila("", "", "2", "Total", q: 0m),                               // 36
            Fila("", "", "", "Subsidio(-)/Contribucion(+)", q: 0m),          // 37
            Fila("", "", "3", "Total", q: 0m),                               // 38
            Fila("", "", "", "Subsidio(-)/Contribucion(+)", q: 0m),          // 39
            Fila("", "", "5", "Total", q: 0m),                               // 40
            Fila("", "NUEVO ESQUEMA", "Total", string.Empty, q: 0m),         // 41
            Fila("", "", "", "Subsidio(-)/Contribucion(+)", q: 0m),          // 42
            Fila("", "", "H", "Total", q: 0m),                               // 43
            Fila("", "", "", "Subsidio(-)/Contribucion(+)", q: 100m),        // 44
            Fila("", "", "O", "Total", q: 100m),                             // 45
            Fila("", "", "", "Subsidio(-)/Contribucion(+)", q: 0m),          // 46
            Fila("", "", "T", "Total", q: 0m),                               // 47
            Fila("", "OCCIDENTE", "Total", string.Empty, q: 100m),           // 48
            Fila("Subs/Cont", "Total", "", string.Empty, q: 100m),           // 49
            Fila("Total", "", "", string.Empty, q: 6_759_471.77m),           // 50
        ],
        fuente: fuente);

    /// <summary>Fixture mínimo: un canal con el cierre indicado y sus agregados coherentes.</summary>
    private static DetalleR2AseInputs ModeloSimple(decimal cierre) =>
        ModeloDe(
        [
            Fila("", "", "", "Vlr Servicio", q: cierre),
            Fila("", "", "E", "Total", q: cierre),
            Fila("", "EMPRESA", "Total", string.Empty, q: cierre),
            Fila("Seccion", "Total", "", string.Empty, q: cierre),
            Fila("Total", "", "", string.Empty, q: cierre),
        ]);

    private static decimal Q(IReadOnlyList<DetalleR2Fila> filas, int indice) =>
        filas[indice].Valor(ReclasificadorDebCredR2.ComponenteDebCred) ?? 0m;

    private static decimal R(IReadOnlyList<DetalleR2Fila> filas, int indice) =>
        filas[indice].Valor(ReclasificadorDebCredR2.ComponenteDebCredOtros) ?? 0m;

    private static string Firma(DetalleR2Fila fila) => $"{fila.A}|{fila.B}|{fila.C}|{fila.D}";
}
