using Remuneracion.Core.Exceptions;
using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Services;

/// <summary>
/// Implementación concreta del motor de cálculo del consolidado para la Fase 1.
/// </summary>
public sealed class CalculoRemuneracion : ICalculoRemuneracion
{
    public ConsolidadoAse Calcular(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4)
    {
        if (ase is null)
        {
            throw new ArgumentNullException(nameof(ase));
        }

        if (r1 is null)
        {
            throw new ArgumentNullException(nameof(r1));
        }

        if (r2 is null)
        {
            throw new ArgumentNullException(nameof(r2));
        }

        if (r4 is null)
        {
            throw new ArgumentNullException(nameof(r4));
        }

        return new ConsolidadoAse
        {
            Ase = ase,
            TotOpt = r1.TotalOportuno,
            R2TotalOportuno = r2.TotalOportuno,
            Extemp = r1.Extemporaneo,
            ReversionR4 = r4.TotalReversiones,
            AjustesSfT = 0m
        };
    }

    /// <inheritdoc />
    public ConsolidadoAse Calcular(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4, decimal ajustesSfT)
    {
        var consolidado = Calcular(ase, r1, r2, r4);
        consolidado.AjustesSfT = ajustesSfT;
        return consolidado;
    }

    public ResultadoRemuneracion CalcularConsolidado(
        Periodo periodo,
        List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4)> datos)
    {
        if (periodo is null)
        {
            throw new ArgumentNullException(nameof(periodo));
        }

        if (datos is null)
        {
            throw new ArgumentNullException(nameof(datos));
        }

        if (datos.Count == 0)
        {
            throw new CalculoInvalidoException("La colección de datos no puede estar vacía.");
        }

        if (datos.Any(d => d.ase is null || d.r1 is null || d.r2 is null || d.r4 is null))
        {
            throw new ArgumentNullException(nameof(datos), "No se admiten tuplas con elementos nulos.");
        }

        var duplicados = datos
            .GroupBy(d => d.ase.Id)
            .FirstOrDefault(g => g.Count() > 1);

        if (duplicados is not null)
        {
            throw new CalculoInvalidoException($"Se detectaron ASE duplicados: {duplicados.Key}.");
        }

        if (periodo.NumeroQuincena is not 1 and not 2)
        {
            throw new CalculoInvalidoException($"Número de quincena inválido: {periodo.NumeroQuincena}. Solo se soportan 1 o 2.");
        }

        if (periodo.NumeroQuincena == 2)
        {
            throw new CalculoInvalidoException(
                "El cálculo para quincena 2 no está soportado en esta fase: requiere modelar SALDOS POR NOTA y RETRIBUCIÓN NEGATIVA antes de continuar.");
        }

        var consolidados = datos
            .Select(d => Calcular(d.ase, d.r1, d.r2, d.r4))
            .ToList();

        return new ResultadoRemuneracion
        {
            Periodo = periodo,
            Consolidados = consolidados,
            Exitoso = true,
            Mensajes = [],
            Timestamp = DateTime.Now
        };
    }

    /// <inheritdoc />
    public ResultadoRemuneracion CalcularConsolidado(
        Periodo periodo,
        List<(Ase ase, RecaudoComponenteR1 r1, SaldosFavorR2 r2, ReversionR4 r4, decimal ajustesSfT)> datos)
    {
        ArgumentNullException.ThrowIfNull(periodo);
        ArgumentNullException.ThrowIfNull(datos);

        if (datos.Count == 0)
        {
            throw new CalculoInvalidoException("La colección de datos no puede estar vacía.");
        }

        if (datos.Any(d => d.ase is null || d.r1 is null || d.r2 is null || d.r4 is null))
        {
            throw new ArgumentNullException(nameof(datos), "No se admiten tuplas con elementos nulos.");
        }

        var duplicados = datos
            .GroupBy(d => d.ase.Id)
            .FirstOrDefault(g => g.Count() > 1);

        if (duplicados is not null)
        {
            throw new CalculoInvalidoException($"Se detectaron ASE duplicados: {duplicados.Key}.");
        }

        if (periodo.NumeroQuincena is not 1 and not 2)
        {
            throw new CalculoInvalidoException($"Número de quincena inválido: {periodo.NumeroQuincena}. Solo se soportan 1 o 2.");
        }

        // HU-11 (2.5, D1): Q2 se desbloquea por este overload que SÍ recibe ajustes por ASE.
        // El overload viejo (sin ajustes) queda intacto como red de seguridad (fail-fast Q2).
        var consolidados = datos
            .Select(d => Calcular(d.ase, d.r1, d.r2, d.r4, d.ajustesSfT))
            .ToList();

        return new ResultadoRemuneracion
        {
            Periodo = periodo,
            Consolidados = consolidados,
            Exitoso = true,
            Mensajes = [],
            Timestamp = DateTime.Now
        };
    }
}
