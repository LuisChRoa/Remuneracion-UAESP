using Remuneracion.Core.Constants;
using Remuneracion.Core.Models;

namespace Remuneracion.Core.Services;

/// <summary>
/// HU-14 (S-3): factoría ÚNICA de <see cref="Ase"/> desde <see cref="CarpetasAse.Prefijos"/>
/// (fuente única de verdad de los 5 nombres). Reemplaza los <c>CrearAse</c> duplicados de
/// <see cref="ProcesadorPeriodo"/> y <c>ValidacionOracleReader.CrearAse</c> (el array
/// hardcodeado del reader desaparece; se elimina la fuente divergente).
/// </summary>
public static class AseFactory
{
    /// <summary>
    /// Crea el <see cref="Ase"/> del id indicado derivando el nombre del prefijo de carpeta
    /// (<c>N-Nombre</c> → <c>Nombre</c>). Id fuera de 1..5 = fail-fast con el rango válido.
    /// </summary>
    /// <param name="id">Id del ASE (1..5).</param>
    /// <returns>ASE con Id, NombreCorto (mayúsculas), NombreCompleto y NumeroCarpeta = id.</returns>
    public static Ase DesdeId(int id)
    {
        if (id < 1 || id > CarpetasAse.Prefijos.Count)
        {
            throw new ArgumentOutOfRangeException(
                nameof(id),
                $"No existe el ASE {id}; el catálogo define los ASE 1..{CarpetasAse.Prefijos.Count} ({string.Join(", ", CarpetasAse.Prefijos)}).");
        }

        var prefijo = CarpetasAse.Prefijos[id - 1];
        var guion = prefijo.IndexOf('-');
        var nombre = guion > 0 ? prefijo[(guion + 1)..] : prefijo;
        return new Ase
        {
            Id = id,
            NombreCorto = nombre.ToUpperInvariant(),
            NombreCompleto = nombre,
            NumeroCarpeta = id
        };
    }
}