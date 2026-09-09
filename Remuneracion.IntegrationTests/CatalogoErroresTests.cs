using Remuneracion.Core.Errors;
using Remuneracion.Core.Exceptions;
using Xunit;

namespace Remuneracion.IntegrationTests;

/// <summary>
/// HU-14 (3.1, PR1 — dominio): catálogo de errores con códigos, mensajes UX por categoría,
/// mapa a códigos de salida (contrato HU-15) y extensión no-rompiente de las 2 excepciones
/// de dominio (constructores previos intactos + propiedad <see cref="ArchivoFuenteNoEncontradoException.Codigo"/>).
/// </summary>
public sealed class CatalogoErroresTests
{
    [Theory]
    [InlineData(CodigoError.FuenteNoEncontrada, CodigosSalida.FuenteOPlantilla)]
    [InlineData(CodigoError.Plantilla, CodigosSalida.FuenteOPlantilla)]
    [InlineData(CodigoError.FormatoFuente, CodigosSalida.FuenteOPlantilla)]
    [InlineData(CodigoError.Validacion, CodigosSalida.Validacion)]
    [InlineData(CodigoError.Escritura, CodigosSalida.Escritura)]
    [InlineData(CodigoError.Inesperado, CodigosSalida.Inesperado)]
    [InlineData(CodigoError.CanceladoPorUsuario, CodigosSalida.CanceladoPorUsuario)]
    public void CodigoSalidaPara_MapaCompleto_DevuelveElContrato(string codigo, int esperado)
    {
        Assert.Equal(esperado, CatalogoErrores.CodigoSalidaPara(codigo));
    }

    [Fact]
    public void CodigoSalidaPara_CodigoDesconocido_FallaSeguroConInesperado()
    {
        Assert.Equal(CodigosSalida.Inesperado, CatalogoErrores.CodigoSalidaPara("ERR-NO-EXISTE"));
    }

    [Theory]
    [InlineData(CodigoError.FuenteNoEncontrada, "Archivo fuente faltante")]
    [InlineData(CodigoError.Plantilla, "Plantilla o ruta de salida no válida")]
    [InlineData(CodigoError.FormatoFuente, "Formato de archivo fuente no compatible")]
    [InlineData(CodigoError.Validacion, "La validación no cierra")]
    [InlineData(CodigoError.Escritura, "Error al generar el archivo de salida")]
    [InlineData(CodigoError.Inesperado, "Error inesperado")]
    [InlineData(CodigoError.CanceladoPorUsuario, "Proceso cancelado")]
    public void Para_TituloPorCategoria_NoVacio(string codigo, string esperado)
    {
        var (titulo, guia) = CatalogoErrores.Para(codigo, null);
        Assert.Equal(esperado, titulo);
        Assert.False(string.IsNullOrWhiteSpace(guia));
    }

    [Fact]
    public void Para_FuenteNoEncontrada_GuiaIncluyeElDetalleDelMensaje()
    {
        var ex = new ArchivoFuenteNoEncontradoException("No se encontró R2 del ASE 3 en C:\\fuentes\\3-Ciudad Limpia.");
        var (_, guia) = CatalogoErrores.Para(CodigoError.FuenteNoEncontrada, ex);
        Assert.Contains("R2", guia, StringComparison.Ordinal);
        Assert.Contains("ASE 3", guia, StringComparison.Ordinal);
        Assert.Contains("reintente", guia, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Para_Inesperado_GuiaRemiteAlLogConRunId()
    {
        var (_, guia) = CatalogoErrores.Para(CodigoError.Inesperado, new InvalidOperationException("boom"));
        Assert.Contains("RunId", guia, StringComparison.Ordinal);
        Assert.Contains("remuneracion_log", guia, StringComparison.OrdinalIgnoreCase);
    }

    // ── Defaults y compatibilidad de constructores (D1: extensión por propiedad, no jerarquía) ──

    [Fact]
    public void ArchivoFuenteNoEncontrado_ConstructorViejo_CodigoDefaultFuenteNoEncontrada()
    {
        var ex = new ArchivoFuenteNoEncontradoException("mensaje");
        Assert.Equal(CodigoError.FuenteNoEncontrada, ex.Codigo);
        Assert.Equal("mensaje", ex.Message);
    }

    [Fact]
    public void ArchivoFuenteNoEncontrado_ConstructorConCodigo_UsaElCodigoExplicito()
    {
        var ex = new ArchivoFuenteNoEncontradoException(CodigoError.Plantilla, "plantilla ausente");
        Assert.Equal(CodigoError.Plantilla, ex.Codigo);
        Assert.Equal("plantilla ausente", ex.Message);
    }

    [Fact]
    public void CalculoInvalido_ConstructorViejo_CodigoDefaultValidacion()
    {
        var ex = new CalculoInvalidoException("mensaje");
        Assert.Equal(CodigoError.Validacion, ex.Codigo);
        Assert.Equal("mensaje", ex.Message);
    }

    [Fact]
    public void CalculoInvalido_ConstructorConCodigo_UsaElCodigoExplicito()
    {
        var ex = new CalculoInvalidoException(CodigoError.Escritura, "fallo de escritura");
        Assert.Equal(CodigoError.Escritura, ex.Codigo);
        Assert.Equal("fallo de escritura", ex.Message);
    }

    [Fact]
    public void Excepciones_ConstructoresConInnerException_Intactos()
    {
        var inner = new InvalidOperationException("causa");
        var exArchivo = new ArchivoFuenteNoEncontradoException("mensaje", inner);
        var exCalculo = new CalculoInvalidoException("mensaje", inner);
        Assert.Same(inner, exArchivo.InnerException);
        Assert.Same(inner, exCalculo.InnerException);
        Assert.Equal(CodigoError.FuenteNoEncontrada, exArchivo.Codigo);
        Assert.Equal(CodigoError.Validacion, exCalculo.Codigo);
    }

    [Fact]
    public void CodigosSalida_ContratoEstable_ParaHu15()
    {
        Assert.Equal(0, CodigosSalida.Ok);
        Assert.Equal(1, CodigosSalida.Validacion);
        Assert.Equal(2, CodigosSalida.FuenteOPlantilla);
        Assert.Equal(3, CodigosSalida.Escritura);
        Assert.Equal(4, CodigosSalida.Inesperado);
        Assert.Equal(5, CodigosSalida.CanceladoPorUsuario);
    }
}