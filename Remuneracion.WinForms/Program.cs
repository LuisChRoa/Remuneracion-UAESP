using Remuneracion.Core.Interfaces;
using Remuneracion.Core.Services;
using Remuneracion.Infrastructure.Excel;
using Remuneracion.Infrastructure.FileSystem;

namespace Remuneracion.WinForms
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();

            var lectorR1R2R4 = new ExcelDataReaderRecaudoReader();
            var hojaLeafReader = new ExcelDataReaderWorkbookLeafInputReader();
            var calculo = new CalculoRemuneracion();
            var validador = new ValidadorBasico();
            var writer = new OpenXmlPlantillaWriter();
            var locator = new ArchivoFuenteLocator();

            // HU-20 (G3): oráculo R10 del período (OBLIGATORIO) — valida el DetRetri calculado en
            // ambas quincenas. HU-13 (2.7): oráculo de LECTURA de validaciones cruzadas (opcional;
            // la regresión corre sin él = HU-12 puro, D3).
            var detRetriR10Reader = new ExcelDataReaderDetRetriR10Reader();
            var validacionOracleReader = new ValidacionOracleReader();

            IProcesadorRemuneracion procesador = new ProcesadorRemuneracion(
                lectorR1R2R4,
                hojaLeafReader,
                calculo,
                validador,
                writer);

            IProcesadorPeriodo procesadorPeriodo = new ProcesadorPeriodo(
                lectorR1R2R4,
                hojaLeafReader,
                calculo,
                validador,
                writer,
                locator,
                detRetriR10Reader,
                validacionOracleReader);

            Application.Run(new Form1(procesador, procesadorPeriodo, locator));
        }
    }
}
