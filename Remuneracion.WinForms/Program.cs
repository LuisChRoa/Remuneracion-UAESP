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
                locator);

            Application.Run(new Form1(procesador, procesadorPeriodo, locator));
        }
    }
}