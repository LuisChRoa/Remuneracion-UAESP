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
            IProcesadorRemuneracion procesador = new ProcesadorRemuneracion(
                lectorR1R2R4,
                hojaLeafReader,
                calculo,
                validador,
                writer);

            var locator = new ArchivoFuenteLocator();
            Application.Run(new Form1(procesador, locator));
        }
    }
}